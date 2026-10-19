<% ' ---------------------------------------------------------------------------
  '  Placement Tracker  -  inc\auth.asp
  '  ---------------------------------------------------------------------------
  '  Login protection: session state, idle/absolute timeouts, brute-force
  '  lockout, one-time login tickets and the credential check.
  '
  '  This file is included at the TOP of inc\helpers.asp, so every page that
  '  includes helpers (which is every page) has these procedures available,
  '  including inc\header.asp and the shared error page.  It contains only
  '  procedure definitions - no top-level statements - so it is harmless on
  '  pages that never check authentication (login.asp, 404.asp).
  '
  '  What is stored where
  '    Session   uid, uname, auth_at, auth_start, login_next, login_reason
  '    Application  one-time login tickets and failed-attempt counters
  '    dbo.Users username, salt, SHA-256(salt + password), is_active
  '
  '  Nothing in this file writes a password anywhere.  Verification happens
  '  inside SQL Server:
  '      SELECT ... WHERE username = ? AND is_active = 1
  '                   AND password_hash = HASHBYTES('SHA2_256', salt + ?)
  '  both placeholders are adVarWChar parameters - the username is lower-cased
  '  in VBScript first, and the password is passed through untouched.
  '    ------------------------------------------------------------------------ %>

<%
' ===========================================================================
'  Tunables.  Both environment variables are read per request, which is how
'  tests\smoke.ps1 shortens the timeout for its own test instance without
'  touching the code.
' ===========================================================================
Const AUTH_DEFAULT_IDLE = 1200   ' 20 minutes without a request
Const AUTH_DEFAULT_MAX  = 28800  ' 8 hours, whatever the activity
Const AUTH_MAX_FAILURES = 5      ' failed sign-ins per account...
Const AUTH_LOCK_SECONDS = 300    ' ...then a 5 minute lock
Const AUTH_TICKET_TTL   = 60     ' a one-time login ticket lives 60 seconds

Function AuthEnvVar(name)
    AuthEnvVar = ""
    On Error Resume Next
    Dim sh
    Set sh = Server.CreateObject("WScript.Shell")
    If Err.Number = 0 Then
        Dim v
        v = sh.Environment("PROCESS")(name) & ""
        Err.Clear
        If v <> "" Then AuthEnvVar = v
        Set sh = Nothing
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Function

' Returns an environment-backed number, clamped to [minValue, maxValue].
' Anything unparsable falls back to the default, so a bad value can never
' disable authentication or produce a 500.
Function AuthSeconds(name, fallback, minValue, maxValue)
    AuthSeconds = fallback
    Dim raw
    raw = Trim(AuthEnvVar(name))
    If raw = "" Then Exit Function
    If Not TestPattern("^[0-9]{1,7}$", raw) Then Exit Function
    Dim n
    n = CLng(raw)
    If n < minValue Then n = minValue
    If n > maxValue Then n = maxValue
    AuthSeconds = n
End Function

Function AuthIdleSeconds()
    AuthIdleSeconds = AuthSeconds("PT_AUTH_IDLE_SECONDS", AUTH_DEFAULT_IDLE, 1, 86400)
End Function

Function AuthMaxSeconds()
    AuthMaxSeconds = AuthSeconds("PT_AUTH_MAX_SECONDS", AUTH_DEFAULT_MAX, 60, 86400)
End Function

' Used once per request (inc\helpers.asp) so the ASP session itself cannot
' outlive the absolute lifetime we enforce in code.
Function AuthSessionMinutes()
    Dim mins
    mins = CLng(Int((AuthMaxSeconds() + 59) / 60))
    If mins < 1 Then mins = 1
    If mins > 1440 Then mins = 1440
    AuthSessionMinutes = mins
End Function

' Seconds since 1970-01-01 in local time.  DateSerial keeps this locale-proof
' (no parsing of "1970-01-01" strings, which differ between en-US and en-GB).
Function AuthNow()
    AuthNow = DateDiff("s", DateSerial(1970, 1, 1), Now)
End Function

' ===========================================================================
'  Session state
' ===========================================================================
Sub AuthClearSession()
    On Error Resume Next
    Session.Contents.Remove("uid")
    Session.Contents.Remove("uname")
    Session.Contents.Remove("auth_at")
    Session.Contents.Remove("auth_start")
    Session.Contents.Remove("login_next")
    Session.Contents.Remove("login_reason")
    On Error GoTo 0
End Sub

' Signs a verified user in.  The CSRF token is rotated at the same time so a
' token minted before authentication cannot be replayed after it.
Sub AuthStartSession(userId, username)
    AuthClearSession
    Session("uid")        = CLng(userId)
    Session("uname")      = username & ""
    Session("auth_start") = AuthNow()
    Session("auth_at")    = AuthNow()
    Session("csrf")       = NewToken()
End Sub

Function AuthUserId()
    AuthUserId = 0
    On Error Resume Next
    Dim v
    v = Session("uid")
    On Error GoTo 0
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    If CLng(v) < 1 Then Exit Function
    AuthUserId = CLng(v)
End Function

Function AuthUsername()
    AuthUsername = ""
    On Error Resume Next
    Dim v
    v = Session("uname")
    On Error GoTo 0
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    AuthUsername = v & ""
End Function

' True only when a user id is present and neither timeout has elapsed.
' The idle clock is refreshed on every authenticated request, so this is a
' sliding window; the absolute clock is anchored at sign-in.
Function IsAuthenticated()
    IsAuthenticated = False

    Dim uid, at, started, nowE, idleSecs, maxSecs
    On Error Resume Next
    uid     = Session("uid")
    at      = Session("auth_at")
    started = Session("auth_start")
    On Error GoTo 0

    If IsEmpty(uid)     Or IsNull(uid)     Then Exit Function
    If IsEmpty(at)      Or IsNull(at)      Then Exit Function
    If IsEmpty(started) Or IsNull(started) Then Exit Function
    If Not IsNumeric(uid) Or Not IsNumeric(at) Or Not IsNumeric(started) Then Exit Function
    If CLng(uid) < 1 Then Exit Function

    nowE    = AuthNow()
    idleSecs = AuthIdleSeconds()
    maxSecs  = AuthMaxSeconds()

    If (nowE - CLng(at)) > idleSecs Then
        AuthClearSession            ' timed out: the UI says so after re-login
        Exit Function
    End If
    If (nowE - CLng(started)) > maxSecs Then
        AuthClearSession
        Exit Function
    End If

    Session("auth_at") = nowE
    IsAuthenticated = True
End Function

' ---------------------------------------------------------------------------
'  EnsureAuthenticated - the gate every protected page calls first.
'
'  Anonymous  -> remember the destination and redirect to /login.asp (302).
'  Expired    -> same, flagged so the login page can explain the timeout.
'  Authenticated -> refresh the idle clock and carry on.
'
'  The redirect happens BEFORE the page reads any id or query parameter, so
'  an unauthenticated caller cannot learn whether a record exists: they get
'  the same 302 whether the id is valid or not.
' ---------------------------------------------------------------------------
Sub EnsureAuthenticated()
    If IsAuthenticated() Then Exit Sub

    Dim wasSignedIn
    wasSignedIn = False
    If AuthUserId() > 0 Then wasSignedIn = True   ' IsAuthenticated just cleared it

    Dim target, qs
    target = Request.ServerVariables("URL") & ""
    qs = Request.QueryString & ""
    If qs <> "" Then target = target & "?" & qs

    AuthClearSession
    If wasSignedIn Then
        Session("login_reason") = "expired"
    Else
        Session("login_reason") = ""
    End If
    Session("login_next") = target

    Response.Redirect "/login.asp"
End Sub

' ===========================================================================
'  Safe redirect target for after sign-in
'
'  The destination comes out of the session (set by EnsureAuthenticated), but
'  it is re-validated here anyway so an attacker who can write into the
'  session cannot bounce the user off-site: only a site-relative link to one
'  of this application's own .asp pages is ever accepted.
' ===========================================================================
Function AuthSafeNext(url)
    AuthSafeNext = "/default.asp"
    url = Trim(url & "")
    If url = "" Then Exit Function
    If url = "/" Then Exit Function
    If Left(url, 1) <> "/" Then Exit Function        ' must be site-relative
    If Left(url, 2) = "//" Then Exit Function        ' not protocol-relative
    If InStr(url, "\") > 0 Then Exit Function        ' no backslashes
    If InStr(url, "..") > 0 Then Exit Function       ' no traversal
    If InStr(url, "//") > 0 Then Exit Function       ' no embedded '//'
    If InStr(url, "#") > 0 Then Exit Function        ' no fragment
    If Not TestPattern("^/(default|new|view|edit|delete|404|login|logout|forgot)\.asp(\?|$)", LCase(url)) Then Exit Function
    AuthSafeNext = url
End Function

' ===========================================================================
'  One-time login ticket
'
'  On success the login POST calls Session.Abandon, which throws away the
'  session the credentials were typed into, and hands the result back through
'  a ticket stored in Application scope.  The follow-up GET therefore runs in
'  a brand new session - the session identifier changes at the moment of
'  authentication, which is what defeats session fixation.
'
'  Application.Lock is held for every read-modify-write below.
' ===========================================================================
Function AuthTicketCreate(userId, username, nextUrl)
    Dim t, value
    t = NewToken()
    value = CStr(CLng(userId)) & "|" & CStr(AuthNow() + AUTH_TICKET_TTL) & _
            "|" & (username & "") & "|" & Server.UrlEncode(nextUrl & "")

    Application.Lock
    AuthPurgeTickets
    Application("tkt_" & t) = value
    Application.Unlock

    AuthTicketCreate = t
End Function

' Returns the user id, or 0 when the ticket is absent, malformed, unknown or
' expired.  The ticket is removed from Application scope as soon as it is read,
' so each one can only be redeemed once.  username/nextUrl are filled ByRef.
Function AuthTicketConsume(ticket, ByRef username, ByRef nextUrl)
    AuthTicketConsume = 0
    username = ""
    nextUrl  = ""

    ticket = Trim(ticket & "")
    If ticket = "" Then Exit Function
    If Not TestPattern("^[0-9a-f]{32,80}$", ticket) Then Exit Function

    Dim key, raw, parts
    key = "tkt_" & ticket

    Application.Lock
    raw = ""
    On Error Resume Next
    raw = Application(key) & ""
    Application.Contents.Remove(key)
    On Error GoTo 0
    Application.Unlock

    If raw = "" Then Exit Function

    parts = Split(raw, "|")
    If UBound(parts) <> 3 Then Exit Function
    If Not IsNumeric(parts(0)) Then Exit Function
    If Not IsNumeric(parts(1)) Then Exit Function
    If CLng(parts(1)) < AuthNow() Then Exit Function      ' expired

    nextUrl  = Server.UrlDecode(parts(3) & "")
    username = parts(2) & ""
    AuthTicketConsume = CLng(parts(0))
End Function

' Drops tickets that are past their TTL.  Caller must hold Application.Lock.
Sub AuthPurgeTickets()
    Dim k, raw, parts, bag
    Set bag = Server.CreateObject("Scripting.Dictionary")

    On Error Resume Next
    For Each k In Application.Contents
        If Left(k, 4) = "tkt_" Then
            raw = ""
            raw = Application(k) & ""
            parts = Split(raw, "|")
            If UBound(parts) <> 3 Then
                bag.Add k, 0                                  ' malformed: drop it
            ElseIf Not IsNumeric(parts(1)) Then
                bag.Add k, 0
            ElseIf (CLng(parts(1)) + 240) < AuthNow() Then
                bag.Add k, 0                                  ' well past its TTL
            End If
        End If
    Next

    For Each k In bag.Keys
        Application.Contents.Remove(k)
    Next
    On Error GoTo 0
    Set bag = Nothing
End Sub

' ===========================================================================
'  Brute-force lockout
'
'  Counted in Application scope and keyed by the submitted username, so an
'  attacker cannot tell a real account from a made-up one: both lock after
'  AUTH_MAX_FAILURES attempts, which also means the counters cannot be used
'  to enumerate valid usernames.
'
'  Stored as  "count|lockedUntilEpoch|lastAttemptEpoch"  (a plain string, so
'  the read-modify-write under Application.Lock is safe).
' ===========================================================================
Function AuthFailKey(username)
    AuthFailKey = "loginfail_" & LCase(Trim(username & ""))
End Function

' Seconds until this username may try again (0 = not locked).
Function AuthLockSeconds(username)
    AuthLockSeconds = 0
    Dim raw, parts, untilWhen
    On Error Resume Next
    raw = Application(AuthFailKey(username)) & ""
    On Error GoTo 0
    If raw = "" Then Exit Function

    parts = Split(raw, "|")
    If UBound(parts) <> 2 Then Exit Function
    If Not IsNumeric(parts(1)) Then Exit Function

    untilWhen = CLng(parts(1))
    If untilWhen <= 0 Then Exit Function
    If untilWhen <= AuthNow() Then Exit Function            ' lock has elapsed
    AuthLockSeconds = untilWhen - AuthNow()
End Function

' Human-readable wait time shown on a locked-out sign-in (used by login.asp).
Function AuthWaitText(seconds)
    Dim mins
    mins = Int((seconds + 59) / 60)
    If mins < 1 Then mins = 1
    If mins = 1 Then
        AuthWaitText = "for 1 minute"
    Else
        AuthWaitText = "for about " & mins & " minutes"
    End If
End Function

Sub AuthRegisterFailure(username)
    Dim key, raw, parts, n, lockUntil, lastTry, nowE
    key  = AuthFailKey(username)
    nowE = AuthNow()

    Application.Lock
    raw = ""
    On Error Resume Next
    raw = Application(key) & ""
    On Error GoTo 0

    n = 0 : lockUntil = 0 : lastTry = 0
    If raw <> "" Then
        parts = Split(raw, "|")
        If UBound(parts) = 2 Then
            If IsNumeric(parts(0)) Then n         = CLng(parts(0))
            If IsNumeric(parts(1)) Then lockUntil = CLng(parts(1))
            If IsNumeric(parts(2)) Then lastTry   = CLng(parts(2))
        End If
    End If

    ' A stale counter (no attempt for two lock windows) starts from scratch.
    If lastTry > 0 And (nowE - lastTry) > (AUTH_LOCK_SECONDS * 2) Then
        n = 0 : lockUntil = 0
    End If
    ' A lock that has since elapsed starts counting again too.
    If lockUntil > 0 And nowE >= lockUntil Then
        n = 0 : lockUntil = 0
    End If

    n = n + 1
    If n >= AUTH_MAX_FAILURES Then lockUntil = nowE + AUTH_LOCK_SECONDS

    Application(key) = CStr(n) & "|" & CStr(lockUntil) & "|" & CStr(nowE)
    Application.Unlock
End Sub

Sub AuthClearFailures(username)
    On Error Resume Next
    Application.Lock
    Application.Contents.Remove(AuthFailKey(username))
    Application.Unlock
    On Error GoTo 0
End Sub

' ===========================================================================
'  Credential check  (the page that calls this must include inc\db.asp)
' ===========================================================================

' How ready is sign-in, exactly?  Returned as a small state code so the UI can
' tell the difference between "the schema was never applied" and "the schema
' is there but nobody has an account yet" - two problems with two different
' fixes, neither of which should surface as a bare 500.
'
'   2 = dbo.Users exists and at least one account is provisioned (normal)
'   1 = dbo.Users exists but is empty  -> run db\provision_admin.ps1
'   0 = dbo.Users does not exist       -> run db\auth.sql first
'  -1 = the query failed unexpectedly  -> the caller raises the generic 500
'
' The page that calls this must include inc\db.asp.
Function AuthSetupState()
    AuthSetupState = -1

    Dim cn, cmd, rs, hasTable, n
    hasTable = False

    On Error Resume Next
    Set cn = OpenDb()

    ' 1. does the table exist?
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = "SELECT OBJECT_ID(N'dbo.Users', N'U')"
    Set rs = cmd.Execute
    If Err.Number <> 0 Then
        Err.Clear
        CloseRs rs
        CloseCn cn
        On Error GoTo 0
        Exit Function
    End If
    If Not rs.EOF Then If Not IsNull(rs(0).Value) Then hasTable = True
    CloseRs rs
    Set rs = Nothing
    cmd.Parameters.RemoveAll
    Set cmd = Nothing

    If Not hasTable Then
        CloseCn cn
        On Error GoTo 0
        AuthSetupState = 0
        Exit Function
    End If

    ' 2. how many accounts?
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = "SELECT COUNT(*) FROM dbo.Users"
    Set rs = cmd.Execute
    If Err.Number <> 0 Then
        Err.Clear
        CloseRs rs
        CloseCn cn
        On Error GoTo 0
        Exit Function
    End If

    n = 0
    If Not rs.EOF Then n = CLng(rs(0).Value)
    CloseRs rs
    CloseCn cn
    On Error GoTo 0

    If n > 0 Then
        AuthSetupState = 2
    Else
        AuthSetupState = 1
    End If
End Function

' Returns the user id of an ACTIVE account with exactly this password, or 0
' when the credentials do not match (0 is also returned for an empty or
' unknown username, so the caller cannot tell a bad name from a bad password).
'
' The comparison runs inside SQL Server as
'     SELECT id FROM dbo.Users
'      WHERE username = ? AND is_active = 1
'        AND password_hash = HASHBYTES('SHA2_256', salt + ?)
' so the plaintext never leaves the parameterised command, no user input is
' concatenated into SQL, and the digest is identical to the one
' db\provision_admin.ps1 writes (verified on this machine).
Function AuthAuthenticate(username, password)
    AuthAuthenticate = 0
    If username = "" Or password = "" Then Exit Function

    Dim cn, cmd, rs, id
    On Error Resume Next
    Set cn = OpenDb()
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = "SELECT id FROM dbo.Users " & _
                      "WHERE username = ? AND is_active = 1 " & _
                      "  AND password_hash = HASHBYTES('SHA2_256', salt + ?)"
    cmd.Parameters.Append cmd.CreateParameter("username", 202, 1, 64,  username)
    cmd.Parameters.Append cmd.CreateParameter("password", 202, 1, 128, password)
    Set rs = cmd.Execute
    If Err.Number <> 0 Then
        Err.Clear
        CloseRs rs
        CloseCn cn
        On Error GoTo 0
        Exit Function                       ' treated as a failed sign-in
    End If
    On Error GoTo 0

    id = 0
    If Not rs.EOF Then
        If Not IsNull(rs(0).Value) Then id = CLng(rs(0).Value)
    End If
    CloseRs rs
    CloseCn cn

    AuthAuthenticate = id
End Function

' Records a successful sign-in (audit value only, never a secret).
Sub AuthRecordLogin(userId)
    On Error Resume Next
    Dim cn, cmd, changed
    Set cn = OpenDb()
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = "UPDATE dbo.Users SET last_login_at = SYSUTCDATETIME() WHERE id = ?"
    cmd.Parameters.Append cmd.CreateParameter("id", 3, 1, 0, CLng(userId))  ' adInteger
    changed = 0
    cmd.Execute changed                                                     ' ByRef records-affected
    CloseCn cn
    On Error GoTo 0
End Sub
%>
