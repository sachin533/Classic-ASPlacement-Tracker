<% ' ---------------------------------------------------------------------------
  '  Placement Tracker  -  inc\helpers.asp
  '  ---------------------------------------------------------------------------
  '  Output encoding, HTTP status helpers, CSRF tokens, flash messages,
  '  formatting helpers and the inline SVG icons shared by every page.
  '
  '  This file MUST be the first include on every page: it switches the
  '  response to UTF-8 before any form field is read.
  '
  '  inc\auth.asp is pulled in HERE rather than by each page, so the login
  '  gate, the session timeout and the signed-in indicator in the header are
  '  available to every page exactly once (including the shared error page).
  '    ------------------------------------------------------------------------ %>
<!--#include virtual="/inc/auth.asp" -->
<%
' -- Response-wide UTF-8 -----------------------------------------------------
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
On Error Resume Next
Session.CodePage = 65001
' Session lifetime is owned by inc\auth.asp: the ASP session must never
' outlive the absolute limit the code enforces (see AuthSessionMinutes).
Session.Timeout = AuthSessionMinutes()
On Error GoTo 0

' -- Shared page state (read by inc\header.asp and inc\errorpage.asp) -------
Dim gTitle, gActiveNav, gShowSearch, gSearchValue, gFilterStatus
Dim gErrCode, gErrTitle, gErrMsg

gTitle       = ""
gActiveNav   = ""
gShowSearch  = False
gSearchValue = ""
gFilterStatus = ""

' ===========================================================================
'  Output helpers
' ===========================================================================

' Encode anything before it is written into HTML.
Function H(v)
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    H = Server.HTMLEncode(CStr(v))
End Function

' Null / Empty tolerant value.
Function Nz(v, fallback)
    If IsNull(v) Or IsEmpty(v) Then
        Nz = fallback
    ElseIf Trim(CStr(v)) = "" Then
        Nz = fallback
    Else
        Nz = v
    End If
End Function

' ===========================================================================
'  Formatting
' ===========================================================================

' Date -> "YYYY-MM-DD" (the format the form submits and the UI shows).
Function FormatIsoDate(d)
    FormatIsoDate = ""
    If IsNull(d) Or IsEmpty(d) Then Exit Function
    If Not IsDate(d) Then Exit Function
    FormatIsoDate = Right("0000" & Year(d), 4) & "-" & _
                    Right("00" & Month(d), 2) & "-" & _
                    Right("00" & Day(d), 2)
End Function

' DATETIME2 (stored in UTC) -> "YYYY-MM-DD HH:MM (UTC)".
Function FormatUtcTimestamp(d)
    FormatUtcTimestamp = "-"
    If IsNull(d) Or IsEmpty(d) Then Exit Function
    If Not IsDate(d) Then Exit Function
    FormatUtcTimestamp = FormatIsoDate(d) & " " & _
                         Right("00" & Hour(d), 2) & ":" & Right("00" & Minute(d), 2) & " (UTC)"
End Function

' "YYYY-MM-DD" for a date, a dash when the column is NULL (Wishlist rows).
Function EmptyDash(v)
    If IsNull(v) Or IsEmpty(v) Then
        EmptyDash = "-"
    ElseIf Not IsDate(v) Then
        EmptyDash = "-"
    Else
        EmptyDash = FormatIsoDate(v)
    End If
End Function

' ===========================================================================
'  Request helpers
' ===========================================================================

' Returns the id as a non-negative Long, or -1 when the value is not a clean
' non-negative integer.  Guards against "1e3", "-1", "1; DROP TABLE", "" etc.
' 0 is deliberately a VALID id: identity values start at 1 on a fresh table,
' but DBCC CHECKIDENT ... RESEED, 0 (see db\seed.sql) legitimately produces
' row id 0, and that row must still be viewable/editable/deletable.
Function ParseId(raw)
    ParseId = -1
    Dim s
    s = Trim(raw & "")
    If s = "" Then Exit Function
    If Not TestPattern("^[0-9]{1,9}$", s) Then Exit Function
    Dim n
    n = CLng(s)
    If n < 0 Then Exit Function
    ParseId = n
End Function

' Returns the dashboard page number as a positive Long.
' Anything that is not a clean positive integer - "", "0", "-3", "abc",
' "2page", "1e3" - silently falls back to page 1 instead of raising an
' error or producing an offset that SQL Server would have to reject.
Function ParsePageNumber(raw)
    ParsePageNumber = 1
    Dim s, n
    s = Trim(raw & "")
    If s = "" Then Exit Function
    If Not TestPattern("^[0-9]{1,9}$", s) Then Exit Function
    On Error Resume Next
    n = CLng(s)
    If Err.Number <> 0 Then
        Err.Clear
        n = 1
    End If
    On Error GoTo 0
    If n < 1 Then Exit Function
    ParsePageNumber = n
End Function

' Regular expression helper.
Function TestPattern(pattern, subject)
    TestPattern = False
    Dim re
    Set re = Server.CreateObject("VBScript.RegExp")
    re.Pattern = pattern
    re.IgnoreCase = False
    re.Global = False
    TestPattern = re.Test(subject & "")
    Set re = Nothing
End Function

' Case-insensitive match against the five allowed statuses; returns the
' canonical spelling, or an empty string when the value is not allowed.
Function CanonicalStatus(raw)
    CanonicalStatus = ""
    Dim allowed, i, s
    allowed = Array("Wishlist", "Applied", "Interview", "Offer", "Rejected")
    s = Trim(raw & "")
    For i = 0 To UBound(allowed)
        If StrComp(allowed(i), s, vbTextCompare) = 0 Then
            CanonicalStatus = allowed(i)
            Exit Function
        End If
    Next
End Function

Function IsHttpUrl(s)
    IsHttpUrl = False

    Dim v, low, hostPart, parts
    v = Trim(s & "")
    If v = "" Then Exit Function
    If Len(v) > 500 Then Exit Function
    If InStr(v, " ") > 0 Then Exit Function
    If TestPattern("[\x00-\x1F]", v) Then Exit Function
    If TestPattern("[<>""']", v) Then Exit Function

    low = LCase(v)
    If Left(low, 8) = "https://" Then
        hostPart = Mid(low, 9)
    ElseIf Left(low, 7) = "http://" Then
        hostPart = Mid(low, 8)
    Else
        Exit Function
    End If

    hostPart = Split(hostPart, "/")(0)
    hostPart = Split(hostPart, "?")(0)
    hostPart = Split(hostPart, "#")(0)

    ' optional port
    If InStr(hostPart, ":") > 0 Then
        Dim portNo
        portNo = Split(hostPart, ":")(1)
        If Not TestPattern("^[0-9]{1,5}$", portNo) Then Exit Function
        hostPart = Split(hostPart, ":")(0)
    End If

    If hostPart = "" Then Exit Function
    If hostPart <> "localhost" Then
        If InStr(hostPart, ".") = 0 Then Exit Function
        If Left(hostPart, 1) = "." Or Right(hostPart, 1) = "." Then Exit Function
        If InStr(hostPart, "..") > 0 Then Exit Function
    End If
    If TestPattern("[^a-z0-9\.\-]", hostPart) Then Exit Function

    IsHttpUrl = True
End Function

' ===========================================================================
'  HTTP status helpers
' ===========================================================================

' Renders a complete, friendly error page with the right status code.
'   code   : 403 / 404 / 405 / 500 ...
'   title  : short headline
'   message: sentence shown under the headline
'
Sub FailPage(code, title, message)
    Response.Clear
    ' Stop IIS from re-processing the status we are about to set; otherwise its
    ' own <httpErrors> mapping fires again and swallows the rendered body.
    On Error Resume Next
    Response.TrySkipIisCustomErrors = True
    On Error GoTo 0

    Select Case CLng(code)
        Case 400
            Response.Status = "400 Bad Request"
        Case 403
            Response.Status = "403 Forbidden"
        Case 404
            Response.Status = "404 Not Found"
        Case 405
            Response.Status = "405 Method Not Allowed"
            On Error Resume Next
            Response.AddHeader "Allow", "POST"
            On Error GoTo 0
        Case 500
            Response.Status = "500 Internal Server Error"
        Case Else
            Response.Status = CStr(code) & " Error"
    End Select

    gErrCode   = CLng(code)
    gErrTitle  = title
    gErrMsg    = message
    gShowSearch = False

    RenderErrorPage CLng(code), title, message
    Response.End
End Sub

Sub NotFound()
    FailPage 404, "Page not found", _
             "We could not find that application. It may have been deleted, " & _
             "or the link may be incorrect."
End Sub

' ===========================================================================
'  CSRF protection
' ===========================================================================

Function NewToken()
    NewToken = LCase(SafeGuid() & SafeGuid())
    If Len(NewToken) < 32 Then
        ' Deterministic-but-unavailable fallback (Scriptlet.TypeLib blocked).
        Dim i, s
        Randomize
        s = ""
        For i = 1 To 64
            s = s & LCase(Hex(Int(Rnd * 16)))
        Next
        NewToken = s & LCase(Hex(Year(Now)) & Hex(Month(Now)) & Hex(Day(Now)))
    End If
End Function

Function SafeGuid()
    SafeGuid = ""
    On Error Resume Next
    Dim tl, g
    Set tl = Server.CreateObject("Scriptlet.TypeLib")
    If Err.Number = 0 Then
        g = Left(tl.GUID, 38)
        g = Replace(Replace(Replace(g, "{", ""), "}", ""), "-", "")
        If Err.Number = 0 Then
            If Len(g) = 32 Then SafeGuid = g
        End If
        Err.Clear
        Set tl = Nothing
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Function

' One token per session, created on demand.
Function CsrfToken()
    Dim t
    On Error Resume Next
    t = Session("csrf")
    On Error GoTo 0
    If IsNull(t) Or IsEmpty(t) Then t = ""
    If t = "" Then
        t = NewToken()
        Session("csrf") = t
    End If
    CsrfToken = t
End Function

' Hidden input for the shared forms.
Function CsrfInput()
    CsrfInput = "<input type=""hidden"" name=""csrf"" value=""" & H(CsrfToken()) & """>"
End Function

' Rejects any state-changing request whose token does not match.  Status 403.
Sub RequireCsrf()
    Dim posted
    posted = Request.Form("csrf") & ""
    If StrComp(posted, CsrfToken(), vbBinaryCompare) <> 0 Then
        FailPage 403, "Request blocked", _
                 "Your session token was missing or is no longer valid. " & _
                 "Go back to the application list, refresh the page and try again."
    End If
End Sub

' ===========================================================================
'  Flash messages (survive the POST/Redirect/GET round trip)
' ===========================================================================
Sub SetFlash(kind, message)
    Session("flash_kind")   = kind & ""
    Session("flash_message") = message & ""
End Sub

Function RenderFlash()
    RenderFlash = ""
    Dim kind, message
    On Error Resume Next
    kind    = Session("flash_kind") & ""
    message = Session("flash_message") & ""
    Session("flash_kind")    = ""
    Session("flash_message") = ""
    On Error GoTo 0

    If message = "" Then Exit Function

    Dim cls
    Select Case kind
        Case "error" : cls = "flash flash-error"
        Case "info"  : cls = "flash flash-info"
        Case Else    : cls = "flash flash-success"
    End Select

    RenderFlash = "<div class=""flash-wrap""><div class=""" & cls & """ role=""status"">" & _
                  "<span class=""flash-icon"" aria-hidden=""true"">" & IconCheck() & "</span>" & _
                  "<span>" & H(message) & "</span>" & _
                  "</div></div>"
End Function

' ===========================================================================
'  Inline SVG icons (feather-style, sized by CSS)
' ===========================================================================
Function IconSearch()
    IconSearch = "<svg class=""icon"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" " & _
                 "stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true"">" & _
                 "<circle cx=""11"" cy=""11"" r=""8""/><line x1=""21"" y1=""21"" x2=""16.65"" y2=""16.65""/></svg>"
End Function

Function IconPlus()
    IconPlus = "<svg class=""icon"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" " & _
               "stroke-width=""2.2"" stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true"">" & _
               "<line x1=""12"" y1=""5"" x2=""12"" y2=""19""/><line x1=""5"" y1=""12"" x2=""19"" y2=""12""/></svg>"
End Function

Function IconBack()
    IconBack = "<svg class=""icon icon-sm"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" " & _
               "stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true"">" & _
               "<line x1=""19"" y1=""12"" x2=""5"" y2=""12""/><polyline points=""12 19 5 12 12 5""/></svg>"
End Function

Function IconEdit()
    IconEdit = "<svg class=""icon icon-sm"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" " & _
               "stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true"">" & _
               "<path d=""M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7""/>" & _
               "<path d=""M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z""/></svg>"
End Function

Function IconTrash()
    IconTrash = "<svg class=""icon icon-sm"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" " & _
                "stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true"">" & _
                "<polyline points=""3 6 5 6 21 6""/><path d=""M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6""/>" & _
                "<path d=""M10 11v6""/><path d=""M14 11v6""/>" & _
                "<path d=""M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2""/></svg>"
End Function

Function IconBriefcase()
    IconBriefcase = "<svg class=""icon icon-lg"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" " & _
                    "stroke-width=""1.8"" stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true"">" & _
                    "<rect x=""2"" y=""7"" width=""20"" height=""14"" rx=""2"" ry=""2""/>" & _
                    "<path d=""M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16""/></svg>"
End Function

Function IconWarning()
    IconWarning = "<svg class=""icon icon-lg"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" " & _
                  "stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true"">" & _
                  "<line x1=""12"" y1=""7"" x2=""12"" y2=""13""/>" & _
                  "<line x1=""12"" y1=""17"" x2=""12.01"" y2=""17""/></svg>"
End Function

Function IconCheck()
    IconCheck = "<svg class=""icon icon-sm"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" " & _
                "stroke-width=""2.4"" stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true"">" & _
                "<polyline points=""20 6 9 17 4 12""/></svg>"
End Function

' --- icons used by the sign-in screens and the signed-in header -------------
Function IconCap()
    IconCap = "<svg class=""icon"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" " & _
              "stroke-width=""1.9"" stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true"">" & _
              "<path d=""M22 10L12 5 2 10l10 5 10-5z""/><path d=""M6 12.5V17c0 1.7 2.7 3 6 3s6-1.3 6-3v-4.5""/>" & _
              "<line x1=""22"" y1=""10"" x2=""22"" y2=""15""/></svg>"
End Function

Function IconColumns()
    IconColumns = "<svg class=""icon"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" " & _
                  "stroke-width=""1.8"" stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true"">" & _
                  "<rect x=""3"" y=""3"" width=""18"" height=""18"" rx=""2"" ry=""2""/>" & _
                  "<line x1=""9"" y1=""3"" x2=""9"" y2=""21""/><line x1=""15"" y1=""3"" x2=""15"" y2=""21""/></svg>"
End Function

Function IconUser()
    IconUser = "<svg class=""icon"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" " & _
               "stroke-width=""1.8"" stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true"">" & _
               "<path d=""M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2""/><circle cx=""12"" cy=""7"" r=""4""/></svg>"
End Function

Function IconLock()
    IconLock = "<svg class=""icon"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" " & _
               "stroke-width=""1.8"" stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true"">" & _
               "<rect x=""3"" y=""11"" width=""18"" height=""11"" rx=""2"" ry=""2""/>" & _
               "<path d=""M7 11V7a5 5 0 0 1 10 0v4""/></svg>"
End Function

Function IconSignIn()
    IconSignIn = "<svg class=""icon"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" " & _
                 "stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true"">" & _
                 "<path d=""M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4""/>" & _
                 "<polyline points=""10 17 15 12 10 7""/><line x1=""15"" y1=""12"" x2=""3"" y2=""12""/></svg>"
End Function

Function IconShield()
    IconShield = "<svg class=""icon"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" " & _
                 "stroke-width=""1.9"" stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true"">" & _
                 "<path d=""M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z""/><polyline points=""9 12 11 14 15 10""/></svg>"
End Function

' Navigation highlighting: which top-bar link is the current page.
Function NavClass(sectionName)
    NavClass = ""
    If StrComp(gActiveNav & "", sectionName, vbTextCompare) = 0 Then NavClass = " is-active"
End Function

Function NavCurrent(sectionName)
    NavCurrent = ""
    If StrComp(gActiveNav & "", sectionName, vbTextCompare) = 0 Then NavCurrent = " aria-current=""page"""
End Function

' Status badge with the palette used across the dashboard and details page.
Function StatusBadge(status)
    Dim key, cls
    key = LCase(Trim(status & ""))
    Select Case key
        Case "wishlist"  : cls = "badge badge-wishlist"
        Case "applied"   : cls = "badge badge-applied"
        Case "interview" : cls = "badge badge-interview"
        Case "offer"     : cls = "badge badge-offer"
        Case "rejected"  : cls = "badge badge-rejected"
        Case Else        : cls = "badge badge-unknown"
    End Select
    StatusBadge = "<span class=""" & cls & """>" & H(status) & "</span>"
End Function
%>
<!--#include virtual="/inc/errorpage.asp" -->
