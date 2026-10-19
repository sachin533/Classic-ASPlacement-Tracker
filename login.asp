<!--#include virtual="/inc/helpers.asp" -->
<!--#include virtual="/inc/db.asp" -->
<%
' ---------------------------------------------------------------------------
'  Sign in - login.asp
'
'  The only page (with /logout.asp and /forgot.asp) that is reachable
'  without a session.  Everything else calls EnsureAuthenticated first.
'
'  Two-step sign-in:
'    POST /login.asp        credentials + CSRF -> issues a one-time ticket
'                           and abandons the session the password was typed
'                           into (session id rotation at the moment of login)
'    GET  /login.asp?ticket=  redeems it, starts the real session and
'                           redirects to wherever the user was headed
'
'  No password is ever written to a file, a log or the response.
' ---------------------------------------------------------------------------

Dim rawUser, username, password, errorMsg, infoMsg
Dim setupState, isPost, lockSecs, userId, ticket, uid, uname, nextUrl

rawUser    = Trim(Request.Form("username") & "")
username   = LCase(rawUser)
password   = Request.Form("password") & ""
errorMsg   = ""
infoMsg    = ""

' --- messages carried in the query string (set after a redirect) ----------
If Trim(Request.QueryString("bye") & "") <> "" Then
    infoMsg = "You have been signed out."
End If
If Trim(Request.QueryString("reason") & "") = "expired" Then
    infoMsg = "Your session timed out. Please sign in again."
End If

' --- step 2: redeem the one-time ticket issued by the POST ---------------
ticket = Trim(Request.QueryString("ticket") & "")
If ticket <> "" Then
    uid = AuthTicketConsume(ticket, uname, nextUrl)
    If uid > 0 Then
        AuthStartSession uid, uname
        Response.Redirect AuthSafeNext(nextUrl)
        Response.End
    End If
    errorMsg = "That sign-in link has expired or was already used. Please sign in again."
End If

' --- already signed in? --------------------------------------------------
If IsAuthenticated() Then
    Response.Redirect "/default.asp"
    Response.End
End If

' --- how ready is sign-in? ------------------------------------------------
setupState = AuthSetupState()
If setupState < 0 Then DbFail      ' unexpected query failure -> generic 500

' --- POST ----------------------------------------------------------------
isPost = (Request.ServerVariables("REQUEST_METHOD") = "POST")

If isPost Then
    RequireCsrf                    ' missing/foreign token -> 403

    If setupState = 0 Then
        errorMsg = "Sign-in storage has not been created yet. " & _
                   "Run db\auth.sql, then db\provision_admin.ps1."

    ElseIf setupState = 1 Then
        errorMsg = "No administrator account has been created yet. " & _
                   "Create one with db\provision_admin.ps1 and sign in again."

    ElseIf Not TestPattern("^[a-z0-9_.\-]{3,64}$", username) Then
        ' Not a well-formed name, so it cannot match any account; say the
        ' same thing as a wrong password so responses stay indistinguishable.
        errorMsg = "That username or password is not correct."

    Else
        lockSecs = AuthLockSeconds(username)
        If lockSecs > 0 Then
            errorMsg = "Too many failed sign-in attempts for this account. " & _
                       "Please wait " & AuthWaitText(lockSecs) & " and try again."
        Else
            userId = AuthAuthenticate(username, password)
            If userId = 0 Then
                AuthRegisterFailure username
                errorMsg = "That username or password is not correct."
            Else
                AuthClearFailures username
                AuthRecordLogin userId

                ' Hand the result to a brand-new session through a ticket
                ' that can only be redeemed once, in the next request.
                nextUrl = AuthSafeNext(Session("login_next") & "")
                ticket  = AuthTicketCreate(userId, username, nextUrl)

                Session.Abandon
                Response.Redirect "/login.asp?ticket=" & ticket
                Response.End
            End If
        End If
    End If
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex">
<title>Sign in | Placement Tracker</title>
<link rel="stylesheet" href="/static/styles.css?v=4">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24'%3E%3Crect width='24' height='24' rx='5' fill='%232563eb'/%3E%3Cpath d='M7 9h10v8H7z' fill='none' stroke='white' stroke-width='1.6'/%3E%3Cpath d='M10 9V7.6h4V9' fill='none' stroke='white' stroke-width='1.6'/%3E%3C/svg%3E">
</head>
<body class="auth-body">

<div class="auth-bg" aria-hidden="true"></div>

<header class="auth-topbar">
  <div class="auth-topbar-inner">
    <span class="brand auth-brand"><span class="brand-mark"><%= IconCap() %></span>Placement Tracker</span>
    <span class="auth-tagline">
      <span class="auth-tagline-mark"><%= IconColumns() %></span>
      <span>Track</span><span class="auth-tagline-sep" aria-hidden="true"></span>
      <span>Manage</span><span class="auth-tagline-sep" aria-hidden="true"></span>
      <span>Grow</span>
    </span>
  </div>
</header>

<main class="auth-main" id="main">
  <section class="auth-card" aria-labelledby="signin-title">

    <div class="auth-avatar" aria-hidden="true"><%= IconCap() %></div>

    <h1 class="auth-title" id="signin-title">Placement Tracker</h1>
    <p class="auth-sub">Sign in to continue to your dashboard</p>

    <% If errorMsg <> "" Then %>
    <p class="auth-alert auth-alert-error" role="alert"><%= H(errorMsg) %></p>
    <% ElseIf infoMsg <> "" Then %>
    <p class="auth-alert auth-alert-info" role="status"><%= H(infoMsg) %></p>
    <% End If %>

    <% If setupState = 0 Then %>
    <p class="auth-alert auth-alert-setup" role="status">
      Sign-in storage does not exist yet, so nobody can sign in.
      From the project root run <code>db\auth.sql</code> and then
      <code>db\provision_admin.ps1</code>.
    </p>
    <% ElseIf setupState = 1 Then %>
    <p class="auth-alert auth-alert-setup" role="status">
      No administrator account exists yet, so nobody can sign in.
      Create one from the project root with
      <code>db\provision_admin.ps1</code>.
    </p>
    <% End If %>

    <form class="auth-form" method="post" action="/login.asp" autocomplete="on">
      <%= CsrfInput() %>

      <div class="auth-field">
        <label class="auth-label" for="username">Username</label>
        <div class="auth-input-wrap">
          <span class="auth-input-icon" aria-hidden="true"><%= IconUser() %></span>
          <input class="auth-input" type="text" id="username" name="username"
                 placeholder="Enter your username"
                 value="<%= H(rawUser) %>"
                 maxlength="64" autocomplete="username" autocapitalize="none"
                 spellcheck="false" required autofocus>
        </div>
      </div>

      <div class="auth-field">
        <div class="auth-label-row">
          <label class="auth-label" for="password">Password</label>
          <a class="auth-forgot" href="/forgot.asp">Forgot password?</a>
        </div>
        <div class="auth-input-wrap">
          <span class="auth-input-icon" aria-hidden="true"><%= IconLock() %></span>
          <input class="auth-input" type="password" id="password" name="password"
                 placeholder="Enter your password"
                 maxlength="128" autocomplete="current-password" required>
        </div>
      </div>

      <button class="auth-submit" type="submit">
        <%= IconSignIn() %><span>Sign In</span>
      </button>
    </form>

    <div class="auth-divider" aria-hidden="true"><span>or</span></div>

    <div class="auth-note">
      <span class="auth-note-icon" aria-hidden="true"><%= IconShield() %></span>
      <div class="auth-note-text">
        <strong>Secure Access</strong>
        <p>Only authorized users can access the Placement Tracker.</p>
      </div>
    </div>

  </section>
</main>

</body>
</html>
