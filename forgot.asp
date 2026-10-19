<!--#include virtual="/inc/helpers.asp" -->
<!--#include virtual="/inc/db.asp" -->
<%
' ---------------------------------------------------------------------------
'  Password help - forgot.asp
'
'  There is deliberately NO self-service password reset: this application has
'  no mail service, no recovery e-mail address is collected anywhere, and
'  inventing a reset flow would be a lie in the interface.  This page says
'  what actually happens instead - the administrator account is (re)created
'  locally by db\provision_admin.ps1, which is the real, working procedure.
'
'  Reachable without a session, because somebody who cannot sign in is
'  exactly the person who needs it.
' ---------------------------------------------------------------------------

If Request.ServerVariables("REQUEST_METHOD") <> "GET" Then
    FailPage 405, "Method not allowed", "This page only answers GET."
End If

Dim setupState, resetLabel
setupState = AuthSetupState()
If setupState < 0 Then DbFail      ' unexpected query failure -> generic 500

resetLabel = "Provisioning"
If setupState = 2 Then resetLabel = "Provisioning or resetting"
%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex">
<title>Reset your password | Placement Tracker</title>
<link rel="stylesheet" href="/static/styles.css?v=4">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24'%3E%3Crect width='24' height='24' rx='5' fill='%232563eb'/%3E%3Cpath d='M7 9h10v8H7z' fill='none' stroke='white' stroke-width='1.6'/%3E%3Cpath d='M10 9V7.6h4V9' fill='none' stroke='white' stroke-width='1.6'/%3E%3C/svg%3E">
</head>
<body class="auth-body">

<div class="auth-bg" aria-hidden="true"></div>

<header class="auth-topbar">
  <div class="auth-topbar-inner">
    <a class="brand auth-brand" href="/login.asp"><span class="brand-mark"><%= IconCap() %></span>Placement Tracker</a>
    <span class="auth-tagline">
      <span class="auth-tagline-mark"><%= IconColumns() %></span>
      <span>Track</span><span class="auth-tagline-sep" aria-hidden="true"></span>
      <span>Manage</span><span class="auth-tagline-sep" aria-hidden="true"></span>
      <span>Grow</span>
    </span>
  </div>
</header>

<main class="auth-main" id="main">
  <section class="auth-card" aria-labelledby="forgot-title">

    <div class="auth-avatar" aria-hidden="true"><%= IconLock() %></div>

    <h1 class="auth-title" id="forgot-title">Password help</h1>
    <p class="auth-sub">There is no e-mail reset for this application</p>

    <div class="auth-note auth-note-block">
      <span class="auth-note-icon" aria-hidden="true"><%= IconShield() %></span>
      <div class="auth-note-text">
        <strong>Why there is no reset link</strong>
        <p>Placement Tracker has no mail service and never collects a
           recovery address, so a "reset link" could not actually be sent.
           Resetting a password is therefore a local administrator action.</p>
      </div>
    </div>

    <div class="auth-steps">
      <p><strong><%= H(resetLabel) %></strong> an account runs from the
         project root and prompts for the password, which is never stored in
         plaintext:</p>
      <pre class="auth-code">powershell -ExecutionPolicy Bypass -File .\db\provision_admin.ps1</pre>
      <p>The script creates <code>dbo.Users</code> if it is missing, then stores a
         random salt together with SHA-256(salt + password) in
         <code>dbo.Users</code>. The password itself is not written to the
         database, to a file, or to the screen.</p>
      <% If setupState < 2 Then %>
      <p class="auth-alert auth-alert-setup" role="status">
        <% If setupState = 0 Then %>Sign-in storage does not exist yet, so nobody can sign in. Run <code>db\auth.sql</code> first.
        <% Else %>No account exists yet, so nobody can sign in until this script is run once.<% End If %>
      </p>
      <% End If %>
    </div>

    <a class="auth-submit auth-submit-link" href="/login.asp">
      <%= IconSignIn() %><span>Back to Sign In</span>
    </a>

  </section>
</main>

</body>
</html>
