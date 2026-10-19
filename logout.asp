<!--#include virtual="/inc/helpers.asp" -->
<%
' ---------------------------------------------------------------------------
'  Sign out - logout.asp
'
'  POST only, with the session CSRF token, exactly like delete.asp: a GET
'  (including a link prefetch or an <img src> planted on another site) can
'  never end somebody's session, so it answers 405 instead.
'
'  Session.Abandon destroys the session server-side, and the browser is sent
'  to /login.asp?bye=1 so the sign-in page can confirm it.
' ---------------------------------------------------------------------------

If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    FailPage 405, "Method not allowed", _
             "You can only sign out from the button in the top bar."
End If

RequireCsrf

Session.Abandon

Response.Redirect "/login.asp?bye=1"
Response.End
%>
