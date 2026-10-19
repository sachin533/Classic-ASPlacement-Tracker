<!--#include virtual="/inc/helpers.asp" -->
<!--#include virtual="/inc/db.asp" -->
<%
' ---------------------------------------------------------------------------
'  Delete - delete.asp
'
'  POST only.  Anything else is a 405 with an Allow: POST header, so a
'  bookmarked or scraped link can never remove a row.  The id and the CSRF
'  token are both re-checked on the server before the DELETE runs.
'
'  Expected body:  id=<integer>&csrf=<session token>
' ---------------------------------------------------------------------------

If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    FailPage 405, "Method not allowed", _
             "Applications can only be deleted from the application list " & _
             "or the details page."
End If

' The method is checked first (405 answers nothing about the data), then the
' session - so an anonymous POST cannot reach the delete path or its CSRF
' check: it is simply sent to the sign-in page.
EnsureAuthenticated

RequireCsrf

Dim id, rec, removed
id = ParseId(Request.Form("id"))
If id < 0 Then NotFound()

Set rec = GetApplication(id)
If rec Is Nothing Then NotFound()

Dim deletedCompany
deletedCompany = rec("company")
Set rec = Nothing

removed = DeleteApplication(id)
If removed = 0 Then NotFound()

SetFlash "success", "Application for " & deletedCompany & " was deleted."
Response.Redirect "/default.asp"
Response.End
%>
