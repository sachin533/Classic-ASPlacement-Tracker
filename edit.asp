<!--#include virtual="/inc/helpers.asp" -->
<!--#include virtual="/inc/db.asp" -->
<!--#include virtual="/inc/validate.asp" -->
<!--#include virtual="/inc/form.asp" -->
<%
' ---------------------------------------------------------------------------
'  Update - edit.asp?id=
'  GET  : the form prefilled from the existing row
'  POST : validate, parameterised UPDATE, then POST/Redirect/GET to view.asp
'         A bad or unknown id is a 404 either way.
' ---------------------------------------------------------------------------

EnsureAuthenticated

Dim id, rec, values, errors, updated
id = ParseId(Request.QueryString("id"))

Set values = Server.CreateObject("Scripting.Dictionary")
Set errors = Server.CreateObject("Scripting.Dictionary")

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

    RequireCsrf

    ' Prefer the id that was posted with the form; fall back to the query
    ' string.  Either way it is re-validated before it touches SQL.
    Dim postedId
    postedId = ParseId(Request.Form("id"))
    If postedId < 0 Then postedId = ParseId(Request.QueryString("id"))

    If postedId < 0 Then NotFound()
    id = postedId

    Set rec = GetApplication(id)
    If rec Is Nothing Then NotFound()

    Set values = CollectApplicationValues()
    Set errors = ValidateApplication(values)

    If errors.Count = 0 Then
        updated = UpdateApplication(id, values)

        If updated > 0 Then
            SetFlash "success", "Application updated successfully."
            Response.Redirect "/view.asp?id=" & id
            Response.End
        End If

        errors("company") = "The application could not be updated. Please try again."
    End If

    Set rec = Nothing
Else
    If id < 0 Then NotFound()
    Set rec = GetApplication(id)
    If rec Is Nothing Then NotFound()

    values("company")       = rec("company")
    values("role")          = rec("role")
    values("status")        = rec("status")
    values("applied_on")    = rec("applied_on_iso")
    values("applied_on_iso") = rec("applied_on_iso")
    values("job_url")       = rec("job_url")
    values("notes")         = rec("notes")
    values("status_canonical") = rec("status")
End If

gTitle       = "Edit Application"
gActiveNav   = "new"
gShowSearch  = False
gSearchValue = ""
%>
<% Call RenderHeader() %>

<a class="back-link" href="/default.asp"><%= IconBack() %>Back to Applications</a>

<div class="page-head">
  <div>
    <h1>Edit Application</h1>
    <p class="subtitle">Update the application details.</p>
  </div>
</div>

<div class="card">
  <% If errors.Count > 0 Then %>
  <div class="form-alert-wrap">
    <p class="form-alert" role="alert"><%= H(ValidationSummary(errors)) %></p>
  </div>
  <% End If %>

  <% RenderApplicationForm "/edit.asp?id=" & id, "Update Application", CStr(id), values, errors %>
</div>

<% Call RenderFooter() %>

