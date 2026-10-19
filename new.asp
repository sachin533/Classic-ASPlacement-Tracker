<!--#include virtual="/inc/helpers.asp" -->
<!--#include virtual="/inc/db.asp" -->
<!--#include virtual="/inc/validate.asp" -->
<!--#include virtual="/inc/form.asp" -->
<%
' ---------------------------------------------------------------------------
'  Create - new.asp
'  GET  : blank form with a CSRF token
'  POST : validate, parameterised INSERT, then POST/Redirect/GET to view.asp
' ---------------------------------------------------------------------------

EnsureAuthenticated

Dim values, errors, isNewId
Set values = Server.CreateObject("Scripting.Dictionary")
Set errors = Server.CreateObject("Scripting.Dictionary")
isNewId = 0

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

    RequireCsrf

    Set values = CollectApplicationValues()
    Set errors = ValidateApplication(values)

    If errors.Count = 0 Then
        isNewId = InsertApplication(values)

        If isNewId > 0 Then
            SetFlash "success", "Application created successfully."
            Response.Redirect "/view.asp?id=" & isNewId
            Response.End
        End If

        ' The insert wrote nothing - treat it as a form error so the values
        ' the user typed are never lost.
        errors("company") = "The application could not be saved. Please try again."
    End If

End If

gTitle       = "Add New Application"
gActiveNav   = "new"
gShowSearch  = False
gSearchValue = ""
%>
<% Call RenderHeader() %>

<a class="back-link" href="/default.asp"><%= IconBack() %>Back to Applications</a>

<div class="page-head">
  <div>
    <h1>Add New Application</h1>
    <p class="subtitle">Enter the details of the job or internship application.</p>
  </div>
</div>

<div class="card">
  <% If errors.Count > 0 Then %>
  <div class="form-alert-wrap">
    <p class="form-alert" role="alert"><%= H(ValidationSummary(errors)) %></p>
  </div>
  <% End If %>

  <% RenderApplicationForm "/new.asp", "Save Application", "", values, errors %>
</div>

<% Call RenderFooter() %>

