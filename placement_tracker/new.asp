<!--#include file="inc/header.asp"-->
<!--#include file="inc/helpers.asp"-->
<!--#include file="inc/db.asp"-->
<!--#include file="inc/validate.asp"-->
<!--#include file="inc/form.asp"-->
<%
' new.asp - create one application.
' GET  shows an empty form. POST validates, inserts with parameters,
'      then follows Post/Redirect/Get to the new record's details page.
Dim vals, errs
Set vals = EmptyVals()

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

  VerifyCsrf

  vals("company")    = Request.Form("company")
  vals("role")       = Request.Form("role")
  vals("status")     = Request.Form("status")
  vals("applied_on") = Request.Form("applied_on")
  vals("job_url")    = Request.Form("job_url")
  vals("notes")      = Request.Form("notes")

  Set errs = ValidateApplication(vals)

  If errs.Count = 0 Then
    Dim cn, cmd, pDate, pUrl, pNotes, rs, newId
    Set cn = OpenDb()
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandText = "INSERT INTO dbo.Applications (company, role, status, applied_on, job_url, notes) " _
                    & "VALUES (?, ?, ?, ?, ?, ?); SELECT SCOPE_IDENTITY() AS new_id;"

    cmd.Parameters.Append cmd.CreateParameter("company", 200, 2, 100, vals("company"))   ' adVarWChar, adParamInput, 100
    cmd.Parameters.Append cmd.CreateParameter("role", 200, 2, 100, vals("role"))
    cmd.Parameters.Append cmd.CreateParameter("status", 200, 2, 20, vals("status"))

    ' Empty date -> NULL (allowed for Wishlist rows only; the CHECK
    ' constraint catches any other combination).
    Set pDate = cmd.CreateParameter("applied_on", 111, 1)                                ' adDBDate
    If Trim(vals("applied_on")) = "" Then pDate.Value = Null Else pDate.Value = CDate(vals("applied_on"))
    cmd.Parameters.Append pDate

    Set pUrl = cmd.CreateParameter("job_url", 200, 2, 500)                               ' adVarWChar
    If Trim(vals("job_url")) = "" Then pUrl.Value = Null Else pUrl.Value = Trim(vals("job_url"))
    cmd.Parameters.Append pUrl

    Set pNotes = cmd.CreateParameter("notes", 200, 2, 1000)                              ' adVarWChar
    If Trim(vals("notes")) = "" Then pNotes.Value = Null Else pNotes.Value = vals("notes")
    cmd.Parameters.Append pNotes

    Dim out
    Set out = cmd.Execute
    newId = CLng(out("new_id"))
    out.Close
    cn.Close
    Set out = Nothing : Set cmd = Nothing : Set cn = Nothing

    SetFlash "Application saved."
    Response.Redirect "view.asp?id=" & newId
  End If

Else
  ' Fresh GET - no errors to show yet.
  Set errs = Server.CreateObject("Scripting.Dictionary")
End If
%>

<a class="back-link" href="/default.asp">&larr; Back to Applications</a>

<div class="page-head">
  <div>
    <h1>Add New Application</h1>
    <p class="subtitle">Enter the details of the job or internship application.</p>
  </div>
</div>

<% Call RenderAppForm("create", vals, errs, "/default.asp") %>

<!--#include file="inc/footer.asp"-->
