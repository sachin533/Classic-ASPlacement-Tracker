<!--#include file="inc/header.asp"-->
<!--#include file="inc/helpers.asp"-->
<!--#include file="inc/db.asp"-->
<!--#include file="inc/validate.asp"-->
<!--#include file="inc/form.asp"-->
<%
' edit.asp?id=N - update one application.
' GET  loads the record into the shared form.
' POST validates and runs a parameterized UPDATE, then redirects to details.
Dim rawId, id
rawId = Trim(Request.QueryString("id"))
If rawId = "" Or Not RegExpIsDigits(rawId) Then NotFound()
id = CLng(rawId)

Dim vals, errs, cn, cmd, rs, isPost
Set vals = EmptyVals()
isPost = (Request.ServerVariables("REQUEST_METHOD") = "POST")

If isPost Then

  VerifyCsrf

  vals("company")    = Request.Form("company")
  vals("role")       = Request.Form("role")
  vals("status")     = Request.Form("status")
  vals("applied_on") = Request.Form("applied_on")
  vals("job_url")    = Request.Form("job_url")
  vals("notes")      = Request.Form("notes")

  Set errs = ValidateApplication(vals)

  If errs.Count = 0 Then
    Dim pDate, pUrl, pNotes, affected
    Set cn = OpenDb()
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandText = "UPDATE dbo.Applications " _
                    & "SET company = ?, role = ?, status = ?, applied_on = ?, job_url = ?, notes = ? " _
                    & "WHERE id = ?"

    cmd.Parameters.Append cmd.CreateParameter("company", 200, 2, 100, vals("company"))
    cmd.Parameters.Append cmd.CreateParameter("role", 200, 2, 100, vals("role"))
    cmd.Parameters.Append cmd.CreateParameter("status", 200, 2, 20, vals("status"))

    Set pDate = cmd.CreateParameter("applied_on", 111, 1)                                ' adDBDate
    If Trim(vals("applied_on")) = "" Then pDate.Value = Null Else pDate.Value = CDate(vals("applied_on"))
    cmd.Parameters.Append pDate

    Set pUrl = cmd.CreateParameter("job_url", 200, 2, 500)
    If Trim(vals("job_url")) = "" Then pUrl.Value = Null Else pUrl.Value = Trim(vals("job_url"))
    cmd.Parameters.Append pUrl

    Set pNotes = cmd.CreateParameter("notes", 200, 2, 1000)
    If Trim(vals("notes")) = "" Then pNotes.Value = Null Else pNotes.Value = vals("notes")
    cmd.Parameters.Append pNotes

    cmd.Parameters.Append cmd.CreateParameter("id", 3, 1, , id)                          ' adInteger

    cmd.Execute affected
    cn.Close
    Set cmd = Nothing : Set cn = Nothing

    ' 0 rows means someone else deleted the record while this form was open.
    If affected = 0 Then NotFound()

    SetFlash "Application updated."
    Response.Redirect "view.asp?id=" & id
  End If

Else
  ' GET - pre-fill the shared form from the database.
  Set errs = Server.CreateObject("Scripting.Dictionary")
  Set cn = OpenDb()
  Set cmd = Server.CreateObject("ADODB.Command")
  Set cmd.ActiveConnection = cn
  cmd.CommandText = "SELECT id, company, role, status, applied_on, job_url, notes " _
                  & "FROM dbo.Applications WHERE id = ?"
  cmd.Parameters.Append cmd.CreateParameter("id", 3, 1, , id)                            ' adInteger
  Set rs = cmd.Execute

  If rs.EOF Then
    rs.Close : cn.Close
    NotFound()
  End If

  vals("company")    = rs("company")
  vals("role")       = rs("role")
  vals("status")     = rs("status")
  vals("applied_on") = FmtDate(rs("applied_on"))   ' yyyy-mm-dd, ready for the date input
  If vals("applied_on") = "-" Then vals("applied_on") = ""
  vals("job_url")    = rs("job_url")
  vals("notes")      = rs("notes")

  rs.Close
  cn.Close
  Set rs = Nothing : Set cmd = Nothing : Set cn = Nothing
End If
%>

<a class="back-link" href="/default.asp">&larr; Back to Applications</a>

<div class="page-head">
  <div>
    <h1>Edit Application</h1>
    <p class="subtitle">Update the application details.</p>
  </div>
</div>

<% Call RenderAppForm("edit", vals, errs, "/view.asp?id=" & id) %>

<!--#include file="inc/footer.asp"-->
