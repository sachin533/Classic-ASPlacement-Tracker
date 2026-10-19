<!--#include file="inc/header.asp"-->
<!--#include file="inc/helpers.asp"-->
<!--#include file="inc/db.asp"-->
<%
' delete.asp - deletes one application.
' POST only (id + CSRF token in the body). A GET answers 405 and removes
' nothing, so no ordinary link click or prefetch can destroy data.
If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then MethodNotAllowed()

VerifyCsrf

Dim rawId, id
rawId = Trim(Request.Form("id"))
If rawId = "" Or Not RegExpIsDigits(rawId) Then NotFound()
id = CLng(rawId)

Dim cn, cmd, affected
Set cn = OpenDb()
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = cn
cmd.CommandText = "DELETE FROM dbo.Applications WHERE id = ?"
cmd.Parameters.Append cmd.CreateParameter("id", 3, 1, , id)   ' adInteger, adParamInput
cmd.Execute affected
cn.Close
Set cmd = Nothing : Set cn = Nothing

' The row was already gone - say so instead of pretending it worked.
If affected = 0 Then NotFound()

SetFlash "Application deleted."
Response.Redirect "default.asp"
%>
