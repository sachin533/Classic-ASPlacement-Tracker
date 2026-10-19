<!--#include virtual="/inc/helpers.asp" -->
<%
' ---------------------------------------------------------------------------
'  404.asp - the site's error page.
'
'  Reached in three ways:
'    1. A browser asks for a URL the site does not serve.  The <rewrite>
'       rules in web.config rewrite it to /404.asp as an ordinary request,
'       which is why this page - and not IIS - decides the status.
'    2. A page calls FailPage() in inc\helpers.asp, which renders the same
'       markup through RenderErrorPage() with the code, title and message
'       passed as arguments.  Server.Execute cannot be used for this: it
'       cannot take a query string (ASP 0173) and it opens a new page scope,
'       so the caller's variables would be invisible.
'    3. A direct request, optionally carrying ptcode / pttitle / ptmsg (used
'       by the smoke tests).  Those values are validated before use.
'
'  Only the status is decided here; the wording defaults to the generic
'  "page not found" text unless the request supplied its own.
' ---------------------------------------------------------------------------

Dim rawCode, rawTitle, rawMsg
rawCode  = Trim(Request.QueryString("ptcode") & "")
rawTitle = Trim(Request.QueryString("pttitle") & "")
rawMsg   = Trim(Request.QueryString("ptmsg") & "")

' ptcode is validated against a strict numeric pattern before it is used.
If TestPattern("^[0-9]{3,4}$", rawCode) Then
    gErrCode = CLng(rawCode)
Else
    gErrCode = 404
End If

Select Case gErrCode
    Case 400 : Response.Status = "400 Bad Request"
    Case 403 : Response.Status = "403 Forbidden"
    Case 405 : Response.Status = "405 Method Not Allowed"
    Case 500 : Response.Status = "500 Internal Server Error"
    Case Else: Response.Status = "404 Not Found"
End Select

If rawTitle = "" Then
    Select Case gErrCode
        Case 403 : rawTitle = "Request blocked"
        Case 405 : rawTitle = "Method not allowed"
        Case 500 : rawTitle = "Something went wrong"
        Case Else: rawTitle = "Page not found"
    End Select
End If

If rawMsg = "" Then
    rawMsg = "The page you are looking for does not exist, the link may " & _
             "be broken, or the application may have been deleted."
End If

RenderErrorPage gErrCode, rawTitle, rawMsg
%>
