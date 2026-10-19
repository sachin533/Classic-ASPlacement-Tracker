<%
' inc/helpers.asp - shared helpers: output encoding, CSRF, flash, 404, badges.
' Include AFTER header.asp (which declares Option Explicit) and once per page.

' ---- Output encoding ------------------------------------------------------
' Every dynamic value printed into HTML MUST pass through H().
Function H(ByVal v)
  If IsNull(v) Or IsEmpty(v) Then v = ""
  H = Server.HTMLEncode(CStr(v))
End Function

' ---- CSRF (Cross-Site Request Forgery) token -------------------------------
' A random secret stored in the visitor's session. Every POST form carries it
' in a hidden field; VerifyCsrf() rejects POSTs that do not match (403).
Function NewCsrfToken()
  Dim tlib
  Set tlib = Server.CreateObject("Scriptlet.TypeLib")
  NewCsrfToken = Replace(Replace(tlib.GUID, "{", ""), "}", "")
  Set tlib = Nothing
End Function

Function CsrfToken()
  If IsEmpty(Session("csrf")) Or Session("csrf") = "" Then
    Session("csrf") = NewCsrfToken()
  End If
  CsrfToken = Session("csrf")
End Function

Sub VerifyCsrf()
  Dim got
  got = Request.Form("csrf")
  If got <> CsrfToken() Then
    Response.Clear
    Response.Status = "403 Forbidden"
    Response.StatusCode = 403
    Response.Write "<!DOCTYPE html><html lang=""en""><head><meta charset=""utf-8"">" _
      & "<title>403 Forbidden</title>" _
      & "<link rel=""stylesheet"" href=""/static/styles.css""></head><body>" _
      & "<main class=""container""><div class=""card center-card"">" _
      & "<h1>403 Forbidden</h1>" _
      & "<p>Your form token was missing or expired. Go back, reload the page, and try again.</p>" _
      & "<p><a class=""btn btn-primary"" href=""/default.asp"">Back to Applications</a></p>" _
      & "</div></main></body></html>"
    Response.End
  End If
End Sub

' ---- Flash message (survives one redirect via Session) ---------------------
Sub SetFlash(ByVal msg)
  Session("flash") = msg
End Sub

Function TakeFlash()
  If Not IsEmpty(Session("flash")) Then
    TakeFlash = Session("flash")
    Session("flash") = Empty
  End If
End Function

' ---- Status handling -------------------------------------------------------
Sub NotFound()
  Response.Clear
  Response.Status = "404 Not Found"
  Response.StatusCode = 404
  Server.Execute "/404.asp"
  Response.End
End Sub

Sub MethodNotAllowed()
  Response.Clear
  Response.Status = "405 Method Not Allowed"
  Response.StatusCode = 405
  Response.Write "<!DOCTYPE html><html lang=""en""><head><meta charset=""utf-8"">" _
    & "<title>405 Method Not Allowed</title>" _
    & "<link rel=""stylesheet"" href=""/static/styles.css""></head><body>" _
    & "<main class=""container""><div class=""card center-card"">" _
    & "<h1>405 Method Not Allowed</h1>" _
    & "<p>This page only accepts POST requests. Nothing was deleted.</p>" _
    & "<p><a class=""btn btn-primary"" href=""/default.asp"">Back to Applications</a></p>" _
    & "</div></main></body></html>"
  Response.End
End Sub

' CSS class for a status badge, matching the reference design.
Function StatusClass(ByVal s)
  Select Case s
    Case "Applied":    StatusClass = "badge badge-applied"
    Case "Interview":  StatusClass = "badge badge-interview"
    Case "Offer":      StatusClass = "badge badge-offer"
    Case "Rejected":   StatusClass = "badge badge-rejected"
    Case "Wishlist":   StatusClass = "badge badge-wishlist"
    Case Else:         StatusClass = "badge badge-wishlist"
  End Select
End Function

' ISO yyyy-mm-dd date display, independent of the server locale
' (FormatDateTime would print 10-08-2024 on an Indian system).Function FmtDate(ByVal d)
  If IsNull(d) Then
    FmtDate = "-"
    Exit Function
  End If
  FmtDate = Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2)
End Function

' yyyy-mm-dd hh:nn:ss display for created_at (stored as UTC by the database).
Function FmtStamp(ByVal d)
  FmtStamp = FmtDate(d) & " " & Right("0" & Hour(d), 2) & ":" & Right("0" & Minute(d), 2) _
           & ":" & Right("0" & Second(d), 2)
End Function

' True when the string is 1-9 plain digits - safe to feed to CLng.
' IsNumeric alone would also accept "1E5" or "-2", which are not ids.
Function RegExpIsDigits(ByVal s)
  Dim re
  Set re = New RegExp
  re.Pattern = "^\d{1,9}$"
  RegExpIsDigits = re.Test(s)
  Set re = Nothing
End Function
%>
