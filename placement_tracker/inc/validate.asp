<%
' inc/validate.asp - shared server-side validation for the application form.
' The SQL CHECK constraints are the last line of defence; this layer gives
' the user friendly messages and keeps bad data out of the database.

' Accepts only a strict yyyy-mm-dd date that really exists (rejects 2024-02-31).
Function IsIsoDate(ByVal s)
  IsIsoDate = False
  If Len(s) <> 10 Then Exit Function
  If Mid(s, 5, 1) <> "-" Or Mid(s, 8, 1) <> "-" Then Exit Function
  If Left(s, 4) < "1900" Or Left(s, 4) > "2999" Then Exit Function
  Dim d
  On Error Resume Next
  d = CDate(s)
  If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0
  IsIsoDate = (Year(d) = CInt(Left(s, 4)) And Month(d) = CInt(Mid(s, 6, 2)) And Day(d) = CInt(Right(s, 2)))
End Function

Function IsAllowedStatus(ByVal s)
  Select Case s
    Case "Wishlist", "Applied", "Interview", "Offer", "Rejected"
      IsAllowedStatus = True
    Case Else
      IsAllowedStatus = False
  End Select
End Function

' Returns an empty Scripting.Dictionary when everything is valid,
' otherwise one entry per bad field: field name -> message.
Function ValidateApplication(ByVal vals)
  Dim errs
  Set errs = Server.CreateObject("Scripting.Dictionary")

  Dim company, role, status, appliedOn, jobUrl, notes
  company  = Trim(vals("company"))
  role     = Trim(vals("role"))
  status   = Trim(vals("status"))
  appliedOn = Trim(vals("applied_on"))
  jobUrl   = Trim(vals("job_url"))
  notes    = vals("notes")

  If Len(company) < 2 Or Len(company) > 100 Then
    errs("company") = "Company must be between 2 and 100 characters."
  End If

  If Len(role) < 2 Or Len(role) > 100 Then
    errs("role") = "Role must be between 2 and 100 characters."
  End If

  If Not IsAllowedStatus(status) Then
    errs("status") = "Please choose one of the allowed statuses."
  End If

  ' A date is required for every status except Wishlist.
  If appliedOn = "" Then
    If status <> "Wishlist" Then
      errs("applied_on") = "Applied date is required unless the status is Wishlist."
    End If
  ElseIf Not IsIsoDate(appliedOn) Then
    errs("applied_on") = "Use a real date in yyyy-mm-dd format."
  End If

  If jobUrl <> "" Then
    If Len(jobUrl) > 500 Then
      errs("job_url") = "Job URL must be 500 characters or fewer."
    ElseIf LCase(Left(jobUrl, 7)) <> "http://" And LCase(Left(jobUrl, 8)) <> "https://" Then
      errs("job_url") = "Job URL must start with http:// or https://."
    End If
  End If

  If Len(notes) > 1000 Then
    errs("notes") = "Notes must be 1,000 characters or fewer."
  End If

  Set ValidateApplication = errs
End Function
%>
