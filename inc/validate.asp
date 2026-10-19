<% ' ---------------------------------------------------------------------------
  '  Placement Tracker  -  inc\validate.asp
  '  ---------------------------------------------------------------------------
  '  Shared server-side validation for new.asp and edit.asp.
  '
  '  There is exactly one rule set: both pages call CollectApplicationValues
  '  to read Request.Form and ValidateApplication to check it, so a value that
  '  is accepted by one page is accepted by the other.
  '
  '  The CHECK constraints in db\schema.sql remain the last line of defence.
  '    ------------------------------------------------------------------------ %>

<%
' ---------------------------------------------------------------------------
'  CollectApplicationValues()
'  Reads the posted fields (or blanks them for a GET) into a Dictionary and
'  adds the normalized keys used by the data layer.
' ---------------------------------------------------------------------------
Function CollectApplicationValues()
    Dim d
    Set d = Server.CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    d("company")    = SafeFormValue("company",    400)
    d("role")       = SafeFormValue("role",       400)
    d("status")     = SafeFormValue("status",     60)
    d("applied_on") = SafeFormValue("applied_on", 40)
    d("job_url")    = SafeFormValue("job_url",    1200)
    d("notes")      = SafeFormValue("notes",      4000)

    ' Normalized keys - filled in by ValidateApplication.
    d("applied_on_iso") = ""
    d("status_canonical") = ""

    Set CollectApplicationValues = d
End Function

' Bounded read of a single posted field.
Function SafeFormValue(fieldName, maxLen)
    Dim v
    v = Request.Form(fieldName) & ""
    If Len(v) > maxLen Then v = Left(v, maxLen)
    SafeFormValue = v
End Function

' ---------------------------------------------------------------------------
'  ValidateApplication(v)
'  v        : Dictionary of raw values (mutated with the normalized values)
'  returns  : Dictionary of field -> message.  Empty means "valid".
' ---------------------------------------------------------------------------
Function ValidateApplication(v)
    Dim errs
    Set errs = Server.CreateObject("Scripting.Dictionary")
    errs.CompareMode = 1

    Dim company, role, status, appliedOn, jobUrl, notes
    company   = Trim(v("company") & "")
    role      = Trim(v("role") & "")
    status    = Trim(v("status") & "")
    appliedOn = Trim(v("applied_on") & "")
    jobUrl    = Trim(v("job_url") & "")
    notes     = v("notes") & ""

    ' ---------------- company -------------------------------------------
    If company = "" Then
        errs("company") = "Company is required."
    ElseIf Len(company) < 2 Then
        errs("company") = "Company must be at least 2 characters."
    ElseIf Len(company) > 100 Then
        errs("company") = "Company must be 100 characters or fewer."
    End If

    ' ---------------- role ----------------------------------------------
    If role = "" Then
        errs("role") = "Role is required."
    ElseIf Len(role) < 2 Then
        errs("role") = "Role must be at least 2 characters."
    ElseIf Len(role) > 100 Then
        errs("role") = "Role must be 100 characters or fewer."
    End If

    ' ---------------- status --------------------------------------------
    Dim canonical
    canonical = CanonicalStatus(status)
    If canonical = "" Then
        errs("status") = "Status must be one of: Wishlist, Applied, Interview, Offer or Rejected."
    End If

    ' ---------------- applied on ----------------------------------------
    Dim appliedDate, appliedIso
    appliedIso = ""
    If canonical = "Wishlist" And appliedOn = "" Then
        ' Wishlist is the only status that may be saved without a date.
    ElseIf appliedOn = "" Then
        errs("applied_on") = "Applied On is required unless the status is Wishlist."
    ElseIf Not IsValidIsoDate(appliedOn) Then
        errs("applied_on") = "Applied On must be a real date in YYYY-MM-DD format."
    Else
        appliedDate = CDate(appliedOn)
        appliedIso  = FormatIsoDate(appliedDate)
    End If

    ' ---------------- job url -------------------------------------------
    If jobUrl <> "" Then
        If Len(jobUrl) > 500 Then
            errs("job_url") = "Job URL must be 500 characters or fewer."
        ElseIf Not IsHttpUrl(jobUrl) Then
            errs("job_url") = "Job URL must be a full http:// or https:// address."
        End If
    End If

    ' ---------------- notes ---------------------------------------------
    If Len(notes) > 1000 Then
        errs("notes") = "Notes must be 1,000 characters or fewer."
    End If

    ' Write the normalized values back so the caller (and the form) use them.
    v("company")          = company
    v("role")             = role
    v("status")           = canonical
    v("status_canonical") = canonical
    v("applied_on")       = appliedIso
    v("applied_on_iso")   = appliedIso
    v("job_url")          = jobUrl
    v("notes")            = notes

    Set ValidateApplication = errs
End Function

' ---------------------------------------------------------------------------
'  Strict "YYYY-MM-DD" date check.  The round trip through FormatIsoDate
'  rejects values the platform would silently roll over (e.g. 2024-02-30).
' ---------------------------------------------------------------------------
Function IsValidIsoDate(s)
    IsValidIsoDate = False

    Dim text
    text = Trim(s & "")
    If text = "" Then Exit Function
    If Not TestPattern("^[0-9]{4}-[0-9]{2}-[0-9]{2}$", text) Then Exit Function

    Dim d, saved
    saved = 0
    On Error Resume Next
    d = CDate(text)
    saved = Err.Number
    Err.Clear
    On Error GoTo 0
    If saved <> 0 Then Exit Function
    If Not IsDate(d) Then Exit Function
    If FormatIsoDate(d) <> text Then Exit Function

    IsValidIsoDate = True
End Function

' ---------------------------------------------------------------------------
'  True when the Dictionary holds at least one error.
' ---------------------------------------------------------------------------
Function HasErrors(errs)
    HasErrors = (errs.Count > 0)
End Function

Function FieldError(errs, fieldName)
    FieldError = ""
    If errs.Exists(fieldName) Then FieldError = errs(fieldName)
End Function

' ---------------------------------------------------------------------------
'  Summary sentence used above the form when validation fails.
' ---------------------------------------------------------------------------
Function ValidationSummary(errs)
    ValidationSummary = ""
    If errs.Count = 0 Then Exit Function
    Dim msg
    If errs.Count = 1 Then
        msg = "Please fix the highlighted field below."
    Else
        msg = "Please fix the " & errs.Count & " highlighted fields below."
    End If
    ValidationSummary = msg
End Function
%>
