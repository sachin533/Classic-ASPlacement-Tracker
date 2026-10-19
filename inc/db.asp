<% ' ---------------------------------------------------------------------------
  '  Placement Tracker  -  inc\db.asp
  '  ---------------------------------------------------------------------------
  '  Connection settings plus the small data-access layer used by every page.
  '
  '  RULES OBSERVED HERE
  '    * Every command that carries user input uses ADODB.Command parameters.
  '      No SQL string is ever built by concatenating a form or query value.
  '    * The connection string is never written to the browser.
  '
  '  CONFIGURATION
  '    Edit PT_SERVER / PT_DATABASE below to point the site somewhere else.
  '    The automated smoke tests override them with the process environment
  '    variables PT_DB_SERVER / PT_DB_NAME so they can run against an isolated
  '    test database; when those variables are absent the defaults are used.
  '    ------------------------------------------------------------------------ %>

<%
' -- provider options -------------------------------------------------------
'   MSOLEDBSQL : Microsoft OLE DB Driver for SQL Server (preferred)
'   SQLOLEDB   : legacy built-in provider (fallback)
Const PT_PROVIDER_PRIMARY = "MSOLEDBSQL"

' -- defaults ---------------------------------------------------------------
Dim PT_SERVER, PT_DATABASE
PT_SERVER   = ".\SQLEXPRESS"
PT_DATABASE = "PlacementTracker"

' -- ADO type constants (avoids a reference to the ADO type library) --------
Const adInteger     = 3      ' adInteger
Const adVarWChar    = 202    ' adVarWChar
Const adDBTimeStamp = 135    ' adDBTimeStamp
Const adOpenStatic  = 3      ' adOpenStatic
Const adLockOptimistic = 1   ' adLockOptimistic
Const adUseClient   = 3      ' adUseClient
' (adExecuteNoRecords is 4; it was declared as 12 here, which ADO reads as
'  adCmdUnknown + adCmdStoredProc and rejects with error 3001.  The calls
'  below therefore run a plain Execute and ignore the empty recordset.)

' ---------------------------------------------------------------------------
'  Optional process-level overrides (used by tests\smoke.ps1).
' ---------------------------------------------------------------------------
Function PtEnvVar(name)
    PtEnvVar = ""
    On Error Resume Next
    Dim sh
    Set sh = Server.CreateObject("WScript.Shell")
    If Err.Number = 0 Then
        Dim v
        v = sh.Environment("PROCESS")(name) & ""
        Err.Clear
        If v <> "" Then PtEnvVar = v
        Set sh = Nothing
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Function

Dim ptEnvServer, ptEnvDatabase
ptEnvServer   = Trim(PtEnvVar("PT_DB_SERVER"))
ptEnvDatabase = Trim(PtEnvVar("PT_DB_NAME"))
If ptEnvServer <> ""   Then PT_SERVER   = ptEnvServer
If ptEnvDatabase <> "" Then PT_DATABASE = ptEnvDatabase

' ---------------------------------------------------------------------------
'  Connection string builder.  Windows integrated security, no credentials
'  stored on disk and nothing that is ever echoed back to the browser.
' ---------------------------------------------------------------------------
Function BuildConnString(provider, server, database)
    Select Case LCase(Trim(provider))
        Case "sqloledb"
            BuildConnString = "Provider=SQLOLEDB;Data Source=" & server & _
                              ";Initial Catalog=" & database & _
                              ";Integrated Security=SSPI;"
        Case Else
            BuildConnString = "Provider=MSOLEDBSQL;Server=" & server & _
                              ";Database=" & database & _
                              ";Integrated Security=SSPI;"
    End Select
End Function

' ---------------------------------------------------------------------------
'  OpenDb() - opens a connection, falling back through the provider list.
' ---------------------------------------------------------------------------
Function OpenDb()
    Dim cn, tried, i, provider, cs
    Set cn = Server.CreateObject("ADODB.Connection")
    cn.CursorLocation = adUseClient
    cn.CommandTimeout = 15
    cn.ConnectionTimeout = 15

    tried = "|" & LCase(PT_PROVIDER_PRIMARY) & "|"

    On Error Resume Next
    Err.Clear
    cn.Open BuildConnString(PT_PROVIDER_PRIMARY, PT_SERVER, PT_DATABASE)
    If Err.Number = 0 Then
        On Error GoTo 0
        Set OpenDb = cn
        Exit Function
    End If
    Err.Clear

    Dim fallbacks
    fallbacks = Array("MSOLEDBSQL", "SQLOLEDB")
    For i = LBound(fallbacks) To UBound(fallbacks)
        provider = LCase(fallbacks(i))
        If InStr(1, tried, "|" & provider & "|", vbBinaryCompare) = 0 Then
            cs = BuildConnString(provider, PT_SERVER, PT_DATABASE)
            cn.Open cs
            If Err.Number = 0 Then
                On Error GoTo 0
                Set OpenDb = cn
                Exit Function
            End If
            tried = tried & "|" & provider & "|"
            Err.Clear
        End If
    Next
    On Error GoTo 0

    Set cn = Nothing
    DbFail
End Function

' ---------------------------------------------------------------------------
'  Generic, non-leaking database failure.  Details go to the IIS log through
'  Response.AppendToLog; the browser only ever sees a friendly message.
' ---------------------------------------------------------------------------
Sub DbFail()
    Dim msg
    msg = Request.ServerVariables("URL") & " DBERROR"
    On Error Resume Next
    Response.AppendToLog " " & msg
    On Error GoTo 0
    FailPage 500, "Something went wrong", _
             "We could not complete that request. Please try again in a moment."
End Sub

' ---------------------------------------------------------------------------
'  String helpers shared by the data layer.
' ---------------------------------------------------------------------------
' ---------------------------------------------------------------------------
'  Escapes the LIKE wildcards a user can type so that they are matched
'  literally instead of acting as wildcards.  The argument is ByVal: an
'  earlier ByRef version rewrote the caller's search term on every call, so
'  the five LIKE parameters drifted apart (and eventually overflowed the
'  parameter size) whenever the query contained %, _ or [.
' ---------------------------------------------------------------------------
Function EscapeLike(ByVal s)
    s = Replace(s, "[", "[[]")
    s = Replace(s, "%", "[%]")
    s = Replace(s, "_", "[_]")
    EscapeLike = s
End Function

Sub CloseCn(cn)
    On Error Resume Next
    If Not cn Is Nothing Then
        If cn.State <> 0 Then cn.Close
        Set cn = Nothing
    End If
    On Error GoTo 0
End Sub

Sub CloseRs(rs)
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
        Set rs = Nothing
    End If
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
'  Appends the six application parameters to a command, in the exact order
'  the placeholders appear in the SQL text.
'
'  IMPORTANT: every statement in this file uses positional "?" placeholders
'  rather than named "@param" ones.  ADO only translates named parameters
'  for some providers - with the Microsoft OLE DB Driver (MSOLEDBSQL) a
'  named parameter is passed through verbatim and SQL Server rejects the
'  statement with "Must declare the scalar variable @id".  Positional
'  placeholders are bound in the order the parameters were appended, so the
'  call sites below must always create them in statement order.
'
'  v is a Scripting.Dictionary with the normalized keys:
'    company, role, status, applied_on_iso ("YYYY-MM-DD" or ""), job_url, notes
' ---------------------------------------------------------------------------
Sub AppendApplicationParameters(cmd, v)
    cmd.Parameters.Append cmd.CreateParameter("company",    adVarWChar, 1, 100,   v("company"))
    cmd.Parameters.Append cmd.CreateParameter("role",       adVarWChar, 1, 100,   v("role"))
    cmd.Parameters.Append cmd.CreateParameter("status",     adVarWChar, 1, 20,    v("status"))

    Dim pDate
    Set pDate = cmd.CreateParameter("applied_on", adDBTimeStamp, 1, 8)
    ' Leaving a parameter's Value unassigned sends an EMPTY value, not NULL,
    ' and '' is not a valid DATE - the value is set to Null explicitly.
    If v("applied_on_iso") <> "" Then
        pDate.Value = CDate(v("applied_on_iso"))
    Else
        pDate.Value = Null
    End If
    cmd.Parameters.Append pDate

    Dim pUrl
    Set pUrl = cmd.CreateParameter("job_url", adVarWChar, 1, 500)
    ' job_url and notes are nullable columns guarded by CHECK constraints, so
    ' an omitted value must arrive as NULL rather than an empty string.
    If v("job_url") <> "" Then
        pUrl.Value = v("job_url")
    Else
        pUrl.Value = Null
    End If
    cmd.Parameters.Append pUrl

    Dim pNotes
    Set pNotes = cmd.CreateParameter("notes", adVarWChar, 1, 1000)
    If v("notes") <> "" Then
        pNotes.Value = v("notes")
    Else
        pNotes.Value = Null
    End If
    cmd.Parameters.Append pNotes
End Sub

' ===========================================================================
'  SEARCH + FILTER + PAGINATION
'
'  Everything the dashboard needs to show one page of results:
'
'    * ApplicationWhere()      - the WHERE fragment, built from fixed literals
'    * AppendFilterParameters  - binds the user's values to the "?" holes
'    * FetchApplicationsPage   - COUNT(*) + one page, on one connection
'
'  RULES OBSERVED HERE
'    * No request value is ever concatenated into SQL.  The WHERE text only
'      ever contains hand-written literals; company / role / status / url /
'      notes / paging values all travel as positional "?" parameters.
'    * The count query and the page query use the SAME ApplicationWhere()
'      fragment, so "of 132" always describes the complete filtered set and
'      never just the rows on screen.
'    * LIKE metacharacters typed by the user (% _ [) are escaped by
'      EscapeLike() so they are matched literally instead of acting as
'      wildcards.
' ===========================================================================

' Rows per dashboard page (spec: 20).
Const PT_PAGE_SIZE = 20

' ---------------------------------------------------------------------------
'  ApplicationWhere(term, statusFilter) -> " WHERE ..." (or "" when there is
'  no filter at all).  Pure string assembly of constants - never user input.
' ---------------------------------------------------------------------------
Function ApplicationWhere(term, statusFilter)
    Dim sql
    sql = ""

    If term <> "" Then
        sql = "(company LIKE ? OR role LIKE ? OR status LIKE ? " & _
              "OR job_url LIKE ? OR notes LIKE ?)"
    End If

    If statusFilter <> "" Then
        If sql <> "" Then sql = sql & " AND "
        sql = sql & "(status = ?)"
    End If

    If sql <> "" Then sql = " WHERE " & sql
    ApplicationWhere = sql
End Function

' ---------------------------------------------------------------------------
'  Appends the parameters for ApplicationWhere() in exactly the order the
'  placeholders appear: the five search LIKEs first, then the status.
' ---------------------------------------------------------------------------
Sub AppendFilterParameters(cmd, term, statusFilter)
    Dim i, p, pattern

    If term <> "" Then
        ' Escape once and reuse: all five LIKEs must see the same pattern.
        pattern = "%" & EscapeLike(term) & "%"
        For i = 1 To 5
            Set p = cmd.CreateParameter("like" & CStr(i), adVarWChar, 1, 600, pattern)
            cmd.Parameters.Append p
            Set p = Nothing
        Next
    End If

    If statusFilter <> "" Then
        Set p = cmd.CreateParameter("fstatus", adVarWChar, 1, 20, statusFilter)
        cmd.Parameters.Append p
        Set p = Nothing
    End If
End Sub

' ---------------------------------------------------------------------------
'  FetchApplicationsPage(term, statusFilter, requestedPage, pageSize)
'
'  Returns a Scripting.Dictionary with:
'    rows    - Dictionary of zero-based row Dictionaries (newest first)
'    total   - every row matching the filters, not just this page
'    page    - the page actually shown, after clamping
'    pages   - total number of pages (at least 1)
'    offset  - rows skipped before this page
'    firstNo - 1-based number of the first row on this page
'    lastNo  - 1-based number of the last row on this page
'
'  Paging uses OFFSET ... FETCH NEXT, which SQL Server 2012+ supports and
'  which, unlike a cursor, lets the engine stop reading after one page.  The
'  ORDER BY carries the unique id as a tie-breaker so the sequence of pages
'  is stable between requests.
' ---------------------------------------------------------------------------
Function FetchApplicationsPage(term, statusFilter, requestedPage, pageSize)
    Dim cn, cmd, rs, sql, list, row, rowIndex, total, pages, page, offsetRows
    Dim firstNo, lastNo, result

    Set result = Server.CreateObject("Scripting.Dictionary")
    result.CompareMode = 1
    Set FetchApplicationsPage = result

    pageSize = CLng(pageSize + 0)
    If pageSize < 1 Then pageSize = PT_PAGE_SIZE
    If pageSize > 200 Then pageSize = 200

    Set list = Server.CreateObject("Scripting.Dictionary")
    list.CompareMode = 1

    Set cn = OpenDb()

    ' --- 1. how many rows match? (identical WHERE to the page query) -------
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1          ' adCmdText
    cmd.CommandText = "SELECT COUNT(*) FROM dbo.Applications" & _
                      ApplicationWhere(term, statusFilter)
    AppendFilterParameters cmd, term, statusFilter

    total = 0
    On Error Resume Next
    Set rs = cmd.Execute
    If Err.Number <> 0 Then
        Err.Clear
        CloseCn cn
        On Error GoTo 0
        DbFail
    End If
    On Error GoTo 0

    If Not rs.EOF Then
        If Not IsNull(rs(0).Value) Then total = CLng(rs(0).Value)
    End If
    CloseRs rs
    Set rs = Nothing

    ' --- 2. clamp the requested page so ?page=99999 can never 500 or 404 ---
    pages = Int((total + pageSize - 1) / pageSize)
    If pages < 1 Then pages = 1

    page = CLng(requestedPage + 0)
    If page < 1 Then page = 1
    If page > pages Then page = pages

    offsetRows = (page - 1) * pageSize
    If offsetRows < 0 Then offsetRows = 0

    ' --- 3. one page of rows ------------------------------------------------
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = "SELECT id, company, role, status, applied_on, job_url, " & _
                      "notes, created_at FROM dbo.Applications" & _
                      ApplicationWhere(term, statusFilter) & _
                      " ORDER BY id DESC " & _
                      "OFFSET ? ROWS FETCH NEXT ? ROWS ONLY"
    AppendFilterParameters cmd, term, statusFilter
    ' Paging values are bound too: offset and page size are integers produced
    ' by the arithmetic above, never raw request text.
    cmd.Parameters.Append cmd.CreateParameter("pt_offset", adInteger, 1, 0, CLng(offsetRows))
    cmd.Parameters.Append cmd.CreateParameter("pt_size",   adInteger, 1, 0, CLng(pageSize))

    On Error Resume Next
    Set rs = cmd.Execute
    If Err.Number <> 0 Then
        Err.Clear
        CloseCn cn
        On Error GoTo 0
        DbFail
    End If
    On Error GoTo 0

    rowIndex = 0
    Do While Not rs.EOF
        Set row = Server.CreateObject("Scripting.Dictionary")
        row.CompareMode = 1

        row("id")         = CLng(rs("id").Value)
        row("company")    = rs("company").Value & ""
        row("role")       = rs("role").Value & ""
        row("status")     = rs("status").Value & ""
        row("job_url")    = rs("job_url").Value & ""
        row("notes")      = rs("notes").Value & ""
        row("applied_on") = EmptyDash(rs("applied_on").Value)

        list.Add CStr(rowIndex), row
        rowIndex = rowIndex + 1
        rs.MoveNext
    Loop

    CloseRs rs
    CloseCn cn

    ' --- 4. package ---------------------------------------------------------
    firstNo = offsetRows + 1
    lastNo  = offsetRows + rowIndex
    If rowIndex = 0 Then
        firstNo = 0
        lastNo  = 0
    End If

    result.Add "rows",    list
    result.Add "total",   total
    result.Add "page",    page
    result.Add "pages",   pages
    result.Add "offset",  offsetRows
    result.Add "firstNo", firstNo
    result.Add "lastNo",  lastNo
End Function

' ---------------------------------------------------------------------------
'  GetApplication(id)
'  Returns a Scripting.Dictionary with the record, or Nothing when the id
'  does not exist.
' ---------------------------------------------------------------------------
Function GetApplication(id)
    Dim cn, cmd, rs, d
    Set GetApplication = Nothing

    Set cn = OpenDb()
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = "SELECT id, company, role, status, applied_on, job_url, notes, created_at " & _
                      "FROM dbo.Applications WHERE id = ?"
    cmd.Parameters.Append cmd.CreateParameter("id", adInteger, 1, 0, CLng(id))

    On Error Resume Next
    Set rs = cmd.Execute
    If Err.Number <> 0 Then
        Err.Clear
        CloseCn cn
        On Error GoTo 0
        DbFail
    End If
    On Error GoTo 0

    If rs.EOF Then
        CloseRs rs
        CloseCn cn
        Exit Function
    End If

    Set d = Server.CreateObject("Scripting.Dictionary")
    d.CompareMode = 1             ' text compare for keys

    d("id")       = CLng(rs("id").Value)
    d("company")  = rs("company").Value & ""
    d("role")     = rs("role").Value & ""
    d("status")   = rs("status").Value & ""
    d("job_url")  = rs("job_url").Value & ""
    d("notes")    = rs("notes").Value & ""

    If IsNull(rs("applied_on").Value) Then
        d("applied_on_iso") = ""
        d("applied_on")     = Null
    Else
        d("applied_on_iso") = FormatIsoDate(rs("applied_on").Value)
        d("applied_on")     = CDate(rs("applied_on").Value)
    End If

    d("created_at")     = rs("created_at").Value
    d("created_at_text") = FormatUtcTimestamp(rs("created_at").Value)

    CloseRs rs
    CloseCn cn
    Set GetApplication = d
End Function

' ---------------------------------------------------------------------------
'  InsertApplication(v) - returns the new id, or 0 when nothing was written.
' ---------------------------------------------------------------------------
Function InsertApplication(v)
    InsertApplication = 0

    Dim cn, cmd
    Set cn = OpenDb()
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = "INSERT INTO dbo.Applications " & _
                      "(company, role, status, applied_on, job_url, notes) " & _
                      "OUTPUT INSERTED.id " & _
                      "VALUES (?, ?, ?, ?, ?, ?)"
    AppendApplicationParameters cmd, v

    ' OUTPUT INSERTED.id returns the new key as a resultset from the INSERT
    ' itself.  SCOPE_IDENTITY() is deliberately NOT used: ADO sends a
    ' parameterised command through sp_prepexec/sp_execute, which is a nested
    ' scope, so a following SCOPE_IDENTITY() call reads NULL.  OUTPUT has no
    ' such problem - it is the same statement, and unlike @@IDENTITY it is
    ' unaffected by concurrent inserts from other sessions.
    Dim rs
    On Error Resume Next
    Set rs = cmd.Execute
    If Err.Number <> 0 Then
        Err.Clear
        CloseCn cn
        On Error GoTo 0
        DbFail
    End If
    On Error GoTo 0

    If Not rs.EOF Then
        If Not IsNull(rs(0).Value) Then InsertApplication = CLng(rs(0).Value)
    End If

    CloseRs rs
    CloseCn cn
End Function

' ---------------------------------------------------------------------------
'  UpdateApplication(id, v) - returns the number of rows changed.
' ---------------------------------------------------------------------------
Function UpdateApplication(id, v)
    UpdateApplication = 0

    Dim cn, cmd
    Set cn = OpenDb()
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = "UPDATE dbo.Applications SET company = ?, " & _
                      "role = ?, status = ?, applied_on = ?, " & _
                      "job_url = ?, notes = ? WHERE id = ?"
    AppendApplicationParameters cmd, v
    cmd.Parameters.Append cmd.CreateParameter("id", adInteger, 1, 0, CLng(id))

    On Error Resume Next
    ' ADODB.Command.Execute(RecordsAffected, Parameters, Options) - the row
    ' count is returned through the FIRST (ByRef/out) argument of Execute.
    ' It is not a property of either ADODB.Command or ADODB.Connection:
    ' reading cn.RecordsAffected / cmd.RecordsAffected raises ADO 3001 or
    ' VBScript 800a01b6 and turned every successful UPDATE into a 500.
    Dim changed
    changed = 0
    cmd.Execute changed
    If Err.Number <> 0 Then
        Err.Clear
        CloseCn cn
        On Error GoTo 0
        DbFail
    End If
    On Error GoTo 0

    UpdateApplication = CLng(changed + 0)
    CloseCn cn
End Function

' ---------------------------------------------------------------------------
'  DeleteApplication(id) - returns the number of rows removed.
' ---------------------------------------------------------------------------
Function DeleteApplication(id)
    DeleteApplication = 0

    Dim cn, cmd
    Set cn = OpenDb()
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = "DELETE FROM dbo.Applications WHERE id = ?"
    cmd.Parameters.Append cmd.CreateParameter("id", adInteger, 1, 0, CLng(id))

    On Error Resume Next
    Dim removed
    removed = 0
    ' Row count comes back through Execute's ByRef RecordsAffected argument.
    cmd.Execute removed
    If Err.Number <> 0 Then
        Err.Clear
        CloseCn cn
        On Error GoTo 0
        DbFail
    End If
    On Error GoTo 0

    DeleteApplication = CLng(removed + 0)
    CloseCn cn
End Function
%>



