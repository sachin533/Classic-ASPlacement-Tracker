<!--#include virtual="/inc/helpers.asp" -->
<!--#include virtual="/inc/db.asp" -->
<%
' ---------------------------------------------------------------------------
'  Dashboard - default.asp
'
'  Query string
'    ?q=...       search term matched against company, role, status,
'                 job_url and notes (server-side, partial, escaped LIKE)
'    ?status=...  one of the five statuses; unknown values are ignored
'    ?page=...    1-based page number.  Anything that is not a clean positive
'                 integer falls back to 1, and a page past the end is clamped
'                 to the last page, so a hand-edited URL can never 500.
'
'  A new search or a changed filter is always submitted WITHOUT a page value,
'  which is what resets the result set to page 1.
' ---------------------------------------------------------------------------

' Every page except login.asp / logout.asp / forgot.asp requires a session:
' an anonymous visitor is sent to /login.asp with a 302 before any query
' parameter is read (see inc\auth.asp).
EnsureAuthenticated

Dim searchRaw, statusRaw, statusFilter, requestedPage
Dim result, apps, total, pages, page, firstNo, lastNo
Dim rowCount, emptyTitle, emptyText, hasFilters, summaryHtml

searchRaw = Trim(Request.QueryString("q") & "")
If Len(searchRaw) > 100 Then searchRaw = Left(searchRaw, 100)

statusRaw    = Trim(Request.QueryString("status") & "")
statusFilter = CanonicalStatus(statusRaw)   ' "" when blank or not one of the five

requestedPage = ParsePageNumber(Request.QueryString("page"))

gTitle        = "Applications"
gActiveNav    = "home"
gShowSearch   = True
gSearchValue  = searchRaw
gFilterStatus = statusFilter

Set result = FetchApplicationsPage(searchRaw, statusFilter, requestedPage, PT_PAGE_SIZE)
Set apps   = result("rows")
total      = result("total")
pages      = result("pages")
page       = result("page")
firstNo    = result("firstNo")
lastNo     = result("lastNo")
rowCount   = apps.Count
hasFilters = (searchRaw <> "" Or statusFilter <> "")

' --- results summary (the only place user text is written into markup) -----
If total = 0 Then
    summaryHtml = "No applications match"
    If searchRaw <> "" Then summaryHtml = summaryHtml & " &quot;" & H(searchRaw) & "&quot;"
    If statusFilter <> "" Then summaryHtml = summaryHtml & " with status " & H(statusFilter)
    summaryHtml = summaryHtml & "."
Else
    summaryHtml = "Showing " & firstNo & "&ndash;" & lastNo & " of " & total & " application"
    If total <> 1 Then summaryHtml = summaryHtml & "s"
    If searchRaw <> "" Then summaryHtml = summaryHtml & " matching &quot;" & H(searchRaw) & "&quot;"
    If statusFilter <> "" Then summaryHtml = summaryHtml & " with status " & H(statusFilter)
End If

If rowCount = 0 Then
    If hasFilters Then
        emptyTitle = "No matching applications"
        emptyText  = "Try a different company, role, status or note - or clear the filters to see everything."
    Else
        emptyTitle = "No applications yet"
        emptyText  = "Start by adding your first job or internship application."
    End If
End If

' ---------------------------------------------------------------------------
'  Link builder - every pager link carries the current query and status so a
'  search never silently disappears when the user moves to another page.
' ---------------------------------------------------------------------------
Function DashHref(pageNo)
    Dim s
    s = "/default.asp?page=" & CLng(pageNo)
    If searchRaw <> "" Then s = s & "&q=" & Server.URLEncode(searchRaw)
    If statusFilter <> "" Then s = s & "&status=" & Server.URLEncode(statusFilter)
    DashHref = s
End Function

' ---------------------------------------------------------------------------
'  Page-number window: 1 .. totalPages for short runs, otherwise the first
'  page, the last page and the two pages either side of the current one,
'  joined with an ellipsis wherever there is a gap of more than one.
' ---------------------------------------------------------------------------
Sub AddPageToken(bag, n)
    If n < 1 Then Exit Sub
    If bag.Exists(CStr(n)) Then Exit Sub
    bag.Add CStr(n), n
End Sub

Function PagerTokens(cur, totalPages)
    Dim bag, arr, i, j, tmp, out, prev

    Set bag = Server.CreateObject("Scripting.Dictionary")
    bag.CompareMode = 1
    AddPageToken bag, 1
    AddPageToken bag, totalPages
    ' The two pages either side of the current one - but never past the end,
    ' otherwise the last page would offer links to pages that do not exist.
    For i = cur - 2 To cur + 2
        If i <= totalPages Then AddPageToken bag, i
    Next

    arr = bag.Items

    ' insertion sort - the list is never longer than seven numbers
    For i = 1 To UBound(arr)
        tmp = arr(i)
        j = i - 1
        Do While j >= 0
            If arr(j) > tmp Then
                arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        arr(j + 1) = tmp
    Next

    out  = ""
    prev = 0
    For i = 0 To UBound(arr)
        If prev <> 0 Then
            If arr(i) - prev > 1 Then out = out & ",..."
        End If
        out = out & "," & CStr(arr(i))
        prev = arr(i)
    Next

    PagerTokens = Mid(out, 2)
End Function

' ---------------------------------------------------------------------------
'  The whole Previous / 1 2 3 ... N / Next control, or "" for zero or one page.
' ---------------------------------------------------------------------------
Function RenderPager(cur, totalPages)
    Dim tokens, i, p, out

    RenderPager = ""
    If totalPages < 2 Then Exit Function

    tokens = Split(PagerTokens(cur, totalPages), ",")
    out = "<nav class=""pager"" aria-label=""Application pages"">"

    If cur <= 1 Then
        out = out & "<span class=""pager-btn is-disabled"" aria-disabled=""true"">Previous</span>"
    Else
        out = out & "<a class=""pager-btn"" rel=""prev"" href=""" & H(DashHref(cur - 1)) & _
              """>Previous</a>"
    End If

    For i = 0 To UBound(tokens)
        If tokens(i) = "..." Then
            out = out & "<span class=""pager-ellipsis"" aria-hidden=""true"">&hellip;</span>"
        Else
            p = CLng(tokens(i))
            If p = cur Then
                out = out & "<span class=""pager-btn is-active"" aria-current=""page"">" & _
                      CStr(p) & "</span>"
            Else
                out = out & "<a class=""pager-btn"" href=""" & H(DashHref(p)) & _
                      """ aria-label=""Go to page " & CStr(p) & """>" & CStr(p) & "</a>"
            End If
        End If
    Next

    If cur >= totalPages Then
        out = out & "<span class=""pager-btn is-disabled"" aria-disabled=""true"">Next</span>"
    Else
        out = out & "<a class=""pager-btn"" rel=""next"" href=""" & H(DashHref(cur + 1)) & _
              """>Next</a>"
    End If

    RenderPager = out & "</nav>"
End Function

' ---------------------------------------------------------------------------
'  Status <option> list for the filter select.
' ---------------------------------------------------------------------------
Function StatusFilterOptions(selected)
    Dim allowed, i, name, out
    allowed = Array("Wishlist", "Applied", "Interview", "Offer", "Rejected")

    out = "<option value="""""
    If Trim(selected & "") = "" Then out = out & " selected=""selected"""
    out = out & ">All statuses</option>" & vbCrLf

    For i = 0 To UBound(allowed)
        name = allowed(i)
        out = out & "<option value=""" & H(name) & """"
        If StrComp(name, Trim(selected & ""), vbTextCompare) = 0 Then _
            out = out & " selected=""selected"""
        out = out & ">" & H(name) & "</option>" & vbCrLf
    Next

    StatusFilterOptions = out
End Function
%>
<% Call RenderHeader() %>

<div class="page-head">
  <div>
    <h1>Applications</h1>
    <p class="subtitle">Track your job and internship applications in one place.</p>
  </div>
  <a class="btn btn-primary" href="/new.asp"><%= IconPlus() %>Add Application</a>
</div>

<% ' --- prominent search + status filter ---------------------------------- %>
<div class="card search-card">
  <form class="search-panel" role="search" method="get" action="/default.asp">
    <div class="search-row">
      <div class="search-field">
        <span class="search-icon" aria-hidden="true"><%= IconSearch() %></span>
        <label class="visually-hidden" for="dash-search">Search company, role, status, job URL or notes</label>
        <input type="search" id="dash-search" name="q" maxlength="100"
               class="search-input js-search-input"
               placeholder="Search company, role, status or notes..."
               value="<%= H(searchRaw) %>" autocomplete="off"
               spellcheck="false">
        <button type="button" class="search-clear js-search-clear"
                aria-label="Clear the search box"<% If Trim(searchRaw) = "" Then %> hidden<% End If %>>&times;</button>
      </div>

      <div class="filter-field">
        <label class="visually-hidden" for="dash-status">Filter by status</label>
        <select id="dash-status" name="status" class="filter-select">
          <%= StatusFilterOptions(statusFilter) %>
        </select>
      </div>

      <button type="submit" class="btn btn-primary search-submit">
        <%= IconSearch() %>Search
      </button>

      <% If hasFilters Then %>
      <a class="btn btn-secondary" href="/default.asp">Clear filters</a>
      <% End If %>
    </div>
  </form>
</div>

<% ' --- results summary + active filters ---------------------------------- %>
<div class="results-bar">
  <p class="results-summary" role="status"><%= summaryHtml %></p>
  <% If hasFilters Then %>
  <div class="active-filters">
    <% If searchRaw <> "" Then %>
    <span class="chip">Search: <strong>&quot;<%= H(searchRaw) %>&quot;</strong></span>
    <% End If %>
    <% If statusFilter <> "" Then %>
    <span class="chip">Status: <strong><%= H(statusFilter) %></strong></span>
    <% End If %>
    <a class="chip-clear" href="/default.asp">Clear all</a>
  </div>
  <% End If %>
</div>

<% If rowCount = 0 Then %>

<div class="card">
  <div class="empty-state">
    <span class="empty-icon"><%= IconBriefcase() %></span>
    <p class="empty-title"><%= H(emptyTitle) %></p>
    <p class="empty-text"><%= H(emptyText) %></p>
    <% If hasFilters Then %>
    <a class="btn btn-secondary" href="/default.asp">Clear search</a>
    <% End If %>
    <a class="btn btn-primary" href="/new.asp"><%= IconPlus() %>Add Application</a>
  </div>
</div>

<% Else %>

<div class="card">
  <div class="table-wrap">
    <table class="app-table">
      <caption class="visually-hidden">Applications<% If hasFilters Then %> matching &quot;<%= H(searchRaw) %>&quot;<% If statusFilter <> "" Then %> with status <%= H(statusFilter) %><% End If %><% End If %>, page <%= page %> of <%= pages %></caption>
      <thead>
        <tr>
          <th scope="col" class="col-num">#</th>
          <th scope="col" class="col-company">Company</th>
          <th scope="col" class="col-role">Role</th>
          <th scope="col" class="col-status">Status</th>
          <th scope="col" class="col-date">Applied On</th>
          <th scope="col" class="col-actions">Actions</th>
        </tr>
      </thead>
      <tbody>
        <%
        Dim i, row, rowId
        For i = 0 To rowCount - 1
            Set row = apps.Item(CStr(i))
            rowId = row("id")
        %>
        <tr>
          <td class="col-num"><%= firstNo + i %></td>
          <td class="col-company"><a href="/view.asp?id=<%= rowId %>"><%= H(row("company")) %></a></td>
          <td class="col-role"><%= H(row("role")) %></td>
          <td class="col-status"><%= StatusBadge(row("status")) %></td>
          <td class="col-date"><%= H(row("applied_on")) %></td>
          <td class="col-actions">
            <span class="row-actions">
              <a href="/view.asp?id=<%= rowId %>">View</a>
              <span class="sep" aria-hidden="true">|</span>
              <a href="/edit.asp?id=<%= rowId %>">Edit</a>
              <span class="sep" aria-hidden="true">|</span>
              <form id="del-form-<%= rowId %>" method="post" action="/delete.asp"
                    class="inline-delete js-delete-form"
                    onsubmit="return confirm('Delete this application? This cannot be undone.');">
                <input type="hidden" name="id" value="<%= rowId %>">
                <%= CsrfInput() %>
                <button type="submit" class="link-btn is-danger js-delete-trigger"
                        form="del-form-<%= rowId %>">Delete</button>
              </form>
            </span>
          </td>
        </tr>
        <%
        Next
        Set row = Nothing
        Set apps = Nothing
        Set result = Nothing
        %>
      </tbody>
    </table>
  </div>
</div>

<div class="pager-wrap">
  <p class="pager-status">Page <strong><%= page %></strong> of <strong><%= pages %></strong></p>
  <%= RenderPager(page, pages) %>
</div>

<% End If %>
<% Call RenderFooter() %>
