<!--#include file="inc/header.asp"-->
<!--#include file="inc/helpers.asp"-->
<!--#include file="inc/db.asp"-->
<%
' default.asp - dashboard: lists every application, optional ?q= search.
Dim q, cn, cmd, rs, rowNum, total
q = Trim(Request.QueryString("q"))

Set cn = OpenDb()
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = cn

If q <> "" Then
  ' Parameterized LIKE - the wildcard comes from the parameter, never from
  ' string concatenation of the SQL statement itself.
  cmd.CommandText = "SELECT id, company, role, status, applied_on FROM dbo.Applications " _
                  & "WHERE company LIKE ? OR role LIKE ? ORDER BY id"
  cmd.Parameters.Append cmd.CreateParameter("q", 200, 1, 302, "%" & q & "%")  ' adVarWChar, adParamInput
  cmd.Parameters.Append cmd.CreateParameter("q2", 200, 1, 302, "%" & q & "%")
Else
  cmd.CommandText = "SELECT id, company, role, status, applied_on FROM dbo.Applications ORDER BY id"
End If
Set rs = cmd.Execute

total = 0
If Not rs.EOF Then total = rs.RecordCount
%>

<div class="page-head">
  <div>
    <h1>Applications<% If total > 0 Then %> (<%= total %>)<% End If %></h1>
    <p class="subtitle">Track your job and internship applications in one place.</p>
  </div>
  <a class="btn btn-primary" href="/new.asp">
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" aria-hidden="true">
      <line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line>
    </svg>
    Add Application
  </a>
</div>

<% If rs.EOF Then %>

<div class="card empty-card">
  <span class="empty-icon" aria-hidden="true">
    <svg width="34" height="34" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7">
      <rect x="2" y="7" width="20" height="14" rx="2"></rect>
      <path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"></path>
      <line x1="2" y1="13" x2="22" y2="13"></line>
    </svg>
  </span>
  <% If q <> "" Then %>
  <h2>No matches found</h2>
  <p>No applications match &ldquo;<%= H(q) %>&rdquo;. Try a different company or role.</p>
  <a class="btn btn-secondary" href="/default.asp">Clear search</a>
  <% Else %>
  <h2>No applications yet</h2>
  <p>Start by adding your first job or internship application.</p>
  <a class="btn btn-primary" href="/new.asp">
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" aria-hidden="true">
      <line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line>
    </svg>
    Add Application
  </a>
  <% End If %>
</div>

<% Else %>

<div class="card table-card">
  <div class="table-wrap">
    <table class="applications">
      <thead>
        <tr>
          <th scope="col">#</th>
          <th scope="col">Company</th>
          <th scope="col">Role</th>
          <th scope="col">Status</th>
          <th scope="col">Applied On</th>
          <th scope="col">Actions</th>
        </tr>
      </thead>
      <tbody>
<%
  rowNum = 0
  showDeleteModal = True
  While Not rs.EOF
    rowNum = rowNum + 1
%>
        <tr>
          <td class="col-num"><%= rowNum %></td>
          <td class="cell-strong"><a href="/view.asp?id=<%= H(rs("id")) %>"><%= H(rs("company")) %></a></td>
          <td><%= H(rs("role")) %></td>
          <td><span class="<%= StatusClass(rs("status")) %>"><%= H(rs("status")) %></span></td>
          <td>
            <% If IsNull(rs("applied_on")) Then %>
            <span class="cell-muted">-</span>
            <% Else %>
            <%= H(FmtDate(rs("applied_on"))) %>
            <% End If %>
          </td>
          <td class="row-actions">
            <a href="/view.asp?id=<%= H(rs("id")) %>">View</a><span class="sep">|</span><a href="/edit.asp?id=<%= H(rs("id")) %>">Edit</a><span class="sep">|</span><a href="/delete.asp?id=<%= H(rs("id")) %>" class="delete-link" data-delete-id="<%= H(rs("id")) %>">Delete</a>
          </td>
        </tr>
<%
    rs.MoveNext
  Wend
%>
      </tbody>
    </table>
  </div>
</div>

<% End If %>

<%
rs.Close
cn.Close
Set rs = Nothing
Set cmd = Nothing
Set cn = Nothing
%>

<!--#include file="inc/footer.asp"-->
