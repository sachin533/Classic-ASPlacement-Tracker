<!--#include file="inc/header.asp"-->
<!--#include file="inc/helpers.asp"-->
<!--#include file="inc/db.asp"-->
<%
' view.asp?id=N - details for one application.
' Unknown, missing or non-numeric ids get a real 404 response.
Dim rawId, id
rawId = Trim(Request.QueryString("id"))

' Strictly digits (1-9 digits). IsNumeric alone would accept "1E5"
' or "-1" and crash CLng or look up nonsense ids.
If rawId = "" Or Not RegExpIsDigits(rawId) Then NotFound()
id = CLng(rawId)

Dim cn, cmd, rs
Set cn = OpenDb()
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = cn
cmd.CommandText = "SELECT id, company, role, status, applied_on, job_url, notes, created_at " _
                & "FROM dbo.Applications WHERE id = ?"
cmd.Parameters.Append cmd.CreateParameter("id", 3, 1, , id)   ' adInteger, adParamInput
Set rs = cmd.Execute

If rs.EOF Then
  rs.Close : cn.Close
  NotFound()
End If
%>

<a class="back-link" href="/default.asp">&larr; Back to Applications</a>

<div class="page-head">
  <h1>Application Details</h1>
  <div class="page-head-actions">
    <a class="btn btn-primary" href="/edit.asp?id=<%= H(rs("id")) %>">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
        <path d="M17 3a2.8 2.8 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5L17 3z"></path>
      </svg>
      Edit
    </a>
    <a class="btn btn-danger" href="/delete.asp?id=<%= H(rs("id")) %>" data-delete-id="<%= H(rs("id")) %>">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
        <path d="M3 6h18"></path><path d="M8 6V4a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1v2"></path>
        <path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"></path>
      </svg>
      Delete
    </a>
  </div>
</div>

<div class="card detail-card">
  <div class="detail-row">
    <span class="detail-label">Company</span>
    <span class="detail-value"><%= H(rs("company")) %></span>
  </div>
  <div class="detail-row">
    <span class="detail-label">Role</span>
    <span class="detail-value"><%= H(rs("role")) %></span>
  </div>
  <div class="detail-row">
    <span class="detail-label">Status</span>
    <span class="detail-value"><span class="<%= StatusClass(rs("status")) %>"><%= H(rs("status")) %></span></span>
  </div>
  <div class="detail-row">
    <span class="detail-label">Applied On</span>
    <span class="detail-value">
      <% If IsNull(rs("applied_on")) Then %><span class="cell-muted">-</span><% Else %><%= H(FmtDate(rs("applied_on"))) %><% End If %>
    </span>
  </div>
  <div class="detail-row">
    <span class="detail-label">Job URL</span>
    <span class="detail-value url">
      <% If IsNull(rs("job_url")) Or rs("job_url") = "" Then %>
      <span class="cell-muted">-</span>
      <% Else %>
      <a href="<%= H(rs("job_url")) %>" rel="noopener noreferrer" target="_blank"><%= H(rs("job_url")) %></a>
      <% End If %>
    </span>
  </div>
  <div class="detail-row">
    <span class="detail-label">Notes</span>
    <span class="detail-value">
      <% If IsNull(rs("notes")) Or rs("notes") = "" Then %><span class="cell-muted">-</span><% Else %><%= H(rs("notes")) %><% End If %>
    </span>
  </div>
  <div class="detail-row">
    <span class="detail-label">Created At</span>
    <span class="detail-value"><%= H(FmtStamp(rs("created_at"))) %> (UTC)</span>
  </div>
</div>

<%
showDeleteModal = True
rs.Close
cn.Close
Set rs = Nothing : Set cmd = Nothing : Set cn = Nothing
%>

<!--#include file="inc/footer.asp"-->
