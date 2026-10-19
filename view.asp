<!--#include virtual="/inc/helpers.asp" -->
<!--#include virtual="/inc/db.asp" -->
<%
' ---------------------------------------------------------------------------
'  Details - view.asp?id=
'  Shows one record.  A missing, malformed or unknown id produces a real
'  HTTP 404 through the shared error page.
' ---------------------------------------------------------------------------

Dim id, rec
id = ParseId(Request.QueryString("id"))
If id < 0 Then NotFound()

Set rec = GetApplication(id)
If rec Is Nothing Then NotFound()

gTitle       = "Application Details"
gActiveNav   = "home"
gShowSearch  = False
gSearchValue = ""
%>
<% Call RenderHeader() %>

<a class="back-link" href="/default.asp"><%= IconBack() %>Back to Applications</a>

<div class="page-head">
  <h1>Application Details</h1>
  <div class="detail-actions">
    <a class="btn btn-primary" href="/edit.asp?id=<%= rec("id") %>"><%= IconEdit() %>Edit</a>
    <button type="submit" form="app-delete-form"
            class="btn btn-danger js-delete-trigger"><%= IconTrash() %>Delete</button>
  </div>
</div>

<div class="card">
  <dl class="detail-list">
    <div class="detail-row">
      <dt>Company</dt>
      <dd><%= H(rec("company")) %></dd>
    </div>
    <div class="detail-row">
      <dt>Role</dt>
      <dd><%= H(rec("role")) %></dd>
    </div>
    <div class="detail-row">
      <dt>Status</dt>
      <dd><%= StatusBadge(rec("status")) %></dd>
    </div>
    <div class="detail-row">
      <dt>Applied On</dt>
      <dd<% If rec("applied_on_iso") = "" Then %> class="is-empty"<% End If %>><%= H(EmptyDash(rec("applied_on_iso"))) %></dd>
    </div>
    <div class="detail-row">
      <dt>Job URL</dt>
      <dd><% If rec("job_url") <> "" Then %><a href="<%= H(rec("job_url")) %>" rel="noopener noreferrer" target="_blank"><%= H(rec("job_url")) %></a><% Else %><span class="is-empty">Not provided</span><% End If %></dd>
    </div>
    <div class="detail-row">
      <dt>Notes</dt>
      <dd<% If rec("notes") = "" Then %> class="is-empty"<% End If %>><% If rec("notes") <> "" Then %><%= H(rec("notes")) %><% Else %>No notes<% End If %></dd>
    </div>
    <div class="detail-row">
      <dt>Created At</dt>
      <dd><%= H(rec("created_at_text")) %></dd>
    </div>
  </dl>
</div>

<form id="app-delete-form" method="post" action="/delete.asp"
      class="visually-hidden js-delete-form"
      onsubmit="return confirm('Are you sure you want to delete this application? This action cannot be undone.');">
  <input type="hidden" name="id" value="<%= rec("id") %>">
  <%= CsrfInput() %>
  <button type="submit">Delete application</button>
</form>

<% Call RenderFooter() %>

