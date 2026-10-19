<!--#include file="inc/header.asp"-->
<!--#include file="inc/helpers.asp"-->
<%
' 404.asp - friendly not-found page.
' Reached through NotFound(), which sets the real 404 status code first
' and then Server.Execute's this file.
%>
<div class="card center-card">
  <span class="empty-icon" aria-hidden="true">
    <svg width="34" height="34" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7">
      <circle cx="11" cy="11" r="7"></circle>
      <line x1="21" y1="21" x2="16.5" y2="16.5"></line>
    </svg>
  </span>
  <h1>Page not found</h1>
  <p>That application does not exist, or the link is wrong.</p>
  <a class="btn btn-primary" href="/default.asp">Back to Applications</a>
</div>

<!--#include file="inc/footer.asp"-->
