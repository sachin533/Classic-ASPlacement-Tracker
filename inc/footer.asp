<% ' ---------------------------------------------------------------------------
  '  Placement Tracker  -  inc\footer.asp
  '  ---------------------------------------------------------------------------
  '  RenderFooter() closes the page and hosts the shared delete-confirmation
  '  dialog.  Like RenderHeader() it is a Sub so the error page can reuse the
  '  identical chrome.
  '    ------------------------------------------------------------------------ %>
<%
Sub RenderFooter()
%>
</main>

<footer class="site-footer">
  <div class="site-footer-inner">
    <span>Placement Tracker</span>
    <span class="dot" aria-hidden="true">&middot;</span>
    <span>Classic ASP + SQL Server</span>
  </div>
</footer>

<!-- =========================================================================
     Shared delete confirmation dialog.  The trigger buttons live inside real
     <form method="post" action="/delete.asp"> elements, so without JavaScript
     the browser simply confirms and posts; with JavaScript the dialog is used
     instead.  The server re-validates the id and the CSRF token either way.
     ====================================================================== -->
<div class="modal-overlay" id="delete-modal" hidden>
  <div class="modal" role="dialog" aria-modal="true"
       aria-labelledby="delete-modal-title" aria-describedby="delete-modal-body">
    <div class="modal-head">
      <h2 id="delete-modal-title">Delete Application</h2>
      <button type="button" class="modal-close" data-modal-close aria-label="Close dialog">&times;</button>
    </div>
    <div class="modal-body" id="delete-modal-body">
      <span class="modal-icon"><%= IconWarning() %></span>
      <div class="modal-text">
        <p class="modal-question">Are you sure you want to delete this application?</p>
        <p class="modal-note">This action cannot be undone.</p>
      </div>
    </div>
    <div class="modal-actions">
      <button type="button" class="btn btn-danger" id="delete-modal-confirm">Yes, Delete</button>
      <button type="button" class="btn btn-secondary" data-modal-close>Cancel</button>
    </div>
  </div>
</div>

<script src="/static/app.js?v=1"></script>
</body>
</html>
<%
End Sub
%>
