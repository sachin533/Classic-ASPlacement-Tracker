<%
' Delete confirmation modal - shared by the dashboard and the details page.
' The form POSTs id + CSRF to /delete.asp; without JavaScript the plain link
' goes to /delete.asp directly, which answers 405 and deletes nothing.
If showDeleteModal Then
%>
<div class="modal-backdrop" id="delete-modal" hidden>
  <form method="post" action="/delete.asp" class="modal-card" id="delete-form">
    <input type="hidden" name="csrf" value="<%= H(CsrfToken()) %>">
    <input type="hidden" name="id" id="delete-id" value="">
    <div class="modal-head">
      <h2>Delete Application</h2>
      <button type="button" class="modal-close" data-modal-close aria-label="Close dialog">&times;</button>
    </div>
    <div class="modal-body">
      <span class="modal-icon" aria-hidden="true">
        <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <circle cx="12" cy="12" r="9"></circle><line x1="12" y1="7" x2="12" y2="13"></line>
          <circle cx="12" cy="16.6" r="0.4" fill="currentColor"></circle>
        </svg>
      </span>
      <p>Are you sure you want to delete this application?<br>
        <span class="muted">This action cannot be undone.</span></p>
    </div>
    <div class="modal-actions">
      <button type="submit" class="btn btn-danger">Yes, Delete</button>
      <button type="button" class="btn btn-secondary" data-modal-close>Cancel</button>
    </div>
  </form>
</div>
<script>
(function () {
  var modal = document.getElementById('delete-modal');
  if (!modal) return;
  var idField = document.getElementById('delete-id');
  function open(id) {
    idField.value = id;
    modal.hidden = false;
  }
  function close() { modal.hidden = true; }
  document.addEventListener('click', function (e) {
    var trigger = e.target.closest ? e.target.closest('[data-delete-id]') : null;
    if (trigger) { e.preventDefault(); open(trigger.getAttribute('data-delete-id')); return; }
    if (e.target.closest && e.target.closest('[data-modal-close]')) close();
  });
  modal.addEventListener('click', function (e) { if (e.target === modal) close(); });
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && !modal.hidden) close(); });
})();
</script>
<% End If %>
</main>
<footer class="site-footer">
  <div class="container">
    <p>Placement Application Tracker &middot; Built with Classic ASP &amp; SQL Server</p>
  </div>
</footer>
</body>
</html>
