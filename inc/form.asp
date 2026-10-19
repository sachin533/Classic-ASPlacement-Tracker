<% ' ---------------------------------------------------------------------------
  '  Placement Tracker  -  inc\form.asp
  '  ---------------------------------------------------------------------------
  '  The create / edit form markup, rendered from one place so that new.asp
  '  and edit.asp can never drift apart.
  '
  '  RenderApplicationForm action, submitLabel, recordId, v, errs
  '    action     : form action URL
  '    submitLabel: "Save Application" / "Update Application"
  '    recordId   : "" for create, the id for edit
  '    v          : Dictionary of values (already normalized when re-displayed)
  '    errs       : Dictionary of field -> message
  '    ------------------------------------------------------------------------ %>

<%
Sub RenderApplicationForm(action, submitLabel, recordId, v, errs)
    Dim hasErrs
    hasErrs = (errs.Count > 0)
%>
<form method="post" action="<%= H(action) %>" class="app-form" novalidate>
  <%= CsrfInput() %>
  <% If recordId <> "" Then %>
  <input type="hidden" name="id" value="<%= H(recordId) %>">
  <% End If %>

  <div class="form-grid">

    <div class="form-field<%= ErrClass(errs, "company") %>">
      <label for="f-company">Company <span class="req" title="Required">*</span></label>
      <input type="text" id="f-company" name="company" maxlength="100"
             placeholder="e.g. Google" autocomplete="organization"
             value="<%= H(v("company")) %>"<%= AriaInvalid(errs, "company") %>>
      <%= ErrBlock(errs, "company") %>
    </div>

    <div class="form-field<%= ErrClass(errs, "role") %>">
      <label for="f-role">Role <span class="req" title="Required">*</span></label>
      <input type="text" id="f-role" name="role" maxlength="100"
             placeholder="e.g. Software Engineer Intern"
             value="<%= H(v("role")) %>"<%= AriaInvalid(errs, "role") %>>
      <%= ErrBlock(errs, "role") %>
    </div>

    <div class="form-field<%= ErrClass(errs, "status") %>">
      <label for="f-status">Status <span class="req" title="Required">*</span></label>
      <select id="f-status" name="status"<%= AriaInvalid(errs, "status") %>>
        <%= StatusOptions(v("status")) %>
      </select>
      <%= ErrBlock(errs, "status") %>
    </div>

    <div class="form-field<%= ErrClass(errs, "applied_on") %>">
      <label for="f-applied-on">Applied On <span class="req" title="Required">*</span></label>
      <input type="date" id="f-applied-on" name="applied_on" placeholder="dd-mm-yyyy"
             value="<%= H(v("applied_on")) %>"<%= AriaInvalid(errs, "applied_on") %>>
      <p class="field-hint">Leave empty for a Wishlist application.</p>
      <%= ErrBlock(errs, "applied_on") %>
    </div>

    <div class="form-field form-field-wide<%= ErrClass(errs, "job_url") %>">
      <label for="f-job-url">Job URL <span class="optional">(optional)</span></label>
      <input type="url" id="f-job-url" name="job_url" maxlength="500"
             placeholder="https://" spellcheck="false"
             value="<%= H(v("job_url")) %>"<%= AriaInvalid(errs, "job_url") %>>
      <%= ErrBlock(errs, "job_url") %>
    </div>

    <div class="form-field form-field-wide<%= ErrClass(errs, "notes") %>">
      <label for="f-notes">Notes <span class="optional">(optional)</span></label>
      <textarea id="f-notes" name="notes" rows="3" maxlength="1000"
                placeholder="Add any notes here..."<%= AriaInvalid(errs, "notes") %>><%= H(v("notes")) %></textarea>
      <%= ErrBlock(errs, "notes") %>
    </div>

  </div>

  <div class="form-actions">
    <button type="submit" class="btn btn-primary"><%= H(submitLabel) %></button>
    <a class="btn btn-secondary" href="/default.asp">Cancel</a>
  </div>
</form>
<%
End Sub

' --- option list -----------------------------------------------------------
Function StatusOptions(selected)
    Dim allowed, i, name, out
    allowed = Array("Wishlist", "Applied", "Interview", "Offer", "Rejected")
    out = ""
    For i = 0 To UBound(allowed)
        name = allowed(i)
        out = out & "<option value=""" & H(name) & """"
        If StrComp(name, Trim(selected & ""), vbTextCompare) = 0 Then out = out & " selected=""selected"""
        out = out & ">" & H(name) & "</option>" & vbCrLf
    Next
    StatusOptions = out
End Function

' --- field error helpers ---------------------------------------------------
Function ErrClass(errs, fieldName)
    ErrClass = ""
    If errs.Exists(fieldName) Then ErrClass = " has-error"
End Function

Function AriaInvalid(errs, fieldName)
    AriaInvalid = ""
    If errs.Exists(fieldName) Then AriaInvalid = " aria-invalid=""true"" aria-describedby=""err-" & H(fieldName) & """"
End Function

Function ErrBlock(errs, fieldName)
    ErrBlock = ""
    If Not errs.Exists(fieldName) Then Exit Function
    ErrBlock = "<p class=""field-error"" id=""err-" & H(fieldName) & """ role=""alert"">" & _
               H(errs(fieldName)) & "</p>"
End Function
%>
