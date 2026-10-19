<%
' inc/form.asp - the create/edit form markup, shared by new.asp and edit.asp.
' Needs header.asp, helpers.asp and validate.asp included before it.

' Fresh value dictionary with every key the form knows about.
Function EmptyVals()
  Dim v
  Set v = Server.CreateObject("Scripting.Dictionary")
  v("company") = ""
  v("role") = ""
  v("status") = "Applied"
  v("applied_on") = ""
  v("job_url") = ""
  v("notes") = ""
  Set EmptyVals = v
End Function

' One text field with label, optional red asterisk, error message and
' the previously entered value (so a failed save never loses input).
Sub FieldText(ByVal inputId, ByVal labelText, ByVal isRequired, ByVal val, ByVal errText, ByVal placeholder, ByVal inputType)
  Dim cls, msg
  cls = "form-control"
  msg = ""
  If errText <> "" Then cls = cls & " invalid"
%>
<div class="form-group">
  <label for="<%= inputId %>"><%= labelText %><% If isRequired Then %><span class="req" aria-hidden="true">*</span><% End If %></label>
  <input type="<%= inputType %>" id="<%= inputId %>" name="<%= inputId %>" class="<%= cls %>"
         value="<%= H(val) %>" placeholder="<%= H(placeholder) %>"<% If isRequired Then %> required<% End If %>>
  <% If errText <> "" Then %><p class="field-error" id="<%= inputId %>-error"><%= H(errText) %></p><% End If %>
</div>
<%
End Sub

' The full application form. mode = "create" or "edit".
Sub RenderAppForm(ByVal mode, ByVal vals, ByVal errs, ByVal cancelUrl)
  Dim hasErr, actionLabel, errKey
  hasErr = (errs.Count > 0)
  If mode = "edit" Then
    actionLabel = "Update Application"
  Else
    actionLabel = "Save Application"
  End If
%>
<% If hasErr Then %>
<div class="alert alert-error" role="alert" tabindex="-1" id="form-errors">
  <strong>Please fix the highlighted fields below.</strong>
</div>
<% End If %>

<form method="post" action="" class="card form-card" novalidate>
  <input type="hidden" name="csrf" value="<%= H(CsrfToken()) %>">
  <div class="form-grid">
    <%
      Call FieldText("company", "Company", True, vals("company"), errs("company"), "e.g. Google", "text")
      Call FieldText("role", "Role", True, vals("role"), errs("role"), "e.g. Software Engineer Intern", "text")
    %>
    <div class="form-group">
      <label for="status">Status<span class="req" aria-hidden="true">*</span></label>
      <select id="status" name="status" class="form-control<% If errs("status") <> "" Then %> invalid<% End If %>">
        <% Dim st : For Each st In Array("Wishlist", "Applied", "Interview", "Offer", "Rejected") %>
        <option value="<%= st %>"<% If vals("status") = st Then %> selected<% End If %>><%= st %></option>
        <% Next %>
      </select>
      <% If errs("status") <> "" Then %><p class="field-error"><%= H(errs("status")) %></p><% End If %>
    </div>
    <%
      Call FieldText("applied_on", "Applied On", True, vals("applied_on"), errs("applied_on"), "yyyy-mm-dd", "date")
      Call FieldText("job_url", "Job URL (optional)", False, vals("job_url"), errs("job_url"), "https://", "url")
    %>
    <div class="form-group form-group-wide">
      <label for="notes">Notes (optional)</label>
      <textarea id="notes" name="notes" rows="3" placeholder="Add any notes here..."
                class="form-control<% If errs("notes") <> "" Then %> invalid<% End If %>"><%= H(vals("notes")) %></textarea>
      <% If errs("notes") <> "" Then %><p class="field-error"><%= H(errs("notes")) %></p><% End If %>
    </div>
  </div>
  <div class="form-actions">
    <button type="submit" class="btn btn-primary"><%= actionLabel %></button>
    <a class="btn btn-secondary" href="<%= H(cancelUrl) %>">Cancel</a>
  </div>
</form>
<%
End Sub
%>
