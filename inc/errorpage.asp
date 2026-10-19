<% ' ---------------------------------------------------------------------------
  '  Placement Tracker  -  inc\errorpage.asp
  '  ---------------------------------------------------------------------------
  '  RenderErrorPage(code, title, message) draws the whole error page: the
  '  standard chrome plus the big status card.
  '
  '  WHY A SUB AND NOT Server.Execute
  '    Server.Execute cannot take a query string in Classic ASP - it hands the
  '    whole string to MapPath and fails with ASP 0173 "Invalid Path
  '    Character" - and it starts a NEW page scope, so values set by the
  '    caller are invisible inside the executed page.  Passing the code and
  '    the text as arguments is therefore the only reliable option.
  '
  '  RenderHeader() / RenderFooter() are included at file level (below) and
  '  called from inside this Sub; a Sub may not be defined inside another Sub.
  '
  '  It must never call Response.End - the caller owns that decision.
  '    ------------------------------------------------------------------------ %>
<!--#include virtual="/inc/header.asp" -->
<!--#include virtual="/inc/footer.asp" -->
<%
Sub RenderErrorPage(code, title, message)
    gErrCode     = code
    gErrTitle    = title
    gErrMsg      = message
    gTitle       = title
    gActiveNav   = ""
    gShowSearch  = False
    gSearchValue = ""

    RenderHeader
%>
<div class="error-wrap">
  <div class="card error-card">
    <p class="error-code"><%= H(code) %></p>
    <h1 class="error-title"><%= H(title) %></h1>
    <p class="error-message"><%= H(message) %></p>
    <div class="error-actions">
      <a class="btn btn-primary" href="/default.asp">Back to Applications</a>
      <a class="btn btn-secondary" href="/new.asp">Add Application</a>
    </div>
  </div>
</div>
<%
    RenderFooter
End Sub
%>
