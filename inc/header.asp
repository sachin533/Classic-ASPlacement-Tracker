<% ' ---------------------------------------------------------------------------
  '  Placement Tracker  -  inc\header.asp
  '  ---------------------------------------------------------------------------
  '  RenderHeader() draws the doctype, <head>, the dark navigation bar and the
  '  compact search box.  It is a Sub (not top-level markup) so that the shared
  '  error page can render exactly the same chrome without duplicating the HTML.
  '
  '  Every page includes this file through inc\errorpage.asp and calls
  '  RenderHeader() AFTER setting gTitle / gActiveNav / gShowSearch.
  '
  '  The bar's form is a GET to /default.asp carrying only q (and the status
  '  filter when one is active).  It deliberately never carries a page field,
  '  so submitting a new search always restarts the results at page 1.
  '    ------------------------------------------------------------------------ %>
<%
Sub RenderHeader()
    Dim title, showSearch, searchValue, filterStatus, clearHidden

    title = ""
    If Not IsEmpty(gTitle) Then title = gTitle & ""
    If title = "" Then
        title = "Placement Tracker"
    Else
        title = title & " | Placement Tracker"
    End If

    showSearch = False
    If Not IsEmpty(gShowSearch) Then showSearch = CBool(gShowSearch)

    searchValue = ""
    If Not IsEmpty(gSearchValue) Then searchValue = gSearchValue & ""

    filterStatus = ""
    If Not IsEmpty(gFilterStatus) Then filterStatus = gFilterStatus & ""

    ' The clear (x) button is rendered open whenever the field already holds
    ' text, so the control is useful even before the script file loads.
    clearHidden = ""
    If Trim(searchValue) = "" Then clearHidden = " hidden"
%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex">
<title><%= H(title) %></title>
<link rel="stylesheet" href="/static/styles.css?v=4">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24'%3E%3Crect width='24' height='24' rx='5' fill='%232563eb'/%3E%3Cpath d='M7 9h10v8H7z' fill='none' stroke='white' stroke-width='1.6'/%3E%3Cpath d='M10 9V7.6h4V9' fill='none' stroke='white' stroke-width='1.6'/%3E%3C/svg%3E">
</head>
<body>
<a class="skip-link" href="#main">Skip to main content</a>

<header class="topbar">
  <div class="topbar-inner">
    <a class="brand" href="/default.asp">Placement Tracker</a>

    <nav class="mainnav" aria-label="Primary">
      <a class="navlink<%= NavClass("home") %>" href="/default.asp"<%= NavCurrent("home") %>>Home</a>
      <a class="navlink<%= NavClass("new") %>" href="/new.asp"<%= NavCurrent("new") %>>New Application</a>
    </nav>

    <% If showSearch Then %>
    <form class="site-search" role="search" method="get" action="/default.asp">
      <label class="visually-hidden" for="site-search">Search company, role, status, job URL or notes</label>
      <input type="search" id="site-search" name="q" maxlength="100"
             placeholder="Search company, role, status or notes..."
             value="<%= H(searchValue) %>"
             autocomplete="off">
      <% If filterStatus <> "" Then %>
      <input type="hidden" name="status" value="<%= H(filterStatus) %>">
      <% End If %>
      <button type="button" class="site-search-clear js-search-clear"<%= clearHidden %>
              aria-label="Clear the search box">&times;</button>
      <button type="submit" class="site-search-btn" aria-label="Search">
        <%= IconSearch() %>
      </button>
    </form>
    <% End If %>

    <% ' Session state comes from inc\auth.asp, pulled in by inc\helpers.asp.
       ' Signed in: the account name plus a POST + CSRF sign-out.
       ' Signed out: this only happens on the404 page or an error page,
       ' because every other page redirects to /login.asp first. %>
    <% If IsAuthenticated() Then %>
    <div class="topbar-account">
      <span class="account-name" title="Signed in as <%= H(AuthUsername()) %>">
        <%= IconUser() %><span class="account-user"><%= H(AuthUsername()) %></span>
      </span>
      <form class="account-logout" method="post" action="/logout.asp">
        <%= CsrfInput() %>
        <button type="submit" class="account-logout-btn">Sign out</button>
      </form>
    </div>
    <% Else %>
    <a class="account-signin" href="/login.asp"><%= IconSignIn() %>Sign in</a>
    <% End If %>
  </div>
</header>

<main id="main" class="page">
<%= RenderFlash() %>
<%
End Sub
%>
