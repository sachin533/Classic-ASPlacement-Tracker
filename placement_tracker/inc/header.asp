<%@ Language="VBScript" CodePage="65001" %>
<%
' inc/header.asp - shared page top: layout, nav, search, flash messages.
' Must be the FIRST include of every page (declares Option Explicit).
Option Explicit
Response.CodePage = 65001
Response.Charset = "utf-8"
Response.Buffer = True

' Set by pages that can delete a record, so footer.asp renders the modal.
Dim showDeleteModal
showDeleteModal = False

' Active-nav highlight based on the current script name.
Dim scriptName, homeClass, newClass
scriptName = LCase(Request.ServerVariables("SCRIPT_NAME"))
homeClass = "nav-link"
newClass = "nav-link"
If InStr(scriptName, "/new.asp") > 0 Or InStr(scriptName, "/edit.asp") > 0 Then
  newClass = "nav-link active"
Else
  homeClass = "nav-link active"
End If

' Search box keeps the current query visible.
Dim searchQuery
searchQuery = Trim(Request.QueryString("q"))
%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Placement Tracker</title>
<link rel="stylesheet" href="/static/styles.css">
</head>
<body>
<header class="site-header">
  <div class="container header-inner">
    <a class="brand" href="/default.asp">Placement Tracker</a>
    <nav class="main-nav" aria-label="Main">
      <a class="<%= homeClass %>" href="/default.asp">Home</a>
      <a class="<%= newClass %>" href="/new.asp">New Application</a>
    </nav>
    <form class="search-form" method="get" action="/default.asp" role="search">
      <label class="visually-hidden" for="site-search">Search company or role</label>
      <input type="search" id="site-search" name="q" placeholder="Search company or role..."
             value="<%= H(searchQuery) %>">
      <button type="submit" aria-label="Search">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" aria-hidden="true">
          <circle cx="11" cy="11" r="7"></circle><line x1="21" y1="21" x2="16.5" y2="16.5"></line>
        </svg>
      </button>
    </form>
  </div>
</header>
<main class="container">
<% Dim flashMsg : flashMsg = TakeFlash() %>
<% If flashMsg <> "" Then %>
<div class="alert alert-success" role="status"><%= H(flashMsg) %></div>
<% End If %>
