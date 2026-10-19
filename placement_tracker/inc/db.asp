<%
' inc/db.asp - connection string and OpenDb()
' One include per page - never include the same file twice in Classic ASP.
Const CONN_STR = "Provider=MSOLEDBSQL;Server=.\SQLEXPRESS;Database=PlacementTracker;Integrated Security=SSPI"

' Fallback if MSOLEDBSQL is not installed (32-bit IIS Express):
' Const CONN_STR = "Provider=SQLOLEDB;Data Source=.\SQLEXPRESS;Initial Catalog=PlacementTracker;Integrated Security=SSPI"

Function OpenDb()
  Dim cn
  Set cn = Server.CreateObject("ADODB.Connection")
  cn.ConnectionTimeout = 5
  cn.Open CONN_STR
  Set OpenDb = cn
End Function
%>
