<!-- #INCLUDE FILE = "includes/Adovbs.inc" -->
<%
Dim IFAconn 
set IFAconn = Server.CreateObject("ADODB.Connection")
IFAconn.Open Application("fraudconnectionstring")
Response.write(vbCrLf & "<!-- mbelcher: -  \\10.18.12.41\Sites\ifa\custom\IFAConnection.asp - LOADED -->" & vbCrLf)
%>