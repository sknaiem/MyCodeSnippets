<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%> 
<% Response.CharSet = "UTF-8" %>
<!--#include file="../console/common.asp"-->
<!--#include file="../console/TL_Reference.asp"-->
<!--#include file="../IFAConnection.asp" -->
<!--#include file="vendor_search_functions.asp" -->
<%    
If NOT trim(session("loggedin"))="True" Then
	response.redirect trim(application("siteurl"))& "../../login.asp"
End If 
'________________ REQUEST.FORM ________________
vendorID = trim(Request.QueryString("ID"))
'_________________________________________

'******************************* Page Processing '*******************************
IF len(vendorID)> 0 THEN
	'If vendorID is provided then call the SP and get data
	'strSB = strSB & "Vendor ID:"&vendorID
	isDeleted = DeleteVendor(vendorID)
	IF isDeleted THEN
		strSB = strSB & "<center><font color='red'><b>"&vendorID&"</b> has been deleted</font><br/>"
		strSB = strSB & "<a href='Search.asp'>Click here to go back to Search Page</a>"
	ELSE
		strSB = strSB & "<center><font color='red'>Could not delete the vendor</font><br/>"
		strSB = strSB & "<a href='Search.asp'>Click here to go back to Search Page</a>"
	END IF
	strDisplay = strSB
ELSE
	strSB = strSB & "<center><font color='red'>Vendor selected is not a valid vendor</font><br/>"
	strSB = strSB & "<a href='Search.asp'>Click here to go back to Search Page</a>"
END IF 
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=TOPS_PAGE_HTML_TITLE%></title>
</head>
<body>
<!-- #include file="../TL_Header.asp" -->
<h1>DELETE VENDOR</h1>
<span>This page is for deleting a vendor from vendor directory.</span>
<div id="DeleteVendor" class="deletevendor">
<%=strDisplay %>
<!-- #include file="../TL_Footer.asp" -->
</body>
</html>
