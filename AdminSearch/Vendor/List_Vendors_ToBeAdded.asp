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

'_________________________________________

'******************************* Page Processing '*******************************
strDisplay = GetVendorsDataNotIncludedInVendorDirectory()
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=TOPS_PAGE_HTML_TITLE%></title>
<style type="text/css">
.vendors_table{
	width:100%;
	cellpadding:2px;
	cellspacing:2px;
}
.vendors_table_header{
	font-weight: bold;
    color: #ffffff;
    padding: 4px;
    vertical-align: middle;
    background-color: #054e86;
    text-align: left;
}
.vendors_table_row{
	background-color:#eeeeee;
}	
</style>
</head>
<body>
<!-- #include file="../TL_Header.asp" -->
<h1>MEMBERS TO BE ADDED IN VENDOR DIRECTORY</h1>
<div id="VendorDetails">
<%=strDisplay %>
<!-- #include file="../TL_Footer.asp" -->
</body>
</html>
