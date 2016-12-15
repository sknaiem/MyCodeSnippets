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
IF vendorID <> "" THEN
	'If vendorID is provided then call the SP and get data
	strSB = GetVendorDetailsForEdit(vendorID)	
	strDisplay = strSB
 END IF 
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=TOPS_PAGE_HTML_TITLE%></title>
<style>
div.row
{
border:1px solid #eeeeee;
background:#eeeeee;
}
span.LeftControl
{
width:30%;
float:left;
}
span.RightControl
{
width:70%;
float:left;
}
</style>
</head>
<body>
<!-- #include file="../TL_Header.asp" -->
 <h1>EDIT VENDOR</h1>
<div id="EditVendor">
<%=strDisplay %>
<%dim i
For Each i in Application.Contents
  Response.Write(i & "<br>")
  Response.Write(Application.Contents(i) & "<br>")
  Response.Write("------------------------------<br/>")
Next

' dim j
' j=Application.Contents.Count
' For i=1 to j
  ' Response.Write(Application.Contents(i) & "<br>")
' Next
%>
</div>
<!-- #include file="../TL_Footer.asp" -->
</body>
</html>
<script language="JavaScript" type="text/JavaScript">
function openAddPhoto(){
	window.open('addLogo_Vendor.asp?addphoto=1&ID=<%=vendorID%>', 'multiPopup', 'toolbar=no,location=no,status=no,scrollbars=no,menubar=no,width=600,height=500,resizable=no' );
}
</script>
