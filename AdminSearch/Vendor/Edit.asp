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
process = trim(Request.Form("process"))
IF process = 1 THEN
	dim model(10)
	model(0) = vendorID
	model(1) = Request.Form("Basic_category")
	model(2) = Request.Form("yearsOfServiceUsed")
	model(3) = Request.Form("chkAttorney")
	model(4) = Request.Form("chkFirm")
	model(5) = Request.Form("txtRecommendedBy")
	model(6) = Request.Form("rdoRecommended")
	model(7) = Request.Form("txtRecommendationComments")
	model(8) = Request.Form("logoPath")
	model(9) = Request.Form("ATTNY_SPECIALTY")	
	model(10) = Request.Form("ATTNY_STATE")

	
ELSE
	IF vendorID <> "" THEN
		const TEMPLATE_PATH = "templates/"
		const ForReading = 1, ForWriting = 2, ForAppending = 8
		strRecommendationsTemplate = GetTemplate2("vendor_recommendations.htm")
		
		'If vendorID is provided then call the SP and get data
		strSB = GetVendorDetailsForEdit(vendorID)	
		strDisplay = strSB
	END IF
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
<form name="frmEditVendor" id="frmEditVendor" action="Edit.asp" method="post"> 
<div id="EditVendor">
<%=strDisplay %>
<%
' dim i
' For Each i in Application.Contents
  ' Response.Write(i & "<br>")
  ' Response.Write(Application.Contents(i) & "<br>")
  ' Response.Write("------------------------------<br/>")
' Next

' dim j
' j=Application.Contents.Count
' For i=1 to j
  ' Response.Write(Application.Contents(i) & "<br>")
' Next
%>
</div>
<input type="hidden" name="process" value="1" />
<input type="submit" name="btnEditVendor" value="Submit" />
</form>
<!-- #include file="../TL_Footer.asp" -->
</body>
</html>
<script language="JavaScript" type="text/JavaScript">
function openAddPhoto(){
	window.open('addLogo_Vendor.asp?addphoto=1&ID=<%=vendorID%>', 'multiPopup', 'toolbar=no,location=no,status=no,scrollbars=no,menubar=no,width=600,height=500,resizable=no' );
}
function confirm_delete(URL)
{
  if (confirm("Dleleting this photo will remove it from your profile and the online directory listing. Click 'OK' to delete the image ")) 
  { 
  window.open(URL, 'multiPopup', 'toolbar=no,location=no,status=no,scrollbars=no,menubar=no,width=600,height=500,resizable=no' ); 
  }
}
</script>
