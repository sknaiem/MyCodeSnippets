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
process = Request.Form("process")
IF process THEN
	dim specialties,i,pipeDelimitedSpecialties,pipeDelimitedStates,isSavedSuccessfully
	specialties = split(trim(Request.Form("ATTNY_SPECIALTY")),",")
	practStates = split(trim(Request.Form("ATTNY_STATE")),",")
	FOR i=0 to UBound(specialties)
		IF i=0 THEN
			pipeDelimitedSpecialties = specialties(i)
		ELSE
			pipeDelimitedSpecialties = pipeDelimitedSpecialties &"|"&specialties(i)
		END IF
	Next
	FOR i=0 to UBound(practStates)
		IF i=0 THEN
			pipeDelimitedStates = practStates(i)
		ELSE
			pipeDelimitedStates = pipeDelimitedStates &"|"&practStates(i)
		END IF
	Next
	category = trim(Request.Form("Basic_category"))
	yearsOfServiceUsed = trim(Request.Form("yearsOfServiceUsed"))
	recommendedFirm = trim(Request.Form("chkFirm"))
	if len(recommendedFirm)<=0 then
		recommendedFirm = 0
	End if	
	recommendedAttorney = trim(Request.Form("chkAttorney"))
	if len(recommendedAttorney)<=0 then
		recommendedAttorney = 0
	End if	
	recommendedBy = trim(Request.Form("txtRecommendedBy"))
	isRecommended = trim(Request.Form("rdoRecommended"))
	recommendationComments = trim(Request.Form("txtRecommendationComments"))
	logoPath = trim(Request.Form("logoPath"))
	dim model(10)
	model(0) = vendorID
	model(1) = category
	model(2) = yearsOfServiceUsed
	model(3) = recommendedFirm
	model(4) = recommendedAttorney
	model(5) = recommendedBy
	model(6) = isRecommended
	model(7) = recommendationComments
	model(8) = logoPath
	model(9) = pipeDelimitedSpecialties
	model(10) = pipeDelimitedStates
	'ToDo: Remove later
	strSB = strSB & "<div><span>Debug information</span><br/>"	
	strSB = strSB & "<span>VendorID:</span><span>"&model(0)&"</span><br/>"
	strSB = strSB & "<span>Category:</span><span>"&model(1)&"</span><br/>"
	strSB = strSB & "<span>ServiceUsed:</span><span>"&model(2)&"</span><br/>"
	strSB = strSB & "<span>IsFirm:</span><span>"&model(3)&"</span><br/>"
	strSB = strSB & "<span>IsAttorney:</span><span>"&model(4)&"</span><br/>"
	strSB = strSB & "<span>RecommendedBy:</span><span>"&model(5)&"</span><br/>"
	strSB = strSB & "<span>IsRecommended:</span><span>"&model(6)&"</span><br/>"
	strSB = strSB & "<span>Rec Comments:</span><span>"&model(7)&"</span><br/>"
	strSB = strSB & "<span>logo Path:</span><span>"&model(8)&"</span><br/>"
	strSB = strSB & "<span>specialties:</span><span>"&model(9)&"</span><br/>"
	strSB = strSB & "<span>pract states:</span><span>"&model(10)&"</span><br/>"
	strDisplay = strSB
	' TODO: Remove above code
	isSavedSuccessfully = AddOrUpdateVendorDetails(model)
	' TODO: show confirmation page.
	Response.Redirect("Confirm.asp?success="&isSavedSuccessfully)
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
<form name="frmEditVendor" id="frmEditVendor" method="post"> 
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
