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
dim strRecommendationsTemplate,memberType
strRecommendationsTemplate = ""
memberType = ""
'________________ REQUEST.FORM ________________
vendorID = trim(Request.QueryString("ID"))
memberType = trim(Request.QueryString("memberType"))
'_________________________________________

'******************************* Page Processing '*******************************
process = Request.Form("process")
IF process THEN
	dim specialties,i,pipeDelimitedSpecialties,pipeDelimitedStates,isSavedSuccessfully,csSpecialties, csPractStates
	csSpecialties = trim(Request.Form("ATTNY_SPECIALTY"))
	csPractStates = trim(Request.Form("ATTNY_STATE"))
	
	IF csSpecialties = "" THEN
		pipeDelimitedSpecialties = null
	ELSE
		specialties = split(csSpecialties,",")	
		FOR i=0 to UBound(specialties)
			IF i=0 THEN
				pipeDelimitedSpecialties = specialties(i)
			ELSE
				pipeDelimitedSpecialties = pipeDelimitedSpecialties &"|"&specialties(i)
			END IF
		Next
	END IF
	IF csPractStates <> "" THEN
		practStates = split(csPractStates,",")
		FOR i=0 to UBound(practStates)
			IF i=0 THEN
				pipeDelimitedStates = practStates(i)
			ELSE
				pipeDelimitedStates = pipeDelimitedStates &"|"&practStates(i)
			END IF
		Next
	ELSE
		pipeDelimitedStates = null
	END IF
	category = trim(Request.Form("Basic_category"))' TODO: if category is not provided then no need to go further.
	yearsOfServiceUsed = trim(Request.Form("yearsOfServiceUsed"))
	IF yearsOfServiceUsed = "" THEN
		yearsOfServiceUsed =  NULL
	END IF
	recommendedFirm = trim(Request.Form("chkFirm"))
	if len(recommendedFirm)<=0 then
		recommendedFirm = NULL
	End if	
	recommendedAttorney = trim(Request.Form("chkAttorney"))
	if len(recommendedAttorney)<=0 then
		recommendedAttorney = NULL
	End if	
	recommendedBy = trim(Request.Form("txtRecommendedBy"))
	IF len(recommendedBy) <=0 THEN
		recommendedBy = NULL
	END IF
	isRecommended = trim(Request.Form("rdoRecommended"))
	recommendationComments = trim(Request.Form("txtRecommendationComments"))
	IF len(recommendationComments)<=0 THEN
		recommendationComments = NULL
	END IF	
	logoPath = trim(Request.Form("logoPath"))
	IF len(logoPath) <=0 THEN
		logoPath = NULL
	END IF
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
	IF vendorID <> "" AND IsNumeric(vendorID) THEN
		const TEMPLATE_PATH = "templates/"
		const ForReading = 1, ForWriting = 2, ForAppending = 8
		strRecommendationsTemplate = GetTemplate2("vendor_recommendations.htm")
		
		'If vendorID is provided then call the SP and get data
		strSB = GetAddvendorPage(vendorID)	
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
.addvendor
{
padding-left:15px;
padding-right:15px;
}
</style>
</head>
<body>
<!-- #include file="../TL_Header.asp" -->
 <h1>ADD VENDOR</h1>
 <span>This page is for adding Member to vendor directory.Once vendor is added to vendor directory,if Logo has to be uploaded then search for the vendor in search page and Edit the vendor, then upload the logo</span>
<form name="frmAddVendor" id="frmAddVendor" method="post"> 
<div id="AddVendor" class="addvendor">
<%=strDisplay %>
</div>
<input type="hidden" name="process" value="1" />
<input type="submit" name="btnSaveVendor" value="Save Vendor" />
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
