<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%> 
<% Response.CharSet = "UTF-8" %>

<% 'nshaik: Case 638 - TL - IMP CUSTOM WORK: IFA Vendor Search %>
<!--#include file="../IFAConnection.asp" -->
<!--#include file="../includes/control.asp"-->
<!--#include file="../includes/siteaccess.asp"-->
<!--#include file="../includes/staticcm.asp"-->
<!--#include file="../includes/custom_cfcc.asp"-->
<!--#include file="../includes/advertising.asp"-->
<!--#include file="../includes/rotate_sub.asp"-->
<!--#include file="../includes/displayformatteditem2.asp"-->
<!--#include file="../includes/member_directory.asp"-->
<!--#include file="../includes/member_control.asp"-->
<!--#include file="Vendor_Search_functions.asp" -->

<%


'Response.Write ("Vendor Search in progress" &  "" & "<br>")
'Response.End

'can make a function here to check if logged in.
If trim(session("sa_id"))="" Then
	response.redirect trim(application("siteurl"))& "/login.asp"
End If 

'________________ REQUEST.FORM ________________
vendorID = trim(Session("sa_memberid"))
'______________________________________________   


'******************************* Page Processing '*******************************
process = Request.Form("process")
IF process THEN
	dim specialties,i,pipeDelimitedSpecialties,pipeDelimitedStates,isSavedSuccessfully, csSpecialties, csPractStates
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
	
	category = trim(Request.Form("Basic_category"))	
	'logoPath = trim(Request.Form("logoPath"))
	dim model(10)
	model(0) = vendorID
	model(1) = category
	model(2) = null'yearsOfServiceUsed
	model(3) = null'recommendedFirm
	model(4) = null'recommendedAttorney
	model(5) = null'recommendedBy
	model(6) = null'isRecommended
	model(7) = null'recommendationComments
	model(8) = null'logoPath
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
	IF vendorID <> "" and IsNumeric(vendorID) THEN
		'If vendorID is provided then call the SP and get data
		strSB = GetVendorDetailsForEdit(vendorID)	
		strDisplay = strSB	
	ELSE  
		strDisplay = "Unable to fetch vendor information"
	END IF
END IF
'********************************************************************************


Sub ThePageTitle()
	showcmn "Main Page Title Tag"
End Sub

Sub Metatags()
End Sub
 
   
Sub Javascripts()
%>
<%	
End Sub

Sub HeadCode()
 custom_head_code 
%>

<!-- MBELCHER: This is used for our Vendor search template -->
<link rel="stylesheet" type="text/css" href="templates/vendor_search.css">
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

<%
End Sub 
  
Sub TopNav()
	custom_topnav
End Sub

Sub LeftSide()
 	If Len(session("sa_id")) > 0 Then
		'AF_Menu '...... we are using the custom menu in designtemplate
	End If
 ShowLoginBox()
End Sub                  

Sub BreadCrumbs()
%>
 <a href="index.asp">Home</a> > <%=pagetitle%>
<% End Sub %>          



<% Sub PageBody() %>
 

 <h1>VENDOR PROFILE</h1>

<form name="frmEditProfile" id="frmEditProfile" method="post">
<div id="VendorProfile">
<%=strDisplay %>
</div>
<input type="hidden" name="process" value="1" />
<input type="submit" name="btnEditProfile" value="Submit" />
</form>
	
<% End Sub  %>
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


<!--#include file="../designtemplate.asp"-->				

     
