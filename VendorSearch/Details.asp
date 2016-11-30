<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%> 
<% Response.CharSet = "UTF-8" %>

<% 'nshaik: Case 638 - TL - IMP CUSTOM WORK: IFA Vendor Search %>
<!--#include file="../IFAConnection.asp" -->
<!--#include file="../../includes/control.asp"-->
<!--#include file="../../includes/siteaccess.asp"-->
<!--#include file="../../includes/staticcm.asp"-->
<!--#include file="../../includes/custom_cfcc.asp"-->
<!--#include file="../../includes/advertising.asp"-->
<!--#include file="../../includes/rotate_sub.asp"-->
<!--#include file="../../includes/displayformatteditem2.asp"-->
<!--#include file="../../includes/member_directory.asp"-->
<!--#include file="../../includes/member_control.asp"-->
<!--#include file="Vendor_Search_functions.asp" -->

<%


'Response.Write ("Vendor Search in progress" &  "" & "<br>")
'Response.End

'can make a function here to check if logged in.
If trim(session("sa_id"))="" Then
	response.redirect trim(application("siteurl"))& "/login.asp"
End If 

const TEMPLATE_PATH = "templates/"
const ForReading = 1, ForWriting = 2, ForAppending = 8
strSearchTemplate = GetTemplate2("vendor_search.htm")
'strConformationTemplate = GetTemplate2("vendor_search_conformation.htm")




'________________ REQUEST.FORM ________________
vendorID = trim(Request.QueryString("ID"))
'_________________________________________



    


'******************************* Page Processing '*******************************
IF vendorID <> "" THEN
	'If vendorID is provided then call the SP and get data
	strSB = GetVendorDetails(vendorID)
	'TODO:If data is not available then display vendor does not exist
	'If vendor exists then display the data.
	strSB = StrSB & "<hr size='1' align='center'>"
	strSB = strSB & populateSearchPage()
	strDisplay = strSB	

ELSE
    'can load in template of form here.
    'strDisplay = " We display the form here."
    'We might have to load in catagories and other things here; not sure yet.
	' dim vendorCategories, stateOrProvince, countries, attorneyPracticingStates, attorneySpecialties
	' vendorCategories = GetVendorCategories()
	' stateOrProvince = GetStatesOrProvince()
	' countries = GetCountries()
	' attorneyPracticingStates = GetAttorneyPracticingStates()
	' attorneySpecialties = GetAttorneySpecilaties()
    ' strSB = strSearchTemplate
	' strSB = replace(strSB,"[CategoryOptions]",vendorCategories)
	' strSB = replace(strSB,"[StateProvinceOptions]",stateOrProvince)
	' strSB = replace(strSB,"[CountryOptions]",countries)
	' strSB = replace(strSB,"[AttorneyPracticingStateOptions]",attorneyPracticingStates)
	' strSB = replace(strSB,"[AttorneySpecialtiesOptions]",attorneySpecialties)
	strDisplay = populateSearchPage()
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
 

 <h1>VENDOR DETAILS</h1>


<div id="VendorDetails">
<%=strDisplay %>
</div>

	
<% End Sub  %>



<!--#include file="../../designtemplate.asp"-->				

     
