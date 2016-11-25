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
<!--#include file="vendor_search_functions.asp" -->


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
process = trim(Request.Form("process"))

'*** Search Criteria ***
'[Basic Search]
Basic_category = trim(Request.Form("Basic_category")) 
Basic_state = trim(Request.Form("Basic_state")) 
Basic_country = trim(Request.Form("Basic_country")) 
Basic_Keyword_Search = trim(Request.Form("Basic_Keyword_Search")) 
'---
 
'[Attorney Specific Information Search] 
ATTNY_NAME = trim(Request.Form("ATTNY_NAME")) 
ATTNY_FIRM = trim(Request.Form("ATTNY_FIRM")) 
ATTNY_STATE = trim(Request.Form("ATTNY_STATE"))
ATTNY_SPECIALTY = trim(Request.Form("ATTNY_SPECIALTY")) 
'---
 '*** END of Search Criteria ***

'_________________________________________



    


'******************************* Page Processing '*******************************
IF process = "1" Then
    TheModel = ""' this maybe be a small set of records or a view that is passed into GET_ShowConformation to be processed; not sure yet.
    'Our form has submitted lets do something here, clean up, validate a little or what ever.

    'Call the processing function
    'Call SomeFunctionName (strSomeVariableName, strSomeVariableName, strSomeVariableName)
    'CALL Calculate_VendorSearch (TheModel)

    'Need somthing here or can do in the function if no records are found.

    'Processed successfuly and now we display a friendsly message. This could be a hidden div that we later make visible or just load a template into display results.
    strDisplay = " We have posted back and all is well"

    'If all goes well we then display results.
    'strDisplay = strConformationTemplate
    'strDisplay= Calculate_VendorSearch ()

ELSE
    'can load in template of form here.
    'strDisplay = " We display the form here."
    'We might have to load in catagories and other things here; not sure yet.
	vendorCategories = GetVendorCategories()
	stateOrProvince = GetStatesOrProvince()
	countries = GetStatesOrProvince()
	attorneyPracticingStates = GetAttorneyPracticingStates()
	attorneySpecialties = GetAttorneySpecilaties()
    strSB = strSearchTemplate
	strSB = replace(strSB,"[CategoryOptions]",vendorCategories)
	strSB = replace(strSB,"[StateProvinceOptions]",stateOrProvince)
	strSB = replace(strSB,"[CountryOptions]",countries)
	strSB = replace(strSB,"[AttorneyPracticingStateOptions]",attorneyPracticingStates)
	strSB = replace(strSB,"[AttorneySpecialtiesOptions]",attorneySpecialties)
	strDisplay = strSB
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

<!-- MBELCHER: This is used for our fraud record template -->
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
 

 <h1>VENDOR SEARCH</h1>


<div id="SearchResults">
<%=strDisplay %>
</div>

	
<% End Sub  %>



<!--#include file="../../designtemplate.asp"-->				

     
