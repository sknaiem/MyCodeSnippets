<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%> 
<% Response.CharSet = "UTF-8" %>
<!--#include file="../console/common.asp"-->
<!--#include file="../console/TL_Reference.asp"-->
<!--#include file="../IFAConnection.asp" -->
<!--#include file="vendor_search_functions.asp" -->
<%
    'MBELCHER: Case 25487 - Administration.

'Response.Write ("Application('fraudconnectionstring') = " &  Application("fraudconnectionstring") & "<br>")
    'loggedin = True
    'http://secure.qa.membershipsoftware.org/ifa/login.asp
If NOT trim(session("loggedin"))="True" Then
	response.redirect trim(application("siteurl"))& "../../login.asp"
End If 

const TEMPLATE_PATH = "templates/"
const ForReading = 1, ForWriting = 2, ForAppending = 8
strSearchTemplate = GetTemplate2("vendor_search.htm")

'________________ REQUEST.FORM ________________
process = trim(Request.Form("process"))
all = trim(Request.Form("all"))
'*** Search Criteria ***
'[Basic Search]
Basic_category = trim(Request.Form("Basic_category")) 
Basic_state = trim(Request.Form("Basic_state")) 
Basic_country = trim(Request.Form("Basic_country")) 
Basic_Keyword_Search = Replace(trim(Request.Form("Basic_Keyword_Search")),"'","''") 
'---
 
'[Attorney Specific Information Search] 
ATTNY_NAME = Replace(trim(Request.Form("ATTNY_NAME")),"'","''") 
ATTNY_FIRM = Replace(trim(Request.Form("ATTNY_FIRM")),"'","''") 
ATTNY_STATE = trim(Request.Form("ATTNY_STATE"))
ATTNY_SPECIALTY = trim(Request.Form("ATTNY_SPECIALTY")) 
'---
 '*** END of Search Criteria ***

'_________________________________________


'******************************* Page Processing '*******************************
IF process = "1" or all = "1" Then
	dim model(8)
	if all = "1" then
		model(0) = "11"		
	else
		model(0) = Basic_category
		model(1) = Basic_state
		model(2) = Basic_country
		model(3) = Basic_Keyword_Search
		model(4) = ATTNY_NAME
		model(5) = ATTNY_FIRM
		model(6) = ATTNY_STATE
		model(7) = ATTNY_SPECIALTY
	end if		
    strSB = DoVendorSearch(model)
	strSB = StrSB & "<hr size='1' align='center'>"
	strSB = strSB & populateSearchPage()
	strDisplay = strSB
ELSE
	strDisplay = populateSearchPage()
END IF
'********************************************************************************
  
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=TOPS_PAGE_HTML_TITLE%></title>
</head>
<body>
<!-- #include file="../TL_Header.asp" -->
<h1>VENDOR SEARCH</h1>
<div id="SearchResults">
<%=strDisplay %>
</div>
<!-- #include file="../TL_Footer.asp" -->
</body>
</html>
