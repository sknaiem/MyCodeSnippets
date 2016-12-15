<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%> 
<% Response.CharSet = "UTF-8" %>
<!--#include file="../../console/common.asp"-->
<!--#include file="../../console/TL_Reference.asp"-->

<%
    'MBELCHER: Case 25487 - Administration.
  
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=TOPS_PAGE_HTML_TITLE%></title>

<link href="gstablestyle.css" rel="stylesheet" type="text/css" />

</head>

<body>
<!-- #include file="TL_Header.asp" -->

<table class="gs-table">
	<tbody>
    	<tr class="gs-title">
        	<td colspan="4">
            	<h1>IFA custom admin page</h1>
            </td>
        </tr>
        <tr class="gs-subheading">
        	<td colspan="4">
            	<h2></h2>
            </td>
        </tr>
        <tr class="gs-instructions">
        	<td colspan="4">
            </td>
        </tr>
       	<tr>
        	<td colspan="4"> 
                
                <ul>
                    <li><a href="admin_Fraud.asp">FRAUD WATCH LIST</a></li>
                    <li><a href="admin_FactorSearch.asp">FACTOR SEARCH</a></li>
                    <li><a href="../../Vendor/Search.asp">VENDOR SEARCH</a></li>
                    <li><a target="_blank" href="../../console/maskedlogin.asp?userid=3">Timberlake Public Test Login</a></li>
                </ul>
                
                       
            </td>
        </tr>
        

        
        
    </tbody>
</table>
<% 
'End IF 
%>
<!-- #include file="TL_Footer.asp" -->












</body>
</html>
