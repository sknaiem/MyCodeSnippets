<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<!--#include file="../console/common.asp"-->
<!--#include file="../console/TL_Reference.asp"-->
<!--#include file="../IFAConnection.asp" -->
<!--#include file="../console/Connection.asp" -->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%showcmn "Main Page Title Tag"%></title>
<link href="css/popup.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--

function PB(desktopURL)
{
 var desktop = window.open( desktopURL, "multiPopup", "toolbar=no,location=no,status=no,scrollbars=yes,menubar=no,width=370,height=200,resizable=yes" );
}

//-->
</script>
<body bgcolor="#FFFFFF">
<%
memberID = trim(Request.QueryString("ID"))
if session("vendor_logo")<>"" and request("addphoto")<>1 then
	Set cmd = Server.CreateObject("ADODB.Command")
	With cmd
	   .ActiveConnection = IFAconn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "SaveVendorLogo" ' Set the name of the Stored Procedure to use 
	   .Parameters.Append .CreateParameter("@MemberID",adInteger,adParamInput,,memberID)
	   .Parameters.Append .CreateParameter("@LogoPath",adVarChar,adParamInput,255,Session("vendor_logo"))	   
	   .Execute
	End With
	conn.execute "insert into User_Modify_Log (userid,field_modified,previous_value,current_value,modify_date,modify_by) values("&memberID&",'Photo_Link','Null','"&session("vendor_logo")&"', '"&now()&"','"&memberID&"')"
%>
	<script language="JavaScript" type="text/JavaScript">
	self.close();
	opener.location.reload(true);
	</script>
<%
else%>
<table>
        <tr>
      <td valign="center" style="padding-left:10px"><h5>Vendor Photo Management</h5></td>
    </tr>
    
     <TR><TD style="padding-left:10px">

   <div class=content>
	Your logo or photo can be uploaded and displayed next to the vendor 
	information in the online directory.<br>
	<br>
	1) Photos must be saved as a .jpg file format<br>
	2) Photos must be 72dpi and in RGB color profile<br>
	3) The filename must be short in length, have no spaces, and  contain numbers and/or letters only<br>
	4) The file can be no larger than 1mb<br> <br>
	</div></TD></TR>
	<FORM method="POST" action="ul_vendorlogo.asp" enctype="multipart/form-data" id="Form4" >
	<TR><TD style="padding-left:10px">	
  <INPUT type="FILE" name="FILE1" size="40" id="File2"><br /><Br /></TD></TR>
  <TR><TD style="padding-left:10px">
	<INPUT type="hidden" name="memberID" value="<%=memberID%>" />
   <INPUT type="submit" name="submit" value="Upload Photo" border="0" >
  </TD></TR>
    </form>
</div>
  </table>
<%end if%>
</body>
<script type="text/javascript">
var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>
<script type="text/javascript">
try {
var pageTracker = _gat._getTracker("UA-4361821-1");
pageTracker._trackPageview();
} catch(err) {}</script>
</html>
