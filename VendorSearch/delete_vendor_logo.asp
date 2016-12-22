<!--#include file="../includes/filesys.inc"-->
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


<%	
	Userid= session("sa_id")
	sitelocation = trim(application("consoleurl"))
	dim noLogo
	noLogo = false
	
	IF IsNumeric(Userid) THEN
		rmdir sitelocation&"vendor_logo\"&Userid
		Set cmd = Server.CreateObject("ADODB.Command")
		With cmd
		   .ActiveConnection = IFAconn 
		   .CommandType = adCmdStoredProc
		   .CommandText = "viewvendor" ' Set the name of the Stored Procedure to use 
		   .Parameters.Append .CreateParameter("@MemberID",adInteger,adParamInput,,UserId)			
		   set Recordset = .Execute
		End With
		IF NOT Recordset.EOF THEN
			logoFileName = Recordset("LogoPath")
			if isnull(logoFileName) or trim(logoFileName)="" then
			  noLogo=true
			end if
			if NOT noLogo then				
				' imagePath = trim(application("consoleurl"))&"/vendor_logo/"&memberId&"/"&trim(logoFileName)
				' strSB = strSB & "<div class='LeftControl'><img class='vendor_logo' src='"&imagePath&"' style='max-height:200px;max-width:200px;'/></div>"
				IFAconn.execute "UPDATE VendorData SET LogoPath=NULL WHERE MemberID="&UserId				
			end if
		END IF
	END IF		
	
	'conn.execute "insert into User_Modify_Log (userid,field_modified,previous_value,current_value,modify_date,modify_by) values("&UserId&",'Photo_Link','"&photoname&"','Null', '"&now()&"','"&Userid&"')"
%>
<script language="JavaScript" type="text/JavaScript">

	opener.location.reload(true);
	self.close();

</script>