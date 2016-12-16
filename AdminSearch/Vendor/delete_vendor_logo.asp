<!--#include file="../console/filesys.inc"-->
<!--#include file="../console/common.asp"-->
<!--#include file="../console/TL_Reference.asp"-->
<!--#include file="../console/Connection.asp" -->

<%
	Userid =trim(Request.QueryString("ID"))
	'Userid= session("sa_id")
	sitelocation = trim(request.servervariables("APPL_PHYSICAL_PATH"))
	rmdir sitelocation&"vendor_logo\"&Userid
		
	set rsphotoname=getrecordset("select * from af_members where userid="&Userid)
	photoname=rsphotoname("Photo_Link")
	'makeconnection

	conn.execute "update af_members set photo_link=Null, photo_status=null where userid="&Userid
	conn.execute "insert into User_Modify_Log (userid,field_modified,previous_value,current_value,modify_date,modify_by) values("&UserId&",'Photo_Link','"&photoname&"','Null', '"&now()&"','"&Userid&"')"
%>
<script language="JavaScript" type="text/JavaScript">

	opener.location.reload(true);
	self.close();

</script>