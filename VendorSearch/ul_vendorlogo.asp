<!--#include file="../console/filesys.inc"-->

<%
'  Variables
'  *********
   Dim mySmartUpload
   Dim intCount, Userid, sitelocation
   on error resume next        
        sitelocation = trim(request.servervariables("APPL_PHYSICAL_PATH"))	
   if len(sitelocation) > 0 then 'len(Userid) > 0 and 
	   
	'  Object creation
	'  ***************
		Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

		'  Upload
		'  ******
			mySmartUpload.allowedFilesList="jpg,gif"
			mySmartUpload.MaxFileSize = 2000000
			mySmartUpload.Upload
			Userid = mySmartUpload.Form("memberID")
			IF IsNumeric(Userid) THEN
				if not isdir(sitelocation&"vendor_logo\"&Userid) then
					mkdir sitelocation&"vendor_logo\"&Userid
				else
					DeleteFileExtensions  "*.jpg,*.gif", sitelocation&"vendor_logo\"&Userid&"\"
				end if
				
				'  Save the files with their original names in a virtual path of the web server
				'  ****************************************************************************
				intCount = mySmartUpload.Save(sitelocation&"vendor_logo\"&Userid&"\")
			ELSE
				response.Write "Cannot upload your file.Not a valid member"
			END IF
			if err.number <> 0 then
				response.Write "<br>Error Message - " & err.Description 
				err.Clear 
				response.End()
			end if
			' sample with a physical path 
			' intCount = mySmartUpload.Save("c:\temp\")

			'  Display the number of files uploaded
			'  ************************************
			session("vendor_logo")=mySmartUpload.Files("FILE1").FileName
			'Response.Write("Name=" & mySmartUpload.Files("FILE1").FileName)
			Response.Write(intCount & " file(s) uploaded.")
			response.redirect("addLogo_Vendor.asp?ID="&Userid)
		%>
	<script language="JavaScript" type="text/JavaScript">

/*	opener.location.reload("addphoto.asp");
	self.close();*/

				</script>
	<%
	else
		response.Write "Cannot upload your file. Please logon..."
		response.write("userid:"&Userid)
		response.write("sitelocation:"&sitelocation)		
			if err.number<> 0 then
				response.Write "<br>Error Message - " & err.Description 
				err.Clear 
			end if
	end if
'   response.end
 'response.Redirect("filelibrary.asp")
%>