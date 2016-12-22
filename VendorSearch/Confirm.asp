<h2>Confirmation</h2>
<%
isSucceeded = Request.QueryString("success")
IF isSucceeded THEN
	Response.Write("Successfully saved")
ELSE
	Response.Write("Saving of vendor failed")
END IF
%>