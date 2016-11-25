<%
Function GetVendorCategories()
	Dim strSB, strSBC
	strSB = ""
	strSBC = ""
	Set cmd = Server.CreateObject("ADODB.Command")
	With cmd
	   .ActiveConnection = CustomConn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "VendorCategoryLookup" ' Set the name of the Stored Procedure to use   
	   set Recordset = .Execute
	End With

	do until Recordset.EOF
		strSB = strSB & "<option option='"&Recordset["CategoryID"]&"'>"&Recordset["CategoryName"]&"</option>"
	loop

	GetVendorCategories = strSB
	'Clean up
	set cmd = nothing
	set Recordset = nothing
End Function

Function GetStatesOrProvince()
	Dim strSB, strSBC
	strSB = ""
	strSBC = ""
	Set cmd = Server.CreateObject("ADODB.Command")
	With cmd
	   .ActiveConnection = CustomConn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "VendorStateProvLookup" ' Set the name of the Stored Procedure to use   
	   set Recordset = .Execute
	End With

	do until Recordset.EOF
		strSB = strSB & "<option option='"&Recordset["state"]&"'>"&Recordset["state"]&"</option>"
	loop

	GetStatesOrProvince = strSB
	'Clean up
	set cmd = nothing
	set Recordset = nothing
End Function

Function GetCountries()
	Dim strSB, strSBC
	strSB = ""
	strSBC = ""
	Set cmd = Server.CreateObject("ADODB.Command")
	With cmd
	   .ActiveConnection = CustomConn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "VendorCountryLookup" ' Set the name of the Stored Procedure to use   
	   set Recordset = .Execute
	End With

	do until Recordset.EOF
		strSB = strSB & "<option option='"&Recordset["countryname"]&"'>"&Recordset["countryname"]&"</option>"
	loop

	GetCountries = strSB
	'Clean up
	set cmd = nothing
	set Recordset = nothing
End Function

Function GetAttorneyPracticingStates()
	Dim strSB, strSBC
	strSB = ""
	strSBC = ""
	Set cmd = Server.CreateObject("ADODB.Command")
	With cmd
	   .ActiveConnection = CustomConn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "AttorneyPracStateLookup" ' Set the name of the Stored Procedure to use   
	   set Recordset = .Execute
	End With

	do until Recordset.EOF
		strSB = strSB & "<option option='"&Recordset["LaocationID"]&"'>"&Recordset["LocationName"]&"</option>"
	loop

	GetAttorneyPracticingStates = strSB
	'Clean up
	set cmd = nothing
	set Recordset = nothing
End Function

Function GetAttorneySpecilaties()
	Dim strSB, strSBC
	strSB = ""
	strSBC = ""
	Set cmd = Server.CreateObject("ADODB.Command")
	With cmd
	   .ActiveConnection = CustomConn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "AttorneySpecialityLookup" ' Set the name of the Stored Procedure to use   
	   set Recordset = .Execute
	End With

	do until Recordset.EOF
		strSB = strSB & "<option option='"&Recordset["SpecialtyID"]&"'>"&Recordset["SpecialtyName"]&"</option>"
	loop

	GetAttorneySpecilaties = strSB
	'Clean up
	set cmd = nothing
	set Recordset = nothing
End Function

' it does the vendor search
' model will have categoryid, stateorprovince, country, specialtyId, searchText, PracticingStateID, name, firm
Function DoVendorSearch(model)
	Dim strSB, strSBC
	strSB = ""
	strSBC = ""
	Set cmd = Server.CreateObject("ADODB.Command")
	With cmd
	   .ActiveConnection = CustomConn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "AttorneySpecialityLookup" ' Set the name of the Stored Procedure to use 
	   .Parameters.Append .CreateParameter("@categoryID",adInteger,adParamInput,,categoryID)
	   .Parameters.Append .CreateParameter("@stateProvince",adVarChar,adParamInput,7,stateOrProvince)
	   .Parameters.Append .CreateParameter("@country",adVarChar,adParamInput,50,country)
	   .Parameters.Append .CreateParameter("@specialtyID",adInteger,adParamInput,,specialtyID)
	   .Parameters.Append .CreateParameter("@searchText",adVarChar,adParamInput,255,searchText)
	   .Parameters.Append .CreateParameter("@practicingStateID",adVarChar,adParamInput,5,practicingStateID)
	   .Parameters.Append .CreateParameter("@name",adVarChar,adParamInput,255,attorneyName)
	   .Parameters.Append .CreateParameter("@firm",adVarChar,adParamInput,255,attorneyFirm)
	   set Recordset = .Execute
	End With

	do until Recordset.EOF
		' ToDo:you got the data now use it to display in the way you want to.
	loop

	DoVendorSearch = strSB
	'Clean up
	set cmd = nothing
	set Recordset = nothing
End Function
%>
