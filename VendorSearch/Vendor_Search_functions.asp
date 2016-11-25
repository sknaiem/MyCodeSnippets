<%
Function GetVendorCategories()
	Dim strSB, strSBC
	strSB = ""
	strSBC = ""
	Set cmd = Server.CreateObject("ADODB.Command")
	With cmd
	   .ActiveConnection = IFAconn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "VendorCategoryLookup" ' Set the name of the Stored Procedure to use   
	   set Recordset = .Execute
	End With
	
	Do UNTIL Recordset.EOF
		strSB = strSB & "<option option='" & Recordset.Fields("CategoryID").value & "'>" & Recordset.Fields("CategoryName").value & "</option>"
		Recordset.MoveNext
	LOOP

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
	   .ActiveConnection = IFAconn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "VendorStateProvLookup" ' Set the name of the Stored Procedure to use   
	   set Recordset = .Execute
	End With

	Do UNTIL Recordset.EOF
		strSB = strSB & "<option option='" & Recordset.Fields("state").value & "'>" & Recordset.Fields("state").value & "</option>"
		Recordset.MoveNext
	Loop

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
	   .ActiveConnection = IFAconn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "VendorCountryLookup" ' Set the name of the Stored Procedure to use   
	   set Recordset = .Execute
	End With

	Do UNTIL Recordset.EOF
		strSB = strSB & "<option option='" & Recordset.Fields("countryname").value & "'>" & Recordset.Fields("countryname").value & "</option>"
		Recordset.MoveNext
	Loop

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
	   .ActiveConnection = IFAconn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "AttorneyPracStateLookup" ' Set the name of the Stored Procedure to use   
	   set Recordset = .Execute
	End With

	Do UNTIL Recordset.EOF
		strSB = strSB & "<option option='"& Recordset.Fields("LocationID").value &"'>" & Recordset.Fields("LocationName").value &"</option>"
		Recordset.MoveNext
	Loop

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
	   .ActiveConnection = IFAconn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "AttorneySpecialityLookup" ' Set the name of the Stored Procedure to use   
	   set Recordset = .Execute
	End With

	DO UNTIL Recordset.EOF
		strSB = strSB & "<option option='" & Recordset.Fields("SpecialtyID").value & "'>"& Recordset.Fields("SpecialtyName").value & "</option>"
		Recordset.MoveNext
	Loop

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
	   .ActiveConnection = IFAconn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "listvendors" ' Set the name of the Stored Procedure to use 
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

	do while not Recordset.EOF
		' ToDo:you got the data now use it to display in the way you want to.
		Recordset.MoveNext
	loop

	DoVendorSearch = strSB
	'Clean up
	set cmd = nothing
	set Recordset = nothing
End Function

'***** nshaik: Loads in an HTML Template that a designer can work on separatly from the programming.
function GetTemplate2(sTemplateFileName)
      dim fso
      dim f
      dim ts
      dim sTemplate
   
      set fso = Server.CreateObject("Scripting.FileSystemObject")
      set f = fso.GetFile(Server.MapPath(TEMPLATE_PATH & sTemplateFileName))
      set ts = f.OpenAsTextStream(ForReading, -2)
      do while not ts.AtEndOfStream
         sTemplate = sTemplate & ts.ReadLine & vbCrLf
      loop
      ts.close
      set ts = nothing
      set f = nothing
      set fso = nothing
      GetTemplate2 = sTemplate
end function 
%>
