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
		strSB = strSB & "<option value='" & Recordset.Fields("CategoryID").value & "'>" & Recordset.Fields("CategoryName").value & "</option>"
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
		strSB = strSB & "<option value='" & Recordset.Fields("state").value & "'>" & Recordset.Fields("state").value & "</option>"
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
		strSB = strSB & "<option value='" & Recordset.Fields("countryname").value & "'>" & Recordset.Fields("countryname").value & "</option>"
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
		strSB = strSB & "<option value='"& Recordset.Fields("LocationID").value &"'>" & Recordset.Fields("LocationName").value &"</option>"
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
		strSB = strSB & "<option value='" & Recordset.Fields("SpecialtyID").value & "'>"& Recordset.Fields("SpecialtyName").value & "</option>"
		Recordset.MoveNext
	Loop

	GetAttorneySpecilaties = strSB
	'Clean up
	set cmd = nothing
	set Recordset = nothing
End Function

' it does the vendor search
' model will have categoryid, stateorprovince, country, specialtyId, searchText, PracticingStateID, name, firm
Function DoVendorSearch()
	dim isVendorSet
	dim isPreferredVendorSet
	dim prevCategory
	dim prevCountry
	dim prevState
	dim strSB
	'Dim strSBC
	dim categoryID,stateOrProvince,country,specialtyID,searchText,practicingStateID,attorneyName,attorneyFirm
	dim vendorType, categoryName, state, city, companyName
	isVendorSet = false
	isPreferredVendorSet = false
	prevCategory = ""
	prevCountry = ""
	prevState = ""
	strSB = ""
	categoryID = 8
	stateOrProvince = ""
	country = ""	
	searchText = ""
	practicingStateID = ""
	attorneyName = ""
	attorneyFirm = ""	
	strSB = ""
	'strSBC = ""
	
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
	
	strSB = strSB & "<table>"
	
	DO UNTIL Recordset.EOF		
		vendorType = Recordset.Fields("VendorType").value
		categoryName = Recordset.Fields("CategoryName").value
		country = Recordset.Fields("Country").value
		state = Recordset.Fields("State").value
		city = Recordset.Fields("City").value
		companyName = Recordset.Fields("CompanyName").value
		
		if vendorType = "Preferred Vendor" and isPreferredVendorSet = false then
			strSB = strSB & "<tr style='background: #2659B6;'><td>IFA Preferred Vendor</td></tr>"
			isPreferredVendorSet = true
		elseif vendorType = "Vendor" and isVendorSet = false then
			strSB = strSB & "<tr style='background: #2659B6;'><td>IFA Vendor</td></tr>"
			isVendorSet = true
		end if
		if prevCategory = "" or (prevCategory <> "" and categoryName <> prevCategory) then
			strSB = strSB & "<tr style='background: tan'><td>"&categoryName&"</td></tr>"
			prevCategory = categoryName
		end if
		if prevCountry = "" or (prevCountry <> "" and country <> prevCountry) then
			strSB = strSB & "<tr><td style='padding:4px;'>"&country&"</td></tr>"
			prevCountry = country
		end if
		if prevState = "" then
		strSB = strSB & "<tr><td style='padding-top:10px;'>"&state&"</td></tr><tr><td>"
		prevState = state
		elseif prevState <> "" and state <> prevState then
		strSB = strSB & "</td></tr><tr><td style='padding-top:10px;'>"&state&"</td></tr><tr><td>"
		prevState = state
		end if
		strSB = strSB & "<div style='padding:6px 0 6px 40px; width:660px;'><a>"&companyName&"</a><br/>"
		strSB = strSB & city&","&state&","&country&"</div>"
		Recordset.MoveNext
	LOOP
	strSB = strSB & "</table>"
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
