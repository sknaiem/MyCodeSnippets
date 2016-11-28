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
Function DoVendorSearch(model)
	dim isVendorSet
	dim isPreferredVendorSet	
	dim prevCategory	
	dim prevCountry	
	dim prevStateOrProvince
	dim strSB
	
	dim categoryID,stateOrProvince,country,specialtyID,searchText,practicingStateID,attorneyName,attorneyFirm
	dim vendorType, categoryName, state, province, city, companyName
	isVendorSet = false
	isPreferredVendorSet = false
	'prevPrefCategory = ""
	prevCategory = ""
	'prevPrefCountry = ""
	prevCountry = ""
	'prevPrefStateOrProvince = ""
	prevStateOrProvince = ""
	strSB = ""
	categoryID = model(0)
	stateOrProvince = model(1)
	country = model(2)
	searchText = model(3)
	attorneyName = model(4)
	attorneyFirm = model(5)
	practicingStateID = ""'model(6) 'ToDo: Need to modify
	specialtyID = 0'model(7)'ToDo: Need to modify
	strSB = ""
	'strSBC = ""
	
	Set cmd = Server.CreateObject("ADODB.Command")
	With cmd
	   .ActiveConnection = IFAconn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "listvendors" ' Set the name of the Stored Procedure to use 
	   .Parameters.Append .CreateParameter("@categoryID",adInteger,adParamInput,,categoryID)
	   .Parameters.Append .CreateParameter("@stateProvince",adVarChar,adParamInput,8,stateOrProvince)
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
		province = Recordset.Fields("Province").value
		city = Recordset.Fields("City").value
		companyName = Recordset.Fields("CompanyName").value
		
		if vendorType = "Preferred Vendor" and isPreferredVendorSet = false then
			strSB = strSB & "<tr style='background: #2659B6;'><td>IFA Preferred Vendor</td></tr>"
			isPreferredVendorSet = true
			prevCategory = ""'reset the previous category everytime vendor type changes
		end if
		if vendorType = "Vendor" and isVendorSet = false then
			strSB = strSB & "<tr style='background: #2659B6;'><td>IFA Vendor</td></tr>"
			isVendorSet = true
			prevCategory = ""'reset the previous category everytime vendor type changes
		end if
		'within  vendor type vendorcategories should not repeat
			if prevCategory = "" or (prevCategory <> "" and categoryName <> prevCategory) then
				if categoryName <> "" then
					strSB = strSB & "<tr style='background: tan'><td>"&categoryName&"</td></tr>"
					prevCategory = categoryName					
				end if
				prevCountry = ""'reset the previous preferred country everytime the category is changed
			end if
		' within vendorcategory countries should not repeat
				if prevCountry = "" or (prevCountry <> "" and country <> prevCountry) then
					strSB = strSB & "<tr><td style='padding:4px;'><h2>"&country&"</h2></td></tr>"
					prevCountry = country
					prevStateOrProvince=""
				end if
		'within a country the state or province name should not repeat
					if prevStateOrProvince = "" then
						if state <> "--" then
							strSB = strSB & "<tr><td style='padding-top:10px;'>"&state&"</td></tr>"
							prevStateOrProvince = state
						elseif province <> "" then
							strSB = strSB & "<tr><td style='padding-top:10px;'>"&province&"</td></tr>"
							prevStateOrProvince = province
						end if			
					elseif prevStateOrProvince <> "" and state <> prevStateOrProvince then
						if state <> "--" then
							strSB = strSB & "<tr><td style='padding-top:10px;'>"&state&"</td></tr>"
							prevStateOrProvince = state
						elseif province <> "" and state <> prevStateOrProvince then
							strSB = strSB & "<tr><td style='padding-top:10px;'>"&province&"</td></tr>"
							prevStateOrProvince = province
						end if			
					end if
		strSB = strSB & "<tr><td>"
		strSB = strSB & "<div style='padding:6px 0 6px 40px; width:660px;'><a>"&companyName&"</a><br/>"
		strSB = strSB & city&","&state&","&country&"</div>"
		strSB = strSB & "</td></tr>"
		Recordset.MoveNext
	LOOP
	strSB = strSB & "</table>"
	DoVendorSearch = strSB
	'Clean up
	set cmd = nothing
	set Recordset = nothing
End Function

function populateSearchPage()
	dim vendorCategories, stateOrProvince, countries, attorneyPracticingStates, attorneySpecialties,strSB
	vendorCategories = GetVendorCategories()
	stateOrProvince = GetStatesOrProvince()
	countries = GetCountries()
	attorneyPracticingStates = GetAttorneyPracticingStates()
	attorneySpecialties = GetAttorneySpecilaties()
    strSB = strSearchTemplate
	strSB = replace(strSB,"[CategoryOptions]",vendorCategories)
	strSB = replace(strSB,"[StateProvinceOptions]",stateOrProvince)
	strSB = replace(strSB,"[CountryOptions]",countries)
	strSB = replace(strSB,"[AttorneyPracticingStateOptions]",attorneyPracticingStates)
	strSB = replace(strSB,"[AttorneySpecialtiesOptions]",attorneySpecialties)
	populateSearchPage = strSB
end function

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
