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
	dim count
	
	dim categoryID,stateOrProvince,country,specialtyID,searchText,practicingStateID,attorneyName,attorneyFirm
	dim vendorType, categoryName, state, province, city, companyName,memberId
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
	
	IF model(6) = "-ALL-" THEN
		practicingStateID = null
	ELSE
		practicingStateID = model(6)
	END IF
	specialtyID = model(7)'ToDo: Need to modify
	strSB = ""
	'strSBC = ""
	
	Set cmd = Server.CreateObject("ADODB.Command")
	With cmd
	   .ActiveConnection = IFAconn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "listvendors" ' Set the name of the Stored Procedure to use 
	   .Parameters.Append .CreateParameter("@categoryID",adInteger,adParamInput,,categoryID)
	   .Parameters.Append .CreateParameter("@stateProvince",adVarChar,adParamInput,50,stateOrProvince)
	   .Parameters.Append .CreateParameter("@country",adVarChar,adParamInput,50,country)
	   .Parameters.Append .CreateParameter("@specialtyID",adInteger,adParamInput,,specialtyID)
	   .Parameters.Append .CreateParameter("@searchText",adVarChar,adParamInput,255,searchText)
	   .Parameters.Append .CreateParameter("@practicingStateID",adInteger,adParamInput,,practicingStateID)
	   .Parameters.Append .CreateParameter("@name",adVarChar,adParamInput,255,attorneyName)
	   .Parameters.Append .CreateParameter("@firm",adVarChar,adParamInput,255,attorneyFirm)
	   set Recordset1 = .Execute ' this is for calculating total records count
	   set Recordset = .Execute 
	End With
	
	count = 0
	DO UNTIL Recordset1.EOF
		count = count + 1
		Recordset1.MoveNext
	LOOP
	
	strSB = strSB & "<table>"
	IF count <= 0 THEN ' if results returned then display the following message
		strSB = strSB & "<tr><td>Sorry no vendors matched your search criteria.</td></tr>"
	END IF
	
	DO UNTIL Recordset.EOF		
		vendorType = Recordset.Fields("VendorType").value
		categoryName = Recordset.Fields("CategoryName").value
		country = Recordset.Fields("Country").value
		state = Recordset.Fields("State").value
		province = Recordset.Fields("Province").value
		city = Recordset.Fields("City").value
		companyName = Recordset.Fields("CompanyName").value
		memberId = Recordset.Fields("MemberID").value
		
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
		strSB = strSB & "<div style='padding:6px 0 6px 40px; width:660px;'><a href='Details.asp?ID="&memberId&"'>"&companyName&"</a><br/>"
		IF city <> "" THEN
			strSB = strSB & city & ","
		END IF
		IF state <> "--" THEN
			strSB = strSB & state & ","
		ELSEIF province <> "" THEN
			strSB = strSB & province & ","
		END IF
		IF country <> "" THEN
			strSB = strSB & country
		END IF
		strSB = strSB & "</div>"
		strSB = strSB & "</td></tr>"
		Recordset.MoveNext
	LOOP
	strSB = strSB & "</table>"
	DoVendorSearch = strSB
	'Clean up
	set cmd = nothing
	set Recordset = nothing
	set Recordset1 = nothing
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

Function GetVendorDetails(id)
	dim count, primaryRecordsetCounter, specialtyCounter, practicingStatesCounter
	dim memberType
	dim strSB
	dim logoPath
	dim companyName, address, cityStateZip, country, companyInformation, specialty, contactName, contactEmail,contactPhone, companyFax, companyUrl
	dim specialtyName
	dim practicingState
	dim recommended, recommFirmAttorney, yearsServiceUsed, showRecomInfo, recommComment, recommendedBy
	dim counts(3)
	
	strSB = ""
	logoPath = ""
	companyName = ""
	address=""
	cityStateZip = ""
	country = ""
	companyInformation = ""
	specialty = ""
	contactName = ""
	contactEmail = ""
	contactPhone = ""
	companyFax = ""
	companyUrl = ""
	specialtyName = ""
	practicingState = ""
	recommended =""
	recommFirmAttorney = ""
	yearsServiceUsed = ""
	showRecomInfo = ""
	recommComment = ""
	recommendedBy = ""
	
	Set cmd = Server.CreateObject("ADODB.Command")
	With cmd
	   .ActiveConnection = IFAconn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "viewvendor" ' Set the name of the Stored Procedure to use 
	   .Parameters.Append .CreateParameter("@MemberID",adInteger,adParamInput,,id)
		set Recordset1 = .Execute
	   set Recordset = .Execute
	End With
	count = 1
	primaryRecordsetCounter = 1
	specialtyCounter = 1
	practicingStatesCounter = 1
	
	
	Do Until Recordset1 Is Nothing        
		counts(count-1)= 0
        Do Until Recordset1.EOF
			counts(count-1)= counts(count-1)+1            
            Recordset1.MoveNext
        Loop
        Set Recordset1 = Recordset1.NextRecordset
        count = count + 1
    Loop
	IF counts(0)<=0 THEN ' no data can be found related to vendor
		strSB =  strSB & "<div>Sorry no vendors matched your search criteria.</div>"
	END IF
	count = 1
	Do UNTIL Recordset IS NOTHING
		if count = 1 then			
			DO UNTIL Recordset.EOF
				memberType = Recordset.Fields("MemberType").value
				logoPath = Recordset.Fields("LogoPath").value
				companyName = Recordset.Fields("CompanyName").value
				isPreferred = instr(1,memberType,"Preferred",1)=1
				address = Recordset.Fields("Address").value
				cityStateZip = Recordset.Fields("CityStateZip").value
				country = Recordset.Fields("Country").value
				companyInformation = Recordset.Fields("CompanyInfo").value
				specialty = Recordset.Fields("Specialties").value
				generalComments = Recordset.Fields("CompanyInfo").value
				contactName = Recordset.Fields("ContactName").value
				contactEmail = Recordset.Fields("ContactEmail").value
				contactPhone = Recordset.Fields("ContactPhone").value
				companyFax = Recordset.Fields("CompanyFax").value
				companyUrl = Recordset.Fields("CompanyURL").value
				IF primaryRecordsetCounter = 1 THEN ' no need to show the same data twice
					IF isPreferred = true THEN
						strSB = strSB & "<div><b>IFA Preferred Vendor</b></div><br/>"
					END IF
					IF logoPath <> "" THEN
						strSB = strSB & "<div><img src='"&logoPath&"'/></div><br/>"
					END IF				
					'company name and address
					strSB = strSB & "<div>"
					strSB = strSB & "<b>"&companyName&"</b><br/>"
					strSB = strSB & address&"<br/>"
					strSB = strSB & cityStateZip & "<br/>"
					strSB = strSB & country
					strSB = strSB & "</div>"
					'company information in short
					IF companyInformation <> "" THEN
					strSB = strSB & "<div>"
					strSB = strSB & "<b>Company Info:</b><br/>"
					strSB = strSB & companyInformation
					strSB = strSB & "</div>"
					END IF
					'specialty
					IF specialty <> "" THEN
					strSB = strSB & "<div>"
					strSB = strSB & "<b>Specialty:</b><br/>"
					strSB = strSB & specialty
					strSB = strSB & "</div>"
					END IF
					'General comments
					IF generalComments <> "" THEN
					strSB = strSB & "<div>"
					strSB = strSB & "<b>General Comments:</b><br/>"
					strSB = strSB & companyInformation
					strSB = strSB & "</div>"
					END IF
					
					' members contact information
					strSB = strSB & "<div>"
					strSB = strSB & "<b>Member(s):</b><br/>"
				END IF				
				
					strSB = strSB & contactName & "<br/>"
					strSB = strSB & "Email: "&contactEmail & "<br/>"
					strSB = strSB & "Phone: "&contactPhone & "<br/>"
				
				IF primaryRecordsetCounter = counts(count-1) THEN ' in last iteration close the div
					strSB = strSB & "</div>"
				END IF
				
				IF primaryRecordsetCounter = counts(count-1) AND (companyFax <> "" OR companyUrl <> "")THEN
					'Contact info of the vendor
					strSB = strSB & "<div>"
					strSB = strSB & "<b>Contact Info:</b><br/>"
					IF companyFax <> "" THEN
						strSB = strSB & "Fax: "& companyFax & "<br/>"
					END IF
					IF companyUrl <> "" THEN
						strSB = strSB & "URL: "& companyUrl & "<br/>"
					END IF
					strSB = strSB & "</div>"
				END IF				
				primaryRecordsetCounter = primaryRecordsetCounter + 1
				Recordset.MoveNext
			LOOP			
		END IF
		IF count = 2  THEN			
			DO UNTIL Recordset.EOF
				specialtyName = Recordset.Fields("SpecialtyName").value
				IF counts(count-1) > 0 THEN 'if there are any specialties then only show it.
					IF specialtyCounter = 1 then 'when first record execute this
					strSB = strSB & "<div>"
					strSB = strSB & "<b>Specialties:</b><br/>"
					End if
					
					strSB = strSB & specialtyName & "<br/>"
					
					IF specialtyCounter = counts(count-1) then 'when last record need to execute this
					strSB = strSB & "</div>"				
					END IF
				END IF
				specialtyCounter = specialtyCounter + 1
				Recordset.MoveNext
			LOOP					
		END IF
		IF count = 3 THEN			
			DO UNTIL Recordset.EOF
				practicingState = Recordset.Fields("PracticingState").value
				IF counts(count-1) > 0 THEN ' if there are some practicing state(s) then only show it
					IF practicingStatesCounter = 1 THEN
					strSB = strSB & "<div>"
					strSB = strSB & "<b>Other Locations:</b><br/>"
					END IF					
					
					strSB = strSB & practicingState & "<br/>"
					
					IF practicingStatesCounter = counts(count-1) THEN				
					strSB = strSB & "</div>"
					END IF
				END IF				
				practicingStatesCounter = practicingStatesCounter + 1
				Recordset.MoveNext
			LOOP			
		END IF
		IF count = 4 THEN			
			IF NOT Recordset.EOF THEN
				
				recommended = Recordset.Fields("Recommended").value
				recommFirmAttorney = Recordset.Fields("RecommFirmAtorney").value
				yearsServiceUsed = Recordset.Fields("YearsServiceUsed").value
				showRecomInfo = Recordset.Fields("ShowRecomInfo").value
				recommComment = Recordset.Fields("RecommComment").value
				recommendedBy = Recordset.Fields("RecommendedBy").value
				IF recommended = "Yes" THEN
					
					strSB = strSB & "<div>"					
					'IF showRecomInfo = 1 THEN
						IF recommFirmAttorney <> "" THEN
							strSB = strSB & "<b>Recommended Firm or Attorney:</b><br/>"
							strSB = strSB & recommFirmAttorney & "<br/>"
						END IF
					'END IF
					IF yearsServiceUsed <> "" THEN
						strSB = strSB & "<b>Years Service Used:</b><br/>"
						strSB = strSB & yearsServiceUsed & "<br/>"
					END IF
					IF comments <> "" THEN
						strSB = strSB & "<b>Comments:</b><br/>"
						strSB = strSB & recommComment & "<br/>"
					END IF
					IF recommendedBy <> "" THEN
						strSB = strSB & "<b>Recommended By:</b><br/>"
						strSB = strSB & recommendedBy & "<br/>"
					END IF
					strSB = strSB & "</div>"
					
				END IF				
				Recordset.MoveNext
			END IF			
		END IF		
		count = count + 1
		Set Recordset = Recordset.NextRecordset
	LOOP	
	GetVendorDetails = strSB
End Function

'***** Loads in an HTML Template that a designer can work on separatly from the programming.
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
