<%
Function GetVendorCategories(categorySelected)
	Dim strSB 
	strSB = ""	
	categoryID = ""
	Set cmd = Server.CreateObject("ADODB.Command")
	With cmd
	   .ActiveConnection = IFAconn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "VendorCategoryLookup" ' Set the name of the Stored Procedure to use   
	   set Recordset = .Execute
	End With
	
	Do UNTIL Recordset.EOF
		categoryID = Recordset.Fields("CategoryID").value
		IF (categorySelected <> "" OR categorySelected <> NULL) AND categorySelected = categoryID THEN
			strSB = strSB & "<option value='" & categoryID & "' selected>" & Recordset.Fields("CategoryName").value & "</option>"
		ELSE
		strSB = strSB & "<option value='" & categoryID & "'>" & Recordset.Fields("CategoryName").value & "</option>"
		END IF
		Recordset.MoveNext
	LOOP

	GetVendorCategories = strSB
	'Clean up
	set cmd = nothing
	set Recordset = nothing
End Function

Function GetStatesOrProvince()
	Dim strSB
	strSB = ""	
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
	Dim strSB
	strSB = ""
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

Function GetAttorneyPracticingStates(arrPracStates)
	Dim strSB
	strSB = ""
	
	Set cmd = Server.CreateObject("ADODB.Command")
	With cmd
	   .ActiveConnection = IFAconn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "AttorneyPracStateLookup" ' Set the name of the Stored Procedure to use   
	   set Recordset = .Execute
	End With
	IF IsArray(arrPracStates) THEN
		Do UNTIL Recordset.EOF
			attorneyLocationName = Recordset.Fields("LocationName").value
			attorneyLocationID = Recordset.Fields("LocationID").value
			practStateAdded = false
			FOR EACH practState IN arrPracStates			
				IF practState = attorneyLocationID THEN
					strSB = strSB & "<option value='"& attorneyLocationID &"' selected>" & attorneyLocationName &"</option>"
					practStateAdded = true				
				END IF							
			NEXT
			IF practStateAdded = false THEN
				strSB = strSB & "<option value='"& attorneyLocationID &"'>" & attorneyLocationName &"</option>"				
			END IF
			Recordset.MoveNext
		Loop
	ELSE
		Do UNTIL Recordset.EOF
			strSB = strSB & "<option value='"& Recordset.Fields("LocationID").value &"'>" & Recordset.Fields("LocationName").value &"</option>"
			Recordset.MoveNext
		Loop
	END IF

	GetAttorneyPracticingStates = strSB
	'Clean up
	set cmd = nothing
	set Recordset = nothing
End Function

Function GetAttorneySpecialties(arrSpecialties)
	Dim strSB
	strSB = ""
	
	Set cmd = Server.CreateObject("ADODB.Command")
	With cmd
	   .ActiveConnection = IFAconn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "AttorneySpecialityLookup" ' Set the name of the Stored Procedure to use   
	   set Recordset = .Execute
	End With
	IF IsArray(arrSpecialties) THEN ' TODO: if array is empty then this will fail
		DO UNTIL Recordset.EOF
			specialtyName = Recordset.Fields("SpecialtyName").value
			specialtyID = Recordset.Fields("SpecialtyID").value
			specialtyAdded = false
			For Each specialty IN arrSpecialties
				IF specialty = specialtyID THEN
					strSB = strSB & "<option value='" & specialtyID & "' selected>"& specialtyName & "</option>"
					specialtyAdded = true
				END IF				
			Next
			IF specialtyAdded = false Then 
				strSB = strSB & "<option value='" & specialtyID & "'>"& specialtyName & "</option>"
			End If
			Recordset.MoveNext
		Loop
	ELSE
		DO UNTIL Recordset.EOF				
			strSB = strSB & "<option value='" & Recordset.Fields("SpecialtyID").value & "'>"& Recordset.Fields("SpecialtyName").value & "</option>"				
			Recordset.MoveNext
		Loop
	END IF

	GetAttorneySpecialties = strSB
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
	   .Parameters.Append .CreateParameter("@searchText",adVarChar,adParamInput,100,searchText)
	   .Parameters.Append .CreateParameter("@practicingStateID",adInteger,adParamInput,,practicingStateID)
	   .Parameters.Append .CreateParameter("@name",adVarChar,adParamInput,100,attorneyName)
	   .Parameters.Append .CreateParameter("@firm",adVarChar,adParamInput,100,attorneyFirm)
	   set Recordset1 = .Execute ' this is for calculating total records count
	   set Recordset = .Execute 
	End With
	
	count = 0
	DO UNTIL Recordset1.EOF
		count = count + 1
		Recordset1.MoveNext
	LOOP
	
	strSB = strSB & "<table class='vendor_searchresults'>"
	IF count <= 0 THEN ' if results returned then display the following message
		strSB = strSB & "<tr class='nodata'><td>Sorry no vendors matched your search criteria.</td></tr>"
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
			strSB = strSB & "<tr class='vendor_heading'><td>IFA Preferred Vendor</td></tr>"
			isPreferredVendorSet = true
			prevCategory = ""'reset the previous category everytime vendor type changes
		end if
		if vendorType = "Vendor" and isVendorSet = false then
			strSB = strSB & "<tr class='vendor_heading'><td>IFA Vendor</td></tr>"
			isVendorSet = true
			prevCategory = ""'reset the previous category everytime vendor type changes
		end if
		'within  vendor type vendorcategories should not repeat
			if prevCategory = "" or (prevCategory <> "" and categoryName <> prevCategory) then
				if categoryName <> "" then
					strSB = strSB & "<tr class='vendor_category'><td>"&categoryName&"</td></tr>"
					prevCategory = categoryName					
				end if
				prevCountry = ""'reset the previous preferred country everytime the category is changed
			end if
		' within vendorcategory countries should not repeat
				if prevCountry = "" or (prevCountry <> "" and country <> prevCountry) then
					strSB = strSB & "<tr><td class='vendor_country'><h2>"&country&"</h2></td></tr>"
					prevCountry = country
					prevStateOrProvince=""
				end if
		'within a country the state or province name should not repeat
					if prevStateOrProvince = "" then
						if state <> "--" then
							strSB = strSB & "<tr><td class='vendor_state'>"&state&"</td></tr>"
							prevStateOrProvince = state
						elseif province <> "" then
							strSB = strSB & "<tr><td class='vendor_state'>"&province&"</td></tr>"
							prevStateOrProvince = province
						end if			
					elseif prevStateOrProvince <> "" and state <> prevStateOrProvince then
						if state <> "--" then
							strSB = strSB & "<tr><td class='vendor_state'>"&state&"</td></tr>"
							prevStateOrProvince = state
						elseif province <> "" and province <> prevStateOrProvince then
							strSB = strSB & "<tr><td class='vendor_state'>"&province&"</td></tr>"
							prevStateOrProvince = province
						end if			
					end if
		strSB = strSB & "<tr><td>"
		strSB = strSB & "<div class='vendor_info'><a href='Details.asp?ID="&memberId&"'>"&companyName&"</a><br/>"
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
	vendorCategories = GetVendorCategories("")
	stateOrProvince = GetStatesOrProvince()
	countries = GetCountries()
	attorneyPracticingStates = GetAttorneyPracticingStates("")
	attorneySpecialties = GetAttorneySpecialties("")
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
				categoryID = Recordset.Fields("CategoryID").value
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
				imagePath = trim(application("consoleurl"))&"/vendor_logo/"&id&"/"&trim(logoPath)
				IF primaryRecordsetCounter = 1 THEN ' no need to show the same data twice
					IF isPreferred = true THEN
						strSB = strSB & "<div class='vendor_preferred'>IFA Preferred Vendor</div><br/>"
					END IF
					IF logoPath <> "" THEN
						strSB = strSB & "<div class='vendor_logo'><img src='"&imagePath&"' style='max-height:200px;max-width:200px;'/></div><br/>"
					END IF				
					'company name and address
					strSB = strSB & "<div class='vendor_demographics'>"
					strSB = strSB & "<b>"&companyName&"</b><br/>"
					strSB = strSB & address&"<br/>"
					strSB = strSB & cityStateZip & "<br/>"
					strSB = strSB & country
					strSB = strSB & "</div>"
					'company information in short
					IF companyInformation <> "" THEN
					strSB = strSB & "<div class='vendor_info'>"
					strSB = strSB & "<b>Company Info:</b><br/>"
					strSB = strSB & companyInformation
					strSB = strSB & "</div>"
					END IF
					'specialty
					IF specialty <> "" THEN
					strSB = strSB & "<div class='vendor_spacialty'>"
					strSB = strSB & "<b>Specialty:</b><br/>"
					strSB = strSB & specialty
					strSB = strSB & "</div>"
					END IF
					'General comments
					IF generalComments <> "" THEN
					strSB = strSB & "<div class='vendor_comments'>"
					strSB = strSB & "<b>General Comments:</b><br/>"
					strSB = strSB & generalComments
					strSB = strSB & "</div>"
					END IF
					
					' members contact information
					strSB = strSB & "<div class='vendor_members'>"
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
					strSB = strSB & "<div class='vendor_contactinfo'>"
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
					strSB = strSB & "<div class='attorney_specialties'>"
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
					strSB = strSB & "<div class='attorney_practicing_states'>"
					strSB = strSB & "<b>Practicing State(s):</b><br/>"
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
				
				'recommended = Recordset.Fields("Recommended").value
				recommFirmAttorney = Recordset.Fields("RecommFirmAtorney").value
				yearsServiceUsed = Recordset.Fields("YearsServiceUsed").value
				showRecomInfo = Recordset.Fields("ShowRecomInfo").value
				recommComment = Recordset.Fields("RecommComment").value
				recommendedBy = Recordset.Fields("RecommendedBy").value
				'IF recommended = "Yes" THEN
				IF categoryID = "8"  THEN ' AND showRecomInfo = 1 TODO: this needs to be displayed based on the showrecommendation flag
					strSB = strSB & "<div class='attorney_recommendation'>"					
					'IF recommended = "Yes" THEN
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

Function GetVendorDetailsForEdit(id)
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
	specialtyCounter = 0
	practicingStatesCounter = 0
	
	
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
				categoryID = Recordset.Fields("CategoryID").value
				isAttorney = categoryID = "8"
				IF primaryRecordsetCounter = 1 THEN ' no need to show the same data twice
					IF isPreferred = true THEN
						strSB = strSB & "<div class='vendor_preferred'>"&memberType&"</div><br/>"
						strSB = strSB & GetLogoUploadControl(id)						
					END IF
					strSB = strSB & "<div class='row'>"					
					strSB = strSB & "<span class='LeftControl'>Category:</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & "<select name='Basic_category'>"
					strSB = strSB & GetVendorCategories(categoryID)'[VENDOR_CATEGORY]
					strSB = strSB & "</select>"
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					strSB = strSB & "<div class='row'>"
					strSB = strSB & "<span class='LeftControl'>Company:</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & companyName
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					strSB = strSB & "<div class='row'>"
					strSB = strSB & "<span class='LeftControl'>Address:</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & address'"<input type='text' name='txtAddress' value='"&address&"' />"
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					strSB = strSB & "<div class='row'>"
					strSB = strSB & "<span class='LeftControl'>City, State, Zip:</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & cityStateZip'"<input type='text' name='txtCityStateZip' value='"&cityStateZip&"' />"
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					strSB = strSB & "<div class='row'>"
					strSB = strSB & "<span class='LeftControl'>Country:</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & country'"<input type='text' name='txtCountry' value='"&country&"' />"
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					strSB = strSB & "<div class='row'>"
					strSB = strSB & "<span class='LeftControl'>Company Information:</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & companyInformation'"<input type='text' name='txtCompanyInformation' value='"&companyInformation&"' />"
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					strSB = strSB & "<div class='row'>"
					strSB = strSB & "<span class='LeftControl'>Specialty:</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & specialty'"<input type='text' name='txtSpecialty' value='"&specialty&"' />"
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					strSB = strSB & "<div class='row'>"
					strSB = strSB & "<span class='LeftControl'>General Comments:</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & generalComments'"<input type='text' name='txtGeneralComments' value='"&generalComments&"' />"
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					strSB = strSB & "<div class='row'>"
					strSB = strSB & "<span class='LeftControl'>Members(S):</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & contactName
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					strSB = strSB & "<div class='row'>"
					strSB = strSB & "<span class='LeftControl'>Contact Email:</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & contactEmail'"<input type='text' name='txtContactEmailPrimary' value='"&contactEmail&"' />"
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					strSB = strSB & "<div class='row'>"
					strSB = strSB & "<span class='LeftControl'>Contact Phone:</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & contactPhone'"<input type='text' name='txtContactPhonePrimary' value='"&contactPhone&"' />"
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					strSB = strSB & "<div class='row'>"
					strSB = strSB & "<span class='LeftControl'>Company Fax:</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & companyFax'"<input type='text' name='txtCompanyFax' value='"&companyFax&"' />"
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					strSB = strSB & "<div class='row'>"
					strSB = strSB & "<span class='LeftControl'>Company url:</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & companyUrl'"<input type='text' name='txtCompanyUrl' value='"&companyUrl&"' />"
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					' 'company name and address
					' strSB = strSB & "<div class='vendor_demographics'>"
					' strSB = strSB & "<b>"&companyName&"</b><br/>"
					' strSB = strSB & address&"<br/>"
					' strSB = strSB & cityStateZip & "<br/>"
					' strSB = strSB & country
					' strSB = strSB & "</div>"
					' 'company information in short
					' IF companyInformation <> "" THEN
					' strSB = strSB & "<div class='vendor_info'>"
					' strSB = strSB & "<b>Company Info:</b><br/>"
					' strSB = strSB & companyInformation
					' strSB = strSB & "</div>"
					' END IF
					' 'specialty
					' IF specialty <> "" THEN
					' strSB = strSB & "<div class='vendor_spacialty'>"
					' strSB = strSB & "<b>Specialty:</b><br/>"
					' strSB = strSB & specialty
					' strSB = strSB & "</div>"
					' END IF
					' 'General comments
					' IF generalComments <> "" THEN
					' strSB = strSB & "<div class='vendor_comments'>"
					' strSB = strSB & "<b>General Comments:</b><br/>"
					' strSB = strSB & generalComments
					' strSB = strSB & "</div>"
					' END IF
					
					' ' members contact information
					' strSB = strSB & "<div class='vendor_members'>"
					' strSB = strSB & "<b>Member(s):</b><br/>"
				END IF				
				
					' strSB = strSB & contactName & "<br/>"
					' strSB = strSB & "Email: "&contactEmail & "<br/>"
					' strSB = strSB & "Phone: "&contactPhone & "<br/>"
				
				' IF primaryRecordsetCounter = counts(count-1) THEN ' in last iteration close the div
					' strSB = strSB & "</div>"
				' END IF
				
				' IF primaryRecordsetCounter = counts(count-1) AND (companyFax <> "" OR companyUrl <> "")THEN
					' 'Contact info of the vendor
					' strSB = strSB & "<div class='vendor_contactinfo'>"
					' strSB = strSB & "<b>Contact Info:</b><br/>"
					' IF companyFax <> "" THEN
						' strSB = strSB & "Fax: "& companyFax & "<br/>"
					' END IF
					' IF companyUrl <> "" THEN
						' strSB = strSB & "URL: "& companyUrl & "<br/>"
					' END IF
					' strSB = strSB & "</div>"
				' END IF				
				primaryRecordsetCounter = primaryRecordsetCounter + 1
				Recordset.MoveNext
			LOOP			
		END IF
		IF count = 2  THEN
			dim arrSpecialties(8)
			DO UNTIL Recordset.EOF				
				arrSpecialties(specialtyCounter) = Recordset.Fields("SpecialtyID").value
				specialtyCounter = specialtyCounter + 1
				Recordset.MoveNext
			LOOP
			IF isAttorney THEN
				strSB = strSB & "<div class='row'>"
				strSB = strSB & "<span class='LeftControl'>Specialties:</span>"
				strSB = strSB & "<span class='RightControl'>"
				strSB = strSB & "<select name='ATTNY_SPECIALTY' multiple>"
				strSB = strSB & GetAttorneySpecialties(arrSpecialties)
				strSB = strSB & "</select>"
				strSB = strSB & "</span></div>"
			END IF
		END IF
		IF count = 3 THEN
			dim arrPracStates(64)
			DO UNTIL Recordset.EOF				
				arrPracStates(practicingStatesCounter) = Recordset.Fields("PracticingStateID").value
				practicingStatesCounter = practicingStatesCounter + 1
				Recordset.MoveNext
			LOOP
			IF isAttorney THEN
				strSB = strSB & "<div class='row'>"
				strSB = strSB & "<span class='LeftControl'>Practicing States:</span>"
				strSB = strSB & "<span class='RightControl'>"
				strSB = strSB & "<select name='ATTNY_STATE' multiple>"
				strSB = strSB & GetAttorneyPracticingStates(arrPracStates)
				strSB = strSB & "</select>"
				strSB = strSB & "</span></div>"
			END IF						
		END IF
		IF count = 4 THEN			
			IF NOT Recordset.EOF THEN
				
				'recommended = Recordset.Fields("Recommended").value
				recommFirmAttorney = Recordset.Fields("RecommFirmAtorney").value
				yearsServiceUsed = Recordset.Fields("YearsServiceUsed").value
				yearsServiceUsedId = Recordset.Fields("YearsOfServiceUsedId").value
				showRecomInfo = Recordset.Fields("ShowRecomInfo").value
				recommComment = Recordset.Fields("RecommComment").value
				recommendedBy = Recordset.Fields("RecommendedBy").value
				IF ISNULL(recommFirmAttorney) THEN
					recommFirmAttorney = ""
				END IF
				IF ISNULL(yearsServiceUsed) THEN
					yearsServiceUsed = ""
				END IF
				IF ISNULL(yearsServiceUsedId) THEN
					yearsServiceUsedId = ""
				END IF
				IF IsNULL(showRecomInfo) THEN
					showRecomInfo = false
				END IF
				IF ISNUll(recommComment) THEN
					recommComment = ""
				END IF
				IF ISNUll(recommendedBy)  THEN
					recommendedBy = ""
				END IF
				IF isAttorney THEN
					strSB = strSB & strRecommendationsTemplate					
					strSB = Replace(strSB,"[RECOMMENDED_BY]",recommendedBy)
					strSB = Replace(strSB,"[RECOMMENDER_COMPANY]",recommendedBy)'TODO: I don't have recommendor company info yet
					strSB = Replace(strSB,"[REC_COMMENTS]",recommComment)
					'IF UCase(recommended) = "YES" THEN
					IF showRecomInfo THEN
						strSB = Replace(strSB,"[RECOMMENDED_YES]","checked=""checked""")
						strSB = Replace(strSB,"[RECOMMENDED_NO]","")
					Else
						strSB = Replace(strSB,"[RECOMMENDED_NO]","checked=""checked""")
						strSB = Replace(strSB,"[RECOMMENDED_YES]","")
					END IF
					IF yearsServiceUsedId = "90" THEN
						strSB = Replace(strSB,"[20_YEAR]","selected")
						strSB = Replace(strSB,"[15_YEAR]","")
						strSB = Replace(strSB,"[10_YEAR]","")
						strSB = Replace(strSB,"[9_YEAR]","")
						strSB = Replace(strSB,"[8_YEAR]","")
						strSB = Replace(strSB,"[7_YEAR]","")
						strSB = Replace(strSB,"[6_YEAR]","")
						strSB = Replace(strSB,"[5_YEAR]","")
						strSB = Replace(strSB,"[4_YEAR]","")
						strSB = Replace(strSB,"[3_YEAR]","")
						strSB = Replace(strSB,"[2_YEAR]","")
						strSB = Replace(strSB,"[1_YEAR]","")
					ELSEIF yearsServiceUsedId = "89" THEN
						strSB = Replace(strSB,"[20_YEAR]","")
						strSB = Replace(strSB,"[15_YEAR]","selected")
						strSB = Replace(strSB,"[10_YEAR]","")
						strSB = Replace(strSB,"[9_YEAR]","")
						strSB = Replace(strSB,"[8_YEAR]","")
						strSB = Replace(strSB,"[7_YEAR]","")
						strSB = Replace(strSB,"[6_YEAR]","")
						strSB = Replace(strSB,"[5_YEAR]","")
						strSB = Replace(strSB,"[4_YEAR]","")
						strSB = Replace(strSB,"[3_YEAR]","")
						strSB = Replace(strSB,"[2_YEAR]","")
						strSB = Replace(strSB,"[1_YEAR]","")
					ELSEIF yearsServiceUsedId = "88" THEN
						strSB = Replace(strSB,"[20_YEAR]","")
						strSB = Replace(strSB,"[15_YEAR]","")
						strSB = Replace(strSB,"[10_YEAR]","selected")
						strSB = Replace(strSB,"[9_YEAR]","")
						strSB = Replace(strSB,"[8_YEAR]","")
						strSB = Replace(strSB,"[7_YEAR]","")
						strSB = Replace(strSB,"[6_YEAR]","")
						strSB = Replace(strSB,"[5_YEAR]","")
						strSB = Replace(strSB,"[4_YEAR]","")
						strSB = Replace(strSB,"[3_YEAR]","")
						strSB = Replace(strSB,"[2_YEAR]","")
						strSB = Replace(strSB,"[1_YEAR]","")
					ELSEIF yearsServiceUsedId = "87" THEN
						strSB = Replace(strSB,"[20_YEAR]","")
						strSB = Replace(strSB,"[15_YEAR]","")
						strSB = Replace(strSB,"[10_YEAR]","")
						strSB = Replace(strSB,"[9_YEAR]","selected")
						strSB = Replace(strSB,"[8_YEAR]","")
						strSB = Replace(strSB,"[7_YEAR]","")
						strSB = Replace(strSB,"[6_YEAR]","")
						strSB = Replace(strSB,"[5_YEAR]","")
						strSB = Replace(strSB,"[4_YEAR]","")
						strSB = Replace(strSB,"[3_YEAR]","")
						strSB = Replace(strSB,"[2_YEAR]","")
						strSB = Replace(strSB,"[1_YEAR]","")
					ELSEIF yearsServiceUsedId = "86" THEN
						strSB = Replace(strSB,"[20_YEAR]","")
						strSB = Replace(strSB,"[15_YEAR]","")
						strSB = Replace(strSB,"[10_YEAR]","")
						strSB = Replace(strSB,"[9_YEAR]","")
						strSB = Replace(strSB,"[8_YEAR]","selected")
						strSB = Replace(strSB,"[7_YEAR]","")
						strSB = Replace(strSB,"[6_YEAR]","")
						strSB = Replace(strSB,"[5_YEAR]","")
						strSB = Replace(strSB,"[4_YEAR]","")
						strSB = Replace(strSB,"[3_YEAR]","")
						strSB = Replace(strSB,"[2_YEAR]","")
						strSB = Replace(strSB,"[1_YEAR]","")
					ELSEIF yearsServiceUsedId = "85" THEN
						strSB = Replace(strSB,"[20_YEAR]","")
						strSB = Replace(strSB,"[15_YEAR]","")
						strSB = Replace(strSB,"[10_YEAR]","")
						strSB = Replace(strSB,"[9_YEAR]","")
						strSB = Replace(strSB,"[8_YEAR]","")
						strSB = Replace(strSB,"[7_YEAR]","selected")
						strSB = Replace(strSB,"[6_YEAR]","")
						strSB = Replace(strSB,"[5_YEAR]","")
						strSB = Replace(strSB,"[4_YEAR]","")
						strSB = Replace(strSB,"[3_YEAR]","")
						strSB = Replace(strSB,"[2_YEAR]","")
						strSB = Replace(strSB,"[1_YEAR]","")
					ELSEIF yearsServiceUsedId = "84" THEN
						strSB = Replace(strSB,"[20_YEAR]","")
						strSB = Replace(strSB,"[15_YEAR]","")
						strSB = Replace(strSB,"[10_YEAR]","")
						strSB = Replace(strSB,"[9_YEAR]","")
						strSB = Replace(strSB,"[8_YEAR]","")
						strSB = Replace(strSB,"[7_YEAR]","")
						strSB = Replace(strSB,"[6_YEAR]","selected")
						strSB = Replace(strSB,"[5_YEAR]","")
						strSB = Replace(strSB,"[4_YEAR]","")
						strSB = Replace(strSB,"[3_YEAR]","")
						strSB = Replace(strSB,"[2_YEAR]","")
						strSB = Replace(strSB,"[1_YEAR]","")
					ELSEIF yearsServiceUsedId = "83" THEN
						strSB = Replace(strSB,"[20_YEAR]","")
						strSB = Replace(strSB,"[15_YEAR]","")
						strSB = Replace(strSB,"[10_YEAR]","")
						strSB = Replace(strSB,"[9_YEAR]","")
						strSB = Replace(strSB,"[8_YEAR]","")
						strSB = Replace(strSB,"[7_YEAR]","")
						strSB = Replace(strSB,"[6_YEAR]","")
						strSB = Replace(strSB,"[5_YEAR]","selected")
						strSB = Replace(strSB,"[4_YEAR]","")
						strSB = Replace(strSB,"[3_YEAR]","")
						strSB = Replace(strSB,"[2_YEAR]","")
						strSB = Replace(strSB,"[1_YEAR]","")
					ELSEIF yearsServiceUsedId = "82" THEN
						strSB = Replace(strSB,"[20_YEAR]","")
						strSB = Replace(strSB,"[15_YEAR]","")
						strSB = Replace(strSB,"[10_YEAR]","")
						strSB = Replace(strSB,"[9_YEAR]","")
						strSB = Replace(strSB,"[8_YEAR]","")
						strSB = Replace(strSB,"[7_YEAR]","")
						strSB = Replace(strSB,"[6_YEAR]","")
						strSB = Replace(strSB,"[5_YEAR]","")
						strSB = Replace(strSB,"[4_YEAR]","selected")
						strSB = Replace(strSB,"[3_YEAR]","")
						strSB = Replace(strSB,"[2_YEAR]","")
						strSB = Replace(strSB,"[1_YEAR]","")
					ELSEIF yearsServiceUsedId = "81" THEN
						strSB = Replace(strSB,"[20_YEAR]","")
						strSB = Replace(strSB,"[15_YEAR]","")
						strSB = Replace(strSB,"[10_YEAR]","")
						strSB = Replace(strSB,"[9_YEAR]","")
						strSB = Replace(strSB,"[8_YEAR]","")
						strSB = Replace(strSB,"[7_YEAR]","")
						strSB = Replace(strSB,"[6_YEAR]","")
						strSB = Replace(strSB,"[5_YEAR]","")
						strSB = Replace(strSB,"[4_YEAR]","")
						strSB = Replace(strSB,"[3_YEAR]","selected")
						strSB = Replace(strSB,"[2_YEAR]","")
						strSB = Replace(strSB,"[1_YEAR]","")
					ELSEIF yearsServiceUsedId = "80" THEN
						strSB = Replace(strSB,"[20_YEAR]","")
						strSB = Replace(strSB,"[15_YEAR]","")
						strSB = Replace(strSB,"[10_YEAR]","")
						strSB = Replace(strSB,"[9_YEAR]","")
						strSB = Replace(strSB,"[8_YEAR]","")
						strSB = Replace(strSB,"[7_YEAR]","")
						strSB = Replace(strSB,"[6_YEAR]","")
						strSB = Replace(strSB,"[5_YEAR]","")
						strSB = Replace(strSB,"[4_YEAR]","")
						strSB = Replace(strSB,"[3_YEAR]","")
						strSB = Replace(strSB,"[2_YEAR]","selected")
						strSB = Replace(strSB,"[1_YEAR]","")
					ELSEIF yearsServiceUsedId = "79" THEN
						strSB = Replace(strSB,"[20_YEAR]","")
						strSB = Replace(strSB,"[15_YEAR]","")
						strSB = Replace(strSB,"[10_YEAR]","")
						strSB = Replace(strSB,"[9_YEAR]","")
						strSB = Replace(strSB,"[8_YEAR]","")
						strSB = Replace(strSB,"[7_YEAR]","")
						strSB = Replace(strSB,"[6_YEAR]","")
						strSB = Replace(strSB,"[5_YEAR]","")
						strSB = Replace(strSB,"[4_YEAR]","")
						strSB = Replace(strSB,"[3_YEAR]","")
						strSB = Replace(strSB,"[2_YEAR]","")
						strSB = Replace(strSB,"[1_YEAR]","selected")
					Else						
						strSB = Replace(strSB,"[20_YEAR]","")
						strSB = Replace(strSB,"[15_YEAR]","")
						strSB = Replace(strSB,"[10_YEAR]","")
						strSB = Replace(strSB,"[9_YEAR]","")
						strSB = Replace(strSB,"[8_YEAR]","")
						strSB = Replace(strSB,"[7_YEAR]","")
						strSB = Replace(strSB,"[6_YEAR]","")
						strSB = Replace(strSB,"[5_YEAR]","")
						strSB = Replace(strSB,"[4_YEAR]","")
						strSB = Replace(strSB,"[3_YEAR]","")
						strSB = Replace(strSB,"[2_YEAR]","")
						strSB = Replace(strSB,"[1_YEAR]","")
					END IF
						
				END IF
				Recordset.MoveNext
			END IF			
		END IF		
		count = count + 1
		Set Recordset = Recordset.NextRecordset
	LOOP	
	GetVendorDetailsForEdit = strSB
End Function

Function GetLogoUploadControl(memberId)
	dim strSB
	dim imagePath,logoFileName
	strSB = ""
	imagePath = ""
	logoFileName = ""	
	IF IsNumeric(memberId) THEN			
		Set cmd = Server.CreateObject("ADODB.Command")
		With cmd
		   .ActiveConnection = IFAconn 
		   .CommandType = adCmdStoredProc
		   .CommandText = "viewvendor" ' Set the name of the Stored Procedure to use 
		   .Parameters.Append .CreateParameter("@MemberID",adInteger,adParamInput,,memberId)			
		   set Recordset = .Execute
		End With
		IF NOT Recordset.EOF THEN
			logoFileName = Recordset("LogoPath")
			if isnull(logoFileName) or trim(logoFileName)="" then
			  noLogo=true
			end if
			if noLogo then
				strSB = strSB & "<div class='row'><input onclick='openAddPhoto()' type='button' value='Add Photo/Logo' /></div>"
			else
				imagePath = trim(application("consoleurl"))&"/vendor_logo/"&memberId&"/"&trim(logoFileName)
				strSB = strSB & "<div class='LeftControl'><img class='vendor_logo' src='"&imagePath&"' style='max-height:200px;max-width:200px;'/></div>"
				strSB = strSB & "<div class='LeftControl'><a href='javascript:confirm_delete(""delete_vendor_logo.asp?ID="&memberId&""")'>Delete</a></div>"
			end if
		END IF
	END IF
	GetLogoUploadControl = strSB	
End Function

Function AddOrUpdateVendorDetails(model)
	AddOrUpdateVendorDetails = false
	Set cmd = Server.CreateObject("ADODB.Command")
	With cmd
	   .ActiveConnection = IFAconn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "SaveVendor" ' Set the name of the Stored Procedure to use 
	   .Parameters.Append .CreateParameter("@MemberID",adInteger,adParamInput,,model(0))
	   .Parameters.Append .CreateParameter("@CategoryID",adInteger,adParamInput,,model(1))
	   .Parameters.Append .CreateParameter("@ServiceYearsID",adInteger,adParamInput,,model(2))
	   .Parameters.Append .CreateParameter("@RecommendedFirm",adBoolean,adParamInput,,model(3))
	   .Parameters.Append .CreateParameter("@RecommendedAttorney",adBoolean,adParamInput,,model(4))
	   .Parameters.Append .CreateParameter("@RecommendedBy",adVarChar,adParamInput,500,model(5))
	   .Parameters.Append .CreateParameter("@ShowRecomInfo",adBoolean,adParamInput,,model(6))
	   .Parameters.Append .CreateParameter("@RecommComment",adVarChar,adParamInput,500,model(7))
	   .Parameters.Append .CreateParameter("@LogoPath",adVarChar,adParamInput,255,model(8))
	   .Parameters.Append .CreateParameter("@Specialties",adVarChar,adParamInput,255,model(9))
	   .Parameters.Append .CreateParameter("@PracticingState",adVarChar,adParamInput,255,model(10))	   
		
	   .Execute
	End With
	
	If Err.number <> 0 Then
		err.Clear 
	Else
		AddOrUpdateVendorDetails = TRUE
	End If
	
	Set cmd.ActiveConnection = Nothing
    If Not cmd Is Nothing then Set cmd = Nothing	
End Function

Function DeleteVendor(vendorId)
	DeleteVendor = false
	Set cmd = Server.CreateObject("ADODB.Command")
	With cmd
	   .ActiveConnection = IFAconn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "DeleteVendor" ' Set the name of the Stored Procedure to use 
	   .Parameters.Append .CreateParameter("@MemberID",adInteger,adParamInput,,vendorId)		
	   .Execute
	End With
	
	If Err.number <> 0 Then
		err.Clear 
	Else
		DeleteVendor = TRUE
	End If
	
	Set cmd.ActiveConnection = Nothing
    If Not cmd Is Nothing then Set cmd = Nothing
	
End Function

Function GetVendorsDataNotIncludedInVendorDirectory()
	dim strSB 
	strSB = ""
	Set cmd = Server.CreateObject("ADODB.Command")
	With cmd
	   .ActiveConnection = IFAconn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "ListVendorsNotInDirectory"
	   set Recordset = .Execute
	End With
	
	strSB = strSB & "<Table class='vendors_table'>"
	strSB = strSB & "<THead><tr class='vendors_table_header'>"
	strSB = strSB & "<TH>Member Type</TD>"
	strSB = strSB & "<TH>Country</TD>"
	strSB = strSB & "<TH>Company Name</TD>"		
	strSB = strSB & "<TH>City</TD>"
	strSB = strSB & "<TH>State</TD>"
	strSB = strSB & "</tr></THead>"
	DO UNTIL Recordset.EOF
		memberId = Recordset.Fields("MemberID").value
		memberType = Recordset.Fields("MemberType").value
		country = Recordset.Fields("Country").value
		companyName = Recordset.Fields("CompanyName").value
		city = Recordset.Fields("City").value
		state = Recordset.Fields("State").value
		strSB = strSB & "<TR class='vendors_table_row'>"
		strSB = strSB & "<TD>"&memberType&"</TD>"
		strSB = strSB & "<TD>"&Country&"</TD>"
		strSB = strSB & "<TD><a href=""Add.asp?ID="&memberId&""">"&CompanyName&"</a></TD>"
		strSB = strSB & "<TD>"&City&"</TD>"
		strSB = strSB & "<TD>"&State&"</TD>"
		strSB = strSB & "</TR>"
		Recordset.MoveNext
	LOOP
	strSB = strSB & "</Table>"
	GetVendorsDataNotIncludedInVendorDirectory = strSB
	'Clean up
	set cmd = nothing
	set Recordset = nothing
End Function

' Function GetVendorsNotIncludedInVendorDirectory()
	' dim Recordset,strSB
	' strSB=""
	' Recordset = GetVendorsDataNotIncludedInVendorDirectory()
	' strSB = strSB & "<Table>"
	' strSB = strSB & "<TH>"
	' strSB = strSB & "<TD>Member Type</TD>"
	' strSB = strSB & "<TD>Country</TD>"
	' strSB = strSB & "<TD>Company Name</TD>"		
	' strSB = strSB & "<TD>City</TD>"
	' strSB = strSB & "<TD>State</TD>"
	' strSB = strSB & "</TH>"
	' DO UNTIL Recordset.EOF
		' memberId = Recordset.Fields("MemberID").value
		' memberType = Recordset.Fields("MemberType").value
		' country = Recordset.Fields("Country").value
		' companyName = Recordset.Fields("CompanyName").value
		' city = Recordset.Fields("City").value
		' state = Recordset.Fields("State").value
		' strSB = strSB & "<TR>"
		' strSB = strSB & "<TD>"&memberType&"</TD>"
		' strSB = strSB & "<TD>"&Country&"</TD>"
		' strSB = strSB & "<TD><a href=""Add_Vendor.asp?ID="&memberId&""">"&CompanyName&"</a></TD>"
		' strSB = strSB & "<TD>"&City&"</TD>"
		' strSB = strSB & "<TD>"&State&"</TD>"
		' strSB = strSB & "</TR>"
	' LOOP
	' GetVendorsNotIncludedInVendorDirectory = strSB
' End Function

Function GetVendorMemberDetailsById(id)
	Set cmd = Server.CreateObject("ADODB.Command")
	With cmd
	   .ActiveConnection = IFAconn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "ListVendorsNotInDirectory"
	   .Parameters.Append .CreateParameter("@memberID",adInteger,adParamInput,,id)
	   set Recordset = .Execute
	End With
	SET GetVendorMemberDetailsById = Recordset
	SET cmd = Nothing
	SET Recordset = Nothing
End Function

Function GetAddvendorPage(id)		
	dim strSB,isAttorney,companyName, city,state,province,country,memberType,Recordset
	strSB = ""
	companyName = ""
	city = ""
	state = ""
	province = ""
	country = ""
	memberType = ""	
	
	SET Recordset = GetVendorMemberDetailsById(id)
	DO UNTIL Recordset.EOF
		companyName = Recordset.Fields("CompanyName")
		city = Recordset.Fields("City")
		state = Recordset.Fields("State")
		province = Recordset.Fields("Province")
		country = Recordset.Fields("Country")
		memberType = Recordset.Fields("MemberType")
		memberTypeID = Recordset.Fields("MemberTypeID")
		
		strSB = strSB & "<div class='row'>"
		strSB = strSB & "<span class='LeftControl'>Company Name:</span>"
		strSB = strSB & "<span class='RightControl'>"&companyName&"</span>"
		strSB = strSB & "</div>"
		strSB = strSB & "<div class='row'>"
		strSB = strSB & "<span class='LeftControl'>City:</span>"
		strSB = strSB & "<span class='RightControl'>"&city&"</span>"
		strSB = strSB & "</div>"
		strSB = strSB & "<div class='row'>"
		strSB = strSB & "<span class='LeftControl'>State/Province:</span>"
		IF state <> "" THEN
			strSB = strSB & "<span class='RightControl'>"&state&"</span>"
		ELSE
			strSB = strSB & "<span class='RightControl'>"&province&"</span>"
		END IF
		strSB = strSB & "</div>"
		strSB = strSB & "<div class='row'>"
		strSB = strSB & "<span class='LeftControl'>Country:</span>"
		strSB = strSB & "<span class='RightControl'>"&country&"</span>"
		strSB = strSB & "</div>"
		Recordset.MoveNext
	LOOP
	isAttorney = false
	isAttorney = memberTypeID = 9 OR memberTypeID = 10 
	strSB = strSB & "<div class='row'>"					
	strSB = strSB & "<span class='LeftControl'>Category:</span>"
	strSB = strSB & "<span class='RightControl'>"
	strSB = strSB & "<select name='Basic_category'>"
	strSB = strSB & GetVendorCategories("")
	strSB = strSB & "</select>"
	strSB = strSB & "</span>"
	strSB = strSB & "</div>"
	IF isAttorney THEN
		strSB = strSB & "<div class='row'>"
		strSB = strSB & "<span class='LeftControl'>Specialties:</span>"
		strSB = strSB & "<span class='RightControl'>"
		strSB = strSB & "<select name='ATTNY_SPECIALTY' multiple>"
		strSB = strSB & GetAttorneySpecialties("")
		strSB = strSB & "</select>"
		strSB = strSB & "</span></div>"
		
		strSB = strSB & "<div class='row'>"
		strSB = strSB & "<span class='LeftControl'>Practicing States:</span>"
		strSB = strSB & "<span class='RightControl'>"
		strSB = strSB & "<select name='ATTNY_STATE' multiple>"
		strSB = strSB & GetAttorneyPracticingStates("")
		strSB = strSB & "</select>"
		strSB = strSB & "</span></div>"
		
		strSB = strSB & strRecommendationsTemplate					
		strSB = Replace(strSB,"[RECOMMENDED_BY]","")
		strSB = Replace(strSB,"[RECOMMENDER_COMPANY]","")
		strSB = Replace(strSB,"[REC_COMMENTS]","")
		
		strSB = Replace(strSB,"[RECOMMENDED_YES]","")
		strSB = Replace(strSB,"[RECOMMENDED_NO]","")
		
		strSB = Replace(strSB,"[20_YEAR]","")
		strSB = Replace(strSB,"[15_YEAR]","")
		strSB = Replace(strSB,"[10_YEAR]","")
		strSB = Replace(strSB,"[9_YEAR]","")
		strSB = Replace(strSB,"[8_YEAR]","")
		strSB = Replace(strSB,"[7_YEAR]","")
		strSB = Replace(strSB,"[6_YEAR]","")
		strSB = Replace(strSB,"[5_YEAR]","")
		strSB = Replace(strSB,"[4_YEAR]","")
		strSB = Replace(strSB,"[3_YEAR]","")
		strSB = Replace(strSB,"[2_YEAR]","")
		strSB = Replace(strSB,"[1_YEAR]","")
	END IF
	GetAddvendorPage = strSB
END Function

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
