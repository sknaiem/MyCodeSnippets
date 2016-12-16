<%
Function GetVendorCategories(categorySelected)
	Dim strSB, strSBC
	strSB = ""
	strSBC = ""
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

Function GetAttorneyPracticingStates(arrPracStates)
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
	IF IsArray(arrPracStates) THEN
		Do UNTIL Recordset.EOF
			attorneyLocationName = Recordset.Fields("LocationName").value
			attorneyLocationID = Recordset.Fields("LocationID").value
			practStateAdded = false
			FOR EACH practState IN arrPracStates			
				IF practState = attorneyLocationName THEN
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
	IF IsArray(arrSpecialties) THEN ' TODO: if array is empty then this will fail
		DO UNTIL Recordset.EOF
			specialtyName = Recordset.Fields("SpecialtyName").value
			specialtyID = Recordset.Fields("SpecialtyID").value
			specialtyAdded = false
			For Each specialty IN arrSpecialties
				IF specialty = specialtyName THEN
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
						strSB = strSB & "<div class='vendor_preferred'>IFA Preferred Vendor</div><br/>"
					END IF
					IF logoPath <> "" THEN
						strSB = strSB & "<div class='vendor_logo'><img src='"&logoPath&"'/></div><br/>"
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
				
				recommended = Recordset.Fields("Recommended").value
				recommFirmAttorney = Recordset.Fields("RecommFirmAtorney").value
				yearsServiceUsed = Recordset.Fields("YearsServiceUsed").value
				showRecomInfo = Recordset.Fields("ShowRecomInfo").value
				recommComment = Recordset.Fields("RecommComment").value
				recommendedBy = Recordset.Fields("RecommendedBy").value
				IF recommended = "Yes" THEN
					
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
				isAttorney = instr(1,memberType,"Attorney",1)> 0
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
					strSB = strSB & "<input type='text' name='txtAddress' value='"&address&"' />"
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					strSB = strSB & "<div class='row'>"
					strSB = strSB & "<span class='LeftControl'>City, State, Zip:</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & "<input type='text' name='txtCityStateZip' value='"&cityStateZip&"' />"
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					strSB = strSB & "<div class='row'>"
					strSB = strSB & "<span class='LeftControl'>Country:</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & "<input type='text' name='txtCountry' value='"&country&"' />"
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					strSB = strSB & "<div class='row'>"
					strSB = strSB & "<span class='LeftControl'>Company Information:</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & "<input type='text' name='txtCompanyInformation' value='"&companyInformation&"' />"
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					strSB = strSB & "<div class='row'>"
					strSB = strSB & "<span class='LeftControl'>Specialty:</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & "<input type='text' name='txtSpecialty' value='"&specialty&"' />"
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					strSB = strSB & "<div class='row'>"
					strSB = strSB & "<span class='LeftControl'>General Comments:</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & "<input type='text' name='txtGeneralComments' value='"&generalComments&"' />"
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
					strSB = strSB & "<input type='text' name='txtContactEmailPrimary' value='"&contactEmail&"' />"
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					strSB = strSB & "<div class='row'>"
					strSB = strSB & "<span class='LeftControl'>Contact Phone:</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & "<input type='text' name='txtContactPhonePrimary' value='"&contactPhone&"' />"
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					strSB = strSB & "<div class='row'>"
					strSB = strSB & "<span class='LeftControl'>Company Fax:</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & "<input type='text' name='txtCompanyFax' value='"&companyFax&"' />"
					strSB = strSB & "</span>"
					strSB = strSB & "</div>"
					strSB = strSB & "<div class='row'>"
					strSB = strSB & "<span class='LeftControl'>Company url:</span>"
					strSB = strSB & "<span class='RightControl'>"
					strSB = strSB & "<input type='text' name='txtCompanyUrl' value='"&companyUrl&"' />"
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
				arrSpecialties(specialtyCounter) = Recordset.Fields("SpecialtyName").value
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
				arrPracStates(practicingStatesCounter) = Recordset.Fields("PracticingState").value
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
				
				recommended = Recordset.Fields("Recommended").value
				recommFirmAttorney = Recordset.Fields("RecommFirmAtorney").value
				yearsServiceUsed = Recordset.Fields("YearsServiceUsed").value
				showRecomInfo = Recordset.Fields("ShowRecomInfo").value
				recommComment = Recordset.Fields("RecommComment").value
				recommendedBy = Recordset.Fields("RecommendedBy").value
				IF isAttorney THEN
					strSB = strSB & strRecommendationsTemplate
					strSB = Replace(strSB,"[RECOMMENDED_BY]",recommendedBy)
					strSB = Replace(strSB,"[RECOMMENDER_COMPANY]",recommendedBy)'TODO: I don't have recommendor company info yet
					strSB = Replace(strSB,"[REC_COMMENTS]",recommComment)
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
	dim imagePath
	strSB = ""
	imagePath = ""
	set rscheckphoto=getrecordset("select * from af_members where userid="&memberId)
	if isnull(rscheckphoto("photo_link")) or trim(rscheckphoto("photo_link"))="" then
	  noLogo=true
	end if
	if noLogo then
		strSB = strSB & "<div class='row'><input onclick='openAddPhoto()' type='button' value='Add Photo/Logo' /></div>"
	else
		imagePath = trim(application("consoleurl"))&"/vendor_logo/"&memberId&"/"&trim(rscheckphoto("photo_link"))
		strSB = strSB & "<div class='LeftControl'><img class='vendor_logo' src='"&imagePath&"' style='max-height:200px;max-width:200px;'/></div>"
		strSB = strSB & "<div class='LeftControl'><a href='javascript:confirm_delete(""delete_vendor_logo.asp?ID="&memberId&""")'>Delete</a></div>"
	end if
	GetLogoUploadControl = strSB	
End Function

Function UpdateVendor(model)
	UpdateVendor = false
	Set cmd = Server.CreateObject("ADODB.Command")
	With cmd
	   .ActiveConnection = IFAconn 
	   .CommandType = adCmdStoredProc
	   .CommandText = "SaveVendor" ' Set the name of the Stored Procedure to use 
	   .Parameters.Append .CreateParameter("@MemberID",adInteger,adParamInput,,id)
	   .Parameters.Append .CreateParameter("@CategoryID",adInteger,adParamInput,,id)
	   .Parameters.Append .CreateParameter("@ServiceYearsID",adInteger,adParamInput,,id)
	   .Parameters.Append .CreateParameter("@RecommendedFirm",adInteger,adParamInput,,id)
	   .Parameters.Append .CreateParameter("@RecommendedAttorney",adInteger,adParamInput,,id)
	   .Parameters.Append .CreateParameter("@RecommendedBy",adInteger,adParamInput,,id)
	   .Parameters.Append .CreateParameter("@ShowRecomInfo",adInteger,adParamInput,,id)
	   .Parameters.Append .CreateParameter("@RecommComment",adInteger,adParamInput,,id)
	   .Parameters.Append .CreateParameter("@LogoPath",adInteger,adParamInput,,id)
	   .Parameters.Append .CreateParameter("@Specialties",adInteger,adParamInput,,id)
	   .Parameters.Append .CreateParameter("@PracticingState",adInteger,adParamInput,,id)	   
		
	   .Execute
	End With
	
	If Err.number <> 0 Then
		err.Clear 
	Else
		UpdateVendor = TRUE
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
	   .Parameters.Append .CreateParameter("@MemberID",adInteger,adParamInput,,id)		
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
