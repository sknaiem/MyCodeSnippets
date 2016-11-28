<%
' pseudo code
' get data from Vendorlist stored proc
' we have data in recordset
dim isVendorSet
dim isPreferredVendorSet
dim prevPrefCategory
dim prevCategory
dim prevPrefCountry
dim prevCountry
dim prevPrefStateOrProvince
dim prevStateOrProvince
dim strSB
isVendorSet = false
isPreferredVendorSet = false
prevPrefCategory = ""
prevCategory = ""
prevPrefCountry = ""
prevCountry = ""
prevPrefStateOrProvince = ""
prevStateOrProvince = ""
strSB = ""
do while not recordset.eof
	vendorType = recordset.Fields("VendorType").value
	categoryName = recordset.Fields("CategoryName").value
	country = recordset.Fields("Country").value
	state = recordset.Fields("State").value
	city = recordset.Fields("City").value
	companyName = recordset.Fields("CompanyName").value
	
	
	if vendorType = "Preferred Vendor" and isPreferredVendorSet = false then
		strSB = strSB & "<tr><td>IFA Preferred Vendor</td></tr>"
		isPreferredVendorSet = true
		prevPrefCategory = ""'reset the previous preferred category everytime vendor type changes
	end if
	'within preferred vendor vendorcategories should not repeat
		if prevPrefCategory = "" or (prevPrefCategory <> "" and categoryName <> prevPrefCategory) then
			strSB = strSB & "<tr><td>"&categoryName&"</td></tr>"
			prevPrefCategory = categoryName
			prevPrefCountry = ""'reset the previous preferred country everytime the category is changed
		end if
	' within vendorcategory countries should not repeat
			if prevPrefCountry = "" or (prevPrefCountry <> "" and country <> prevPrefCountry) then
				strSB = strSB & "<tr><td>"&country&"</td></tr>"
				prevPrefCountry = country
				prevPrefStateOrProvince=""
			end if
	'within a country the state or province name should not repeat
				if prevPrefStateOrProvince = "" then
					if state <> "--" then
						strSB = strSB & "<tr><td style='padding-top:10px;'>"&state&"</td></tr><tr><td>"
						prevPrefStateOrProvince = state
					elseif province <> "" then
						strSB = strSB & "<tr><td style='padding-top:10px;'>"&province&"</td></tr><tr><td>"
						prevPrefStateOrProvince = province
					end if			
				elseif prevPrefStateOrProvince <> "" and state <> prevPrefStateOrProvince then
					if state <> "--" then
						strSB = strSB & "</td></tr><tr><td style='padding-top:10px;'>"&state&"</td></tr><tr><td>"
						prevPrefStateOrProvince = state
					elseif province <> "" and state <> prevPrefStateOrProvince then
						strSB = strSB & "</td></tr><tr><td style='padding-top:10px;'>"&province&"</td></tr><tr><td>"
						prevPrefStateOrProvince = province
					end if			
				end if
	
	
	if vendorType = "Vendor" and isVendorSet = false then
		strSB = strSB & "<tr><td>IFA Vendor</td></tr>"
		isVendorSet = true
		prevCategory = ""'reset the previous preferred category everytime vendor type changes
	end if
	'within vendor vendorcategories should not repeat
		if prevCategory = "" or (prevCategory <> "" and categoryName <> prevCategory) then
			strSB = strSB & "<tr><td>"&categoryName&"</td></tr>"
			prevCategory = categoryName
			prevCountry = ""'reset the previous preferred country everytime the category is changed
		end if
	' within vendorcategory countries should not repeat
			if prevCountry = "" or (prevCountry <> "" and country <> prevCountry) then
				strSB = strSB & "<tr><td>"&country&"</td></tr>"
				prevCountry = country
				prevStateOrProvince=""
			end if
	'within a country the state or province name should not repeat
				if prevStateOrProvince = "" then
					if state <> "--" then
						strSB = strSB & "<tr><td style='padding-top:10px;'>"&state&"</td></tr><tr><td>"
						prevStateOrProvince = state
					elseif province <> "" then
						strSB = strSB & "<tr><td style='padding-top:10px;'>"&province&"</td></tr><tr><td>"
						prevStateOrProvince = province
					end if			
				elseif prevStateOrProvince <> "" and state <> prevStateOrProvince then
					if state <> "--" then
						strSB = strSB & "</td></tr><tr><td style='padding-top:10px;'>"&state&"</td></tr><tr><td>"
						prevStateOrProvince = state
					elseif province <> "" and state <> prevStateOrProvince then
						strSB = strSB & "</td></tr><tr><td style='padding-top:10px;'>"&province&"</td></tr><tr><td>"
						prevStateOrProvince = province
					end if			
				end if
	
	
	strSB = strSB & "<div><a>"&companyName&"</a><br/>"
	strSB = strSB & city&","&state&","&country&"</div>"
	recordset.MoveNext
loop
%>