<%
' pseudo code
' get data from Vendorlist stored proc
' we have data in recordset
dim isVendorSet
dim isPreferredVendorSet
dim prevCategory
dim prevCountry
dim prevState
dim strSB
isVendorSet = false
isPreferredVendorSet = false
prevCategory = ''
prevCountry = ''
prevState = ''
strSB = ''
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
	else if vendorType = "Vendor" and isVendorSet = false then
		strSB = strSB & "<tr><td>IFA Vendor</td></tr>"
		isVendorSet = true
	end if
	if prevCategory = '' or (prevCategory <> '' and categoryName <> prevCategory) then
		strSB = strSB & "<tr><td>"&categoryName&"</td></tr>"
		prevCategory = categoryName
	end if
	if prevCountry = '' or (prevCountry <> '' and country <> prevCountry) then
		strSB = strSB & "<tr><td>"&country&"</td></tr>"
		prevCountry = country
	end if
	if prevState = '' then
	strSB = strSB & "<tr><td>"&state&"</td></tr><tr><td>"
	prevState = state
	else if prevState <> '' and state <> prevState then
	strSB = strSB & "</td></tr><tr><td>"&state&"</td></tr><tr><td>"
	prevState = state
	end if
	strSB = strSB & "<div><a>"&companyName&"</a><br/>"
	strSB = strSB & city&","&state&","&country&"</div>"
	recordset.MoveNext
loop
%>