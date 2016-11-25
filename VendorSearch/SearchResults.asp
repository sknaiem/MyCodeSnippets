<%
' pseudo code
' get data from Vendorlist stored proc
' we have data in recordset
dim isVendorSet = false
dim isPreferredVendorSet = false
dim prevCategory
dim prevCountry
dim prevState
do while not recordset.eof
if recordset[isPreferredVendor] = isPreferred and isPreferredVendorSet = false
<tr>
<td>IFA Preferred Vendor</td>
</tr>
isPreferredVendorSet = true
else if recordset[isPreferredvendor]=false and isVendorSet = false
<tr>
<td>IFA Vendor</td>
</tr>
isVendorSet = true
end if
if prevCategory = '' or prevCategory <> '' and recordset[category] <> prevCategory then
<tr>
<td>recordset[category]</td>
</tr>
prevCategory = recordset[category]
end if
if prevCountry = '' or prevCountry <> '' and recordset[country] <> prevCountry then
<tr>
<td>recordset[country]</td>
</tr>
prevCountry = recordset[country]
end if
if prevState = '' then
<tr>
<td>recordset[state]</td>
</tr>
<tr>
<td>
prevState = recordset[state]
else if prevState <> '' and recordset[state] <> prevState
</td>
</tr>
<tr>
<td>recordset[state]</td>
</tr>
<tr>
<td>
prevState = recordset[state]
end if
<div>
<a>recordset[VendorName]</a>
recordset[City],recordset[state],recordset[country]
</div>
loop
%>