<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
Dim sSQL,rs,success,cnn,errmsg,index,allcountries
success=true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
method=trim(request("method"))
if method<>"" then shipType=int(method)
shipmet = "USPS"
if shipType=4 then shipmet = "UPS"
if shipType=6 then shipmet = yyCanPos
if shipType=7 then shipmet = "FedEx"
if shipType=8 then shipmet = "FedEx SmartPost"
if shipType=9 then shipmet = "DHL"
if shipType=10 then shipmet = "Australia Post"
function checkisdocument(st,serv)
	checkisdocument=""
	if st=9 then
		if serv="2" OR serv="5" OR serv="6" OR serv="7" OR serv="9" OR serv="B" OR serv="C" OR serv="D" OR serv="G" OR serv="I" OR serv="K" OR serv="L" OR serv="N" OR serv="R" OR serv="S" OR serv="T" OR serv="U" OR serv="W" OR serv="X" then
			checkisdocument=" <strong>(document)</strong>"
		end if
	end if
end function
if getpost("posted")="1" then
	if getpost("doadmin")<>"" then
		if shipType=3 then
			sSQL="UPDATE adminshipping SET adminUSPSUser='"&escape_string(getpost("adminUSPSUser"))&"' WHERE adminShipID=1"
			ect_query(sSQL)
		elseif shipType=4 then
			' sSQL="UPDATE adminshipping SET adminUPSAccount='"&escape_string(getpost("adminUPSAccount"))&"',adminUPSNegotiated="&getpost("UPSNegotiated")&" WHERE adminShipID=1"
			sSQL="UPDATE adminshipping SET adminUPSNegotiated="&getpost("UPSNegotiated")&" WHERE adminShipID=1"
			ect_query(sSQL)
		elseif shipType=6 then
			sSQL="UPDATE adminshipping SET adminCanPostUser='"&escape_string(getpost("adminCanPostUser"))&"' WHERE adminShipID=1"
			ect_query(sSQL)
		elseif shipType=8 then
			sSQL="UPDATE adminshipping SET smartPostHub='"&escape_string(getpost("smartPostHub"))&"' WHERE adminShipID=1"
			ect_query(sSQL)
		elseif shipType=9 then
			sSQL="UPDATE adminshipping SET DHLSiteID='"&escape_string(getpost("DHLSiteID"))&"',DHLSitePW='"&escape_string(getpost("DHLSitePW"))&"',DHLAccountNo='"&escape_string(getpost("DHLAccountNo"))&"' WHERE adminShipID=1"
			ect_query(sSQL)
		elseif shipType=10 then
			sSQL="UPDATE adminshipping SET AusPostAPI='"&escape_string(getpost("AusPostAPI"))&"' WHERE adminShipID=1"
			ect_query(sSQL)
		elseif shipType=99 then
			sSQL="UPDATE adminshipping SET shipStationUser='"&escape_string(getpost("shipStationUser"))&"',shipStationPass='"&escape_string(getpost("shipStationPass"))&"' WHERE adminShipID=1"
			ect_query(sSQL)
		end if
	else
		if shipType=3 OR shipType=10 then
			for index=1+IIfVr(shipType=10,600,0) to 50+IIfVr(shipType=10,600,0)
				if getpost("methodshow"&index)<>"" then
					sSQL = "UPDATE uspsmethods SET uspsShowAs='"&escape_string(getpost("methodshow"&index))&"',"
					if getpost("methodfsa"&index)="ON" then
						sSQL = sSQL & "uspsFSA=1,"
					else
						sSQL = sSQL & "uspsFSA=0,"
					end if
					if getpost("methoduse"&index)="ON" then
						sSQL = sSQL & "uspsUseMethod=1 WHERE uspsID="&index
					else
						sSQL = sSQL & "uspsUseMethod=0 WHERE uspsID="&index
					end if
					ect_query(sSQL)
				end if
			next
		elseif shipType=4 OR shipType=6 OR shipType=7 OR shipType=8 OR shipType=9 then
			indexadd=0
			if shipType=6 then
				indexadd=100
			elseif shipType=7 then
				indexadd=200
			elseif shipType=8 then
				indexadd=300
			elseif shipType=9 then
				indexadd=400
			end if
			for index=100+indexadd to 155+indexadd
				if getpost("methodshow"&index)<>"" then
					sSQL = "UPDATE uspsmethods SET "
					if getpost("methodfsa"&index)="ON" then
						sSQL = sSQL & "uspsFSA=1,"
					else
						sSQL = sSQL & "uspsFSA=0,"
					end if
					if getpost("methoduse"&index)="ON" then
						sSQL = sSQL & "uspsUseMethod=1 WHERE uspsID="&index
					else
						sSQL = sSQL & "uspsUseMethod=0 WHERE uspsID="&index
					end if
					ect_query(sSQL)
				end if
			next
		end if
	end if
	print "<meta http-equiv=""refresh"" content=""2; url=adminuspsmeths.asp"">"
else
	sSQL = "SELECT uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal,uspsFSA FROM uspsmethods "
	if shipType=3 then
		sSQL = sSQL & " WHERE uspsID < 100"
	elseif shipType=4 then
		sSQL = sSQL & " WHERE uspsID > 100 AND uspsID < 200"
	elseif shipType=6 then
		sSQL = sSQL & " WHERE uspsID > 200 AND uspsID < 300"
	elseif shipType=7 then
		sSQL = sSQL & " WHERE uspsID > 300 AND uspsID < 400"
	elseif shipType=8 then
		sSQL = sSQL & " WHERE uspsID > 400 AND uspsID < 500"
	elseif shipType=9 then
		sSQL = sSQL & " WHERE uspsID > 500 AND uspsID < 600"
	elseif shipType=10 then
		sSQL = sSQL & " WHERE uspsID > 600 AND uspsID < 700"
	end if
	sSQL = sSQL & " ORDER BY uspsLocal DESC, uspsOrder, uspsShowAs, uspsID"
	rs.open sSQL,cnn,0,1
	allmethods=rs.getrows
	rs.close
end if
cpurl="https://" & IIfVr(canadaposttestmode,"ct.","") & "soa-gw.canadapost.ca/ot/soap/merchant/registration"
if getget("token-id")<>"" AND getget("registration-status")<>"" then
	print "<h2>Canada Post Registration</h2>"
	if getget("registration-status")="SUCCESS" then
		sXML="<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:reg=""http://www.canadapost.ca/ws/soap/merchant/registration"">" & _
		"<soapenv:Header><wsse:Security soapenv:mustUnderstand=""1"" xmlns:wsse=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"" xmlns:wsu=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd""><wsse:UsernameToken><wsse:Username>" & IIfVr(canadaposttestmode,"3e726c38d754ea80","2de7ca2bc2f0a552") & "</wsse:Username><wsse:Password>" & IIfVr(canadaposttestmode,"a47e23e9d34ee61fda2199","1d3ac063ca9baccdc2ca69") & "</wsse:Password></wsse:UsernameToken></wsse:Security></soapenv:Header>" & _
		"<soapenv:Body><reg:get-merchant-registration-info-request><locale>EN</locale><token-id>" & getget("token-id") & "</token-id></reg:get-merchant-registration-info-request></soapenv:Body></soapenv:Envelope>"
		iscanadapost=TRUE
		CanadaPostCalculate=callxmlfunction(cpurl, sXML, xmlres, "", "WinHTTP.WinHTTPRequest.5.1", errormsg, FALSE)
		iscanadapost=FALSE
		' print replace(sXML,"<","<br />&lt;")&"<hr>"
		' print replace(xmlres,"<","<br />&lt;")&"<br>"
		set xmlDoc=Server.CreateObject("MSXML2.DOMDocument")
		xmlDoc.validateOnParse=FALSE
		xmlDoc.loadXML(xmlres)
		set obj1=xmlDoc.getElementsByTagName("customer-number")
		if obj1.length>0 then customernumber = obj1.item(0).firstChild.nodeValue
		set obj1=xmlDoc.getElementsByTagName("merchant-username")
		if obj1.length>0 then merchantusername = obj1.item(0).firstChild.nodeValue
		set obj1=xmlDoc.getElementsByTagName("merchant-password")
		if obj1.length>0 then merchantpassword = obj1.item(0).firstChild.nodeValue
		sSQL="UPDATE adminshipping SET adminCanPostUser='"&escape_string(customernumber)&"',adminCanPostLogin='"&escape_string(merchantusername)&"',adminCanPostPass='"&escape_string(merchantpassword)&"' WHERE adminShipID=1"
		ect_query(sSQL)
		print "<div style=""text-align:center;margin:15px"">The Canada Post Registration system has completed successfully.</div>"
		print "<div style=""text-align:center;margin:15px""><a href=""admin.asp""><strong>"&yyAdmHom&"</strong></a></div>"
	else
		print "<div style=""text-align:center;font-weight:bold;margin:15px"">Sorry - An error occurred!</div>"
	end if
elseif getget("canadapost")="register" then
	print "<h2>Canada Post Registration</h2>"
	thetokenid=""
	sXML="<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:reg=""http://www.canadapost.ca/ws/soap/merchant/registration"">" & _
	"<soapenv:Header><wsse:Security soapenv:mustUnderstand=""1"" xmlns:wsse=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"" xmlns:wsu=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd""><wsse:UsernameToken><wsse:Username>" & IIfVr(canadaposttestmode,"3e726c38d754ea80","2de7ca2bc2f0a552") & "</wsse:Username><wsse:Password>" & IIfVr(canadaposttestmode,"a47e23e9d34ee61fda2199","1d3ac063ca9baccdc2ca69") & "</wsse:Password></wsse:UsernameToken></wsse:Security></soapenv:Header>" & _
	"<soapenv:Body><reg:get-merchant-registration-token-request></reg:get-merchant-registration-token-request></soapenv:Body></soapenv:Envelope>"
	iscanadapost=TRUE
	CanadaPostCalculate=callxmlfunction(cpurl, sXML, xmlres, "", "WinHTTP.WinHTTPRequest.5.1", errormsg, FALSE)
	iscanadapost=FALSE
'	print cpurl & "<br>"
'	print replace(sXML,"<","<br />&lt;")&"<hr>"
'	print replace(xmlres,"<","<br />&lt;")&"<br>"
	set xmlDoc=Server.CreateObject("MSXML2.DOMDocument")
	xmlDoc.validateOnParse=FALSE
	xmlDoc.loadXML(xmlres)
	set obj1=xmlDoc.getElementsByTagName("token-id")
	if obj1.length>0 then thetokenid = obj1.item(0).firstChild.nodeValue
	if thetokenid="" then
		set obj1=xmlDoc.getElementsByTagName("faultcode")
		if obj1.length>0 then faultcode = obj1.item(0).firstChild.nodeValue
		set obj1=xmlDoc.getElementsByTagName("faultstring")
		if obj1.length>0 then faultstring = obj1.item(0).firstChild.nodeValue
		print "<div style=""text-align:center;margin:15px"">There was an error connecting with the Canada Post Registration Server. Please ask your host to make sure the following URL is not blocked by the server firewall...</div>"
		print "<div style=""text-align:center;margin:15px"">" & cpurl & "</div>"
		print "<div style=""text-align:center;margin:15px"">If you have done this and still get an error, please quote the following when contacting <a href=""https://www.ecommercetemplates.com/support"" target=""_blank"">support at Ecommerce Templates</a>.</div>"
		print "<div style=""text-align:center;margin:15px"">" & faultcode & " : " & faultstring & "</div>"
	else
%>
<form method="post" id="canposform" action="https://www.canadapost.ca/cpotools/apps/drc/merchant">
<div style="text-align:center;margin:15px">In a few seconds you will be taken to the Canada Post website to complete the registration process...</div>
<div style="text-align:center;margin:15px">If that does not happen automatically, please <a href="javascript:document.getElementById('canposform').submit()">Click Here</a></div>
<input type="hidden" name="return-url" value="<%=IIfVr(request.servervariables("HTTPS")="on","https://","http://")&request.servervariables("HTTP_HOST")&request.servervariables("URL")%>" />
<input type="hidden" name="token-id" value="<%=thetokenid%>" />
<input type="hidden" name="platform-id" value="0008107483" />
</form>
<script>
setTimeout('document.getElementById("canposform").submit()', 4000);
</script>
<%	end if
elseif getget("royalmail")="setup" then %>
<p>&nbsp;</p>
<p align="center">Proceeding will replace all your weight based shipping tables with Royal Mail 2013 rates</p>
<p align="center">Product weights are assumed to be in metric (kg)</p>
<p align="center">Please note, clicking below will wipe all your current postal zone inforamtion and cannot be undone.</p>
<p>&nbsp;</p>
<form method="post" action="adminuspsmeths.asp">
<input type="hidden" name="royalmail" value="dosetup" />
<p align="center">
<table border="0" width="100%">
<tr><td align="right" width="25%"><input type="checkbox" name="addrecorded" value="ON" /></td><td align="left">Add Recorded Signed For option to First, Second and Standard Parcel rates? (&pound;1.10 extra)</td></tr>
<tr><td align="right"><input type="checkbox" name="addinternationalsigned" value="ON" /></td><td align="left">Add International Signed For option to International rates? (&pound;5.30 extra)</td></tr>
<tr><td align="right"><input type="checkbox" name="addspecial" value="ON" /></td><td align="left">Add Special Delivery 9am and 1pm Services?</td></tr>
</table>
</p>
<p>&nbsp;</p>
<p align="center"><input type="submit" value="Apply Royal Mail Rates" /></p>
</form>
<p>&nbsp;</p>
<%
elseif getpost("royalmail")="dosetup" then %>
<p>&nbsp;</p>
<p align="center">The process has completed successfully</p>
<p align="center">You still need to select "Weight Based Shipping" as your shipping method in the admin main settings page.</p>
<p>&nbsp;</p>
<%	addrecorded=FALSE
	addinternationalsigned=FALSE
	if getpost("addrecorded")="ON" then addrecorded=TRUE
	if getpost("addinternationalsigned")="ON" then addinternationalsigned=TRUE
	sub doaddrate(zczone,zcweight,zcrate,zcrate2,zcrate3,zcrate4)
		if zczone=1 AND addrecorded AND zcrate>0 then zcrate=zcrate+1.10
		if zczone=1 AND addrecorded AND zcrate2>0 then zcrate2=zcrate2+1.10
		if zczone>1 AND addinternationalsigned AND zcrate>0 then zcrate=zcrate+5.30
		ect_query("INSERT INTO zonecharges (zcZone,zcWeight,zcRate,zcRate2,zcRate3,zcRate4) VALUES (" & zczone & "," & zcweight & "," & zcrate & "," & zcrate2 & "," & zcrate3 & "," & zcrate4 & ")")
	end sub
	function addpostalzone(zoneid,pzname,pzmultishipping,pzmethodname1,pzmethodname2,pzmethodname3,pzmethodname4)
		sSQL="UPDATE postalzones SET pzName='" & pzname & "',pzMultiShipping=" & pzmultishipping & ",pzMethodName1='" & pzmethodname1 & "',pzMethodName2='" & pzmethodname2 & "',pzMethodName3='" & pzmethodname3 & "',pzMethodName4='" & pzmethodname4 & "' WHERE pzID=" & zoneid
		ect_query(sSQL)
		addpostalzone=zoneid
	end function

	ect_query("DELETE FROM zonecharges")
	ect_query("UPDATE admin SET adminUSZones=0")
	ect_query("UPDATE countries SET countryZone=99999")
	
	zoneid = addpostalzone(1,"Great Britain",IIfVr(getpost("addspecial")="ON",3,1),"First Class","Second Class","Special Delivery Next Day (1:00pm)","Special Delivery Next Day (9:00am)")
	call doaddrate(1,0.10, 0.90, 0.69, 6.22,17.64)
	call doaddrate(1,0.25, 1.20, 1.10, 6.22,17.64)
	call doaddrate(1,0.50, 1.60, 1.40, 6.95,19.92)
	call doaddrate(1,0.75, 2.30, 1.90, 6.95,19.92)
	call doaddrate(1,1.00, 3.00, 2.60, 8.25,21.60)
	call doaddrate(1,2.00, 6.85, 5.60,11.00,26.16)
	call doaddrate(1,5.00,15.10,13.35,11.00,-99999)
	call doaddrate(1,10.0,21.25,19.65,25.80,-99999)
	call doaddrate(1,20.0,32.40,27.70,40.00,-99999)
	call doaddrate(1,20.0001,-99999,-99999,-99999,-99999)
	 
	ect_query("UPDATE countries SET countryZone=" & zoneid & " WHERE countryID IN (107,142,201,214,216)")
	
	zoneid = addpostalzone(2,"Europe",0,"Standard Shipping","","","")
	call doaddrate(2,0.10, 3.00,0,0,0)
	call doaddrate(2,0.25, 3.50,0,0,0)
	call doaddrate(2,0.50, 4.95,0,0,0)
	call doaddrate(2,0.75, 6.40,0,0,0)
	call doaddrate(2,1.00, 7.85,0,0,0)
	call doaddrate(2,1.25, 9.30,0,0,0)
	call doaddrate(2,1.50,10.75,0,0,0)
	call doaddrate(2,1.75,12.20,0,0,0)
	call doaddrate(2,2.00,13.65,0,0,0)
	for indexar=1 to 12
		call doaddrate(2,2+(indexar/4.0),13.65+(indexar*1.45),0,0,0)
	next
	call doaddrate(2,5.01,-99999,0,0,0)
	ect_query("UPDATE countries SET countryZone=" & zoneid & " WHERE countryID IN (4,6,12,15,16,21,22,28,32,46,48,49,50,59,62,64,65,70,71,73,74,75,86,87,91,93,97,103,108,109,110,112,118,123,124,133,143,152,153,156,157,163,175,170,171,175,182,183,186,194,195,199,203,205,217,218,219,221,223)")
	
	zoneid = addpostalzone(3,"World Zone 1",0,"Standard Shipping","","","")
	call doaddrate(3,0.10, 3.50,0,0,0)
	call doaddrate(3,0.25, 4.50,0,0,0)
	call doaddrate(3,0.50, 7.20,0,0,0)
	call doaddrate(3,0.75, 9.90,0,0,0)
	call doaddrate(3,1.00,12.60,0,0,0)
	call doaddrate(3,1.25,15.30,0,0,0)
	call doaddrate(3,1.50,18.00,0,0,0)
	call doaddrate(3,1.75,20.70,0,0,0)
	call doaddrate(3,2.00,23.40,0,0,0)
	for indexar=1 to 12
		call doaddrate(3,2+(indexar/4.0),23.4+(indexar*2.7),0,0,0)
	next
	call doaddrate(3,5.01,-99999,0,0,0)
	ect_query("UPDATE countries SET countryZone=" & zoneid & " WHERE countryZone=99999")
	
	zoneid = addpostalzone(4,"World Zone 2",0,"Standard Shipping","","","")
	call doaddrate(4,0.10, 3.50,0,0,0)
	call doaddrate(4,0.25, 4.70,0,0,0)
	call doaddrate(4,0.50, 7.55,0,0,0)
	call doaddrate(4,0.75,10.40,0,0,0)
	call doaddrate(4,1.00,13.25,0,0,0)
	call doaddrate(4,1.25,16.10,0,0,0)
	call doaddrate(4,1.50,18.95,0,0,0)
	call doaddrate(4,1.75,21.80,0,0,0)
	call doaddrate(4,2.00,24.65,0,0,0)
	for indexar=1 to 12
		call doaddrate(4,2+(indexar/4.0),24.65+(indexar*2.85),0,0,0)
	next
	call doaddrate(4,5.01,-99999,0,0,0)
	ect_query("UPDATE countries SET countryZone=" & zoneid & " WHERE countryID IN (14,63,67,99,111,131,135,136,140,141,147,151,162,169,172,190,191,197)")
elseif getget("royalmail")="register" then %>
	<h2>Royal Mail Registration</h2>
	<div style="text-align:center;margin:50px">
	
	Registering with the Royal Mail is not necessary as there is no Online Shipping Rates service and instead we have setup the rates using our Weight Based Shipping tables.<br /><br />
	To apply Royal Mail rates to your weight based shipping tables, please click the button below.<br /><br />
	<input type="button" value="Setup Royal Mail Rates" onclick="document.location='adminuspsmeths.asp?royalmail=setup'" />
	<br /><br >
	After doing this you must select &quot;Weight Based Shipping&quot; from the admin main settings page.<br /><br />
	There are more details about setting up Royal Mail rates here...<br />
	<a href="https://www.ecommercetemplates.com/help/royal-mail.asp" target="_blank">https://www.ecommercetemplates.com/help/royal-mail.asp</a>
	
	</div>
<%
elseif getget("dhl")="register" then %>
	<h2>DHL Registration</h2>
	<div style="text-align:center;margin:50px">
	
	To register with DHL you need to contact your DHL Account Manager to apply for a Site ID and Password. Once you have received these, return to the shipping methods admin page here 
	in the Ecommerce Plus admin and click on the &quot;DHL Admin&quot; button where you can enter these along with your DHL Account Number.<br /><br />
	There are more details about setting up shipping rates with DHL here...<br />
	<a href="https://www.ecommercetemplates.com/help/dhl.asp" target="_blank">https://www.ecommercetemplates.com/help/dhl.asp</a>
	
	</div>
<%
elseif getget("auspost")="register" then %>
	<h2>Australia Post Registration</h2>
	<div style="text-align:center;margin:50px">
	
	To register with Australia Post please click on the link below. This will take you to the Australia Post website where you can register for an API Key. Once you 
	have received your API Key, return to the shipping methods admin page here in the Ecommerce Plus admin and click on the &quot;Australia Post Admin&quot; button where you can enter your API Key.
	<br /><br />
	<a href="https://auspost.com.au/forms/pacpcs-registration.html" target="_blank">https://auspost.com.au/forms/pacpcs-registration.html</a>
	<br /><br />
	
	</div>
<%
elseif getget("admin")<>"" then
	sSQL = "SELECT adminUSPSUser,adminUPSUser,adminUPSPw,adminUPSAccess,adminUPSAccount,adminUPSNegotiated,adminCanPostUser,DHLSiteID,DHLSitePW,DHLAccountNo,AusPostAPI,smartPostHub,shipStationUser,shipStationPass FROM adminshipping WHERE adminShipID=1"
	rs.open sSQL,cnn,0,1
%>
		  <form method="post" action="adminuspsmeths.asp">
<%	call writehiddenvar("doadmin", "1")
	call writehiddenvar("method", getget("admin"))
	call writehiddenvar("posted", "1") %>
			<table width="100%" border="0" cellspacing="2" cellpadding="3">
<%	if getget("admin")="3" then %>
			  <tr>
                <td colspan="2" align="center"><h2>USPS Admin</h2><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yyIfUSPS%><br /></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong><%=yyUname%>: </strong></td>
				<td width="50%" align="left"><input type="text" size="15" name="adminUSPSUser" value="<%=rs("adminUSPSUser")%>" /></td>
			  </tr>
<%	elseif getget("admin")="4" then %>
			  <tr>
                <td colspan="2" align="center"><h2>UPS Admin</h2><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><p>To obtain your UPS Rate Code you need to use the registration form <a href="adminupslicense.asp"><strong>here</strong></a>.</p>
				<p>To use UPS Negotiated Rates, you need to register first and specify your UPS Shipper Number in the registration form. Then forward your UPS Rate Code and Shipper Number to your UPS Account Manager who will enable UPS Negotiated Rates once approved.</p></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong>UPS Rate Code: </strong></td>
				<td width="50%" align="left"><%=upsdecode(rs("adminUPSUser"), "")%></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong>UPS Shipper Number: </strong></td>
				<td width="50%" align="left"><%=rs("adminUPSAccount")%></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong>UPS Access Key: </strong></td>
				<td width="50%" align="left"><%=rs("adminUPSAccess")%></td>
			  </tr>
			  <tr>
				<td width="50%" align="right"><strong>Use Negotiated Rates: </strong></td>
				<td width="50%" align="left"><select size="1" name="UPSNegotiated">
					<option value="0">Use Published Rates</option>
<%		if trim(rs("adminUPSUser")&"")<>"" AND trim(rs("adminUPSAccount")&"")<>"" then print "<option value=""1"""&IIfVr(cint(rs("adminUPSNegotiated"))<>0, " selected=""selected""", "")&">Use Negotiated Rates</option>" %>
					</select>
				</td>
			  </tr>
<%	elseif getget("admin")="6" then %>
			  <tr>
                <td colspan="2" align="center"><h2><%=yyCanPos%> Admin</h2><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><hr width="70%" /><%=yyEnMerI%></td>
			  </tr>
			  <tr>
				<td colspan="2" align="center"><strong><%=yyRetID%>: </strong><input type="text" size="36" name="adminCanPostUser" value="<%=rs("adminCanPostUser")%>" /></td>
			  </tr>
<%	elseif getget("admin")="9" then %>
			  <tr>
                <td colspan="2" align="center"><h2>DHL Admin</h2><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2">&nbsp;</td>
			  </tr>
			  <tr>
				<td align="right" width="45%"><strong>Site ID: </strong></td><td><input type="text" size="36" name="DHLSiteID" value="<%=rs("DHLSiteID")%>" /></td>
			  </tr>
			  <tr>
				<td align="right"><strong>Site Password: </strong></td><td><input type="password" size="36" name="DHLSitePW" value="<%=rs("DHLSitePW")%>" /></td>
			  </tr>
			  <tr>
				<td align="right"><strong>Account Number: </strong></td><td><input type="text" size="36" name="DHLAccountNo" value="<%=rs("DHLAccountNo")%>" /></td>
			  </tr>
<%	elseif getget("admin")="10" then %>
			  <tr>
                <td colspan="2" align="center"><h2>Australia Post Admin</h2><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2">&nbsp;</td>
			  </tr>
			  <tr>
				<td align="right" width="40%"><strong>API Key: </strong></td><td><input type="text" size="36" name="AusPostAPI" value="<%=rs("AusPostAPI")%>" /></td>
			  </tr>
<%	elseif getget("admin")="8" then %>
			  <tr>
                <td colspan="2" align="center"><h2>FedEx SmartPost Admin</h2><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2">&nbsp;</td>
			  </tr>
			  <tr>
				<td align="right" width="40%"><strong>SmartPost Hub ID: </strong></td><td><input type="text" size="36" name="smartPostHub" value="<%=rs("smartPostHub")%>" /></td>
			  </tr>
<%	elseif getget("admin")="99" then %>
			  <tr>
                <td colspan="2" align="center"><h2>Ship Station Admin</h2><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2">&nbsp;</td>
			  </tr>
			  <tr>
				<td align="right" width="40%"><strong>Ship Station Username: </strong></td><td><input type="text" size="36" name="shipStationUser" value="<%=rs("shipStationUser")%>" /></td>
			  </tr>
			  <tr>
				<td align="right" width="40%"><strong>Ship Station Password: </strong></td><td><input type="password" size="36" name="shipStationPass" value="<%=rs("shipStationPass")%>" /></td>
			  </tr>
<%	end if %>
			  <tr>
				<td width="100%" align="center" colspan="2">&nbsp;</td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2"><input type="submit" value="<%=yySubmit%>" /> <input type="reset" value="<%=yyReset%>" /></td>
			  </tr>
<%	if getget("admin")="4" then %>
			  <tr>
				<td width="100%" align="center" colspan="2"><br /><span style="font-size:10px">Please note: Subsequent registrations for UPS OnLine&reg; Tools will change the UPS Rate Code
within this application. In the event Negotiated Rates functionality was enabled under a previous UPS Rate Code, the
Negotiated Rates functionality will be disabled.</span></td>
			  </tr>
<%	end if %>
			  <tr>
				<td width="100%" align="center" colspan="2"><br />&nbsp;<br />&nbsp;<br /><a href="adminuspsmeths.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
			  </tr>
			</table>
		  </form>
<%	rs.close
elseif method="" then %>
			<table width="100%" border="0" cellspacing="2" cellpadding="3">
			  <tr>
                <td align="center"><h2><%=yyShpAdm%></h2>
			<table width="100%" class="stackable admin-table-a sta-white">
			  <tr>
				<th class="cobhl" height="30"><strong>Shipping Carrier</strong></th>
				<th class="cobhl"><strong>Registration</strong></th>
				<th class="cobhl"><strong>Administration</strong></th>
				<th class="cobhl"><strong>Shipping Method</strong></th>
			  </tr>
			  <tr>
				<td class="cobhl" height="30"><strong>Ship Station</strong></td>
				<td class="cobll">&nbsp;</td>
				<td class="cobll"><input type="button" value="Ship Station Admin" onclick="document.location='adminuspsmeths.asp?admin=99'" /></td>
				<td class="cobll">&nbsp;</td>
			  </tr>
			  <tr>
				<td class="cobhl" height="30"><strong>Australia Post</strong></td>
				<td class="cobll"><input type="button" value="<%=replace(yyRegUPS,"UPS","Australia Post")%>" onclick="document.location='adminuspsmeths.asp?auspost=register'" /></td>
				<td class="cobll"><input type="button" value="Australia Post Admin" onclick="document.location='adminuspsmeths.asp?admin=10'" /></td>
				<td class="cobll"><input type="button" value="<%=yyEdit&" "&yyShpMet%>" onclick="document.location='adminuspsmeths.asp?method=10'" /></td>
			  </tr>
			  <tr>
				<td class="cobhl" height="30"><strong>Canada Post</strong></td>
				<td class="cobll"><input type="button" value="<%=replace(yyRegUPS,"UPS","Canada Post")%>" onclick="document.location='adminuspsmeths.asp?canadapost=register'" /></td>
				<td class="cobll"><input type="button" value="Canada Post Admin" onclick="document.location='adminuspsmeths.asp?admin=6'" /></td>
				<td class="cobll"><input type="button" value="<%=yyEdit&" "&yyShpMet%>" onclick="document.location='adminuspsmeths.asp?method=6'" /></td>
			  </tr>
			  <tr>
				<td class="cobhl" height="30"><strong>DHL</strong></td>
				<td class="cobll"><input type="button" value="<%=replace(yyRegUPS,"UPS","DHL")%>" onclick="document.location='adminuspsmeths.asp?dhl=register'" /></td>
				<td class="cobll"><input type="button" value="DHL Admin" onclick="document.location='adminuspsmeths.asp?admin=9'" /></td>
				<td class="cobll"><input type="button" value="<%=yyEdit&" "&yyShpMet%>" onclick="document.location='adminuspsmeths.asp?method=9'" /></td>
			  </tr>
			  <tr>
				<td class="cobhl" height="30"><strong>FedEx</strong></td>
				<td class="cobll"><input type="button" value="<%=replace(yyRegUPS,"UPS","FedEx")%>" onclick="document.location='adminfedexlicense.asp'" /></td>
				<td class="cobll">&nbsp;</td>
				<td class="cobll"><input type="button" value="<%=yyEdit&" "&yyShpMet%>" onclick="document.location='adminuspsmeths.asp?method=7'" /></td>
			  </tr>
			  <tr>
				<td class="cobhl" height="30"><strong>FedEx SmartPost</strong></td>
				<td class="cobll"><input type="button" value="<%=replace(yyRegUPS,"UPS","FedEx")%>" onclick="document.location='adminfedexlicense.asp'" /></td>
				<td class="cobll"><input type="button" value="FedEx SmartPost Admin" onclick="document.location='adminuspsmeths.asp?admin=8'" /></td>
				<td class="cobll"><input type="button" value="<%=yyEdit&" "&yyShpMet%>" onclick="document.location='adminuspsmeths.asp?method=8'" /></td>
			  </tr>
			  <tr>
				<td class="cobhl" height="30"><strong>Royal Mail</strong></td>
				<td class="cobll"><input type="button" value="<%=replace(yyRegUPS,"UPS","Royal Mail")%>" onclick="document.location='adminuspsmeths.asp?royalmail=register'" /></td>
				<td class="cobll"><input type="button" value="Setup Royal Mail Rates" onclick="document.location='adminuspsmeths.asp?royalmail=setup'" /></td>
				<td class="cobll"><input type="button" value="<%=yyEdit&" Postal Zones"%>" onclick="document.location='adminzones.asp'" /></td>
			  </tr>
			  <tr>
				<td class="cobhl" height="30"><strong>UPS</strong></td>
				<td class="cobll"><input type="button" value="<%=yyRegUPS%>" onclick="document.location='adminupslicense.asp'" /></td>
				<td class="cobll"><input type="button" value="UPS Admin" onclick="document.location='adminuspsmeths.asp?admin=4'" /></td>
				<td class="cobll"><input type="button" value="<%=yyEdit&" "&yyShpMet%>" onclick="document.location='adminuspsmeths.asp?method=4'" /></td>
			  </tr>
			  <tr>
				<td class="cobhl" height="30"><strong>USPS</strong></td>
				<td class="cobll"><input type="button" value="Register with USPS" onclick="window.open('https://reg.usps.com/register','USPSSignup','')" /></strong></td>
				<td class="cobll"><input type="button" value="USPS Admin" onclick="document.location='adminuspsmeths.asp?admin=3'" /></td>
				<td class="cobll"><input type="button" value="<%=yyEdit&" "&yyShpMet%>" onclick="document.location='adminuspsmeths.asp?method=3'" /></td>
			  </tr>
			  <tr>
				<td class="cobhl" height="30"><strong>Weight / Price Based</strong></td>
				<td class="cobll">&nbsp; </td>
				<td class="cobll">&nbsp; </td>
				<td class="cobll"><input type="button" value="<%=yyEdit&" Postal Zones"%>" onclick="document.location='adminzones.asp'" /></td>
			  </tr>
			</table>
			
			<br />&nbsp;<br />&nbsp;<br /><a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;
			
				</td>
			  </tr>
			</table>
			<br />&nbsp;
            
<%
elseif getpost("posted")="1" AND success then %>
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <a href="adminuspsmeths.asp"><strong><%=yyClkHer%></strong></a>.<br /><br />&nbsp;
                </td>
			  </tr>
			</table>
<%
else %>
		  <form method="post" action="adminuspsmeths.asp">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="method" value="<%=method%>" />
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
			  <tr> 
                <td width="100%" colspan="5" align="center"><strong><%=yyUsUpd & " " & shipmet & " " & yyShpMet%>.</strong><br />&nbsp;</td>
			  </tr>
<%	if not success then %>
			  <tr> 
                <td width="100%" colspan="5" align="center"><br /><span style="color:#FF0000"><%=errmsg%></span>
                </td>
			  </tr>
<%	end if %>
			  <tr>
				<td colspan="5"><ul><%
	if shipType=4 then
		print "<li><span style=""font-size:10px"">"&yyUSS3&"</span></li>"
	elseif shipType=3 then
		print "<li><span style=""font-size:10px"">"&yyUSS1&"</span></li>"
	end if %>
			<li><span style="font-size:10px">You can use this page to set which <%=shipmet%> shipping methods qualify for free shipping discounts by checking the FSA (Free Shipping Available) checkbox.</span></li>
				<li><span style="font-size:10px"><%
			print replace(yyUSS2,"USPS",shipmet)
			if shipType=3 then %>
				<a href="http://www.usps.com">http://www.usps.com</a>
<%			elseif shipType=4 then %>
				<a href="http://www.ups.com">http://www.ups.com</a>
<%			elseif shipType=6 then %>
				<a href="http://www.canadapost.ca" target="_blank">http://www.canadapost.ca</a>.
<%			elseif shipType=9 then %>
				<a href="http://www.dhl.com" target="_blank">http://www.dhl.com</a>.
<%			elseif shipType=10 then %>
				<a href="http://auspost.com.au" target="_blank">http://auspost.com.au</a>.
<%			else %>
				<a href="http://www.fedex.com" target="_blank">http://www.fedex.com</a>.
<%			end if %>
				</span></li>
				</ul></td>
			  </tr>
<%	if shipType=3 OR shipType=10 then
		for index=0 to UBOUND(allmethods,2) %>
			  <tr>
			    <td align="right"><%=IIfVr(shipType=10,"AusPost Method",yyUSPSMe)%>:</td>
				<td align="left"><span style="font-size:10px;font-weight:bold"><%
				if shipType=3 then
					if allmethods(0,index)="1" then
						print "Express Mail"
					elseif allmethods(0,index)="2" then
						print "Priority Mail"
					elseif allmethods(0,index)="3" then
						print "Parcel Select Ground"
					elseif allmethods(0,index)="14" then
						print "Media Mail"
					elseif allmethods(0,index)="15" then
						print "Bound Printed Matter"
					elseif allmethods(0,index)="16" then
						print "First Class Mail"
					elseif allmethods(0,index)="17" then
						print "First Class Commercial"
					elseif allmethods(0,index)="18" then
						print "Priority Commercial"
					elseif allmethods(0,index)="19" then
						print "Priority CPP"
					elseif allmethods(0,index)="20" then
						print "Priority Mail Express Commercial"
					elseif allmethods(0,index)="21" then
						print "Priority Mail Express CPP"
					elseif allmethods(0,index)="22" then
						print "Retail Ground"
					elseif allmethods(0,index)="30" then
						print "Global Express Guaranteed Document"
					elseif allmethods(0,index)="31" then
						print "Global Express Guaranteed Non-Document Rectangular"
					elseif allmethods(0,index)="32" then
						print "Global Express Guaranteed Non-Document Non-Rectangular"
					elseif allmethods(0,index)="33" then
						print "Express Mail International (EMS)"
					elseif allmethods(0,index)="34" then
						print "Express Mail International (EMS) Flat Rate Envelope"
					elseif allmethods(0,index)="35" then
						print "Priority Mail International"
					elseif allmethods(0,index)="36" then
						print "Priority Mail International Flat Rate Envelope"
					elseif allmethods(0,index)="37" then
						print "Priority Mail International Regular Flat-Rate Boxes"
					elseif allmethods(0,index)="38" then
						print "First Class Mail International Letters"
					elseif allmethods(0,index)="39" then
						print "First Class Mail International Large Envelope"
					elseif allmethods(0,index)="40" then
						print "First Class Mail International Package"
					elseif allmethods(0,index)="41" then
						print "Priority Mail International Large Flat-Rate Box"
					elseif allmethods(0,index)="42" then
						print "Priority Mail International Small Flat Rate Box"
					elseif allmethods(0,index)="43" then
						print "Express Mail International Legal Flat Rate Envelope"
					elseif allmethods(0,index)="44" then
						print "Priority Mail International Small Flat Rate Envelope"
					elseif allmethods(0,index)="45" then
						print "Priority Mail International DVD Flat Rate Box"
					elseif allmethods(0,index)="46" then
						print "Express Mail International Flat Rate Box"
					end if
				else ' Australia Post
				'	if allmethods(1,index)="AUS_PARCEL_REGULAR" then
				'		print "Parcel Post"
				'	elseif allmethods(1,index)="AUS_PARCEL_REGULAR_SATCHEL_3KG" then
				'		print "Parcel Post Medium (3Kg) Satchel"
				'	elseif allmethods(1,index)="AUS_PARCEL_EXPRESS_SATCHEL_3KG" then
				'		print "Express Post Medium (3Kg) Satchel"
				'	elseif allmethods(1,index)="AUS_PARCEL_EXPRESS" then
				'		print "Express Post"
				'	else
					if allmethods(1,index)="INT_PARCEL_COR_OWN_PACKAGING" then
						print "International Post Courier"
					elseif allmethods(1,index)="INT_PARCEL_EXP_OWN_PACKAGING" then
						print "International Post Express"
					elseif allmethods(1,index)="INTL_SERVICE_ECI_D" then
						print "Express Courier International Documents"
					elseif allmethods(1,index)="INT_PARCEL_STD_OWN_PACKAGING" then
						print "International Post Standard"
					elseif allmethods(1,index)="INT_PARCEL_AIR_OWN_PACKAGING" then
						print "International Air"
					elseif allmethods(1,index)="INT_PARCEL_SEA_OWN_PACKAGING" then
						print "International Sea"
					elseif allmethods(1,index)="INT_LETTER_REG_SMALL" then
						print "International Economy Letter Small"
					elseif allmethods(1,index)="INT_LETTER_REG_LARGE" then
						print "International Economy Letter Large"
					else
						print allmethods(1,index)
					end if
				end if%></span></td>
				<td align="center"><%=yyUseMet%></td>
				<td align="center"><acronym title="<%=yyFSApp%>"><%=yyFSA%></acronym></td>
				<td align="center"><%=yyType%></td>
			  </tr>
			  <tr>
				<td align="right"><%=yyShwAs%>:</td>
			    <td align="left"><input type="text" name="methodshow<%=allmethods(0,index)%>" value="<%=allmethods(2,index)%>" size="36" /></td>
				<td align="center"><input type="checkbox" name="methoduse<%=allmethods(0,index)%>" value="ON" <% if Int(allmethods(3,index))=1 then print "checked=""checked"""%> /></td>
				<td align="center"><input type="checkbox" name="methodfsa<%=allmethods(0,index)%>" value="ON" <% if Int(allmethods(5,index))=1 then print "checked=""checked"""%> /></td>
				<td align="center"><%if Int(allmethods(4,index))=1 then print "<span style=""color:#FF0000"">Domestic</span>" else print "<span style=""color:#0000FF"">Internat.</span>"%></td>
			  </tr>
			  <tr>
				<td colspan="5" align="center"><hr width="80%" /></td>
			  </tr>
<%		next
	else
		for index=0 to UBOUND(allmethods,2) %>
			  <tr>
			    <td align="right"><input type="hidden" name="methodshow<%=allmethods(0,index)%>" value="1" /><strong><%=yyShipMe%>:</strong></td>
				<td align="left"> &nbsp; <%=allmethods(2,index) & checkisdocument(shipType,allmethods(1,index)) %></td>
				<td align="center"><strong><%=IIfVr(shipType=4 OR shipType=6 OR shipType=7 OR shipType=9,yyUseMet,"&nbsp;")%></strong></td>
				<td align="center"><acronym title="<%=yyFSApp%>"><%=yyFSA%></acronym></td>
				<td>&nbsp;</td>
			  </tr>
			  <tr>
				<td colspan="2">&nbsp;</td>
				<td align="center"><input type="<%=IIfVr(shipType=4 OR shipType=6 OR shipType=7 OR shipType=9,"checkbox","hidden")%>" name="methoduse<%=allmethods(0,index)%>" value="ON" <% if Int(allmethods(3,index))=1 then print "checked=""checked"""%> /></td>
				<td align="center"><input type="checkbox" name="methodfsa<%=allmethods(0,index)%>" value="ON" <% if Int(allmethods(5,index))=1 then print "checked=""checked"""%> /></td>
				<td>&nbsp;</td>
			  </tr>
			  <tr>
				<td colspan="5" align="center"><hr width="80%" /></td>
			  </tr>
<%		next
	end if %>
			  <tr> 
                <td width="100%" colspan="5" align="center"><br /><input type="submit" value="<%=yySubmit%>" /><br />&nbsp;</td>
			  </tr>
            </table>
		  </form>
<%
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>