<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
Response.Charset="8859-1"
success=true
Set rs=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
countryCode=origCountryCode
if upstestmode=TRUE then upsurl="wwwcie.ups.com" else upsurl="onlinetools.ups.com"
if upstestmode then
	registerurl="https://wwwcie.ups.com/webservices/Registration"
else
	registerurl="https://onlinetools.ups.com/webservices/Registration"
end if
allowedcountries="'AR','AU','AT','BE','BR','CA','CL','CN','CO','CR','DK','DO','FI','FR','DE','GR','GT','HK','IN','IE','IL','IT','JP','MY','MX','NL','NZ','NO','PA','PH','PT','PR','SG','KR','ES','SE','CH','TW','TH','GB','US'"
Function ParseUPSLicenseOutput(xmlDoc, rootNodeName, byRef thetext, byRef errormsg)
Dim noError, nodeList, e, i, j, k, l, n, t, t2
	noError=True
	errormsg=""
	gotxml=false
	thetext=""
	if dumpshippingxml then print replace(replace(xmlDoc.xml,"</","&lt;/"),"<","<br />&lt;")&"<hr />"
	set errnodes=xmlDoc.getElementsByTagName("err:Severity")
	if errnodes.length > 0 then
		if errnodes.Item(0).firstChild.nodeValue="Hard" then
			noError=FALSE
			set errdesc=xmlDoc.getElementsByTagName("err:Description")
			errormsg=errdesc.Item(0).firstChild.nodeValue
		end if
	end if
	if noError then
		set t2=xmlDoc.getElementsByTagName(rootNodeName).Item(0)
		for j=0 to t2.childNodes.length - 1
			Set n=t2.childNodes.Item(j)
			if n.nodename="Response" OR n.nodename="common:Response" then
				for i=0 To n.childNodes.length - 1
					Set e=n.childNodes.Item(i)
					if e.nodeName="common:ResponseStatus" then
						for k=0 To e.childNodes.length - 1
							Set t=e.childNodes.Item(k)
							Select Case t.nodeName
								Case "common:Code"
									noError=int(t.firstChild.nodeValue)=1
								Case "common:Description"
									' errormsg=errormsg & t.firstChild.nodeValue
							end Select
						next
					elseif e.nodeName="ResponseStatusCode" then
						noError=Int(e.firstChild.nodeValue)=1
					elseif e.nodeName="Error" then
						errormsg=""
						for k=0 To e.childNodes.length - 1
							Set t=e.childNodes.Item(k)
							Select Case t.nodeName
								Case "ErrorSeverity"
									if t.firstChild.nodeValue="Transient" then errormsg="This is a temporary error. Please wait a few moments then refresh this page.<br />" & errormsg
								Case "ErrorDescription"
									errormsg=errormsg & t.firstChild.nodeValue
							end Select
						next
					end if
					' print "The Nodename is : " & e.nodeName & ":" & e.firstChild.nodeValue & "<br />"
				next
			elseif n.nodename="AccessLicenseNumber" then
				thetext=n.firstChild.nodeValue
			elseif n.nodename="AccessLicenseText" then
				SESSION("adminUPSLicense")=n.firstChild.nodeValue
				thetext=n.firstChild.nodeValue
			'	if mysqlserver then rs.CursorLocation=3
			'	rs.open "SELECT * FROM admin WHERE adminID=1",cnn,1,3,&H0001
			'	rs.Fields("adminUPSLicense")=n.firstChild.nodeValue
			'	rs.Update
			'	rs.close
			elseif n.nodename="UserId" then
				thetext=n.firstChild.nodeValue
			end if
		next
	end if
	ParseUPSLicenseOutput=noError
end Function
sub registrationsuccess() %>
	<form method="post" name="licform" action="admin.asp">
	  <input type="hidden" name="upsstep" value="5" />
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr> 
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr>
				<td rowspan="3" width="70" align="center" valign="top"><img src="../images/upslogo.png" border="0" alt="UPS" /><br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</td>
                <td width="100%" align="center"><strong><%=yyUPSWiz%> - <% if success then print yyRegSucc else print yyError %></strong><br />&nbsp;
                </td>
			  </tr>
<%	if success then
		sSQL="UPDATE adminshipping SET adminUPSUser='"&upsencode(saveuser, "")&"',adminUPSpw='"&upsencode(thepw, "")&"'"
		ect_query(sSQL)
		SESSION("adminUPSLicense")=""
%>
			  <tr> 
                <td width="100%" align="left">
				  <p><strong><%=yyRegSucc%> !</strong></p>
				  <p><%=yyUPSLi5%></p>
				  <p><%=yyUPSLi6%> <a href="http://www.ups.com/content/us/en/bussol/browse/cat/developer_kit.html" target="_blank">http://www.ups.com/content/us/en/bussol/browse/cat/developer_kit.html</a>.</p>
				  <p><%=yyUPSLi7%> <a href="adminmain.asp"><%=yyAdmMai%></a>.</p>
				  <p><%=yyUPSLi8%> <a href="http://www.ups.com/content/us/en/bussol/browse/internet_shipping.html" target="_blank"><%=yyClkHer%></a>.</p>
				  <p>&nbsp;</p>
				  <p align="center"><input type="submit" value="<%=yyDone%>" /></p>
				  <p>&nbsp;</p>
                </td>
			  </tr>
<%	else %>
			  <tr> 
                <td width="100%" align="center"><p><%=yySorErr%></strong></p>
				<p>&nbsp;</p>
				<p><%=errormsg%></p>
				<p>&nbsp;</p>
				<p><%=yyTryBac%> <a href="javascript:history.go(-1)"><%=yyClkHer%></a>.</p>
				<p>&nbsp;</p>
                </td>
			  </tr>
<%	end if %>
			  <tr> 
                <td colspan="2" width="100%" align="center">
				  <p>&nbsp;</p><p><span style="font-size:10px"><%=yyUPStm%></span></p>
                </td>
			  </tr>
            </table>
          </td>
        </tr>
      </table>
	</form>
<%
end sub
if getpost("reregister")="3" then
	sXML="<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:upss=""http://www.ups.com/XMLSchema/XOLTWS/UPSS/v1.0"" xmlns=""http://www.ups.com/XMLSchema/XOLTWS/Registration/v1.0"" xmlns:common=""http://www.ups.com/XMLSchema/XOLTWS/Common/v1.0"">" & _
		"<soapenv:Header><upss:UPSSecurity><upss:UsernameToken><upss:Username>vince2002</upss:Username><upss:Password>Ups332211</upss:Password></upss:UsernameToken><upss:ServiceAccessToken><upss:AccessLicenseNumber>DB9341F6791A3D7A</upss:AccessLicenseNumber></upss:ServiceAccessToken></upss:UPSSecurity></soapenv:Header>"
	sXML=sXML & "<soapenv:Body>"
	sXML=sXML & "<ManageAccountRequest>" & _ 
			"<ShipperAccount><AccountNumber>" & getpost("upsaccount") & "</AccountNumber>"
	sXML=sXML & "<PostalCode>" & getpost("postcode") & "</PostalCode>" & _
			"<CountryCode>" & getpost("country") & "</CountryCode>"
	if getpost("invoicenumber")<>"" then
		sXML=sXML & "<InvoiceInfo>"
		sXML=sXML & "<InvoiceNumber>" & xmlencode(getpost("invoicenumber")) & "</InvoiceNumber>"
		if getpost("invoicedate")<>"" then sXML=sXML & "<InvoiceDate>" & xmlencode(getpost("invoicedate")) & "</InvoiceDate>"
		if getpost("invoicecurrency")<>"" then sXML=sXML & "<CurrencyCode>" & xmlencode(getpost("invoicecurrency")) & "</CurrencyCode>"
		if getpost("invoiceamount")<>"" then sXML=sXML & "<InvoiceAmount>" & xmlencode(getpost("invoiceamount")) & "</InvoiceAmount>"
		if getpost("invoicecontrolid")<>"" then sXML=sXML & "<ControlID>" & xmlencode(getpost("invoicecontrolid")) & "</ControlID>"
		sXML=sXML & "</InvoiceInfo>"
	end if
	sXML=sXML & "</ShipperAccount></ManageAccountRequest>" & _
		"</soapenv:Body></soapenv:Envelope>"
	xmlres="xml"
	if dumpshippingxml then print replace(replace(sXML,"</","&lt;/"),"<","<br />&lt;")&"<hr />"
	if callxmlfunction(registerurl, sXML, xmlres, "", "Msxml2.ServerXMLHTTP", errormsg, FALSE) then
		Session.LCID=1033
		success=ParseUPSLicenseOutput(xmlres, "reg:ManageAccountResponse", theuser, errormsg)
		Session.LCID=saveLCID
	end if
	call registrationsuccess()
elseif getpost("reregister")="1" then %>
	<form method="post" name="licform" action="adminupslicense.asp">
		<input type="hidden" name="reregister" value="2" />
<%	sSQL="SELECT adminUPSAccount FROM adminshipping WHERE adminShipID=1"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		adminUPSAccount=trim(rs("adminUPSAccount")&"")
	end if
	rs.close
	call writehiddenidvar("noupsaccount", "1") %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr> 
          <td width="100%" align="center">
			<p>&nbsp;</p>
            <table width="500" border="0" cellspacing="0" cellpadding="2">
			  <tr>
				<td rowspan="6" width="70" align="left" valign="top"><img src="../images/upslogo.png" border="0" alt="UPS" /><br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</td>
                <td width="100%" align="center" colspan="2"><p><strong><%=yyUPSWiz%> - </strong></p>
				<p>Please enter your UPS Shipper Number below then click &quot;Continue&quot;.<br />&nbsp;
                </td>
			  </tr>
			  <tr> 
                <td align="right"><p><strong>UPS Shipper Number:</strong></td>
				<td><input type="text" name="upsaccount" value="<%=adminUPSAccount%>" size="20" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyCountry%> : </strong></td>
				<td><select name="country" size="1">
<option value=''><%=yySelect%></option>
<%
sSQL="SELECT countryName,countryCode FROM countries WHERE countryCode IN (" & allowedcountries & ") ORDER BY countryName"
rs.open sSQL,cnn,0,1
do while not rs.EOF
	print "<option value="""&rs("countryCode")&""""
	if origCountryCode=rs("countryCode") then response.write " selected=""selected"""
	print ">"&rs("countryName")&"</option>"
	rs.MoveNext
loop
rs.close
%>
				</select></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyPCode%> : </strong></td>
				<td><input type="text" name="postcode" size="15" value="<%=origZip%>" /></td>
			  </tr>
			  <tr> 
                <td align="right">&nbsp;</td>
				<td><p>&nbsp;</p><p><input type="submit" value="<%=yyContin%>" /></p></td>
			  </tr>
			  <tr> 
                <td colspan="2" width="100%" align="center">
				  <p>&nbsp;</p><p><span style="font-size:10px"><%=yyUPStm%></span></p>
                </td>
			  </tr>
            </table>
          </td>
        </tr>
      </table>
	</form>
<%
elseif getpost("upsstep")="4" AND getpost("upsaccount")="" AND getpost("noupsaccount")<>"1" then %>
	<form method="post" name="licform" action="adminupslicense.asp">
<%	for each objItem in request.form
		if objItem<>"upsaccount" then call writehiddenidvar(objItem, request.form(objItem))
	next
	call writehiddenidvar("noupsaccount", "1") %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr> 
          <td width="100%" align="center">
			<p>&nbsp;</p>
            <table width="500" border="0" cellspacing="0" cellpadding="2">
			  <tr>
				<td rowspan="3" width="70" align="left" valign="top"><img src="../images/upslogo.png" border="0" alt="UPS" /><br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</td>
                <td width="100%" align="center"><p><strong><%=yyUPSWiz%> - </strong></p>
				<p>You did not provide a UPS Shipper Number with your registration. 
				If you have a UPS Shipper Number, please enter it in the space provided below before clicking &quot;Continue&quot;.<br />&nbsp;
                </td>
			  </tr>
			  <tr> 
                <td align="left"><p><strong>UPS Shipper Number:</strong>
				<input type="text" name="upsaccount" value="" size="20" /></p>
				<p>&nbsp;</p>
				<p><input type="submit" value="<%=yyContin%>" /></td>
			  </tr>
			  <tr> 
                <td colspan="2" width="100%" align="center">
				  <p>&nbsp;</p><p><span style="font-size:10px"><%=yyUPStm%></span></p>
                </td>
			  </tr>
            </table>
          </td>
        </tr>
      </table>
	</form>
<%
elseif getpost("reregister")="2" OR (getpost("upsstep")="4" AND getpost("upsaccount")<>"" AND getpost("invoicenumber")="" AND getpost("noinvoicenumber")<>"1") then %>
	<form method="post" name="licform" action="adminupslicense.asp">
<%	for each objItem in request.form
		if objItem="reregister" then
			call writehiddenvar("reregister", "3")
		else
			call writehiddenvar(objItem, request.form(objItem))
		end if
	next
	call writehiddenidvar("noinvoicenumber", "1") %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr> 
          <td width="100%" align="center">
			<p>&nbsp;</p>
            <table width="600" border="0" cellspacing="0" cellpadding="2">
			  <tr>
				<td rowspan="3" width="70" align="left" valign="top"><img src="../images/upslogo.png" border="0" alt="UPS" /><br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</td>
                <td width="100%" align="center"><p><strong><%=yyUPSWiz%> - </strong></p>
				<p>If you have received a UPS invoice in the past, please enter the invoice information in the space provided below 
				before clicking &quot;Continue&quot;. This will authenticate your account and allow you to view negotiated rates.<br />&nbsp;
                </td>
			  </tr>
			  <tr> 
                <td align="left">
					<table>
						<tr><td align="right"><strong>Invoice Number:</strong></td>
						<td><input type="text" name="invoicenumber" value="" size="20" /></td></tr>
						<tr><td align="right"><strong>Invoice Date:</strong></td>
						<td><input type="text" name="invoicedate" value="" size="20" /> (E.g. 20120225)</td></tr>
						<tr><td align="right"><strong>Invoice Amount:</strong></td>
						<td><input type="text" name="invoiceamount" value="" size="20" /></td></tr>
						<tr><td align="right"><strong>Invoice Currency:</strong></td>
						<td><select name="invoicecurrency" size="1">
<option value=''><%=yySelect%></option>
<%
sSQL="SELECT DISTINCT countryCurrency FROM countries WHERE countryCode IN (" & allowedcountries & ") ORDER BY countryCurrency"
rs.open sSQL,cnn,0,1
do while not rs.EOF
	print "<option value='"&rs("countryCurrency")&"'>"&rs("countryCurrency")&"</option>"
	rs.MoveNext
loop
rs.close
%>
				</select></td></tr>
						<tr><td align="right"><strong>Invoice Control ID:</strong></td>
						<td><input type="text" name="invoicecontrolid" value="" size="20" /></td></tr>
						<tr><td align="right">&nbsp;</td>
						<td>Optional, but this value is required if it is present on your invoice.</td></tr>
						<tr><td>
						<p>&nbsp;</p>
						<p><input type="submit" value="<%=yyContin%>" /></td></tr>
					</table>
				</td>
			  </tr>
			  <tr> 
                <td colspan="2" width="100%" align="center">
				  <p>&nbsp;</p><p><span style="font-size:10px"><%=yyUPStm%></span></p>
                </td>
			  </tr>
            </table>
          </td>
        </tr>
      </table>
	</form>
<%
elseif getpost("upsstep")="4" then
	Set docXML=Server.CreateObject("MSXML2.DOMDocument")
'	sSQL="SELECT adminUPSLicense FROM admin WHERE adminID=1"
'	rs.open sSQL,cnn,0,1
	sXML="<?xml version=""1.0"" encoding="""&UCASE(adminencoding)&"""?>"
	sXML=sXML & "<AccessLicenseRequest xml:lang=""en-US""><Request><TransactionReference><CustomerContext>Ecomm Plus UPS Reg</CustomerContext><XpciVersion>1.0001</XpciVersion></TransactionReference>"
	sXML=sXML & "<RequestAction>AccessLicense</RequestAction><RequestOption>AllTools</RequestOption></Request>"
	sXML=sXML & "<CompanyName>" & getpost("company") & "</CompanyName>"
	sXML=sXML & "<Address><AddressLine1>" & getpost("address") & "</AddressLine1>"
	if getpost("address2")<>"" then sXML=sXML & "<AddressLine2>" & getpost("address2") & "</AddressLine2>"
	sXML=sXML & "<City>" & getpost("city") & "</City>"
	if getpost("country")="US" OR getpost("country")="CA" then
		sXML=sXML & "<StateProvinceCode>" & getpost("usstate") & "</StateProvinceCode>"
	else
		sXML=sXML & "<StateProvinceCode>XX</StateProvinceCode>"
	end if
	if getpost("postcode")<>"" then sXML=sXML & "<PostalCode>" & getpost("postcode") & "</PostalCode>"
	sXML=sXML & "<CountryCode>" & getpost("country") & "</CountryCode></Address>"
	sXML=sXML & "<PrimaryContact><Name>" & getpost("contact") & "</Name><Title>" & getpost("ctitle") & "</Title>"
	sXML=sXML & "<EMailAddress>" & getpost("email") & "</EMailAddress><PhoneNumber>" & getpost("telephone") & "</PhoneNumber></PrimaryContact>"
	sXML=sXML & "<CompanyURL>" & getpost("websiteurl") & "</CompanyURL>"
'	if getpost("upsaccount")<>"" then sXML=sXML & "<ShipperNumber>" & getpost("upsaccount") & "</ShipperNumber>"
	sXML=sXML & "<DeveloperLicenseNumber>BB9341E83CC05B12</DeveloperLicenseNumber>"
	sXML=sXML & "<AccessLicenseProfile><CountryCode>" & getpost("countryCode") & "</CountryCode><LanguageCode>" & getpost("languageCode") & "</LanguageCode>"
	sXML=sXML & "<AccessLicenseText>" & replace(SESSION("adminUPSLicense"),"&","&amp;") & "</AccessLicenseText>"
	sXML=sXML & "</AccessLicenseProfile>"
	sXML=sXML & "<OnLineTool><ToolID>RateXML</ToolID><ToolVersion>1.0</ToolVersion></OnLineTool><OnLineTool><ToolID>TrackXML</ToolID><ToolVersion>1.0</ToolVersion></OnLineTool>"
	sXML=sXML & "<ClientSoftwareProfile><SoftwareInstaller>" & getpost("upsrep") & "</SoftwareInstaller><SoftwareProductName>Ecommerce Plus Templates</SoftwareProductName><SoftwareProvider>ViciSoft SL</SoftwareProvider><SoftwareVersionNumber>2.5</SoftwareVersionNumber></ClientSoftwareProfile>"
	sXML=sXML & "</AccessLicenseRequest>"
	docXML.loadXML(sXML)
	if dumpshippingxml then print replace(replace(sXML,"</","&lt;/"),"<","<br />&lt;")&"<hr />"
'	rs.close
	if proxyserver<>"" then
		set objHttp=Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
		objHttp.setProxy 2, proxyserver
	else
		set objHttp=Server.CreateObject("Msxml2.ServerXMLHTTP")
	end if
	objHttp.open "POST", "https://"&upsurl&"/ups.app/xml/License", false
	objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	on error resume next
	err.number=0
	objHttp.Send docXML
	on error goto 0
	If err.number <> 0 OR objHttp.status <> 200 Then
		errormsg="Error, couldn't connect to UPS server<br />" & objHttp.status & ": (" & objHttp.statusText & ") " & err.number & ": (" & err.Description & ")"
		success=false
	else
		saveLCID=Session.LCID
		Session.LCID=1033
		success=ParseUPSLicenseOutput(objHttp.responseXML, "AccessLicenseResponse", accessnumber, errormsg)
		Session.LCID=saveLCID
	end If
	set objHttp=nothing
	if success then
		sSQL="UPDATE adminshipping SET adminUPSAccess='"&accessnumber&"'"
		if getpost("upsaccount")<>"" then sSQL=sSQL & ",adminUPSAccount='"&escape_string(getpost("upsaccount"))&"'"
		sSQL=sSQL&",adminUPSNegotiated=0 WHERE adminShipID=1"
		ect_query(sSQL)
		if getpost("myupsuser")<>"" AND getpost("myupspw")<>"" then
			saveuser=getpost("myupsuser")
			thepw=getpost("myupspw")
		else
			noloops=0
			Randomize
			upperbound="999999"
			lowerbound="100000"
			thepw="ecp" & Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
			theuser="ecu" & Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
			do while theuser<>"" AND success AND noloops < 5
				saveuser=theuser
				sXML="<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:upss=""http://www.ups.com/XMLSchema/XOLTWS/UPSS/v1.0"" xmlns=""http://www.ups.com/XMLSchema/XOLTWS/Registration/v1.0"" xmlns:common=""http://www.ups.com/XMLSchema/XOLTWS/Common/v1.0"">" & _
					"<soapenv:Header><upss:UPSSecurity><upss:UsernameToken><upss:Username>vince2002</upss:Username><upss:Password>Ups332211</upss:Password></upss:UsernameToken><upss:ServiceAccessToken><upss:AccessLicenseNumber>DB9341F6791A3D7A</upss:AccessLicenseNumber></upss:ServiceAccessToken></upss:UPSSecurity></soapenv:Header>"
				sXML=sXML & "<soapenv:Body><RegisterRequest>"
				sXML=sXML & "<Username>"&theuser&"</Username><Password>"&thepw&"</Password>" & _
					"<CompanyName>" & getpost("company") & "</CompanyName><CustomerName>" & getpost("contact") & "</CustomerName>" & _
					"<Title>" & getpost("ctitle") & "</Title>"
				sXML=sXML & "<Address><AddressLine>" & getpost("address") & "</AddressLine>"
				if getpost("address2")<>"" then sXML=sXML & "<AddressLine>" & getpost("address2") & "</AddressLine>"
				sXML=sXML & "<City>" & getpost("city") & "</City>"
				if getpost("country")="US" OR getpost("country")="CA" then
					sXML=sXML & "<StateProvinceCode>" & getpost("usstate") & "</StateProvinceCode>"
				else
					sXML=sXML & "<StateProvinceCode>XX</StateProvinceCode>"
				end if
				sXML=sXML & "<PostalCode>" & getpost("postcode") & "</PostalCode><CountryCode>" & getpost("country") & "</CountryCode>" & _
					"</Address>" & _
					"<PhoneNumber>" & getpost("telephone") & "</PhoneNumber>" & _
					"<EmailAddress>" & getpost("email") & "</EmailAddress>" & _
					"<NotificationCode>00</NotificationCode>"
				if getpost("upsaccount")<>"" then
					sXML=sXML & "<ShipperAccount>"
					if getpost("upsaccount")<>"" then sXML=sXML & "<AccountNumber>" & getpost("upsaccount") & "</AccountNumber>"
					sXML=sXML & "<PostalCode>" & getpost("postcode") & "</PostalCode>" & _
						"<CountryCode>" & getpost("country") & "</CountryCode>"
					if getpost("invoicenumber")<>"" then
						sXML=sXML & "<InvoiceInfo>"
						sXML=sXML & "<InvoiceNumber>" & xmlencode(getpost("invoicenumber")) & "</InvoiceNumber>"
						if getpost("invoicedate")<>"" then sXML=sXML & "<InvoiceDate>" & xmlencode(getpost("invoicedate")) & "</InvoiceDate>"
						if getpost("invoicecurrency")<>"" then sXML=sXML & "<CurrencyCode>" & xmlencode(getpost("invoicecurrency")) & "</CurrencyCode>"
						if getpost("invoiceamount")<>"" then sXML=sXML & "<InvoiceAmount>" & xmlencode(getpost("invoiceamount")) & "</InvoiceAmount>"
						if getpost("invoicecontrolid")<>"" then sXML=sXML & "<ControlID>" & xmlencode(getpost("invoicecontrolid")) & "</ControlID>"
						sXML=sXML & "</InvoiceInfo>"
					end if
					sXML=sXML & "</ShipperAccount>"
				end if
				sXML=sXML & "<SuggestUsernameIndicator>Y</SuggestUsernameIndicator>" & _
					"</RegisterRequest></soapenv:Body></soapenv:Envelope>"
				xmlres="xml"
				if dumpshippingxml then print replace(replace(sXML,"</","&lt;/"),"<","<br />&lt;")&"<hr />"
				if callxmlfunction(registerurl, sXML, xmlres, "", "Msxml2.ServerXMLHTTP", errormsg, FALSE) then
					Session.LCID=1033
					success=ParseUPSLicenseOutput(xmlres, "reg:RegisterResponse", theuser, errormsg)
					Session.LCID=saveLCID
				end if
				noloops=noloops+1
			loop
		end if
	end if
	Set docXML=nothing
	call registrationsuccess()
elseif getpost("upsstep")="3" AND getpost("doagree")="1" then
%>
<script>
<!--
function checkforamp(checkObj){
  checkStr=checkObj.value;
  for (i=0;  i < checkStr.length;  i++){
	if (checkStr.charAt(i) == "&"){
	  alert("Please do not use the ampersand \"&\" character in any field.");
	  checkObj.focus();
	  return(false);
	}
  }
  return(true);
}
function formvalidator(theForm){
  if(theForm.contact.value == ""){
    alert("<%=jscheck(yyPlsEntr&" """&yyConNam)%>\".");
    theForm.contact.focus();
    return (false);
  }
  if(!checkforamp(theForm.contact)) return(false);
  if(theForm.ctitle.value == ""){
    alert("<%=jscheck(yyPlsEntr&" """&yyTitle)%>\".");
    theForm.ctitle.focus();
    return (false);
  }
  if(!checkforamp(theForm.ctitle)) return(false);
  if(theForm.company.value == ""){
    alert("<%=jscheck(yyPlsEntr&" """&yyComNam)%>\".");
    theForm.company.focus();
    return (false);
  }
  if(!checkforamp(theForm.company)) return(false);
  if(theForm.address.value == ""){
    alert("<%=jscheck(yyPlsEntr&" """&yyStrAdd)%>\".");
    theForm.address.focus();
    return (false);
  }
  if(!checkforamp(theForm.address)) return(false);
  if(theForm.city.value == ""){
    alert("<%=jscheck(yyPlsEntr&" """&yyCity)%>\".");
    theForm.city.focus();
    return (false);
  }
  if(!checkforamp(theForm.city)) return(false);
  var cntry=theForm.country[theForm.country.selectedIndex].value;
  if(cntry=="US" || cntry=="CA"){
	if (theForm.usstate.selectedIndex == 0){
      alert("<%=jscheck(yyPlsSel&" """&yyState)%>\".");
      theForm.usstate.focus();
      return (false);
	}
  }
  if(theForm.country.selectedIndex == 0){
    alert("<%=jscheck(yyPlsSel&" """&yyCountry)%>\".");
    theForm.country.focus();
    return (false);
  }
  if(cntry!='CL' && cntry!='CO' && cntry!='CR' && cntry!='DO' && cntry!='GT' && cntry!='HK' && cntry!='IE' && cntry!='PA'){
	if (theForm.postcode.value == ""){
	  alert("<%=jscheck(yyPlsEntr&" """&yyPCode)%>\".");
	  theForm.postcode.focus();
	  return (false);
	}
  }
  if(!checkforamp(theForm.postcode)) return(false);
  if(theForm.telephone.value == ""){
    alert("<%=jscheck(yyPlsEntr&" """&yyTelep)%>\".");
    theForm.telephone.focus();
    return (false);
  }
  if(theForm.telephone.value.length < 10 || theForm.telephone.value.length > 14){
    alert("<%=jscheck(yyValTN)%>");
    theForm.telephone.focus();
    return (false);
  }
  var checkOK="0123456789";
  var checkStr=theForm.telephone.value;
  var allValid=true;
  for (i=0;  i < checkStr.length;  i++){
    ch=checkStr.charAt(i);
    for (j=0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length){
      allValid=false;
      break;
    }
  }
  if(!allValid){
    alert("<%=jscheck(yyOnDig&" """&yyTelep)%>\".");
    theForm.telephone.focus();
    return (false);
  }
  if(theForm.websiteurl.value == ""){
    alert("<%=jscheck(yyPlsEntr&" """&yyWebURL)%>\".");
    theForm.websiteurl.focus();
    return (false);
  }
  if(!checkforamp(theForm.websiteurl)) return(false);
  var checkStr=theForm.websiteurl.value;
  var gotDot=false;
  for (i=0;  i < checkStr.length;  i++){
    ch=checkStr.charAt(i);
	if (ch == ".") gotDot=true;
  }
  if(!(gotDot)){
    alert("<%=jscheck(yyValEnt&" """&yyWebURL)%>\".");
    theForm.websiteurl.focus();
    return (false);
  }
  if(theForm.email.value == ""){
    alert("<%=jscheck(yyPlsEntr&" """&yyEmail)%>\".");
    theForm.email.focus();
    return (false);
  }
  var checkStr=theForm.email.value;
  var gotDot=false;
  var gotAt=false;
  for (i=0;  i < checkStr.length;  i++){
    ch=checkStr.charAt(i);
    if (ch == "@") gotAt=true;
	if (ch == ".") gotDot=true;
  }
  if (!(gotDot && gotAt)){
    alert("<%=jscheck(yyValEnt&" """&yyEmail)%>\".");
    theForm.email.focus();
    return (false);
  }
  if(theForm.upsrep[0].checked==false && theForm.upsrep[1].checked==false){
    alert("<%=jscheck(yyUPSrep)%>");
    return (false);
  }
  return (true);
}
//-->
</script>
	<form method="post" name="licform" action="adminupslicense.asp" onsubmit="return formvalidator(this)">
	  <input type="hidden" name="upsstep" value="4" />
	  <input type="hidden" name="countryCode" value="<%=getpost("countryCode")%>" />
	  <input type="hidden" name="languageCode" value="<%=getpost("languageCode")%>" />
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr> 
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr>
				<td rowspan="20" width="70" align="center" valign="top"><img src="../images/upslogo.png" border="0" alt="UPS" /><br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</td>
                <td width="100%" align="center" colspan="2"><strong><%=yyUPSWiz%> - <%=yyStep%> 2</strong><br />&nbsp;
                </td>
			  </tr>
			  <tr> 
                <td width="40%" align="right"><strong><%=yyConNam%> : </strong></td>
				<td width="60%"><input type="text" name="contact" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyTitle%> : </strong></td>
				<td><input type="text" name="ctitle" size="10" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyComNam%> : </strong></td>
				<td><input type="text" name="company" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyStrAdd%> : </strong></td>
				<td><input type="text" name="address" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyAddr2%> : </strong></td>
				<td><input type="text" name="address2" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyCity%> : </strong></td>
				<td><input type="text" name="city" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyState%> <%=yyUSCan%> : </strong></td>
				<td><select name="usstate" size="1">
<option value=''><%=yyOutUS%></option>
<option value='AL'>Alabama</option>
<option value='AK'>Alaska</option>
<option value='AB'>Alberta</option>
<option value='AZ'>Arizona</option>
<option value='AR'>Arkansas</option>
<option value='BC'>British Columbia</option>
<option value='CA'>California</option>
<option value='CO'>Colorado</option>
<option value='CT'>Connecticut</option>
<option value='DE'>Delaware</option>
<option value='DC'>District Of Columbia</option>
<option value='FL'>Florida</option>
<option value='GA'>Georgia</option>
<option value='HI'>Hawaii</option>
<option value='ID'>Idaho</option>
<option value='IL'>Illinois</option>
<option value='IN'>Indiana</option>
<option value='IA'>Iowa</option>
<option value='KS'>Kansas</option>
<option value='KY'>Kentucky</option>
<option value='LA'>Louisiana</option>
<option value='ME'>Maine</option>
<option value='MB'>Manitoba</option>
<option value='MD'>Maryland</option>
<option value='MA'>Massachusetts</option>
<option value='MI'>Michigan</option>
<option value='MN'>Minnesota</option>
<option value='MS'>Mississippi</option>
<option value='MO'>Missouri</option>
<option value='MT'>Montana</option>
<option value='NE'>Nebraska</option>
<option value='NV'>Nevada</option>
<option value='NB'>New Brunswick</option>
<option value='NH'>New Hampshire</option>
<option value='NJ'>New Jersey</option>
<option value='NM'>New Mexico</option>
<option value='NY'>New York</option>
<option value='NF'>Newfoundland</option>
<option value='NC'>North Carolina</option>
<option value='ND'>North Dakota</option>
<option value='NT'>Northwest Territories</option>
<option value='NS'>Nova Scotia</option>
<option value='NU'>Nunavut</option>
<option value='OH'>Ohio</option>
<option value='OK'>Oklahoma</option>
<option value='ON'>Ontario</option>
<option value='OR'>Oregon</option>
<option value='PA'>Pennsylvania</option>
<option value='PE'>Prince Edward Island</option>
<option value='QC'>Quebec</option>
<option value='RI'>Rhode Island</option>
<option value='SK'>Saskatchewan</option>
<option value='SC'>South Carolina</option>
<option value='SD'>South Dakota</option>
<option value='TN'>Tennessee</option>
<option value='TX'>Texas</option>
<option value='UT'>Utah</option>
<option value='VT'>Vermont</option>
<option value='VA'>Virginia</option>
<option value='WA'>Washington</option>
<option value='WV'>West Virginia</option>
<option value='WI'>Wisconsin</option>
<option value='WY'>Wyoming</option>
<option value='YT'>Yukon</option>
</select></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyCountry%> : </strong></td>
				<td><select name="country" size="1" onchange="document.getElementById('upslink').href='http://www.ups.com/content/XXXX/en/index.jsx'.replace('XXXX',this[this.selectedIndex].value.toLowerCase())">
<option value=''><%=yySelect%></option>
<%
sSQL="SELECT countryName,countryCode FROM countries WHERE countryCode IN (" & allowedcountries & ") ORDER BY countryName"
rs.open sSQL,cnn,0,1
do while not rs.EOF
	print "<option value='"&rs("countryCode")&"'>"&rs("countryName")&"</option>"
	rs.MoveNext
loop
rs.close
%>
				</select></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyPCode%> : </strong></td>
				<td><input type="text" name="postcode" size="15" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyTelep%> : </strong></td>
				<td><input type="text" name="telephone" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyWebURL%> : </strong></td>
				<td><input type="text" name="websiteurl" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyEmail%> : </strong></td>
				<td><input type="text" name="email" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyUPSac%> : </strong></td>
				<td><input type="text" name="upsaccount" size="15" maxlength="6" /></td>
			  </tr>
			  <tr> 
                <td align="center" colspan="2">
					<%=yyUPSsr%><br /><input type="radio" name="upsrep" value="yes" /> <strong><%=yyYes%></strong> <input type="radio" name="upsrep" value="no" /> <strong><%=yyNo%></strong>
				</td>
			  </tr>
<%	if FALSE then %>
			  <tr> 
                <td align="center" width="100%" colspan="2">If you are already in posession of a My UPS User ID and Password please enter this below. Otherwise just leave blank and one will be created for you.</td>
			  </tr>
			  <tr> 
                <td align="right"><strong>My UPS User ID : </strong></td>
				<td><input type="text" name="myupsuser" size="20" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong>My UPS Password : </strong></td>
				<td><input type="text" name="myupspw" size="20" /></td>
			  </tr>
<%	end if %>
			  <tr>
                <td width="100%" align="center" colspan="2"><br />&nbsp;<input type="submit" name="agree" value="&nbsp;&nbsp;<%=yyNext%>&nbsp;&nbsp;" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="button" value="<%=yyCancel%>" onclick="window.location='admin.asp';" />
                </td>
			  </tr>
			  <tr> 
                <td align="center" colspan="2"><p><span style="font-size:10px"><%=yyUPSop%> <a href="http://www.ups.com/content/us/en/index.jsx" target="_blank" id="upslink"><%=yyClkHer%></a> <%=yyUPScl%><br />
				<%=yyUPSMI%> <a href="http://www.ups.com/content/us/en/bussol/browse/cat/developer_kit.html" target="_blank"><%=yyClkHer%></a>.<br />
				<%=yyUPshp%> <a href="http://www.ups.com/content/us/en/bussol/browse/internet_shipping.html" target="_blank"><%=yyClkHer%></a></span></p>
				</td>
			  </tr>
			  <tr> 
                <td colspan="3" width="100%" align="center">
				  <p>&nbsp;</p><p><span style="font-size:10px"><%=yyUPStm%></span></p>
                </td>
			  </tr>
            </table>
          </td>
        </tr>
      </table>
	</form>
<%
elseif getpost("upsstep")="2" then
	languageCode="EN"
	if countryCode="AR" OR countryCode="ES" OR countryCode="MX" OR countryCode="CA" OR countryCode="DO" OR countryCode="GT" OR countryCode="CR" OR countryCode="CO" OR countryCode="PA" OR countryCode="PR" OR countryCode="CL" then
		languageCode="ES"
	elseif countryCode="AT" OR countryCode="DE" then
		languageCode="DE"
	elseif countryCode="PT" OR countryCode="BR" then
		languageCode="PT"
	elseif countryCode="FR" OR countryCode="CH" OR countryCode="BE" then
		languageCode="FR"
	elseif countryCode="CN" OR countryCode="HK" then
		languageCode="ZH"
	elseif countryCode="DK" then
		languageCode="DA"
	elseif countryCode="FI" then
		languageCode="FI"
	elseif countryCode="GR" then
		languageCode="EL"
	elseif countryCode="IN" then
		languageCode="GU"
	elseif countryCode="IL" then
		languageCode="IW"
	elseif countryCode="IT" then
		languageCode="IT"
	elseif countryCode="JP" then
		languageCode="JA"
	elseif countryCode="MY" then
		languageCode="MS"
	elseif countryCode="NL" then
		languageCode="NL"
	elseif countryCode="NO" then
		languageCode="NO"
	elseif countryCode="KR" then
		languageCode="KO"
	elseif countryCode="SE" then
		languageCode="SV"
	elseif countryCode="TH" then
		languageCode="TH"
	end if
	sXML="<?xml version=""1.0"" encoding="""&UCASE(adminencoding)&"""?>"
	sXML=sXML & "<AccessLicenseAgreementRequest><Request><RequestOption>AllTools</RequestOption><TransactionReference><CustomerContext>Ecomm Plus UPS License</CustomerContext><XpciVersion>1.0001</XpciVersion></TransactionReference>"
	sXML=sXML & "<RequestAction>AccessLicense</RequestAction></Request><DeveloperLicenseNumber>BB9341E83CC05B12</DeveloperLicenseNumber>"
	sXML=sXML & "<AccessLicenseProfile><CountryCode>"&countryCode&"</CountryCode><LanguageCode>"&languageCode&"</LanguageCode></AccessLicenseProfile>"
	sXML=sXML & "<OnLineTool><ToolID>RateXML</ToolID><ToolVersion>1.0</ToolVersion></OnLineTool><OnLineTool><ToolID>TrackXML</ToolID><ToolVersion>1.0</ToolVersion></OnLineTool></AccessLicenseAgreementRequest>"
	if proxyserver<>"" then
		set objHttp=Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
		objHttp.setProxy 2, proxyserver
	else
		set objHttp=Server.CreateObject("Msxml2.ServerXMLHTTP")
	end if
	objHttp.open "POST", "https://"&upsurl&"/ups.app/xml/License", false
	objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	on error resume next
	err.number=0
	errnum=0
	if dumpshippingxml then print replace(replace(sXML,"</","&lt;/"),"<","<br />&lt;")&"<hr />"
	objHttp.Send sXML
	errnum=err.number
	errtxt=err.Description
	on error goto 0
	lictext=""
	if errnum <> 0 then
		errormsg="Error, couldn't connect to UPS server<br />Error: " & errnum & ": (" & errtxt & ")"
		success=false
	elseif objHttp.status <> 200 then
		errormsg="Error, couldn't connect to UPS server<br />" & objHttp.status & ": (" & objHttp.statusText & ")"
		success=false
	else
		saveLCID=Session.LCID
		Session.LCID=1033
		success=ParseUPSLicenseOutput(objHttp.responseXML, "AccessLicenseAgreementResponse", lictext, errormsg)
		Session.LCID=saveLCID
	end if
	set objHttp=nothing
%>
<script>
<!--
var origlictext="";
function printlicense(){
	var prnttext='<html><body>\n';
	if(origlictext != document.licform.lictext.value){
		alert("It appears that the license text has been modified. Cannot print license.");
		return;
	}
	prnttext += document.licform.lictext.value.replace(/\n|\r\n/g,'<br />');
	prnttext += '</body></'+'html>';
	var newwin=window.open("","printlicense",'menubar=no, scrollbars=yes, width=500, height=400, directories=no,location=no,resizable=yes,status=no,toolbar=no');
	newwin.document.open();
	newwin.document.write(prnttext);
	newwin.document.close();
	newwin.print();
}
function checkaccept(theForm){
  if(origlictext != document.licform.lictext.value){
	alert("It appears that the license text has been modified. Cannot proceed.");
	return (false);
  }
  if (theForm.doagree[0].checked == false){
    alert("<%=jscheck(yyUPSLi4)%>");
    return (false);
  }
  return (true);
}
var hasscrolled=false;
function checkscroll(tarea){
	if(tarea.offsetHeight+tarea.scrollTop+1>=tarea.scrollHeight){
		hasscrolled=true;
	}
}
function checkhasscrolled(radbut){
	if(! hasscrolled){
		radbut.checked=false;
		alert("You must scroll through the whole license agreement before you can select this option.");
	}
}
//-->
</script>
	<form method="post" name="licform" action="adminupslicense.asp" onsubmit="return checkaccept(this)">
	  <input type="hidden" name="upsstep" value="3" />
	  <input type="hidden" name="countryCode" value="<%=countryCode%>" />
	  <input type="hidden" name="languageCode" value="<%=languageCode%>" />
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr> 
          <td width="100%">
            <table width="100%" border="0" cellspacing="2" cellpadding="0">
			  <tr>
                <td width="100%" align="center"><img src="../images/upslogo.png" border="0" align="middle" alt="" />&nbsp;&nbsp;<strong><%=yyUPSWiz%> - <%=yyStep%> 1</strong><br />&nbsp;
                </td>
			  </tr>
<%	if success then %>
			  <tr> 
                <td width="100%" align="center"><textarea cols="80" rows="20" name="lictext" onscroll="checkscroll(this)"><%=lictext%></textarea><br /><br />
				<p><%=yyUPSTer%></p>
				<p><%=yyAgree%> <input type="radio" name="doagree" value="1" onclick="checkhasscrolled(this)" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=yyNoAgre%> <input type="radio" name="doagree" value="0" /></p>
				<p>&nbsp;</p>
                </td>
			  </tr>
<script>
<!--
var origlictext=document.licform.lictext.value;
//-->
</script>
<%	else %>
			  <tr> 
                <td width="100%" align="center"><p><%=yySorErr%></strong></p>
				<p>&nbsp;</p>
				<p><%=errormsg%></p>
				<p>&nbsp;</p>
                </td>
			  </tr>
<%	end if %>
			  <tr> 
                <td width="100%" align="center"><% if success then %><input type="button" value="&nbsp;<%=yyPrint%>&nbsp;" onclick="printlicense();" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="agree" value="&nbsp;&nbsp;<%=yyNext%>&nbsp;&nbsp;" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<% end if %>
				<input type="button" value="<%=yyCancel%>" onclick="window.location='admin.asp';" />
                </td>
			  </tr>
			  <tr> 
                <td align="center"><p><span style="font-size:10px"><%=yyUPSop%> <a href="http://www.ups.com/content/us/en/index.jsx" target="_blank"><%=yyClkHer%></a> <%=yyUPScl%><br />
				<%=yyUPSMI%> <a href="http://www.ups.com/content/us/en/bussol/browse/cat/developer_kit.html" target="_blank"><%=yyClkHer%></a>.<br />
				<%=yyUPshp%> <a href="http://www.ups.com/content/us/en/bussol/browse/internet_shipping.html" target="_blank"><%=yyClkHer%></a>.</span></p>
				</td>
			  </tr>
			  <tr> 
                <td width="100%" align="center">
				  <p>&nbsp;</p><p><span style="font-size:10px"><%=yyUPStm%></span></p>
                </td>
			  </tr>
            </table>
          </td>
        </tr>
      </table>
	</form>
<%
else
	isregistered=FALSE
	sSQL = "SELECT adminUPSUser,adminUPSpw FROM adminshipping WHERE adminShipID=1"
	rs.open sSQL,cnn,0,1
	if not rs.EOF then
		upsUser=trim(rs("adminUPSUser")&"")
		upsPw=trim(rs("adminUPSpw")&"")
	end if
	rs.close
	if trim(upsUser&"")<>"" AND trim(upsPw&"")<>"" then isregistered=TRUE %>
	<form method="post" id="licform" action="adminupslicense.asp">
	  <input type="hidden" name="upsstep" value="2" />
	  <input type="hidden" name="reregister" id="reregister" value="" />
      <table border="0" cellspacing="3" cellpadding="3" width="100%" align="center">	
		  <tr>
			<td rowspan="5" width="70" align="center" valign="top"><img src="../images/upslogo.png" border="0" alt="" /><br />&nbsp;</td>
			<td width="100%" align="center"><strong><%=yyUPSWiz%></strong><br />&nbsp;
			</td>
		  </tr>
<%	if isregistered then %>
		  <tr> 
			<td width="100%"><p>&nbsp;</p></td>
		  </tr>
		  <tr> 
			<td width="100%">You have already successfully completed the UPS licensing and registration wizard. If you would like to re-register then please 
			click the "Re-register" button below. If you would just like to update your UPS account information then please click the "Update Account" button below.
			<p>&nbsp;</p>
			</td>
		  </tr>
		  <tr> 
			<td width="100%" align="center"><input type="submit" name="agree" onclick="document.getElementById('reregister').value='';" value="&nbsp;&nbsp;Re-Register&nbsp;&nbsp;" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="button" value="Update Account" onclick="document.getElementById('reregister').value=1;document.getElementById('licform').submit()" />
			</td>
		  </tr>
<%	else %>
		  <tr> 
			<td width="100%"><ul><li><%=yyUPSLi1%><br /><br /></li>
			<li><%=yyUPSLi2%><br /><br /></li>
			<li><%=yyUPSLi3%> <%=yyNoCou%> <a href="adminmain.asp"><%=yyClkHer%></a>.<br /><br /></li>
			<li><%=yyUPSMI%> <a href="http://www.ups.com/content/us/en/bussol/browse/cat/developer_kit.html" target="_blank"><%=yyClkHer%></a>.<br /><br /></li>
			<li><%=yyUPshp%> <a href="http://www.ups.com/content/us/en/bussol/browse/internet_shipping.html" target="_blank"><%=yyClkHer%></a>.</li>
			</ul>
			<p>&nbsp;</p>
			</td>
		  </tr>
		  <tr> 
			<td width="100%" align="center"><input type="submit" name="agree" value="&nbsp;&nbsp;<%=yyNext%>&nbsp;&nbsp;" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="button" value="<%=yyCancel%>" onclick="window.location='admin.asp';" />
			</td>
		  </tr>
		  <tr>
			<td align="center" colspan="2"><p><span style="font-size:10px"><%=yyUPSop%> <a href="http://www.ups.com/content/us/en/index.jsx" target="_blank"><%=yyClkHer%></a> <%=yyUPScl%><br />
			<%=yyUPSMI%> <a href="http://www.ups.com/content/us/en/bussol/browse/cat/developer_kit.html" target="_blank"><%=yyClkHer%></a>.<br />
			<%=yyUPshp%> <a href="http://www.ups.com/content/us/en/bussol/browse/internet_shipping.html" target="_blank"><%=yyClkHer%></a>.</span></p>
			</td>
		  </tr>
<%	end if %>
		  <tr>
			<td width="100%" align="center">
			  <p>&nbsp;</p><p><span style="font-size:10px"><%=yyUPStm%></span></p>
			</td>
		  </tr>
      </table>
	</form>
<%
end if
cnn.Close
set rs=nothing
set cnn=nothing
%>