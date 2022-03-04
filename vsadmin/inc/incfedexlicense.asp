<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
Response.Charset = "8859-1"
success=true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
countryCode = origCountryCode
if getget("act")="version" then %>
	<form method="post" name="licform" action="admin.asp">
	  <input type="hidden" name="upsstep" value="5" />
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr> 
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr>
				<td rowspan="3" width="70" align="center" valign="top"><img src="../images/fedexlogo.png" border="0" alt="FedEx" /><br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</td>
                <td width="100%" align="center"><strong><%=yyFdxWiz%> - Updating FedEx® version information.</strong><br />&nbsp;
                </td>
			  </tr>
			  <tr> 
                <td width="100%" align="left">
				  <p>&nbsp;</p>
				  <p>Please wait while we update your FedEx version information.</p>
				  <p>&nbsp;</p>
				  <p>Step 1, getting location id. <span name="step1span" id="step1span"><strong>Please wait!</strong></span></p>
				  <p>&nbsp;</p>
				  <p>Step 2, updating version. <span name="step2span" id="step2span"><strong>Please wait!</strong></span></p>
				  <p>&nbsp;</p>
				  <p align="center" name="donebutton" id="donebutton" style="display:none"><input type="submit" value="<%=yyDone%>" /></p>
				  <p>&nbsp;</p>
                </td>
			  </tr>
			  <tr> 
                <td colspan="2" width="100%" align="center">
				  <p><br />&nbsp;</p>
				  <p><span style="font-size:10px"><%=fedexcopyright%></span></p>
                </td>
			  </tr>
            </table>
          </td>
        </tr>
      </table>
	</form>
<%	response.flush
	sSQL = "SELECT FedexAccountNo,FedexMeter FROM adminshipping WHERE adminShipID=1"
	rs.open sSQL,cnn,0,1
	if not rs.EOF then
		fedexacctno = trim(rs("FedexAccountNo")&"")
		fedexmeter = trim(rs("FedexMeter")&"")
	end if
	rs.close
	sSQL = "SELECT adminVersion,adminZipCode,countryCode FROM admin INNER JOIN countries ON admin.adminCountry=countries.countryID WHERE adminID=1"
	rs.open sSQL,cnn,0,1
	if not rs.EOF then
		version = trim(rs("adminVersion")&"")
		zipcode = trim(rs("adminZipCode")&"")
		countrycode = trim(rs("countryCode")&"")
	end if
	rs.close
	versionarray = split(version, " v", 2)
	version = versionarray(1)
	versionarray = split(version, ".")
	if int(versionarray(0)<10) then version = "0" & versionarray(0) & versionarray(1) & "0" else version = versionarray(0) & versionarray(1) & "0"
	sXML = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns=""http://fedex.com/ws/packagemovementinformationservice/v4"">"
	sXML = sXML & "<soapenv:Header/><soapenv:Body><PostalCodeInquiryRequest><WebAuthenticationDetail><CspCredential><Key>mKOUqSP4CS0vxaku</Key><Password>IAA5db3Pmhg3lyWW6naMh4Ss2</Password></CspCredential>"
	sXML = sXML & "<UserCredential><Key>" & fedexuserkey & "</Key><Password>" & fedexuserpwd & "</Password></UserCredential>"
	sXML = sXML & "</WebAuthenticationDetail><ClientDetail><AccountNumber>" & fedexacctno & "</AccountNumber><MeterNumber>" & fedexmeter & "</MeterNumber><ClientProductId>IBTP</ClientProductId><ClientProductVersion>3272</ClientProductVersion></ClientDetail>"
	sXML = sXML & "<TransactionDetail><CustomerTransactionId>123xyz</CustomerTransactionId></TransactionDetail>"
	sXML = sXML & "<Version><ServiceId>pmis</ServiceId><Major>4</Major><Intermediate>0</Intermediate><Minor>0</Minor></Version>"
	sXML = sXML & "<CarrierCode>FDXE</CarrierCode><PostalCode>" & zipcode & "</PostalCode><CountryCode>" & countrycode & "</CountryCode></PostalCodeInquiryRequest></soapenv:Body></soapenv:Envelope>"
	
	xmlDoc="xml"
	success = callxmlfunction(fedexurl, sXML, xmlDoc, "", "Msxml2.ServerXMLHTTP", errormsg, FALSE)
	
	if success then
		Set t2 = xmlDoc.getElementsByTagName("LocationId").Item(0)
		locationid = t2.firstChild.nodeValue
		print "<script>document.getElementById('step1span').innerHTML='<strong>Completed!</strong>';</script>"
		response.flush
		sXML = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:v2=""http://fedex.com/ws/registration/v2""><soapenv:Header/><soapenv:Body><v2:VersionCaptureRequest><v2:WebAuthenticationDetail>"
		sXML = sXML & "<v2:CspCredential><v2:Key>mKOUqSP4CS0vxaku</v2:Key><v2:Password>IAA5db3Pmhg3lyWW6naMh4Ss2</v2:Password></v2:CspCredential>"
		sXML = sXML & "<v2:UserCredential><v2:Key>" & fedexuserkey & "</v2:Key><v2:Password>" & fedexuserpwd & "</v2:Password></v2:UserCredential></v2:WebAuthenticationDetail>"
		sXML = sXML & "<v2:ClientDetail><v2:AccountNumber>" & fedexacctno & "</v2:AccountNumber><v2:MeterNumber>" & fedexmeter & "</v2:MeterNumber>"
		sXML = sXML & "<v2:ClientProductId>IBTP</v2:ClientProductId><v2:ClientProductVersion>3272</v2:ClientProductVersion>"
		sXML = sXML & "<v2:Region>US</v2:Region></v2:ClientDetail><v2:TransactionDetail><v2:CustomerTransactionId>Version Capture Request</v2:CustomerTransactionId></v2:TransactionDetail>"
		sXML = sXML & "<v2:Version><v2:ServiceId>fcas</v2:ServiceId><v2:Major>2</v2:Major><v2:Intermediate>1</v2:Intermediate><v2:Minor>0</v2:Minor></v2:Version>"
		sXML = sXML & "<v2:OriginLocationId>" & trim(locationid) & "</v2:OriginLocationId>"
		sXML = sXML & "<v2:VendorProductPlatform>Windows OS</v2:VendorProductPlatform></v2:VersionCaptureRequest></soapenv:Body></soapenv:Envelope>"
		
		xmlDoc="xml"
		success = callxmlfunction(fedexurl, sXML, xmlDoc, "", "Msxml2.ServerXMLHTTP", errormsg, FALSE)
	
		print "<script>document.getElementById('step2span').innerHTML='<strong>Completed!</strong>';document.getElementById('donebutton').style.display='block';</script>"
	else
		print "<script>document.getElementById('step2span').innerHTML='<strong>Failed!</strong>';document.getElementById('donebutton').style.display='block';</script>"
	end if
elseif getpost("upsstep")="3" then
	call splitname(getpost("contact"), firstname, lastname)
	sXML="<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:v2=""http://fedex.com/ws/registration/v2""><soapenv:Header/><soapenv:Body>" & _
		"<v2:RegisterWebCspUserRequest><v2:WebAuthenticationDetail><v2:CspCredential>" & _
		"<v2:Key>mKOUqSP4CS0vxaku</v2:Key>" & _
		"<v2:Password>IAA5db3Pmhg3lyWW6naMh4Ss2</v2:Password>" & _
		"</v2:CspCredential></v2:WebAuthenticationDetail>" & _
		"<v2:ClientDetail>" & _
		"<v2:AccountNumber>" & getpost("fedexaccount") & "</v2:AccountNumber>" & _
		"<v2:ClientProductId>IBTP</v2:ClientProductId>" & _
		"<v2:ClientProductVersion>3272</v2:ClientProductVersion>" & _
		"</v2:ClientDetail>" & _
		"<v2:Version><v2:ServiceId>fcas</v2:ServiceId><v2:Major>2</v2:Major><v2:Intermediate>1</v2:Intermediate><v2:Minor>0</v2:Minor></v2:Version>" & _
		"<v2:BillingAddress>" & _
		"<v2:StreetLines>" & getpost("address") & "</v2:StreetLines>" & _
		"<v2:City>" & getpost("city") & "</v2:City>" & _
		"<v2:StateOrProvinceCode>" & getpost("usstate") & "</v2:StateOrProvinceCode>" & _
		"<v2:PostalCode>" & getpost("postcode") & "</v2:PostalCode>" & _
		"<v2:CountryCode>" & getpost("country") & "</v2:CountryCode>" & _
		"</v2:BillingAddress>" & _
		"<v2:UserContactAndAddress>" & _
		"<v2:Contact>" & _
		"<v2:PersonName>" & _
		"<v2:FirstName>" & firstname & "</v2:FirstName>" & _
		"<v2:LastName>" & lastname & "</v2:LastName>" & _
		"</v2:PersonName>" & _
		"<v2:CompanyName>" & getpost("company") & "</v2:CompanyName>" & _
		"<v2:PhoneNumber>" & getpost("telephone") & "</v2:PhoneNumber>" & _
		"<v2:EMailAddress>" & getpost("email") & "</v2:EMailAddress>" & _
		"</v2:Contact>" & _
		"<v2:Address>" & _
		"<v2:StreetLines>" & getpost("address") & "</v2:StreetLines>" & _
		"<v2:City>" & getpost("city") & "</v2:City>" & _
		"<v2:StateOrProvinceCode>" & getpost("usstate") & "</v2:StateOrProvinceCode>" & _
		"<v2:PostalCode>" & getpost("postcode") & "</v2:PostalCode>" & _
		"<v2:CountryCode>" & getpost("country") & "</v2:CountryCode>" & _
		"</v2:Address>" & _
		"</v2:UserContactAndAddress>" & _
		"</v2:RegisterWebCspUserRequest>" & _
		"</soapenv:Body>" & _
		"</soapenv:Envelope>"

	xmlres="xml"
	success = callxmlfunction(fedexurl, sXML, xmlres, "", "Msxml2.ServerXMLHTTP", errormsg, FALSE)	
	errormsg="Unknown Error"

	if success then
		set obj2 = xmlres.getElementsByTagName("Severity")
		if obj2.length > 0 then
			if obj2.item(0).hasChildNodes then success = trim(obj2.item(0).firstChild.nodeValue&"")="SUCCESS"
		end if
	end if
	if success then
		userkey=""
		userpwd=""
		set obj2 = xmlres.getElementsByTagName("Key")
		if obj2.length > 0 then
			if obj2.item(0).hasChildNodes then userkey = obj2.item(0).firstChild.nodeValue
		end if
		set obj2 = xmlres.getElementsByTagName("Password")
		if obj2.length > 0 then
			if obj2.item(0).hasChildNodes then userpwd = obj2.item(0).firstChild.nodeValue
		end if
		if NOT (userkey<>"" AND userpwd<>"") then
			success=FALSE
		end if
	else
		set obj2 = xmlres.getElementsByTagName("Message")
		if obj2.length > 0 then
			if obj2.item(0).hasChildNodes then errormsg = obj2.item(0).firstChild.nodeValue
		end if
	end if

	if success then
		sXML="<soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">" & _
			"<soap:Body>" & _
			"<SubscriptionRequest xmlns=""http://fedex.com/ws/registration/v2"">" & _
			"<WebAuthenticationDetail>" & _
			"<CspCredential>" & _
			"<Key>mKOUqSP4CS0vxaku</Key>" & _
			"<Password>IAA5db3Pmhg3lyWW6naMh4Ss2</Password>" & _
			"</CspCredential>" & _
			"<UserCredential>" & _
			"<Key>"&userkey&"</Key>" & _
			"<Password>"&userpwd&"</Password>" & _
			"</UserCredential>" & _
			"</WebAuthenticationDetail>" & _
			"<ClientDetail>" & _
			"<AccountNumber>" & getpost("fedexaccount") & "</AccountNumber>" & _
			"<MeterNumber/>" & _
			"<ClientProductId>IBTP</ClientProductId><ClientProductVersion>3272</ClientProductVersion>" & _
			"</ClientDetail>" & _
			"<Version><ServiceId>fcas</ServiceId><Major>2</Major><Intermediate>1</Intermediate><Minor>0</Minor></Version>" & _
			"<CspSolutionId>100</CspSolutionId>" & _
			"<CspType>CERTIFIED_SOLUTION_PROVIDER</CspType>" & _
			"<Subscriber>" & _
			"<AccountNumber>" & getpost("fedexaccount") & "</AccountNumber>" & _
			"<Contact>" & _
			"<PersonName>" & getpost("contact") & "</PersonName>" & _
			"<CompanyName/>" & _
			"<PhoneNumber>" & getpost("telephone") & "</PhoneNumber>" & _
			"<FaxNumber/>" & _
			"<EMailAddress>" & getpost("email") & "</EMailAddress>" & _
			"</Contact><Address>" & _
			"<StreetLines>" & getpost("address") & "</StreetLines>" & _
			"<City>" & getpost("city") & "</City>" & _
			"<StateOrProvinceCode>" & getpost("usstate") & "</StateOrProvinceCode>" & _
			"<PostalCode>" & getpost("postcode") & "</PostalCode>" & _
			"<CountryCode>" & getpost("country") & "</CountryCode>" & _
			"</Address></Subscriber>" & _
			"<AccountShippingAddress>" & _
			"<StreetLines>" & getpost("address") & "</StreetLines>" & _
			"<City>" & getpost("city") & "</City>" & _
			"<StateOrProvinceCode>" & getpost("usstate") & "</StateOrProvinceCode>" & _
			"<PostalCode>" & getpost("postcode") & "</PostalCode>" & _
			"<CountryCode>" & getpost("country") & "</CountryCode>" & _
			"</AccountShippingAddress>" & _
			"</SubscriptionRequest></soap:Body></soap:Envelope>"
			
		xmlres="xml"
		success = callxmlfunction(fedexurl, sXML, xmlres, "", "Msxml2.ServerXMLHTTP", errormsg, FALSE)
		errormsg="Unknown Error"
	
		success=FALSE
		set obj2 = xmlres.getElementsByTagName("Severity")
		if obj2.length > 0 then
			if obj2.item(0).hasChildNodes then success = trim(obj2.item(0).firstChild.nodeValue&"")="SUCCESS"
		end if
		if success then
			fedexmeter=""
			set obj2 = xmlres.getElementsByTagName("MeterNumber")
			if obj2.length > 0 then
				if obj2.item(0).hasChildNodes then fedexmeter = obj2.item(0).firstChild.nodeValue
			end if
			if fedexmeter="" then success=FALSE
		else
			set obj2 = xmlres.getElementsByTagName("Message")
			if obj2.length > 0 then
				if obj2.item(0).hasChildNodes then errormsg = obj2.item(0).firstChild.nodeValue
			end if
		end if
	end if
%>
	<form method="post" name="licform" action="admin.asp">
	  <input type="hidden" name="upsstep" value="5" />
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr> 
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr>
				<td rowspan="3" width="70" align="center" valign="top"><img src="../images/fedexlogo.png" border="0" alt="FedEx" /><br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</td>
                <td width="100%" align="center"><strong><%=yyFdxWiz%> - <% if success then print yyRegSucc else print yyError %></strong><br />&nbsp;
                </td>
			  </tr>
<%	if success then
		sSQL = "UPDATE adminshipping SET FedexAccountNo='"&getpost("fedexaccount")&"',FedexMeter='"&fedexmeter&"',FedexUserKey='"&userkey&"',FedexUserPwd='"&userpwd&"'"
		ect_query(sSQL)
%>
			  <tr> 
                <td width="100%" align="left">
				  <p><strong><%=yyRegSucc%> !</strong></p>
				  <p>Thank you for registering.</p>
				  <p>To learn more about FedEx&reg; shipping services please go to <a href="http://www.fedex.com" target="_blank">fedex.com</a>.</p>
				  <p>To begin using FedEx shipping calculations please don't forget to select FedEx Shipping from the <strong>Shipping Type</strong> dropdown in the page <a href="adminmain.asp"><%=yyAdmMai%></a>.</p>
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
				  <p><br />&nbsp;</p>
				  <p><span style="font-size:10px"><%=fedexcopyright%></span></p>
                </td>
			  </tr>
            </table>
          </td>
        </tr>
      </table>
	</form>
<%
elseif getpost("upsstep")="2" then
%>
<script>
<!--
function checkforamp(checkObj){
  checkStr = checkObj.value;
  for (i = 0;  i < checkStr.length;  i++){
	if (checkStr.charAt(i) == "&"){
	  alert("Please do not use the ampersand \"&\" character in any field.");
	  checkObj.focus();
	  return(false);
	}
  }
  return(true);
}
function formvalidator(theForm)
{
  if(theForm.contact.value == ""){
    alert("<%=jscheck(yyPlsEntr&" """&yyConNam)%>\".");
    theForm.contact.focus();
    return (false);
  }
  if(!checkforamp(theForm.contact)) return(false);
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
  var cntry = theForm.country[theForm.country.selectedIndex].value;
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
  if (theForm.postcode.value == ""){
	alert("<%=jscheck(yyPlsEntr&" """&yyPCode)%>\".");
	theForm.postcode.focus();
	return (false);
  }
  if(!checkforamp(theForm.postcode)) return(false);
  if(theForm.telephone.value == ""){
    alert("<%=jscheck(yyPlsEntr&" """&yyTelep)%>\".");
    theForm.telephone.focus();
    return (false);
  }
  if(theForm.telephone.value.length < 6 || theForm.telephone.value.length > 16){
    alert("<%=jscheck(yyValTN)%>");
    theForm.telephone.focus();
    return (false);
  }
  var checkOK = "0123456789";
  var checkStr = theForm.telephone.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++){
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if(!allValid){
    alert("<%=jscheck(yyOnDig&" """&yyTelep)%>\".");
    theForm.telephone.focus();
    return (false);
  }
  if (theForm.email.value == ""){
	alert("<%=jscheck(yyPlsEntr&" """&yyEmail)%>\".");
	theForm.email.focus();
	return (false);
  }
  if(!checkforamp(theForm.fedexaccount)) return(false);
  if(theForm.fedexaccount.value == ""){
    alert("<%=jscheck(yyPlsEntr)%> \"FedEx Account Number\".");
    theForm.fedexaccount.focus();
    return (false);
  }
  var checkOK = "0123456789";
  var checkStr = theForm.fedexaccount.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++){
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if(!allValid){
    alert("<%=jscheck(yyOnDig)%> \"FedEx Account Number\".");
    theForm.fedexaccount.focus();
    return (false);
  }
  return (true);
}
//-->
</script>
	<form method="post" name="licform" action="adminfedexlicense.asp" onsubmit="return formvalidator(this)">
	  <input type="hidden" name="upsstep" value="3" />
	  <input type="hidden" name="countryCode" value="<%=getpost("countryCode")%>" />
	  <input type="hidden" name="languageCode" value="<%=getpost("languageCode")%>" />
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr> 
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr>
				<td rowspan="18" width="70" align="center" valign="top"><img src="../images/fedexlogo.png" border="0" alt="FedEx" /><br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</td>
                <td width="100%" align="center" colspan="2"><strong><%=yyFdxWiz%></strong><br />&nbsp;
                </td>
			  </tr>
			  <tr> 
                <td width="40%" align="right"><strong><%=redasterix&yyConNam%> : </strong></td>
				<td width="60%"><input type="text" name="contact" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyComNam%> : </strong></td>
				<td><input type="text" name="company" size="15" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong>Department : </strong></td>
				<td><input type="text" name="department" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=redasterix&yyStrAdd%> : </strong></td>
				<td><input type="text" name="address" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=yyAddr2%> : </strong></td>
				<td><input type="text" name="address2" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=redasterix&yyCity%> : </strong></td>
				<td><input type="text" name="city" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=redasterix&yyState%> <%=yyUSCan%> : </strong></td>
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
                <td align="right"><strong><%=redasterix&yyCountry%> : </strong></td>
				<td><select name="country" size="1">
<option value=''><%=yySelect%></option>
<%
' sSQL = "SELECT countryName,countryCode FROM countries WHERE countryCode IN ('AR','AU','AT','BE','BR','CA','CL','CN','CO','CR','DK','DO','FI','FR','DE','GR','GT','HK','IN','IE','IL','IT','JP','MY','MX','NL','NZ','NO','PA','PH','PT','PR','SG','KR','ES','SE','CH','TW','TH','GB','US') ORDER BY countryName"
sSQL = "SELECT countryName,countryCode FROM countries WHERE countryCode IN ('US','CA') ORDER BY countryName"
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
                <td align="right"><strong><%=redasterix&yyPCode%> : </strong></td>
				<td><input type="text" name="postcode" size="15" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=redasterix&yyTelep%> : </strong></td>
				<td><input type="text" name="telephone" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong>Pager Number : </strong></td>
				<td><input type="text" name="pager" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong>Fax Number : </strong></td>
				<td><input type="text" name="fax" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=redasterix&yyEmail%> : </strong></td>
				<td><input type="text" name="email" size="30" /></td>
			  </tr>
			  <tr> 
                <td align="right"><strong><%=redasterix%>FedEx Account Number : </strong></td>
				<td><input type="text" name="fedexaccount" size="30" /></td>
			  </tr>
			  <tr>
                <td width="100%" align="center" colspan="2"><br />&nbsp;<input type="submit" name="agree" value="&nbsp;&nbsp;<%=yyNext%>&nbsp;&nbsp;" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="button" value="<%=yyCancel%>" onclick="window.location='admin.asp';" />
                </td>
			  </tr>
			  <tr> 
                <td colspan="2" width="100%" align="center">
				  <p><br />&nbsp;</p>
				  <p><span style="font-size:10px"><%=fedexcopyright%></span></p>
                </td>
			  </tr>
            </table>
          </td>
        </tr>
      </table>
	</form>
<%
else %>
	<form method="post" action="adminfedexlicense.asp">
	  <input type="hidden" name="upsstep" value="2" />
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr> 
          <td width="100%">
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr>
				<td rowspan="5" width="70" align="center" valign="top"><img src="../images/fedexlogo.png" border="0" alt="" /><br />&nbsp;</td>
                <td width="100%" align="center"><strong><%=yyFdxWiz%></strong><br />&nbsp;
                </td>
			  </tr>
<%	isregistered=FALSE
	sSQL = "SELECT FedexAccountNo,FedexMeter FROM adminshipping WHERE adminShipID=1"
	rs.open sSQL,cnn,0,1
	if not rs.EOF then
		if trim(rs("FedexAccountNo")&"")<>"" AND trim(rs("FedexMeter")&"")<>"" then isregistered=true
	end if
	rs.close
	if isregistered then %>
			  <tr> 
                <td width="100%">You have already successfully completed the FedEx licensing and registration wizard. If you would like to re-register then please 
				click the "Re-register" button below. If you would just like to update your Ecommerce Plus version information with 
				FedEx then please click the "Update Version" button below.
				<p>&nbsp;</p>
                </td>
			  </tr>
			  <tr> 
                <td width="100%" align="center"><input type="submit" name="agree" value="&nbsp;&nbsp;Re-Register&nbsp;&nbsp;" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="button" value="Update Version" onclick="window.location='adminfedexlicense.asp?act=version';" />
                </td>
			  </tr>
<%	else %>
			  <tr> 
                <td width="100%"><ul><li>This wizard will assist you in completing the necessary licensing and registration requirements to activate and use the FedEx&reg; Rates and Tracking services from your Ecommerce Plus Template.<br /><br /></li>
				<li>If you do not wish to use any of the functions that utilize the FedEx Rates and Tracking services, click the Cancel button and those functions will not be enabled. If, at a later time, you wish to use these services, return to this section and complete the FedEx licensing and registration process.<br /><br /></li>
				<li>For more information about FedEx services, please <a href="http://www.fedex.com" target="_blank"><%=yyClkHer%></a>.<br /><br /></li>
				</ul>
				<p>&nbsp;</p>
                </td>
			  </tr>
			  <tr> 
                <td width="100%" align="center"><input type="submit" name="agree" value="&nbsp;&nbsp;<%=yyNext%>&nbsp;&nbsp;" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="button" value="<%=yyCancel%>" onclick="window.location='admin.asp';" />
                </td>
			  </tr>
<%	end if %>
			  <tr> 
                <td align="center" colspan="2"><p><span style="font-size:10px"><br />To open a FedEx account, please <a href="https://www.fedex.com/us/OADR/index.html?link=4" target="_blank"><strong><%=yyClkHer%></strong></a><br /></span></p></td>
			  </tr>
			  <tr> 
                <td width="100%" align="center">
				  <p><br />&nbsp;</p>
				  <p><span style="font-size:10px"><%=fedexcopyright%></span></p>
                </td>
			  </tr>
            </table>
          </td>
        </tr>
      </table>
	</form>
<%
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>