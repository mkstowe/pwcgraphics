<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
Dim sSQL,rs,alldata,success,cnn,rowcounter,errmsg,data1name,data2name,isenabled,demomode,vsdetails,bitname(12)
success=true
demomodeavailable=true
if maxloginlevels="" then maxloginlevels=5
Set rs=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
ect_query("UPDATE payprovider SET payProvAvailable=0,payProvEnabled=0 WHERE payProvID=20")
alreadygotadmin=getadminsettings()
if getpost("act")="domodify" AND is_numeric(getpost("id")) then
	sSQL="SELECT payProvName FROM payprovider WHERE payProvID="&getpost("id")
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then payprovname=rs("payProvName") else payprovname="NOT KNOWN"
	rs.close
	call logevent(SESSION("loginuser"),"PAYPROVIDER",TRUE,"adminpayprov.asp","MODIFY "&payprovname)
	isenabled=0
	demomode=0
	if getpost("isenabled")="1" then isenabled=1
	if getpost("demomode")="1" then demomode=1
	if is_numeric(getpost("pphandlingcharge")) then handlingcharge=getpost("pphandlingcharge") else handlingcharge=0
	if is_numeric(getpost("pphandlingpercent")) then handlingpercent=getpost("pphandlingpercent") else handlingpercent=0
	sSQL="UPDATE payprovider SET payProvShow='"&escape_string(getpost("showas"))&"',payProvEnabled="&isenabled&",payProvDemo="&demomode&",payProvLevel="&getpost("payProvLevel")&",ppHandlingCharge=" & handlingcharge & ",ppHandlingPercent="& handlingpercent &","
	if getpost("id")="5" then ' WorldPay
		sSQL=sSQL & "payProvData1='"&escape_string(getpost("data1"))&"',payProvData2='"&escape_string(getpost("data2"))&"&"&escape_string(getpost("data3"))&"'"
	elseif getpost("id")="7" OR getpost("id")="22" OR (getpost("id")="8" AND (getpost("data1")<>"" OR getpost("data4")<>"")) then ' Payflow Pro OR PayPal Advanced
		sSQL=sSQL & "payProvData1='"&escape_string(getpost("data1"))&"&"&escape_string(getpost("data2"))&"&"&escape_string(getpost("data3"))&"&"&escape_string(getpost("data4"))&"'"
	elseif getpost("id")="8" then ' Payflow Link
		sSQL=sSQL & "payProvData1='"&escape_string(getpost("data2"))&"',payProvData2='"&escape_string(getpost("data3"))&"',payProvData3=''"
	elseif getpost("id")="9" then ' PayPoint.net
		sSQL=sSQL & "payProvData1='"&escape_string(getpost("data1"))&"',payProvData2='"&escape_string(getpost("data2")&"&"&urlencode(getpost("data2supp")))&"',payProvData3='"&escape_string(getpost("data3"))&"'"
	elseif getpost("id")="10" then ' Capture Card
		data1=""
		for index=1 to 20
			if getpost("cardtype" & index)="X" then
				data1=data1&"X"
			else
				data1=data1&"O"
			end if
		next
		sSQL=sSQL & "payProvData1='"&data1&"'"
		if getpost("data2")<>"" then
			admincert=getpost("data2")
			admincert=replace(admincert,"-----BEGIN PUBLIC KEY-----","")
			admincert=replace(admincert,"-----END PUBLIC KEY-----","")
			admincert=trim(admincert)
			certlength=len(admincert)
			if certlength>500 OR instr(admincert,"PRIVATE")>0 then
				success=FALSE
				errmsg="Please do not upload the private key. You should only upload the public key."
			else
				sSQL2="UPDATE admin SET adminCert='"&escape_string(admincert)&"' WHERE adminID=1"
				ect_query(sSQL2)
			end if
		end if
	elseif getpost("id")="18" OR getpost("id")="19" then ' PayPal Pro
		if getpost("ppexpressab")="AB" then
			sSQL=sSQL & "payProvData1='@AB@"&escape_string(getpost("ppexpressabemail"))&"',payProvData2='"&escape_string(urlencode(getpost("data2")))&"&"&IIfVr(getpost("billmelater")="1","1","0")&"',payProvData3='"&escape_string(getpost("data3"))&"'"
		else
			sSQL=sSQL & "payProvData1='"&escape_string(getpost("data1"))&"',payProvData2='"&escape_string(urlencode(getpost("data2")))&"&"&IIfVr(getpost("billmelater")="1","1","0")&"',payProvData3='"&escape_string(getpost("data3"))&"'"
		end if
	elseif getpost("id")="21" then ' Amazon Pay
		thedata2=replace(getpost("data2"),"&","") & "&" & replace(getpost("data2b"),"&","")
		sSQL=sSQL & "payProvData1='" & escape_string(getpost("data1")) & "',payProvData2='" & escape_string(thedata2) & "',payProvData3='" & escape_string(getpost("data3")) & "'"
	else
		thedata1=getpost("data1")
		thedata2=getpost("data2")
		thedata3=getpost("data3")
		thedata4=getpost("data4")
		if secretword<>"" AND (getpost("id")="3" OR getpost("id")="13") then
			thedata1=upsencode(thedata1, secretword)
			thedata2=upsencode(thedata2, secretword)
		elseif getpost("id")="27" then ' PayPal Checkout
			thedata3=getpost("buttonshape")&"|"&getpost("buttonsize")&"|"&getpost("buttoncolor")&"|"&getpost("buttonlayout")&"|"&getpost("paypalcredit")&"|"&getpost("paypalcards")&"|"&getpost("paypalelv")&"|"&getpost("paypalvenmo")
		end if
		sSQL=sSQL & "payProvData1='"&escape_string(thedata1)&"',payProvData2='"&escape_string(thedata2)&"',payProvData3='"&escape_string(thedata3)&"',payProvData4='"&escape_string(thedata4)&"'"
	end if
	databits=0
	for index=1 to 12
		if getpost("databit"&index)="1" then databits=databits+(2 ^ (index-1))
	next
	sSQL=sSQL&",payProvBits='" & databits & "'"
	sSQL=sSQL&",payProvFlag1='" & IIfVr(is_numeric(getpost("payProvFlag1")),getpost("payProvFlag1"),0) & "'"
	for index=2 to adminlanguages+1
		if (adminlangsettings AND 128)=128 then
			sSQL=sSQL & ",payProvShow" & index & "='"&escape_string(getpost("showas" & index))&"'"
		end if
	next
	for index=1 to adminlanguages+1
		languageid=index
		if index=1 OR (adminlangsettings AND 4096)=4096 then
			pprovheaders=getpost("pprovheaders" & index)
			pprovdropshipheaders=getpost("pprovdropshipheaders" & index)
			if NOT (htmlemails AND (htmleditor="froala" OR htmleditor="ckeditor")) then
				pprovheaders=replace(pprovheaders, vbCrLf, "<br />")
				pprovdropshipheaders=replace(pprovdropshipheaders, vbCrLf, "<br />")
			end if
			sSQL=sSQL & "," & getlangid("pProvHeaders",4096) & "='"&escape_string(pprovheaders)&"'"
			sSQL=sSQL & "," & getlangid("pProvDropShipHeaders",4096) & "='"&escape_string(pprovdropshipheaders)&"'"
		end if
	next
	if getpost("transtype")<>"" then sSQL=sSQL & ",payProvMethod=" & getpost("transtype")
	sSQL=sSQL & " WHERE payProvID="&getpost("id")
	ect_query(sSQL)
	if getpost("id")="18" OR getpost("id")="19" then ' PayPal Pro
		sSQL="UPDATE payprovider SET payProvDemo="&demomode&",payProvData1='"&escape_string(getpost("data1"))&"',payProvData2='"&escape_string(urlencode(getpost("data2")))&"&"&IIfVr(getpost("billmelater")="1","1","0")&"',payProvData3='"&escape_string(getpost("data3"))&"',payProvMethod=" & getpost("transtype")
		if getpost("ppexpressab")="AB" then
			sSQL="UPDATE payprovider SET payProvDemo="&demomode&",payProvEnabled=0"
		elseif getpost("id")="18" then
			if isenabled=1 then sSQL=sSQL & ",payProvEnabled=1"
			sSQL=sSQL & " WHERE payProvID=19"
		end if
		if getpost("id")="19" then sSQL=sSQL & " WHERE payProvID=18"
		ect_query(sSQL)
	end if
	if success then
		print "<meta http-equiv=""refresh"" content=""1; url=adminpayprov.asp"
		if getpost("offerpaypal")="ON" then
			print "?act=modify&from=wizard2&id=1"
		else
			if getpost("from")="wizard" OR getpost("from")="wizard2" then print "?act=alternate" else print "?act=list"
		end if
		print """ />"
	end if
elseif getpost("act")="changepos" AND is_numeric(getpost("selectedq")) AND is_numeric(getpost("newval")) then
	currentorder=int(getpost("selectedq"))
	neworder=int(getpost("newval"))
	sSQL="SELECT payProvID FROM payprovider ORDER BY payProvEnabled DESC,payProvOrder"
	rs.open sSQL,cnn,0,1
	alldata=rs.getrows
	rs.close
	for rowcounter=0 to ubound(alldata,2)
		theorder=rowcounter+1
		if currentorder=theorder then
			theorder=neworder
		elseif (currentorder>theorder) AND (neworder <= theorder) then
			theorder=theorder + 1
		elseif (currentorder < theorder) AND (neworder>=theorder) then
			theorder=theorder - 1
		end if
		sSQL="UPDATE payprovider SET payProvOrder="&theorder&" WHERE payProvID="&alldata(0,rowcounter)
		ect_query(sSQL)
	next
	print "<meta http-equiv=""refresh"" content=""1; url=adminpayprov.asp?act=list"">"
end if
%>
<script>
/* <![CDATA[ */
function modrec(id) {
	document.mainform.id.value=id;
	document.mainform.act.value="modify";
	document.mainform.submit();
}
function switchheader(id){
	var thestyle=document.getElementById(id).style.display;
	if(thestyle=='block')
		document.getElementById(id).style.display='none';
	else{
		document.getElementById(id).style.display='block';
<%	if htmleditor="froala" then print "eval(""dfe_""+id.replace(/span/,""pprov"")+""()"");" %>
	}
}
/* ]]> */
</script>
<%	if getpost("act")="domodify" AND success then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
				<%=yyNoAuto%> <a href="adminpayprov.asp<%=IIfVr(getpost("offerpaypal")="ON", "?act=modify&from=wizard2&id=1", "")%>"><strong><%=yyClkHer%></strong></a>.<br /><br />&nbsp;
                </td>
			  </tr>
			</table></td>
        </tr>
      </table>
<%	elseif getpost("act")="domodify" then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><span style="color:#FF0000;font-weight:bold"><%=yyOpFai%></span><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a><br />&nbsp;</td>
			  </tr>
			</table></td>
        </tr>
      </table>
<%	elseif request("act")="alternate" then
		sSQL="SELECT payProvID,payProvEnabled FROM payprovider WHERE payProvID in (4,10,14,19)"
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			if rs("payProvID")=4 then emailenabled=(cint(rs("payProvEnabled"))<>0)
			if rs("payProvID")=10 then capturecardenabled=(cint(rs("payProvEnabled"))<>0)
			if rs("payProvID")=14 then customenabled=(cint(rs("payProvEnabled"))<>0)
			if rs("payProvID")=19 then ppexpressenabled=(cint(rs("payProvEnabled"))<>0)
			rs.movenext
		loop
		rs.close
%>
		  <form name="mainform" method="post" action="adminpayprov.asp?from=wizard2">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="domodify" />
			<input type="hidden" name="id" value="" />
			<input type="hidden" name="from" value="wizard2" />
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr> 
                <td width="100%" colspan="2" align="center"><strong><%=yyPPAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="2" align="center">
				  <table width="80%" border="0" cellspacing="2" cellpadding="2" bgcolor="#FFFFFF">
					<tr>
					  <td align="left" valign="top" bgcolor="#FFFFFF">
						<div>&nbsp;</div>
						<div style="font-size:18px;font-weight:bold;">Set-up Other Payment Options</div>
					  </td>
					</tr>
				  </table>
				  <br />&nbsp;<br />
				  <table width="80%" border="0" cellspacing="2" cellpadding="2" bgcolor="#BFC9E0">
					<tr>
					  <td align="left" valign="top" bgcolor="#<%if ppexpressenabled then print "E6E6E6" else print "FFFFFF"%>">
						<div style="font-size:12px;font-weight:bold;margin:5px;"><input type="checkbox" <%if NOT ppexpressenabled then print "onclick=""modrec(19)""" else print "checked=""checked"" disabled=""disabled"""%>/> <a <%if NOT ppexpressenabled then print "href=""javascript:modrec(19)"""%>>PayPal Express Checkout</a></div>
						<div style="font-size:11px;margin:15px;">According to Jupiter Research, 23% of online shoppers consider PayPal one of their favourite ways to pay online.* 
						Accepting PayPal in addition to Credit Cards is proven to increase your sales.**
						<p align="right" style="margin:0px;font-weight:bold;"><a href="" onclick="newwin=window.open('http://www.paypal.com/en_US/m/demo/wppro/paypal_demo_load_560x355.html','PayPalDemo','menubar=no,scrollbars=yes,width=578,height=372,directories=no,location=no,resizable=yes,status=no,toolbar=no');return false;">See quick demo...</a></p>
						</div>
					  </td>
					</tr>
				  </table>
				  <br />&nbsp;<br />
				  <table width="80%" border="0" cellspacing="2" cellpadding="2" bgcolor="#BFC9E0">
					<tr>
					  <td align="left" valign="top" bgcolor="#FFFFFF">
						<div style="font-size:12px;font-weight:bold;margin:5px;"><input type="checkbox" onclick="modrec(4)" <% if emailenabled then print "checked=""checked"""%>/> <a href="javascript:modrec(4)">Email Payment Method</a></div>
						<div style="font-size:11px;margin:15px;">This payment method will simply collect order information and notify the store owner by email if required. It 
						can be used for instace for a Cash-On-Delivery type payment method.</div>
					  </td>
					</tr>
				  </table>
				  <br />&nbsp;<br />
				  <table width="80%" border="0" cellspacing="2" cellpadding="2" bgcolor="#BFC9E0">
					<tr>
					  <td align="left" valign="top" bgcolor="#FFFFFF">
						<div style="font-size:12px;font-weight:bold;margin:5px;"><input type="checkbox" onclick="modrec(10)" <% if capturecardenabled then print "checked=""checked"""%>/> <a href="javascript:modrec(10)">Capture Card</a></div>
						<div style="font-size:11px;margin:15px;">This payment method will collect credit card numbers and store them in your database. Unless you are really sure 
						of what you are doing you are highly recommended to use an online payment gateway or PayPal. Using this method means you are responsible for your 
						the security of your customers credit card details.
						<p align="right" style="margin:0px;font-weight:bold;"><a href="https://www.ecommercetemplates.com/help/ecommplus/capture_card.asp" target="_blank">More Details...</a></p>
						</div>
					  </td>
					</tr>
				  </table>
				  <br />&nbsp;<br />
				  <table width="80%" border="0" cellspacing="2" cellpadding="2" bgcolor="#BFC9E0">
					<tr>
					  <td align="left" valign="top" bgcolor="#FFFFFF">
						<div style="font-size:12px;font-weight:bold;margin:5px;"><input type="checkbox" onclick="modrec(14)" <% if customenabled then print "checked=""checked"""%>/> <a href="javascript:modrec(14)">Custom Payment Provider</a></div>
						<div style="font-size:11px;margin:15px;">Select this method to configure a custom payment provider. Please click on the link to see a list of the customer payment providers supported.
						<p align="right" style="margin:0px;font-weight:bold;"><a href="https://www.ecommercetemplates.com/help/ecommplus/capture_card.asp" target="_blank">Custom Payment Providers...</a></p>
						</div>
					  </td>
					</tr>
				  </table>
				  <table width="80%" border="0" cellspacing="2" cellpadding="2" bgcolor="#FFFFFF">
					<tr>
					  <td align="left" valign="top" bgcolor="#FFFFFF">
						<p><span style="font-size:11px;font-weight:bold;"><a href="adminpayprov.asp?act=list">See full list of payment processors</a></span></p>
					  </td>
					</tr>
				  </table>
				  <br />&nbsp;
				</td>
			  </tr>
			</table>
		  </form>
<%	elseif request("act")="modify" AND is_numeric(request("id")) then
		sSQL="SELECT payProvID,payProvName,payProvShow,payProvDemo,payProvEnabled,payProvData1,payProvData2,payProvData3,payProvData4,payProvFlag1,payProvMethod,payProvLevel,ppHandlingCharge,ppHandlingPercent,payProvShow2,payProvShow3,payProvBits FROM payprovider WHERE payProvAvailable=1 AND payProvID=" & request("id")
		rs.open sSQL,cnn,0,1
			payProvID=trim(rs("payProvID")&"")
			payProvName=trim(rs("payProvName")&"")
			payProvShow=trim(rs("payProvShow")&"")
			payProvDemo=rs("payProvDemo")
			payProvEnabled=rs("payProvEnabled")
			payProvData1=trim(rs("payProvData1")&"")
			payProvData2=trim(rs("payProvData2")&"")
			payProvData3=trim(rs("payProvData3")&"")
			payProvData4=trim(rs("payProvData4")&"")
			payProvFlag1=rs("payProvFlag1")
			payProvMethod=rs("payProvMethod")
			payProvLevel=rs("payProvLevel")
			payProvShow2=trim(rs("payProvShow2")&"")
			payProvShow3=trim(rs("payProvShow3")&"")
			pphandlingcharge=rs("ppHandlingCharge")
			pphandlingpercent=rs("ppHandlingPercent")
			payProvBits=rs("payProvBits")
		rs.close
		data2name="" : data3name="" : signuppage=""
		hasauthtype=FALSE
		if payProvID=1 then ' PayPal
			payProvName="PayPal Payments Standard"
			signuppage="https://www.paypal.com/us/webapps/mpp/referral/paypal-payments-standard?partner_id=39HT54MCDMV8E"
			data1name=yyEmail
			data2name="Identity Token<br /><span style=""font-size:10px"">(Only when using PDT)</span>"
			demomodeavailable=true
			yyDemoMo="Sandbox"
		elseif payProvID=2 then ' 2Checkout
			signuppage="https://www.2checkout.com/referral?r=etemplates"
			data1name=yyAccNum
			data2name=yyMD5H
			warning1=TRUE
		elseif payProvID=3 OR payProvID=13 then ' Authorize.net
			signuppage="https://go.evopayments.us/ecommercetemplates"
			data1name=yyMercLID
			data2name=yyTrnKey
			if payProvID=3 then data3name="MD5 Hash / Signature Key" else bitname(1)="Accept eChecks"
			if secretword<>"" then
				payProvData1=upsdecode(payProvData1, secretword)
				payProvData2=upsdecode(payProvData2, secretword)
			end if
		elseif payProvID=4 OR payProvID=17 then ' Email
			'data1name=yyEAOrd
			data1name=""
			demomodeavailable=false
		elseif payProvID=5 then ' World Pay
			signuppage="https://business.worldpay.com/partner/ecommerce-templates"
			data1name=yyAccNum
			data2name=yyMD5H
			warning1=TRUE
		elseif payProvID=6 then ' NOCHEX
			signuppage="https://secure.nochex.com/apply/merchant.aspx?partner_id=213200427"
			data1name=yyEmail
		elseif payProvID=7 then ' Payflow Pro
			signuppage="https://www.paypal.com/us/webapps/mpp/referral/paypal-payflow-gateway?partner_id=39HT54MCDMV8E"
			payProvName="PayPal Payflow Pro"
		elseif payProvID=8 then ' Payflow Link
			signuppage="https://www.paypal.com/us/webapps/mpp/referral/paypal-payflow-gateway?partner_id=39HT54MCDMV8E"
			payProvName="PayPal Payflow Link"
		elseif payProvID=9 then ' PayPoint.net
			data1name=yyMercID
			data2name=yyMD5H
			warning1=TRUE
		elseif payProvID=10 then ' Capture Card
			demomodeavailable=false
		elseif payProvID=11 OR payProvID=12 then ' PSiGate
			data1name=yyMercID
		elseif payProvID=14 then ' Custom Payment Processor
			data1name="Data 1"
			data2name="Data 2"
			data3name="Data 3"
		elseif payProvID=15 then ' Netbanx
			signuppage="http://www1.netbanx.com/campaign/REF_ECOMT_PROG.html"
			data1name=yyMercID
			data2name="Checksum"
			demomodeavailable=false
		elseif payProvID=16 then ' Linkpoint
			signuppage="http://ecommercetemplates.cardpay-solutions.com/"
			data1name=yyNumSto
			data2name=yyOwnSit
		elseif payProvID=18 OR payProvID=19 then ' PayPal Payment Pro
			data2arr=split(trim(payProvData2&""),"&")
			if UBOUND(data2arr)>=0 then data2pwd=urldecode(data2arr(0))
			if UBOUND(data2arr)>=1 then wantbillmelater=(data2arr(1)="1") else wantbillmelater=FALSE
			data2hash=payProvData3
			if payProvID=18 then payProvName="PayPal Direct Payments" else payProvName="PayPal Express Payments"
			signuppage="https://www.paypal.com/us/webapps/mpp/referral/paypal-payments-pro?partner_id=39HT54MCDMV8E"
			data1name=yyApiAcN
			data2name=yyApiPw&".<br /><span style=""font-size:10px"">("&yyNoPPP&")</span>"
			yyDemoMo="Sandbox"
		elseif payProvID=20 then ' Google Checkout
			signuppage="http://checkout.google.com/sell?promo=sectem"
			data1name=yyGMerID
			data2name=yyGMerKe
			yyDemoMo="Sandbox"
		elseif payProvID=21 then ' Amazon Pay
			signuppage=""
			data1name="Client ID"
			data2name="AWS Access Key"
			data3name="Secret Access Key"
			' data3name="Account ID (Optional)"
			yyDemoMo="Sandbox"
		elseif payProvID=22 then ' PayPal Advanced
			signuppage="https://www.paypal.com/us/webapps/mpp/referral/paypal-payments-advanced?partner_id=39HT54MCDMV8E"
			payProvName="PayPal Payments Advanced"
		elseif payProvID=23 then ' Stripe.com
			data1name="Secret Key"
			data2name="Publishable Key"
			data3name="Store Name"
			'data4name="Webhook Signing Secret"
			demomodeavailable=FALSE
		elseif payProvID=24 then ' SagePay
			signuppage="https://support.sagepay.com/apply/default.aspx?PartnerID=%7B7B0AD331-0388-44EA-BE3A-D05D3FB9FE28%7D"
			data1name="Vendor name"
			data2name="Encryption Password"
		elseif payProvID=27 then ' PayPal Checkout
			data1name="Client ID"
			data2name="Password"
		elseif payProvID=28 then ' SquareUp
			data1name="Application ID"
			data2name="Access Token"
			data3name="Location ID (Required for SCA)"
		elseif payProvID=29 then ' NMI
			data1name="API Key"
		elseif payProvID=30 then ' eWay
			data1name="API Key"
			data2name="Password"
		elseif payProvID=31 then ' Pay360
			hasauthtype=TRUE
			data1name="instId"
			data2name="API Username"
			data3name="API Password"
		elseif payProvID=32 then ' Global Payments
			hasauthtype=TRUE
			data1name="Merchant ID"
			data2name="Shared Secret"
			data3name="Account (Optional)"
		else
			data1name="Data 1"
		end if
		if htmlemails<>TRUE then htmleditor=""
		if htmleditor="ckeditor" then %>
<script src="ckeditor/ckeditor.js"></script>
<%		end if %>
<script>
/* <![CDATA[ */
function validateform(){
	if(document.getElementById("data1")) document.getElementById("data1").disabled=false;
	if(document.getElementById("data2")) document.getElementById("data2").disabled=false;
	if(document.getElementById("data3")) document.getElementById("data3").disabled=false;
	return true;
}
function disablepaypalapi(disbd){
	if(disbd){
		document.getElementById("data1span").style.color='#A0A0A0';
		document.getElementById("data2span").style.color='#A0A0A0';
		document.getElementById("data3span").style.color='#A0A0A0';
		document.getElementById("data1").disabled=true;
		document.getElementById("data2").disabled=true;
		document.getElementById("data3").disabled=true;
		document.getElementById("ppexpressabemail").disabled=false;
	}else{
		document.getElementById("data1span").style.color='#000000';
		document.getElementById("data2span").style.color='#000000';
		document.getElementById("data3span").style.color='#000000';
		document.getElementById("data1").disabled=false;
		document.getElementById("data2").disabled=false;
		document.getElementById("data3").disabled=false;
		document.getElementById("ppexpressabemail").disabled=true;
	}
}
function advertisingopts(){
	var advtext='<html><head><title>Advertising Options</title><link rel="stylesheet" type="text/css" href="adminstyle.css" /></head><body>' +
		'<div id="header1"><p align="center" style="font-weight:bold;margin:30px;font-size:20px">PayPal Credit Advertising Options</p></div>' +
		'<div id="main"><p>You can advertise PayPal Credit to your customers in order to encourage them to use this financing option. In order to do so you will need to get your PayPal publisher id from the PayPal site. This widget will help you do so.</p>' +
		'<p>Once you have your publisher id, please check <a href="https://www.ecommercetemplates.com/help/ecommplus/paypal-express-checkout.asp" style="font-weight:bold" target="_blank">this page</a> to view details about how to setup PayPal Credit advertising banners on your site.</p>' +
		'<p>&nbsp;</p><p align="center"><input type="button" value="Please Click Here To Get Your Publisher ID" onclick="document.location=\'https://financing.paypal.com/ppfinportal/cart/index?dcp=54d773b600a9fe642a805cb9f8c514d3634acbc7\'"></p>' +
		'<p>&nbsp;</p>' +
		'<p align="center"><input type="button" value="<%=replace(yyClsWin,"'","\'")%>" onclick="window.close()"></p></div></body></'+'html>';
	newwin=window.open('','AdvOpts','menubar=no,scrollbars=no,width=400,height=550,resizable=yes,status=no,toolbar=no,location=no');
	newwin.document.open();
	newwin.document.write(advtext);
	newwin.document.close();
	return false;
}
/* ]]> */
</script>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">

		  <form name="mainform" method="post" action="adminpayprov.asp" onsubmit="return validateform()">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="domodify" />
			<input type="hidden" name="id" value="<%=payProvID%>" />
			<input type="hidden" name="from" value="<%=getget("from")%>" />
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr> 
                <td width="100%" colspan="2" align="center"><strong><%=yyPPAdm%></strong><br />&nbsp;</td>
			  </tr>
<%		if getget("from")="wizard2" then %>
			  <tr> 
                <td width="100%" colspan="2" align="center">
				  <table width="80%" border="0" cellspacing="2" cellpadding="2" bgcolor="#BFC9E0">
					<tr>
					  <td align="left" valign="top" bgcolor="#FFFFFF">
						<div style="font-size:12px;margin:5px;">
						You can now setup your PayPal account details. If you don't yet have a PayPal account and wish to create one please
						 <a href="<%=signuppage%>" target="_blank"><strong><%=yyClkHer%></strong></a><br />&nbsp;
						</div>
					  </td>
					</tr>
				  </table><br />&nbsp;
				</td>
			  </tr>
<%		elseif signuppage<>"" then %>
			  <tr> 
                <td width="100%" colspan="2" align="center"><%=yySignUp%> <a href="<%=signuppage%>" target="_blank"><strong><%=yyClkHer%></strong></a><br />&nbsp;</td>
			  </tr>
<%		end if %>
			  <tr>
				<td align="right" width="50%"><strong><%=yyPPName%> : </strong></td>
				<td align="left"><%=payProvName%></td>
			  </tr>
			  <tr>
				<td align="right"><%=yyShwAs%> : </td>
				<td align="left"><input type="text" name="showas" value="<%=htmlspecials(payProvShow)%>" size="25" /></td>
			  </tr>
<%	for index=2 to adminlanguages+1
		if index=2 then showthis=payProvShow2
		if index=3 then showthis=payProvShow3
		if (adminlangsettings AND 128)=128 then %>
			  <tr>
				<td align="right"><%=yyShwAs & " " & index%> : </td>
				<td align="left"><input type="text" name="showas<%=index%>" value="<%=htmlspecials(showthis)%>" size="25" /></td>
			  </tr>
<%		end if
	next
	if payProvID=7 then ' PayFlow Pro
%>			  <tr>
				<td colspan="2" align="center"><%=yyPPExp%></td>
			  </tr>
<%	end if %>
			  <tr>
				<td align="right"><%=yyEnable%> : </td>
				<td align="left"><input type="checkbox" name="isenabled" value="1" <%if payProvEnabled=1 then print "checked=""checked"""%> /></td>
			  </tr>
<%	if demomodeavailable then %>
			  <tr>
				<td align="right"><%=yyDemoMo%> : </td>
				<td align="left"><input type="checkbox" name="demomode" value="1" <%if payProvDemo=1 then print "checked=""checked"""%> /></td>
			  </tr>
<%	end if
	disableapi=FALSE
	if payProvID=19 then
		if instr(payProvData1,"@AB@")>0 OR (payProvData2="" AND payProvData3="") then disableapi=TRUE
		if disableapi then
			paypalemail=replace(payProvData1&"","@AB@","")
			if disableapi AND paypalemail="" then paypalemail=emailAddr
			payProvData1=""
		end if %>
			  <tr>
				<td align="right">Enable PayPal Credit:</td>
				<td align="left"><input type="checkbox" name="billmelater" value="1" <%if wantbillmelater then print "checked=""checked"""%> /> (<a href="#" onclick="return advertisingopts()">Please click here to view advertising options</a>).</td>
			  </tr>
			  <tr>
				<td colspan="2" align="center">
			<table>
			  <tr>
				<td align="right"><input type="radio" name="ppexpressab" value="AB" onclick="disablepaypalapi(true)" <% if disableapi=TRUE then print "checked=""checked"" "%>/></td>
				<td align="left"><%=yyPPEmal%> : <input type="text" name="ppexpressabemail" id="ppexpressabemail" value="<% if disableapi=TRUE then print paypalemail%>" <% if disableapi=FALSE then print "disabled=""disabled"" " %>size="35" /></td>
			  </tr>
			  <tr>
				<td align="right"><input type="radio" name="ppexpressab" value="" onclick="disablepaypalapi(false)" <% if disableapi=FALSE then print "checked=""checked"" "%>/></td>
				<td align="left"><%=yyPPAPIC%><br />
					(<%=yyCanLat%>)</td>
			  </tr>
			</table>
				</td>
			  </tr>
<%	end if
	if payProvID=7 OR payProvID=8 OR payProvID=22 then ' PayFlow Pro OR PayFlow Link OR PayPal Advanced
		if payProvID=8 AND instr(payProvData1,"&")=0 then
			vs1=""
			vs2=payProvData1
			vs3=payProvData2
			vs4=""
		else
			vsdetails=split(payProvData1,"&")
			if UBOUND(vsdetails)>0 then
				vs1=vsdetails(0)
				vs2=vsdetails(1)
				vs3=vsdetails(2)
				vs4=vsdetails(3)
			end if
		end if
%>
			  <tr>
				<td colspan="2" align="center">Please Note: The login information below is the same login you use for PayPal Manager.</td>
			  </tr>
			  <tr>
				<td align="right"><%=yyPartner%> : </td>
				<td align="left"><input type="text" name="data3" value="<%=vs3%>" size="25" /> <input type="button" value="?" title="Your Partner Name is &quot;PayPal&quot;" /></td>
			  </tr>
			  <tr>
				<td align="right"><%=yyVendor%> : </td>
				<td align="left"><input type="text" name="data2" value="<%=vs2%>" size="25" /> <input type="button" value="?" title="This is the login name you created when<%=vbCrLf%>signing up for PayPal <% if payProvID=7 then print "PayFlow Pro" else print "Payments Advanced"%>" /></td>
			  </tr>
			  <tr>
				<td align="right"><%=yyUserID%> : </td>
				<td align="left"><input type="text" name="data1" value="<%=vs1%>" size="25" /> <input type="button" value="?" title="Instead of entering a Merchant Login, you can<%=vbCrLf%>enter a User Login. A User Login is what PayPal<%=vbCrLf%>recommends because it provides enhanced<%=vbCrLf%>security and prevents service interruption if you<%=vbCrLf%>change your Merchant Login password. You<%=vbCrLf%>can set up a User Login and profile in PayPal<%=vbCrLf%>Manager." /></td>
			  </tr>
			  <tr>
				<td align="right"><%=yyPass%> : </td>
				<td align="left"><input type="text" name="data4" value="<%=vs4%>" size="25" /> <input type="button" value="?" title="This is the password you created when signing<%=vbCrLf%>up for PayPal <% if payProvID=7 then print "PayFlow Pro" else print "Payments Advanced"%> or the<%=vbCrLf%>password you created for API calls." /></td>
			  </tr>
<%	elseif payProvID=10 then ' Capture Card
%>			  <tr>
				<td align="center" colspan="2"><hr width="50%" /><%=yyAccCar%><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td align="right">Visa : </td>
				<td align="left"><input type="checkbox" name="cardtype1" value="X" <% if Mid(payProvData1,1,1)="X" then print "checked=""checked""" %> /></td>
			  </tr>
			  <tr>
				<td align="right">Mastercard : </td>
				<td align="left"><input type="checkbox" name="cardtype2" value="X" <% if Mid(payProvData1,2,1)="X" then print "checked=""checked""" %> /></td>
			  </tr>
			  <tr>
				<td align="right">American Express : </td>
				<td align="left"><input type="checkbox" name="cardtype3" value="X" <% if Mid(payProvData1,3,1)="X" then print "checked=""checked""" %> /></td>
			  </tr>
			  <tr>
				<td align="right">Diners Club : </td>
				<td align="left"><input type="checkbox" name="cardtype4" value="X" <% if Mid(payProvData1,4,1)="X" then print "checked=""checked""" %> /></td>
			  </tr>
			  <tr>
				<td align="right">Discover : </td>
				<td align="left"><input type="checkbox" name="cardtype5" value="X" <% if Mid(payProvData1,5,1)="X" then print "checked=""checked""" %> /></td>
			  </tr>
			  <tr>
				<td align="right">En Route : </td>
				<td align="left"><input type="checkbox" name="cardtype6" value="X" <% if Mid(payProvData1,6,1)="X" then print "checked=""checked""" %> /></td>
			  </tr>
			  <tr>
				<td align="right">JCB : </td>
				<td align="left"><input type="checkbox" name="cardtype7" value="X" <% if Mid(payProvData1,7,1)="X" then print "checked=""checked""" %> /></td>
			  </tr>
			  <tr>
				<td align="right">Maestro/Switch/Solo : </td>
				<td align="left"><input type="checkbox" name="cardtype8" value="X" <% if Mid(payProvData1,8,1)="X" then print "checked=""checked""" %> /></td>
			  </tr>
			  <tr>
				<td align="right">Bankcard (AUS / NZ) : </td>
				<td align="left"><input type="checkbox" name="cardtype9" value="X" <% if Mid(payProvData1,9,1)="X" then print "checked=""checked""" %> /></td>
			  </tr>
			  <tr>
				<td align="right">Laser (IRL) : </td>
				<td align="left"><input type="checkbox" name="cardtype10" value="X" <% if Mid(payProvData1,10,1)="X" then print "checked=""checked""" %> /></td>
			  </tr>
			  <tr>
				<td align="center" colspan="2"><hr width="50%" /><%=yyNewCer%><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td colspan="2" align="center"><textarea name="data2" rows="7" cols="82"></textarea></td>
			  </tr>
<%	elseif data1name<>"" then %>
			  <tr>
				<td align="right"><span id="data1span"<% if disableapi=TRUE then print " style=""color:#A0A0A0;""" %>><%=data1name%> : </span></td>
				<td align="left"><input type="text" name="data1" id="data1" value="<%=payProvData1%>" <% if disableapi=TRUE then print "disabled=""disabled"" " %>size="35" /></td>
			  </tr>
<%	end if
	if payProvID=5 then
		data2arr=split(trim(payProvData2&""),"&",2)
		if UBOUND(data2arr)>=0 then data2md5=data2arr(0)
		if UBOUND(data2arr)>0 then data2cbp=data2arr(1) else data2cbp=""
%>
			  <tr>
				<td align="right">MD5 Secret (Optional) : </td>
				<td align="left"><input type="text" name="data2" value="<%=data2md5%>" size="25" /></td>
			  </tr>
			  <tr>
				<td align="right">Payment Response password (Optional) : </td>
				<td align="left"><input type="text" name="data3" value="<%=data2cbp%>" size="25" /></td>
			  </tr>
<%	elseif payProvID=9 then ' PayPoint.net
		data2arr=split(trim(payProvData2&""),"&",2)
		if UBOUND(data2arr)>=0 then data2md5=data2arr(0)
		if UBOUND(data2arr)>0 then data2template=urldecode(data2arr(1)) else data2template=""
%>
			  <tr>
				<td align="right"><%=yyMD5H%> : </td>
				<td align="left"><input type="text" name="data2" value="<%=data2md5%>" size="25" /></td>
			  </tr>
			  <tr>
				<td align="right">Payment Template (Optional) : </td>
				<td align="left"><input type="text" name="data2supp" value="<%=data2template%>" size="25" /></td>
			  </tr>
			  <tr>
				<td align="right">Callback URL on SSL Connection : </td>
				<td align="left"><select name="data3" size="1"><option value=""><%=yyNo%></option><option value="1" <% if payProvData3="1" then print "selected=""selected"""%>><%=yyYes%></option></select></td>
			  </tr>
<%	elseif payProvID=16 then %>
			  <tr>
				<td align="right">Shared Secret (Connect 2.0 Only) : </td>
				<td align="left"><input type="text" name="data3" value="<%=payProvData3%>" size="25" /></td>
			  </tr>
			  <tr>
				<td align="right"><%=data2name%> : </td>
				<td align="left"><select name="data2" size="1"><option value="0"><%=yyLPSit%></option><option value="1" <% if payProvData2="1" then print "selected=""selected"""%>><%=yyYesOS%></option></select></td>
			  </tr>
<%	elseif payProvID=18 OR payProvID=19 then %>
			  <tr>
				<td align="right"><% if payProvID=18 then print whv("billmelater",IIfVr(wantbillmelater,"1","0")) %><span id="data2span"<% if disableapi=TRUE then print " style=""color:#A0A0A0;""" %>><%=data2name%> : </span></td>
				<td align="left"><input type="text" name="data2" id="data2" value="<%=data2pwd%>" <% if disableapi=TRUE then print "disabled=""disabled"" " %>size="25" /></td>
			  </tr>
			  <tr>
				<td align="right"><span id="data3span"<% if disableapi=TRUE then print " style=""color:#A0A0A0;""" %>><%=yySigHas&".<br /><span style=""font-size:10px"">("&yyOn3tok&")</span>"%> : </span></td>
				<td align="left"><input type="text" name="data3" id="data3" value="<%=data2hash%>" <% if disableapi=TRUE then print "disabled=""disabled"" " %>size="45" /></td>
			  </tr>
<%	elseif payProvID=21 then
			data2arr=split(trim(payProvData2&""),"&",2)
			if UBOUND(data2arr)>=0 then data2=data2arr(0)
			if UBOUND(data2arr)>0 then sellerid=data2arr(1)
%>			  <tr>
				<td align="right">Seller ID : </td>
				<td align="left"><input type="text" name="data2b" value="<%=sellerid%>" size="25" /></td>
			  </tr>
			  <tr>
				<td align="right"><%=data2name%> : </td>
				<td align="left"><input type="text" name="data2" value="<%=data2%>" size="25" /></td>
			  </tr>
<%	elseif data2name<>"" then %>
			  <tr>
				<td align="right"><%=data2name%> : </td>
				<td align="left"><input type="text" name="data2" value="<%=payProvData2%>" size="35" /></td>
			  </tr>
<%	end if
	if data3name<>"" then %>
			  <tr>
				<td align="right"><%=data3name%> : </td>
				<td align="left"><input type="text" name="data3" value="<%=payProvData3%>" size="35" /></td>
			  </tr>
<%	end if
	if data4name<>"" then %>
			  <tr>
				<td align="right"><%=data4name%> : </td>
				<td align="left"><input type="text" name="data4" value="<%=payProvData4%>" size="35" /></td>
			  </tr>
<%	end if
	if data5name<>"" then %>
			  <tr>
				<td align="right"><%=data5name%> : </td>
				<td align="left"><input type="text" name="data5" value="<%=payProvData5%>" size="35" /></td>
			  </tr>
<%	end if
	if payProvID=3 then %>
			  <tr>
				<td align="right">Integration Method : </td>
				<td align="left"><select name="payProvFlag1" size="1">
						<option value="0">Use SIM Integration Method (Deprecated)</option>
						<option value="1"<% if payProvFlag1 then print " selected=""selected"""%>>Use Accept Hosted Integration Method (Recommended)</option>
					</select></td>
			  </tr>
<%	end if
	if payProvID=23 then %>
			  <tr>
				<td align="right">Integration Method : </td>
				<td align="left"><select name="payProvFlag1" size="1">
						<option value="0">Old Stripe Checkout</option>
						<option value="1"<% if payProvFlag1 then print " selected=""selected"""%>>New SCA-Ready Stripe Checkout</option>
					</select></td>
			  </tr>
			  <tr>
				<td align="right">Payment Methods : </td>
				<td align="left">
					<div>
						<div><input type="checkbox" name="databit1" value="1" <% if (payProvBits AND 1)=1 OR payProvBits=0 then print "checked=""checked"" "%>/> Accept Credit Cards</div>
						<div><input type="checkbox" name="databit2" value="1" <% if (payProvBits AND 2)=2 then print "checked=""checked"" "%>/> Accept iDEAL</div>
						<div><input type="checkbox" name="databit3" value="1" <% if (payProvBits AND 4)=4 then print "checked=""checked"" "%>/> Accept Bancontact</div>
						<div><input type="checkbox" name="databit4" value="1" <% if (payProvBits AND 8)=8 then print "checked=""checked"" "%>/> Accept Giropay</div>
						<div><input type="checkbox" name="databit5" value="1" <% if (payProvBits AND 16)=16 then print "checked=""checked"" "%>/> Accept Przelewy24</div>
						<div><input type="checkbox" name="databit6" value="1" <% if (payProvBits AND 32)=32 then print "checked=""checked"" "%>/> Accept EPS</div>
						<div><input type="checkbox" name="databit7" value="1" <% if (payProvBits AND 64)=64 then print "checked=""checked"" "%>/> Accept FPX</div>
						<div><input type="checkbox" name="databit8" value="1" <% if (payProvBits AND 128)=128 then print "checked=""checked"" "%>/> Accept BACS Debit</div>
						<div><input type="checkbox" name="databit9" value="1" <% if (payProvBits AND 256)=256 then print "checked=""checked"" "%>/> Accept Klarna Payments</div>
					</div>
				</td>
			  </tr>
<%	else
		for index=1 to 10
			if bitname(index)<>"" then %>
			  <tr>
				<td align="right"><%=bitname(index)%> : </td>
				<td align="left"><select name="databit<%=index%>" size="1">
						<option value=""><%=yyNo%></option>
						<option value="1"<% if (payProvBits AND (2 ^ (index-1)))=2 ^ (index-1) then print " selected=""selected"""%>><%=yyYes%></option>
					</select></td>
			  </tr>
<%			end if
		next
	end if
	if hasauthtype OR payProvID=1 OR payProvID=3 OR payProvID=5 OR payProvID=7 OR payProvID=8 OR payProvID=9 OR payProvID=11 OR payProvID=12 OR payProvID=13 OR payProvID=14 OR payProvID=16 OR payProvID=18 OR payProvID=19 OR payProvID=21 OR payProvID=22 OR payProvID=23 OR payProvID=24 OR payProvID=27 OR payProvID=28 then ' Pay Providers we can set authorization type
		yyAuthOr=""
		if payProvID=1 OR payProvID=7 OR payProvID=8 OR payProvID=18 OR payProvID=19 OR payProvID=22 OR payProvID=27 then
			yyAuthCp="Sale"
			yyAuthOn="Authorization"
		end if
		if payProvID=27 then yyAuthOr="Order"
%>			  <tr>
				<td align="right"><%=yyTrnTyp%> : </td>
				<td align="left"><select name="transtype" size="1">
					<option value="0"><%=yyAuthCp%></option>
					<option value="1" <% if payProvMethod="1" then print "selected=""selected""" %>><%=yyAuthOn%></option>
<%		if yyAuthOr<>"" then %>
					<option value="2" <% if payProvMethod="2" then print "selected=""selected""" %>><%=yyAuthOr%></option>
<%		end if %>
					</select></td>
			  </tr>
<%	end if
	if payProvID=27 then
		buttonstyle=split(payProvData3,"|")
		if UBOUND(buttonstyle)>=0 then buttonstyle0=buttonstyle(0) else buttonstyle0=""
		if UBOUND(buttonstyle)>=1 then buttonstyle1=buttonstyle(1) else buttonstyle1=""
		if UBOUND(buttonstyle)>=2 then buttonstyle2=buttonstyle(2) else buttonstyle2=""
		if UBOUND(buttonstyle)>=3 then buttonstyle3=buttonstyle(3) else buttonstyle3=""
		if UBOUND(buttonstyle)>=4 then buttonstyle4=buttonstyle(4) else buttonstyle4=""
		if UBOUND(buttonstyle)>=5 then buttonstyle5=buttonstyle(5) else buttonstyle5=""
		if UBOUND(buttonstyle)>=6 then buttonstyle6=buttonstyle(6) else buttonstyle6=""
%>
			  <tr>
				<td align="right">Button Style : </td>
				<td align="left">
					<div>
						<select size="1" name="buttonshape">
							<option value="">Button Shape...</option>
							<option value=""<% if buttonstyle0="" then print " selected=""selected"""%>>Pill (Recommended)</option>
							<option value="rect"<% if buttonstyle0="rect" then print " selected=""selected"""%>>Rectangle</option>
						</select>
					</div>
					<div style="padding:3px 0 0 0">
						<select size="1" name="buttonsize">
							<option value="">Button Size...</option>
							<option value=""<% if buttonstyle1="" then print " selected=""selected"""%>>Small (Recommended)</option>
							<option value="medium"<% if buttonstyle1="medium" then print " selected=""selected"""%>>Medium</option>
							<option value="large"<% if buttonstyle1="large" then print " selected=""selected"""%>>Large</option>
							<option value="responsive"<% if buttonstyle1="responsive" then print " selected=""selected"""%>>Responsive</option>
						</select>
					</div>
					<div style="padding:3px 0 0 0">
						<select size="1" name="buttoncolor">
							<option value="">Button Color...</option>
							<option value=""<% if buttonstyle2="" then print " selected=""selected"""%>>Gold (Recommended)</option>
							<option value="blue"<% if buttonstyle2="blue" then print " selected=""selected"""%>>Blue (First Alternative)</option>
							<option value="silver"<% if buttonstyle2="silver" then print " selected=""selected"""%>>Silver (Second Alternative)</option>
							<option value="white"<% if buttonstyle2="white" then print " selected=""selected"""%>>White (Second Alternative)</option>
							<option value="black"<% if buttonstyle2="black" then print " selected=""selected"""%>>Black (Third Alternative)</option>
						</select>
					</div>
					<div style="padding:3px 0 0 0">
						<select size="1" name="buttonlayout">
							<option value="">Button Layout...</option>
							<option value="horizontal"<% if buttonstyle3="horizontal" then print " selected=""selected"""%>>Horizontal</option>
							<option value="vertical"<% if buttonstyle3="vertical" then print " selected=""selected"""%>>Vertical</option>
						</select>
					</div>
				</td>
			  </tr>
			  <tr>
				<td align="right">Funding Sources : </td>
				<td align="left">
					<div>
						<select size="1" name="paypalcredit">
							<option value=""<% if buttonstyle4="" then print " selected=""selected"""%>>Display PayPal Credit (If available)</option>
							<option value="hide"<% if buttonstyle4="hide" then print " selected=""selected"""%>>Hide PayPal Credit</option>
						</select>
					</div>
					<div style="padding:3px 0 0 0">
						<select size="1" name="paypalcards">
							<option value=""<% if buttonstyle5="" then print " selected=""selected"""%>>Display Card Sources (If available)</option>
							<option value="hide"<% if buttonstyle5="hide" then print " selected=""selected"""%>>Hide Card Sources</option>
						</select>
					</div>
					<div style="padding:3px 0 0 0">
						<select size="1" name="paypalelv">
							<option value=""<% if buttonstyle6="" then print " selected=""selected"""%>>Display Elektronisches Lastschriftverfahren (If available)</option>
							<option value="hide"<% if buttonstyle6="hide" then print " selected=""selected"""%>>Hide Elektronisches Lastschriftverfahren</option>
						</select>
					</div>
					<div style="padding:3px 0 0 0">
						<select size="1" name="paypalvenmo">
							<option value="display">Display VENMO (If available)</option>
							<option value=""<% if buttonstyle7="" then print " selected=""selected"""%>>Hide VENMO</option>
						</select>
					</div>
				</td>
			  </tr>
<%	end if %>
			  <tr>
				<td align="right"><%=yyLiLev%> : </td>
				<td align="left"><select name="payProvLevel" size="1">
				<option value="0"><%=yyNoRes%></option>
				<%	for index=1 to maxloginlevels
						print "<option value="""&index&""""
						if payProvLevel=index then print " selected=""selected"""
						print ">" & yyLiLev & " " & index & "</option>"
					next%></select></td>
			  </tr>
			  <tr>
				<td align="right"><%=yyHanChg%> : </td>
				<td align="left"><input type="text" name="pphandlingcharge" size="5" value="<%=pphandlingcharge%>" /></td>
			  </tr>
			  <tr>
				<td align="right"><%=yyHanChg & " (" & yyPercen & ")"%> : </td>
				<td align="left"><input type="text" name="pphandlingpercent" size="5" value="<%=pphandlingpercent%>" /></td>
			  </tr>
<%		for index=0 to adminlanguages
			languageid=index+1
			if index=0 OR (adminlangsettings AND 4096)=4096 then
				sSQL="SELECT "&getlangid("pProvHeaders",4096)&" FROM payprovider WHERE payProvID=" & request("id")
				rs.open sSQL,cnn,0,1
					theheader=trim(rs(getlangid("pProvHeaders",4096))&"")
				rs.close
				theheader=replace(theheader, "%nl%", "<br />")
				theheader=replace(theheader, "<br>", "<br />")
				if NOT (htmlemails AND (htmleditor="froala" OR htmleditor="ckeditor")) then
					theheader=replace(theheader, "<br />", vbCrLf)
				else
					theheader=replace(theheader,"<","&lt;")
				end if %>
			  <tr>
				<td align="right"><%=yyEmlHdr & " " & IIfVr(index>0, index+1, "")%> :</td><td align="left"><input type="button" value="&nbsp;Edit&nbsp;" onclick="switchheader('spanheaders<%=(index+1)%>')" /></td>
			  </tr>
			  <tr>
				<td align="<%=IIfVr(htmleditor<>"","left","center")%>" colspan="2"><span id="spanheaders<%=(index+1)%>" style="display:none"><textarea id="pprovheaders<%=(index+1)%>" name="pprovheaders<%=(index+1)%>" cols="70" rows="6"><%=theheader%></textarea></span></td>
			  </tr>
<%				sSQL="SELECT "&getlangid("pProvDropShipHeaders",4096)&" FROM payprovider WHERE payProvID=" & request("id")
				rs.open sSQL,cnn,0,1
					theheader=trim(rs(getlangid("pProvDropShipHeaders",4096))&"")
				rs.close
				theheader=replace(theheader, "%nl%", "<br />")
				theheader=replace(theheader, "<br>", "<br />")
				if NOT (htmlemails AND (htmleditor="froala" OR htmleditor="ckeditor")) then
					theheader=replace(theheader, "<br />", vbCrLf)
				else
					theheader=replace(theheader,"<","&lt;")
				end if %>
			  <tr>
				<td align="right"><%=yyDrSppr & " " & yyEmlHdr & " " & IIfVr(index>0, index+1, "")%> :</td><td align="left"><input type="button" value="&nbsp;Edit&nbsp;" onclick="switchheader('spandropshipheaders<%=(index+1)%>')" /></td>
			  </tr>
			  <tr>
				<td align="<%=IIfVr(htmleditor<>"","left","center")%>" colspan="2"><span id="spandropshipheaders<%=(index+1)%>" style="display:none"><textarea id="pprovdropshipheaders<%=(index+1)%>" name="pprovdropshipheaders<%=(index+1)%>" cols="70" rows="6"><%=theheader%></textarea></span></td>
			  </tr>
<%			end if
		next %>
			  <tr>
				<td colspan="2">&nbsp;</td>
			  </tr>
<%	if getget("from")="wizard" AND payProvID<>1 AND payProvID<>18 AND payProvID<>19 then %>
			  <tr>
				<td colspan="2" align="center">
				  <table width="80%" border="0" cellspacing="2" cellpadding="2" bgcolor="#BFC9E0">
					<tr>
					  <td align="left" valign="top" bgcolor="#FFFFFF">
						<img src="adminimages/paypalexample.gif" border="0" style="float:right;margin:5px;" />
					    <div style="font-size:14px;font-weight:bold;margin:5px;"><input type="checkbox" name="offerpaypal" value="ON" checked="checked" />&nbsp;Offer the option to pay with PayPal</div>
						<div style="font-size:12px;color:#3263B3;margin:5px;">According to Jupiter Research, 23% of online shoppers consider PayPal one of their favorite 
						ways to pay online.*<br />
						Accepting PayPal in addition to credit cards is proven to increase your sales.**</div>
						<div style="font-size:12px;margin:5px;">*<span style="font-style:italic;"> Payment Preferences Online</span>, Jupiter Research, September 2000<br />
						** Applies to online businesses doing up to $10 million/year in online sales. Based on a Q4 2007 survey of PayPal shoppers conducted by Northstar Research, and PayPal internal data on Express Checkout transactions.</div>
					  </td>
					</tr>
				  </table><br />&nbsp;
				</td>
			  </tr>
<%	end if %>
			  <tr>
				<td align="center" colspan="2"><input type="submit" value="<%=yySubmit%>" /> <input type="reset" value="<%=yyReset%>" /></td>
			  </tr>
			  <tr>
				<td colspan="2">&nbsp;</td>
			  </tr>
			</table>
		  </form>
<%	if htmleditor="ckeditor" then
		pathtovsadmin=request.servervariables("URL")
		slashpos=instrrev(pathtovsadmin, "/")
		if slashpos>0 then pathtovsadmin=left(pathtovsadmin, slashpos-1)
		print "<script>function loadeditors(){"
		streditor="var pprovheaders=CKEDITOR.replace('pprovheaders',{extraPlugins : 'stylesheetparser,autogrow',autoGrow_maxHeight : 800,removePlugins : 'resize', toolbarStartupExpanded : false, toolbar : 'Basic', filebrowserBrowseUrl :'ckeditor/filemanager/browser/default/browser.html?Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserImageBrowseUrl : 'ckeditor/filemanager/browser/default/browser.html?Type=Image&Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserFlashBrowseUrl :'ckeditor/filemanager/browser/default/browser.html?Type=Flash&Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=File',filebrowserImageUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=Image',filebrowserFlashUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=Flash'});" & vbCrLf
		streditor=streditor & "pprovheaders.on('instanceReady',function(event){var myToolbar='Basic';event.editor.on( 'beforeMaximize', function(){if(event.editor.getCommand('maximize').state == CKEDITOR.TRISTATE_ON && myToolbar != 'Basic'){pprovheaders.setToolbar('Basic');myToolbar='Basic';pprovheaders.execCommand('toolbarCollapse');}else if(event.editor.getCommand('maximize').state == CKEDITOR.TRISTATE_OFF && myToolbar != 'Full'){pprovheaders.setToolbar('Full');myToolbar='Full';pprovheaders.execCommand('toolbarCollapse');}});event.editor.on('contentDom', function(e){event.editor.document.on('blur', function(){if(!pprovheaders.isToolbarCollapsed){pprovheaders.execCommand('toolbarCollapse');pprovheaders.isToolbarCollapsed=true;}});event.editor.document.on('focus',function(){if(pprovheaders.isToolbarCollapsed){pprovheaders.execCommand('toolbarCollapse');pprovheaders.isToolbarCollapsed=false;}});});pprovheaders.fire('contentDom');pprovheaders.isToolbarCollapsed=true;});"
		for index=1 to adminlanguages+1
			if index=1 OR (adminlangsettings AND 4096)=4096 then
				print replace(streditor, "pprovheaders", "pprovheaders" & index)
				print replace(streditor, "pprovheaders", "pprovdropshipheaders" & index)
			end if
		next
		print "}window.onload=function(){loadeditors();}</script>"
	elseif htmleditor="froala" then
		for index=1 to adminlanguages+1
			if index=1 OR (adminlangsettings AND 4096)=4096 then
				call displayfroalaeditor("pprovheaders" & index,yyEmlHdr,"",TRUE,FALSE,1,TRUE)
				call displayfroalaeditor("pprovdropshipheaders" & index,yyDrSppr & " " & yyEmlHdr,"",TRUE,FALSE,1,TRUE)
			end if
		next
	end if %>
		  </td>
		</tr>
      </table>
<%	elseif getpost("act")="changepos" then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%" align="center">
			<p>&nbsp;</p>
			<p>&nbsp;</p>
			<p>&nbsp;</p>
			<p><%=yyUpdat%> . . . . . . . </p>
			<p>&nbsp;</p>
			<p><%=yyNoFor%> <a href="adminpayprov.asp"><%=yyClkHer%></a>.</p>
			<p>&nbsp;</p>
			<p>&nbsp;</p>
		  </td>
		</tr>
      </table>
<%	elseif getget("act")="cccards" then %>
	<table border="0" cellspacing="0" cellpadding="0" width="100%">
	  <tr>
		<td align="center">
		  <table width="80%" height="100%" border="0" cellspacing="0" cellpadding="2">
			<tr>
			  <td align="left"><p style="font-size:18px;font-weight:bold;">Choose a solution to accept credit card payments</p>
					<p>&nbsp;</p>
			  </td>
			</tr>
		  </table>
		  <table width="80%" border="0" cellspacing="2" cellpadding="2" bgcolor="#BFC9E0">
			<tr>
			  <td width="50%" align="left" valign="top" bgcolor="#FFFFFF">
			  <p style="font-size:16px;font-weight:bold;">&nbsp;All-in-one Solution</p>
			  <div onclick="selectopt('allin1')" style="border:1px;font-size:12px;font-weight:bold;background-color:#E6E9F5;padding:4px;min-height:50px;border-style:solid;border-width:1px;">
			  <input type="radio" name="solntype" value="ALL1" id="allin1" /> 
			  I want an all-in-one payment solution that includes a payment gateway and an internet merchant account.</div>
			  &nbsp;
			  <div id="allin1div" style="border:1px;padding:4px;background-color:#E6E6E6;border-style:solid;border-width:1px;">
			  <div id="allin1div2" style="padding:8px;font-size:12px;font-weight:bold;color:#A0A0A0;">Choose your all-in-one solution</div>
			  <p>
			    <ul>
				  <li><a href="adminpayprov.asp?act=modify&from=wizard&id=18" class="allin1">PayPal Payments Pro</a><br />&nbsp;</li>
				  <li><a href="adminpayprov.asp?act=modify&from=wizard&id=1" class="allin1">PayPal Payments Standard</a><br />&nbsp;</li>
				  <li><a href="adminpayprov.asp?act=modify&from=wizard&id=2" class="allin1">2Checkout</a><br />&nbsp;</li>
				  <li><a href="adminpayprov.asp?act=modify&from=wizard&id=21" class="allin1">Amazon Simple Pay</a><br />&nbsp;</li>
				  <li><a href="adminpayprov.asp?act=modify&from=wizard&id=20" class="allin1">Google Checkout</a><br />&nbsp;</li>
				  <li><a href="adminpayprov.asp?act=modify&from=wizard&id=15" class="allin1">Netbanx</a><br />&nbsp;</li>
				  <li><a href="adminpayprov.asp?act=modify&from=wizard&id=6" class="allin1">Nochex</a><br />&nbsp;</li>
				  <li><a href="adminpayprov.asp?act=modify&from=wizard&id=5" class="allin1">RBS WorldPay</a><br />&nbsp;</li>
				</ul>
			  </p>
			  </div>
			  </td>
			  
			  <td align="left" valign="top" bgcolor="#FFFFFF">
			  <p style="font-size:16px;font-weight:bold;">&nbsp;Solution for existing merchant account</p>
			  <div onclick="selectopt('exis')" style="border:1px;font-size:12px;font-weight:bold;background-color:#E6E9F5;padding:4px;min-height:50px;border-style:solid;border-width:1px;">
			  <input type="radio" name="solntype" value="EXIS" id="exis" /> 
			  I prefer a payment gateway that works with my existing merchant account.</div>
			  &nbsp;
			  <div id="exisdiv" style="border:1px;padding:4px;background-color:#E6E6E6;border-style:solid;border-width:1px;">
			  <div id="exisdiv2" style="padding:8px;font-size:12px;font-weight:bold;color:#A0A0A0;">Choose your Gateway</div>
			  <p>
			    <ul>
				  <li><a href="adminpayprov.asp?act=modify&from=wizard&id=7" class="exis">PayPal Payflow Pro</a><br />&nbsp;</li>
				  <li><a href="adminpayprov.asp?act=modify&from=wizard&id=8" class="exis">PayPal Payflow Link</a><br />&nbsp;</li>
				  <li><a href="adminpayprov.asp?act=modify&from=wizard&id=13" class="exis">Authorize.net (AIM)</a><br />&nbsp;</li>
				  <li><a href="adminpayprov.asp?act=modify&from=wizard&id=3" class="exis">Authorize.net (SIM)</a><br />&nbsp;</li>
				  <li><a href="adminpayprov.asp?act=modify&from=wizard&id=16" class="exis">Linkpoint</a><br />&nbsp;</li>
				  <li><a href="adminpayprov.asp?act=modify&from=wizard&id=9" class="exis">PayPoint</a><br />&nbsp;</li>
				  <li><a href="adminpayprov.asp?act=modify&from=wizard&id=11" class="exis">PSiGate</a><br />&nbsp;</li>
				  <li><a href="adminpayprov.asp?act=modify&from=wizard&id=12" class="exis">PSiGate (SSL)</a><br />&nbsp;</li>
				</ul>
			  </p>
			  </div>
			  </td>
			</tr>
		  </table>
		  <table width="80%" border="0" cellspacing="0" cellpadding="2">
			<tr>
			  <td align="left">
<%			if false then %>
				<p style="font-size:11px;font-weight:bold;">Don't see what you are looking for?</p>
				<p style="font-size:11px;font-weight:bold;"><a href="adminpayprov.asp?act=list">See full list of payment processors</a></p>
<%			end if %>
<p>&nbsp;</p>
			  </td>
			</tr>
		  </table>
		  <br /><a href="admin.asp"><%=yyAdmHom%></a><br />&nbsp;
		</td>
	  </tr>
	</table>
<script>
/* <![CDATA[ */
function disableAnchor(obj, disable){
  if(disable){
    var href=obj.href;
    if(href && href!="" && href!=null){
       obj.href_bak=href;
    }
	obj.vrdibled=true;
    obj.removeAttribute('href');
    obj.style.color="gray";
  }else{
	obj.vrdibled=false;
    obj.setAttribute('href',obj.href_bak);
    obj.style.color="";
  }
}
function selectopt(optid){
	document.getElementById(optid).checked=true;
	var thediv=document.getElementById(optid+'div');
	document.getElementById(optid+'div2').style.color='#000000';
	thediv.style.backgroundColor='#FFFFFF';
	var opts=thediv.getElementsByTagName('a');
	i=0;
	while(opt=opts[i++]){
		disableAnchor(opt,false);
	}
	if(optid=='allin1') otheropt='exis'; else otheropt='allin1';
	var thediv=document.getElementById(otheropt+'div');
	document.getElementById(otheropt+'div2').style.color='#A0A0A0';
	thediv.style.backgroundColor='#E6E6E6';
	var opts=thediv.getElementsByTagName('a');
	i=0;
	while(opt=opts[i++]){
		disableAnchor(opt,true);
	}
}
var thediv=document.getElementById('allin1div');
var opts=thediv.getElementsByTagName('a');
i=0;
while(opt=opts[i++]){
	disableAnchor(opt,true);
}
var thediv=document.getElementById('exisdiv');
var opts=thediv.getElementsByTagName('a');
i=0;
while(opt=opts[i++]){
	disableAnchor(opt,true);
}
/* ]]> */
</script>
<%	elseif getget("act")="ccpaypal" then %>
	<table border="0" cellspacing="0" cellpadding="0" width="100%">
	  <tr>
		<td align="center">
		  <table width="80%" height="100%" border="0" cellspacing="0" cellpadding="2">
			<tr>
			  <td align="left"><p style="font-size:18px;font-weight:bold;">Choose a solution to accept credit cards and PayPal</p>
					<p>&nbsp;</p>
			  </td>
			</tr>
		  </table>
		  <table width="80%" border="0" cellspacing="2" cellpadding="2" bgcolor="#BFC9E0">
			<tr>
			  <td align="left" valign="top" bgcolor="#FFFFFF" width="50%">
			  <p style="font-size:16px;font-weight:bold;">&nbsp;PayPal Payments Standard</p>
			  <div style="font-size:12px;font-weight:bold;background-color:#E6E9F5;padding:4px;">Easy to get started, no monthly fees.<br />&nbsp;<br />
			  <p align="right" style="margin:0px;"><a href="" onclick="newwin=window.open('http://www.paypal.com/en_US/m/demo/demo_wps/demo_WPS.html','PayPalDemo','menubar=no,scrollbars=yes,width=598,height=380,directories=no,location=no,resizable=yes,status=no,toolbar=no');return false;">See demo</a></p>
			  </div>
			  <p>
			    <ul>
				  <li>Accept Visa, MasterCard, American Express, Discover, PayPal and more at one low rate.<br />&nbsp;</li>
				  <li>Buyers enter credit card information on secure PayPal pages and immediately return to your site. Your buyers do NOT need a PayPal account.<br />&nbsp;</li>
				  <li>Start selling as soon as you sign up.<br />&nbsp;</li>
				</ul>
			  </p>
			  <p style="font-size:12px;font-weight:bold;">Pricing</p>
			  <p>
			    <ul>
				  <li>No monthly fees.<br />&nbsp;</li>
				  <li>No setup or cancellation fees.<br />&nbsp;</li>
				  <li>Transaction fees: 1.9% - 2.9% + $0.30 USD<br />
				  (Based on sales volume)<br />&nbsp;</li>
				</ul>
			  </p>
			  <div align="center"><input type="button" value="Select" onclick="document.location='adminpayprov.asp?act=modify&from=wizard&id=1'" /></div>
			  </td>
			  
			  <td align="left" valign="top" bgcolor="#FFFFFF">
			  <p style="font-size:16px;font-weight:bold;">&nbsp;PayPal Payments Pro</p>
			  <div style="font-size:12px;font-weight:bold;background-color:#E6E9F5;padding:4px;">Advanced e-commerce solution for established businesses.
			  <p align="right" style="margin:0px;"><a href="" onclick="newwin=window.open('http://www.paypal.com/en_US/m/demo/wppro/paypal_demo_load_560x355.html','PayPalDemo','menubar=no,scrollbars=yes,width=578,height=372,directories=no,location=no,resizable=yes,status=no,toolbar=no');return false;">See demo</a></p>
			  </div>
			  <p>
			    <ul>
				  <li>Accept Visa, MasterCard, American Express, Discover, PayPal and more at one low rate.<br />&nbsp;</li>
				  <li>Buyers enter credit card info directly on your site, and do NOT need a PayPal account.<br />&nbsp;</li>
				  <li>Business credit application required to start selling. Decision usually comes within 24 hours.<br />&nbsp;</li>
				  <li>Includes Virtual Terminal - accept payments for orders taken via phone, fax and mail.<br />&nbsp;</li>
				</ul>
			  </p>
			  <p style="font-size:12px;font-weight:bold;">Pricing</p>
			  <p>
			    <ul>
				  <li>$30 per month.<br />&nbsp;</li>
				  <li>No setup or cancellation fees.<br />&nbsp;</li>
				  <li>Transaction fees: 2.2% - 2.9% + $0.30 USD<br />
				  (Based on sales volume)<br />&nbsp;</li>
				</ul>
			  </p>
			  <p align="center"><input type="button" value="Select" onclick="document.location='adminpayprov.asp?act=modify&from=wizard&id=18'" /></p>
			  </td>
			</tr>
		  </table>
		  <table width="80%" border="0" cellspacing="0" cellpadding="2">
			<tr>
			  <td align="left">
				<p style="font-size:11px;font-weight:bold;">Don't see what you are looking for?</p>
				<p style="font-size:11px;font-weight:bold;"><a href="adminpayprov.asp?act=list">See full list of payment processors</a></p>
			  </td>
			</tr>
		  </table>
		  <br /><a href="admin.asp"><%=yyAdmHom%></a><br />&nbsp;
		</td>
	  </tr>
	</table>
<%	else
function writeposition(currpos,maxpos)
	Dim reqtext,i
	reqtext="<select name='newpos" & currpos & "' size='1' onchange='javascript:validate_index("&currpos&");'>"
	for i=1 to maxpos
		reqtext=reqtext & "<option value='"&i&"'"
		if currpos=i then reqtext=reqtext&" selected=""selected"""
		reqtext=reqtext & ">"&i&"</option>"
	next
	writeposition=reqtext & "</select>"
end function
%>
<script>
/* <![CDATA[ */
function validate_index(currindex)
{
	var i=eval("document.mainform.newpos"+currindex+".selectedIndex")+1;
	document.mainform.newval.value=i;
	document.mainform.selectedq.value=currindex;
	document.mainform.act.value="changepos";
	if(i==document.mainform.selectedq.value){
		return (false);
	}
	document.mainform.submit();
}
/* ]]> */
</script>
	<form name="mainform" method="post" action="adminpayprov.asp">
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%" align="center">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="modify" />
			<input type="hidden" name="id" value="1" />
			<input type="hidden" name="selectedq" value="1" />
			<input type="hidden" name="newval" value="1" />
            <table width="700" border="0" cellspacing="0" cellpadding="2">
<%	if getget("act")<>"list" then %>
			  <tr>
                <td width="100%" colspan="4" align="left">
				<div>&nbsp;</div>
				<div style="font-size:18px;font-weight:bold;">Set-up credit card processing</div>
				<div>&nbsp;</div>
				<p>
				  <ul>
					<li><span style="font-size:13px;font-weight:bold;"><a href="adminpayprov.asp?act=ccpaypal">Accept Credit Cards and PayPal</a></span><br /><br />
					<a href="adminpayprov.asp?act=ccpaypal">
					<img border="0" src="adminimages/logo_ccVisa.gif" alt="Visa" />
					<img border="0" src="adminimages/logo_ccMC.gif" alt="Mastercard" />
					<img border="0" src="adminimages/logo_ccAmex.gif" alt="American Express" />
					<img border="0" src="adminimages/logo_ccDiscover.gif" alt="Discover" />
					<img border="0" src="adminimages/logo_ccEcheck.gif" alt="eCheck" />
					<img border="0" src="adminimages/PayPal_mark_37x23.gif" alt="PayPal" />
					</a><br />&nbsp;
					</li>
					<li><span style="font-size:13px;font-weight:bold;"><a href="adminpayprov.asp?act=cccards">Accept Credit Cards</a></span><br /><br />
					<a href="adminpayprov.asp?act=cccards">
					<img border="0" src="adminimages/logo_ccVisa.gif" alt="Visa" />
					<img border="0" src="adminimages/logo_ccMC.gif" alt="Mastercard" />
					<img border="0" src="adminimages/logo_ccAmex.gif" alt="American Express" />
					<img border="0" src="adminimages/logo_ccDiscover.gif" alt="Discover" />
					</a><br />&nbsp;
					</li>
				  </ul>
				</p>
				<p>&nbsp;</p>
				<p><span style="font-size:12px;">Note: You will be able to add additional payment options later in this set-up process</span></p>
				<p>&nbsp;</p>
				<p><span style="font-size:11px;font-weight:bold;"><a href="adminpayprov.asp?act=list">See full list of payment processors</a></span></p>
				</td>
			  </tr>
<%	else %>
			  <tr> 
                <td width="100%" colspan="4" align="center"><strong><%=yyPPAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td width="8%" align="center"><strong><%=yyID%></strong></td>
				<td width="8%" align="center"><strong><%=yyOrder%></strong></td>
				<td width="49%" align="left">&nbsp;<strong><%=yyPPName%></strong></td>
				<td width="25%" align="center"><strong><%=yyConf%></strong></td>
				<td width="10%" align="center"><strong><%=yyHlpFil%></strong></td>
			  </tr>
<%		showenabled=TRUE
		for index=0 to 1
			sSQL="SELECT payProvID,payProvName,payProvShow,payProvDemo,payProvEnabled,payProvData1,payProvData2,payProvMethod,payProvShow2,payProvShow3 FROM payprovider WHERE payProvAvailable=1"
			if showenabled then
				sSQL=sSQL & " AND payProvEnabled=1 ORDER BY payProvOrder"
			else
				sSQL=sSQL & " AND payProvEnabled=0 AND payProvID<>9 ORDER BY payProvName"
			end if
			rs.open sSQL,cnn,0,1
			alldata=""
			if NOT rs.EOF then alldata=rs.getrows
			rs.close
			if isarray(alldata) then
				if showenabled then enabledProv=UBOUND(alldata,2)+1 else enabledProv=0
				for rowcounter=0 to UBOUND(alldata,2)
					helplink=""
					if alldata(0,rowcounter)=1 then helplink="https://www.ecommercetemplates.com/help/ecommplus/paypal.asp" : alldata(1,rowcounter)="PayPal Payments Standard"
					if alldata(0,rowcounter)=2 then helplink="https://www.ecommercetemplates.com/help/ecommplus/2checkout.asp"
					if alldata(0,rowcounter)=3 then helplink="https://www.ecommercetemplates.com/help/ecommplus/authorizenet.asp"
					if alldata(0,rowcounter)=4 then helplink=""
					if alldata(0,rowcounter)=5 then helplink="https://www.ecommercetemplates.com/help/ecommplus/worldpay.asp"
					if alldata(0,rowcounter)=6 then helplink="https://www.ecommercetemplates.com/help/ecommplus/nochex.asp"
					if alldata(0,rowcounter)=7 then helplink="https://www.ecommercetemplates.com/help/ecommplus/paypal-payflow-pro.asp" : alldata(1,rowcounter)="PayPal Payflow Pro"
					if alldata(0,rowcounter)=8 then helplink="https://www.ecommercetemplates.com/help/ecommplus/paypal-payflow-link.asp" : alldata(1,rowcounter)="PayPal Payflow Link"
					if alldata(0,rowcounter)=9 then helplink="https://www.ecommercetemplates.com/help/ecommplus/paypoint.asp"
					if alldata(0,rowcounter)=10 then helplink=""
					if alldata(0,rowcounter)=11 then helplink="https://www.ecommercetemplates.com/help/ecommplus/psigate.asp"
					if alldata(0,rowcounter)=12 then helplink="https://www.ecommercetemplates.com/help/ecommplus/psigate.asp"
					if alldata(0,rowcounter)=13 then helplink="https://www.ecommercetemplates.com/help/ecommplus/authorizenet.asp"
					if alldata(0,rowcounter)=14 then helplink=""
					if alldata(0,rowcounter)=15 then helplink="https://www.ecommercetemplates.com/help/ecommplus/netbanx.asp"
					if alldata(0,rowcounter)=16 then helplink="https://www.ecommercetemplates.com/help/ecommplus/linkpoint.asp"
					if alldata(0,rowcounter)=17 then helplink=""
					if alldata(0,rowcounter)=18 then helplink="https://www.ecommercetemplates.com/help/ecommplus/paypal-pro.asp" : alldata(1,rowcounter)="PayPal Direct Payments"
					if alldata(0,rowcounter)=19 then helplink="https://www.ecommercetemplates.com/help/ecommplus/paypal-express-checkout.asp" : alldata(1,rowcounter)="PayPal Express Payments"
					if alldata(0,rowcounter)=20 then helplink=""
					if alldata(0,rowcounter)=21 then helplink=""
					if alldata(0,rowcounter)=22 then helplink="https://www.ecommercetemplates.com/help/ecommplus/paypal-advanced.asp" : alldata(1,rowcounter)="PayPal Payments Advanced"
					if alldata(0,rowcounter)=23 then helplink="https://www.ecommercetemplates.com/help/ecommplus/stripe.asp"
					if alldata(0,rowcounter)=24 then helplink="https://www.ecommercetemplates.com/help/ecommplus/sagepay.asp" : alldata(1,rowcounter)="Opayo / SagePay"
					if alldata(0,rowcounter)=25 then helplink="https://www.ecommercetemplates.com/help/ecommplus/braintree.asp"
					if alldata(0,rowcounter)=27 then helplink="https://www.ecommercetemplates.com/help/ecommplus/paypal-smart-buttons.asp"
					if alldata(0,rowcounter)=28 then helplink="https://www.ecommercetemplates.com/help/ecommplus/squareup.asp"
					if alldata(0,rowcounter)=29 then helplink="https://www.ecommercetemplates.com/help/ecommplus/nmi.asp"
					if alldata(0,rowcounter)=30 then helplink="https://www.ecommercetemplates.com/help/ecommplus/eway.asp"
					if alldata(0,rowcounter)=31 then helplink="https://www.ecommercetemplates.com/help/ecommplus/pay360.asp"
					if alldata(0,rowcounter)=32 then helplink="https://www.ecommercetemplates.com/help/ecommplus/globalpayments.asp"
					
					if bgcolor="altdark" then bgcolor="altlight" else bgcolor="altdark" %>
				  <tr class="<%=bgcolor%>">
					<td align="center"><%=alldata(0,rowcounter)%></td>
					<td align="center"><%if alldata(4,rowcounter)=1 then print writeposition(rowcounter+1,enabledProv) else print "-" %></td>
					<td align="left">&nbsp;&nbsp;<%if alldata(3,rowcounter)=1 then print "<span style=""color:#FF0000"">" %><%if alldata(4,rowcounter)=1 then print "<strong>" %><%=alldata(1,rowcounter)%><%if alldata(4,rowcounter)=1 then print "</strong>" %><%if alldata(3,rowcounter)=1 then print "</span>" %></td>
					<td align="center"><input type="button" value="<%=yyModify%>" onclick="modrec('<%=alldata(0,rowcounter)%>')" /></td>
					<td align="center"><% if helplink="" then print "&nbsp;" else print "<a href=""" & helplink & """ class=""online_help_large"" target=""_blank"">?</a>"%></td>
				  </tr>
<%				next
			end if
			showenabled=FALSE
		next %>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br /><%=yyPPEx1%><br />
				  <%=yyPPEx2%>&nbsp;</td>
			  </tr>
<%	end if %>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br /><a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
			  </tr>
            </table>
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
