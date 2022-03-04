<%@LANGUAGE="VBScript"%>
<!--#include file="inc/md5.asp"-->
<!--#include file="db_conn_open.asp"-->
<!--#include file="inc/languagefile.asp"-->
<!--#include file="includes.asp"-->
<!--#include file="inc/incemail.asp"-->
<!--#include file="inc/incfunctions.asp"-->
<%
'=========================================
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
Dim str,Txn_id,Payment_status,objHttp
Dim emailtxt
ordGrandTotal=0 : ordTotal=0 : ordStateTax=0 : ordHSTTax=0 : ordCountryTax=0 : ordShipping=0 : ordHandling=0 : ordDiscount=0
affilID="" : ordCity="" : ordState="" : ordCountry="" : ordDiscountText="" : emailtxt=""
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs2=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
if htmlemails=true then emlNl="<br />" else emlNl=vbCrLf
alreadygotadmin=getadminsettings()
if request.form("nbx_merchant_reference")<>"" AND request.form("nbx_netbanx_reference")<>"" AND request.form("nbx_checksum")<>"" AND request.form("nbx_status")="passed" then ' Netbanx
	ordIDarr=split(request.form("nbx_merchant_reference"), ".")
	ordID=ordIDarr(0)
	if is_numeric(ordID) AND getpayprovdetails(15,data1,data2,data3,demomode,ppmethod) then
		Txn_id=request.form("nbx_netbanx_reference")
		thechecksum=request.form("nbx_checksum")
		checksumstring=request.form("nbx_payment_amount")&request.form("nbx_currency_code")&request.form("nbx_merchant_reference")&request.form("nbx_netbanx_reference")&data2
		calculatedsum=hex_sha1(checksumstring)
		' sha1_hex($amount.$currency.$ref.$nbx_reference.$secret_key)
		if thechecksum=calculatedsum OR trim(data2)="" then
			cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
			cnn.Execute("UPDATE orders SET ordAuthNumber='"&escape_string(Txn_id)&"',ordStatus=3,ordAVS='"&escape_string(request.form("nbx_houseno_auth")&"-"&request.form("nbx_postcode_auth"))&"',ordCVV='"&escape_string(request.form("nbx_CVV_auth"))&"',ordAuthStatus='' WHERE ordPayProvider=15 AND ordID="&ordID)
			call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
			print "NETBANXOK:1"
		else
			emailtxt=emailtxt & "Checksum mismatch: " & thechecksum & " : " & calculatedsum & emlNl & checksumstring & emlNl
		end if
	end if
else ' Nochex
	str=Request.Form
	' assign posted variables to local variables
	ordID=trim(Replace(Request.Form("order_id"),"'",""))
	Txn_id=Replace(Request.Form("transaction_id"),"'","")
	card_address_check=Replace(Request.Form("card_address_check"),"'","")
	card_postcode_check=Replace(Request.Form("card_postcode_check"),"'","")
	card_security_code=Replace(Request.Form("card_security_code"),"'","")
	' post back to NOCHEX system to validate
	if proxyserver<>"" then
		set objHttp=Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
		objHttp.setProxy 2, proxyserver
	else
		set objHttp=Server.CreateObject("Msxml2.ServerXMLHTTP")
	end if
	Receiver_email=Replace(Request.Form("to_email"),"'","")
	Payment_gross=Replace(Request.Form("amount"),"'","")
	Payer_email=Replace(Request.Form("from_email"),"'","")
	objHttp.open "POST", "https://www.nochex.com/apcnet/apc.aspx", false
	objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objHttp.Send str
	paypalres=objHttp.responseText
	paypalstatus=objHttp.status
	if paypalstatus<>200 then
		' HTTP error handling
	elseif paypalres="AUTHORISED" then
		' check that Payment_status=Completed
		' check that Txn_id has not been previously processed
		' check that Receiver_email is an email address in your PayPal account
		sSQL="UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID
		cnn.Execute(sSQL)
		if lcase(request.form("status"))<>"live" then authstatus="Pending: DEMO MODE" else authstatus=""
		sSQL="UPDATE orders SET ordStatus=3,ordAuthNumber='"&Txn_id&"',ordAVS='"&card_address_check&":"&card_postcode_check&"',ordCVV='"&card_security_code&"',ordAuthStatus='"&authstatus&"' WHERE ordPayProvider=6 AND ordID="&ordID
		cnn.Execute(sSQL)
		call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
	elseif paypalres="DECLINED" then
		' log for manual investigation
	else
		' error
	end if
	set objHttp=nothing
	emailtxt=emailtxt & "usenewnochexcallback : " & usenewnochexcallback & emlNl
	emailtxt=emailtxt & "SQL : " & sSQL & emlNl
	emailtxt=emailtxt & "Status : " & paypalstatus & emlNl
	emailtxt=emailtxt & "Result : " & paypalres & emlNl
end if
if debugmode=TRUE then
	for each objItem In Request.Form
		emailtxt=emailtxt & objItem & " : " & Request.Form(objItem) & emlNl
	next
	call DoSendEmailEO(emailAddr,emailAddr,"","ncconfirm.asp debug",emailtxt,emailObject,themailhost,theuser,thepass)
end if
cnn.close
set cnn=nothing
%>