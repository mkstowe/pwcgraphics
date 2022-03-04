<%@LANGUAGE="VBScript"%>
<!--#include file="inc/md5.asp"-->
<!--#include file="db_conn_open.asp"-->
<!--#include file="inc/languagefile.asp"-->
<!--#include file="inc/incemail.asp"-->
<!--#include file="inc/incfunctions.asp"-->
<%
'=========================================
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
set rs = Server.CreateObject("ADODB.RecordSet")
set rs2 = Server.CreateObject("ADODB.RecordSet")
set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
emlNl="<br />"
ordGrandTotal=0 : ordTotal=0 : ordStateTax=0 : ordHSTTax=0 : ordCountryTax=0 : ordShipping=0 : ordHandling=0 : ordDiscount=0
ordTransSession="" : affilID = "" : ordCity="" : ordState = "" : ordCountry = "" : ordDiscountText="" : emailtxt="Stripe Webhook Debug" & emlNl
function already_authorized(tid)
	already_authorized=FALSE
	sSQL = "SELECT ordAuthNumber,ordAuthStatus,ordTransSession FROM orders WHERE ordID=" & tid
	rs2.open sSQL,cnn,0,1
	if NOT rs2.EOF then
		if trim(rs2("ordAuthNumber")&"")<>"" AND trim(rs2("ordAuthNumber")&"")<>"no ipn" AND trim(rs2("ordAuthNumber")&"")<>"CHECK MANUALLY" then already_authorized=TRUE
		ordTransSession=rs2("ordTransSession")
	end if
	rs2.close
end function
function ANSIToUnicode(ByRef pbinBinaryData)
	Dim lbinData	' Binary Data (ANSI)
	Dim llngLength	' Length of binary data (byte count)
	Dim lobjRs		' Recordset
	Dim lstrData	' Unicode Data
	Set lobjRs=Server.CreateObject("ADODB.Recordset")
	if VarType(pbinBinaryData)=8 then
		llngLength=LenB(pbinBinaryData)
		if llngLength=0 then
			lbinData=ChrB(0)
		else
			Call lobjRs.Fields.Append("BinaryData", adLongVarBinary, llngLength)
			Call lobjRs.Open()
			Call lobjRs.AddNew()
			Call lobjRs.Fields("BinaryData").AppendChunk(pbinBinaryData & ChrB(0)) ' + Null terminator
			Call lobjRs.Update()
			lbinData=lobjRs.Fields("BinaryData").GetChunk(llngLength)
			Call lobjRs.Close()
		end if
	else
		lbinData=pbinBinaryData
	end if
	llngLength=LenB(lbinData)
	if llngLength=0 then
		lstrData=""
	else
		Call lobjRs.Fields.Append("BinaryData", 201, llngLength)
		Call lobjRs.Open()
		Call lobjRs.AddNew()
		Call lobjRs.Fields("BinaryData").AppendChunk(lbinData)
		Call lobjRs.Update()
		lstrData=lobjRs.Fields("BinaryData").Value
		Call lobjRs.Close()
	end if
	Set lobjRs=nothing
	ANSIToUnicode=lstrData
end function

if getpayprovdetx(IIfVr(getget("pprov")="31",31,23),data1,data2,data3,data4,data5,data6,ppflag1,ppflag2,ppflag3,ppbits,demomode,ppmethod) then
	biData=request.binaryread(request.totalbytes)
	payload=ANSIToUnicode(biData)
	emailtxt=emailtxt&"PAYLOAD: " & payload & emlNl
	if getget("pprov")="31" then
		ordID=get_json_val(payload,"merchantRef","")
		signature=sha256(ordID & adminSecret & "pay360")
		if is_numeric(ordID) AND signature=getget("signature") then
			transactionid=get_json_val(payload,"transactionId","")
			ect_query("UPDATE orders SET ordTransID='" & escape_string(transactionid) & "' WHERE ordID=" & ordID & " AND ordStatus<3 AND ordPayProvider=31")
		end if
	else
		success=TRUE
		data4=""
		if trim(data4&"")<>"" then
			sig_header=request.servervariables("HTTP_STRIPE_SIGNATURE")
			header_array=split(sig_header,",")
			payload_sign="" : time_sig=""
			for each val in header_array
				pairs=split(val,"=")
				if pairs(0)="t" then time_sig=pairs(1)
				if pairs(0)="v1" then payload_sign=pairs(1)
			next
			hashedBody=hex_hmac_sha256(data4, time_sig & "." & payload)
			if payload_sign="" then
				success=FALSE
				emailtxt=emailtxt&"Blank payload sign" & emlNl
				print "Blank payload sign"
				response.status="500 Internal Server Error"
			elseif payload_sign<>hashedBody then
				success=FALSE
				emailtxt=emailtxt&"Payload sign mismatch" & emlNl
				print "Payload sign mismatch"
				response.status="500 Internal Server Error"
			end if
		end if
		if success then
			webhooktype=get_json_val(payload,"type","")
			if instr(payload,"""checkout.session.completed""")>0 then
				paymentintent=get_json_val(payload,"payment_intent","")
				xmlfnheaders=array(array("User-Agent","Stripe/v1 RubyBindings/1.12.0"),array("Authorization","Bearer "&data1),array("Content-Type","application/x-www-form-urlencoded"))
				if callxmlfunction("https://api.stripe.com/v1/payment_intents/"&paymentintent,"",jres,"","Msxml2.ServerXMLHTTP", errtext, FALSE) then
					emailtxt=emailtxt&"PAYMENT INTENT: " & jres & emlNl
					paymentstatus=get_json_val(jres,"status","payment_method_types")
					if (ppmethod<>1 AND paymentstatus="succeeded") OR (ppmethod=1 AND paymentstatus="requires_capture") then
						ordavs1=replace(get_json_val(jres,"address_line1_check",""),"null","")
						ordavs2=replace(get_json_val(jres,"address_postal_code_check",""),"null","")
						ordcvv=replace(get_json_val(jres,"cvc_check",""),"null","")
						ordID=""
						sSQL="SELECT ordID FROM orders WHERE ordPayProvider=23 AND ordStatus<3 AND ordTransID='" & escape_string(paymentintent) & "'"
						rs.open sSQL,cnn,0,1
						if NOT rs.EOF then ordID=rs("ordID")
						rs.close
						if ordID<>"" then
							do_send_emails=NOT already_authorized(ordID)
							ect_query("UPDATE cart SET cartCompleted=1 WHERE cartOrderID=" & ordID)
							ect_query("UPDATE orders SET ordAVS='" & escape_string(trim(ordavs1&":"&ordavs2))&"',ordCVV='"&escape_string(ordcvv)&"',ordStatus=3,ordAuthNumber='Authorized' WHERE ordID=" & ordID & " AND ordStatus<3")
							if do_send_emails then call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
						end if
					end if
				end if
			end if
		end if
	end if
end if
if debugmode=TRUE then
	call DoSendEmailEO(emailAddr,emailAddr,"","stripewebhook.asp debug",emailtxt,emailObject,themailhost,theuser,thepass)
end if
cnn.close
set cnn=nothing
%>