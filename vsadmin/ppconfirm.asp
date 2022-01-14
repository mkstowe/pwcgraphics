<%@LANGUAGE="VBScript"%>
<%
anetsignature=""
for each sobj in request.servervariables
	if instr(ucase(replace(replace(sobj,"-",""),"_","")),"XANETSIGNATURE")>0 then
		anetsignature=request.servervariables(sobj)
		exit for
	end if
next
%>
<!--#include file="inc/md5.asp"-->
<!--#include file="db_conn_open.asp"-->
<!--#include file="inc/languagefile.asp"-->
<%	if anetsignature="" then %>
<!--#include file="includes.asp"-->
<%	end if %>
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
ordGrandTotal=0 : ordTotal=0 : ordStateTax=0 : ordHSTTax=0 : ordCountryTax=0 : ordShipping=0 : ordHandling=0 : ordDiscount=0
affilID = "" : ordCity="" : ordState = "" : ordCountry = "" : ordDiscountText="" : emailtxt="PPCONFIRM Debug<br>"
function already_authorized(tid)
	already_authorized=FALSE
	sSQL = "SELECT ordAuthNumber FROM orders WHERE ordID=" & tid
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		if trim(rs("ordAuthNumber")&"")<>"" AND trim(rs("ordAuthNumber")&"")<>"no ipn" AND trim(rs("ordAuthNumber")&"")<>"CHECK MANUALLY" then already_authorized=TRUE
	end if
	rs.close
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
set rs = Server.CreateObject("ADODB.RecordSet")
set rs2 = Server.CreateObject("ADODB.RecordSet")
set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
isamazonpayment = FALSE
if anetsignature<>"" then
	if getpayprovdetails(3,data1,data2,data3,demomode,ppmethod) then
		biData=request.binaryread(request.totalbytes)
		payload=ANSIToUnicode(biData)
		hashedBody=UCASE(calcHMACSha512(data3,payload,"TEXT","TEXT"))
		if anetsignature="sha512=" & hashedBody then
			transid=get_json_val(payload,"id","")
			sjson="{""getTransactionDetailsRequest"":{" & _
						"""merchantAuthentication"":{""name"":" & json_encode(data1) & ",""transactionKey"":" & json_encode(data2) & "}," & _
						"""transId"":" & json_encode(transid) & "}}"
			success=callxmlfunction("https://api" & IIfVs(demomode,"test") & ".authorize.net/xml/v1/request.api",sjson,jres,"","Msxml2.ServerXMLHTTP",vsRESPMSG,FALSE)
			if success then
				ordID=get_json_val(jres,"invoiceNumber","")
				authcode=get_json_val(jres,"authCode","")
				responsecode=get_json_val(jres,"responseCode","")
				ordavs=get_json_val(jres,"AVSResponse","")
				ordcvv=get_json_val(jres,"cardCodeResponse","")
				statusdesc=get_json_val(jres,"responseReasonDescription","")
				if NOT is_numeric(ordID) then
					emailtxt=emailtxt&"Missing Order ID" & emlNl
				elseif responsecode="1" then
					do_send_emails=NOT already_authorized(ordID)
					cnn.execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID=" & escape_string(ordID))
					cnn.execute("UPDATE orders SET ordAVS='" & ordavs & "',ordCVV='" & ordcvv & "',ordStatus=3,ordAuthNumber='" & authcode & "',ordAuthStatus='',ordTransID='" & transid & "' WHERE ordPayProvider IN (3) AND ordID=" & escape_string(ordID) & " AND ordStatus<3")
					if do_send_emails then call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
				else
					cnn.execute("UPDATE orders SET ordPrivateStatus='" & escape_string(strip_tags2("(" & responsecode & ") " & statusdesc)) & "' WHERE ordPayProvider IN (3) AND ordAuthNumber='' AND ordID=" & escape_string(ordID))
				end if
			end if		
		else
			emailtxt=emailtxt&"Auth.net Signature Mismatch" & emlNl
			emailtxt=emailtxt&"anetsignature:" & anetsignature & emlNl
			emailtxt=emailtxt&"hashedBody: sha512=" & hashedBody & emlNl
		end if
	end if	
else
	if trim(request.form("transactionId"))<>"" AND trim(request.form("status"))<>"" AND trim(request.form("referenceId"))<>"" AND trim(request.form("signature"))<>"" then isamazonpayment=TRUE
	if isamazonpayment then
		call getpayprovdetails(21,data1,data2,data3,demomode,ppmethod)
		ordID = replace(request.form("referenceId"), "'", "")
		Txn_id = replace(request.form("transactionId"), "'", "")
		avs = ""
		cvv = ""
		receipt_id = ""
		sigchk=""
		dim sigarr(50)
		signum=0
		for each objElem in Request.Form
			if NOT objElem="signature" then sigarr(signum) = objElem & request.form(objElem) : signum = signum + 1
		next
		stillchecking=TRUE
		do while stillchecking
			stillchecking=FALSE
			for index=0 to signum-2
				if sigarr(index)>sigarr(index+1) then sigtmp = sigarr(index) : sigarr(index) = sigarr(index+1) : sigarr(index+1) = sigtmp : stillchecking=TRUE
			next
		loop
		for index=0 to signum-1
			sigchk = sigchk & sigarr(index)
		next
		b64pad="="
		thesig = b64_hmac_sha1(data2,sigchk)
		if (request.form("status")="PS" OR request.form("status")="PR" OR request.form("status")="PI") AND is_numeric(ordID) then
			do_send_emails = NOT already_authorized(ordID)
			if request.form("status")="PR" then authstatus="Pending: Settle" else authstatus="Pending: Check Processor"
			if request.form("status")="PI" then authstatus="Pending: Review"
			ect_query("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
			ect_query("UPDATE orders SET ordAVS='"&avs&"',ordCVV='"&cvv&"',ordStatus=3,ordAuthNumber='Check Processor',ordAuthStatus='"&authstatus&"',ordTransID='"&Txn_id&"' WHERE ordPayProvider=21 AND ordID="&ordID)
			if do_send_emails then call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
		end if
	else
		Receiver_email = Request.Form("receiver_email")
		Item_number = Request.Form("item_number")
		invoice = Request.Form("invoice")
		Payment_status = Request.Form("payment_status")
		session.LCID = 1033
		mc_gross = cdbl(Request.Form("mc_gross"))+0.01
		session.LCID = saveLCID
		Txn_id = Request.Form("txn_id")
		ordID = trim(replace(IIfVr(invoice<>"",invoice,request.form("custom")),"'",""))
		Payer_email = Request.Form("payer_email")
		receipt_id = trim(request.form("receipt_id"))
		address_status = lcase(trim(request.form("address_status")))
		pending_reason = Request.Form("pending_reason")
		if address_status="confirmed" then
			avs = "Y"
		elseif address_status="unconfirmed" then
			avs = "N"
		else
			avs = "U"
		end if
		payer_status = lcase(trim(request.form("payer_status")))
		if payer_status="verified" then
			cvv = "Y"
		elseif payer_status="unverified" then
			cvv = "N"
		else
			cvv = "U"
		end if
		str = Request.Form & "&cmd=_notify-validate"
		xmlfnheaders=array(array("Content-Type","application/x-www-form-urlencoded"),array("Host","www.paypal.com"),array("Connection","close"))
		if trim(request.form("txn_type"))="express_checkout" AND trim(request.form("parent_txn_id"))<>"" then
			sSQL = "SELECT ordID FROM orders WHERE ordDate>" & vsusdate(date()-31) &" AND ordAuthNumber='" & escape_string(request.form("parent_txn_id")) & "'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then ordID = rs("ordID")
			rs.close
		end if
		if (instr(ordID,":")=0 AND is_numeric(ordID)) OR getget("ppdebug")="true" OR getget("ppdebug")="tls" then
			if NOT is_numeric(ordID) OR getget("ppdebug")="true" OR getget("ppdebug")="tls" then ordID=0
			payprov=1
			sSQL = "SELECT ordPayProvider FROM orders WHERE ordID=" & ordID
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then payprov=rs("ordPayProvider")
			rs.close
			call getpayprovdetails(payprov,data1,data2,data3,demomode,ppmethod)
			emailtxt = emailtxt & "demomode: " & demomode & emlNl
			endpoint="https://ipnpb." & IIfVs(demomode OR getget("ppdebug")="tls","sandbox.") & "paypal.com/cgi-bin/webscr"
			if getget("ppdebug")="true" OR getget("ppdebug")="tls" then print "Testing URL: " & endpoint & "<br />"
			if callxmlfunction(endpoint, str, paypalres, "", "WinHTTP.WinHTTPRequest.5.1", errormsg, FALSE) then
				if paypalres="VERIFIED" then
					amount=0
					orderexists=TRUE
					sSQL = "SELECT ordShipping,ordStateTax,ordCountryTax,ordHandling,ordTotal,ordDiscount FROM orders WHERE ordID=" & ordID
					rs.open sSQL,cnn,0,1
						if NOT rs.EOF then amount=(rs("ordShipping")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordTotal")+rs("ordHandling"))-rs("ordDiscount") else orderexists=FALSE
					rs.close
					if orderexists then
						if Payment_status="Completed" OR Payment_status="Pending" then
							ect_query("UPDATE orders SET ordAVS='"&avs&"',ordCVV='"&cvv&"' WHERE ordPayProvider=1 AND ordID="&ordID)
						end if
						if (Payment_status="Completed" OR Payment_status="Pending") AND amount > (mc_gross) then
							ect_query("UPDATE cart SET cartCompleted=2 WHERE cartCompleted=0 AND cartOrderID="&ordID)
							ect_query("UPDATE orders SET ordAuthNumber='"&Txn_id&"',ordAuthStatus='Pending: Total paid " & request.form("mc_currency") & " " & mc_gross & "' WHERE ordPayProvider IN (1,18,19,22) AND ordID="&ordID)
						elseif Payment_status="Completed" then
							do_send_emails = NOT already_authorized(ordID)
							ect_query("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
							ect_query("UPDATE orders SET ordAuthNumber='"&Txn_id&"',ordAuthStatus='',ordTransID='"&receipt_id&"' WHERE ordPayProvider IN (1,18,19,22) AND ordID="&ordID)
							ect_query("UPDATE orders SET ordStatus=3 WHERE ordStatus<3 AND ordPayProvider IN (1,18,19,22) AND ordID="&ordID)
							if do_send_emails then call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
						elseif Payment_status="Pending" then
							if pending_reason="authorization" then pending_reason="Capture"
							do_send_emails = NOT already_authorized(ordID)
							ect_query("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
							ect_query("UPDATE orders SET ordStatus=3,ordAuthNumber='"&Txn_id&"',ordAuthStatus='Pending: " & escape_string(pending_reason) & "' WHERE ordPayProvider IN (1,18,19,22) AND ordID="&ordID)
							if do_send_emails then call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
						end if
					end if
				elseif paypalres="INVALID" then
					' log for manual investigation
				else
					if debugmode=TRUE then response.write paypalres ' error
				end if
				emailtxt = emailtxt & "Result : " & paypalres & emlNl
				if getget("ppdebug")="true" OR getget("ppdebug")="tls" then
					print "Result : " & paypalres & "<br />"
					if paypalres="INVALID" then
						print "This is a good/correct result as it shows that communication with the PayPal server was successful and the transaction was of course rejected as invalid.<br />"
					else
						print "This is not a correct response and may indicate problems with communication with the PayPal server.<br />"
					end if
				end if
			else
				if getget("ppdebug")="true" OR getget("ppdebug")="tls" then print "Error : " & errormsg & "<br />"
			end if
		end if
	end if ' isamazonpayment
end if
if debugmode=TRUE then
	if anetsignature="" then
		for each objItem In Request.Form
			emailtxt = emailtxt & objItem & " : " & Request.Form(objItem) & emlNl
		next
	end if
	Call DoSendEmailEO(emailAddr,emailAddr,"","ppconfirm.asp debug",emailtxt,emailObject,themailhost,theuser,thepass)
end if
cnn.close
set cnn=nothing
%>