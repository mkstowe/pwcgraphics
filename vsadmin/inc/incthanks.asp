<!--#include file="incemail.asp"-->
<!--#include file="md5.asp"-->
<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if sDSN="" then response.write "Database connection not set" : response.end
Dim rs,rs2,sSQL,orderText,errtext,ordGrandTotal,ordTotal,ordID
if request.totalbytes > 10000 then response.end
success=FALSE
ordGrandTotal=0 : ordTotal=0 : ordStateTax=0 : ordHSTTax=0 : ordCountryTax=0 : ordShipping=0 : ordHandling=0 : ordDiscount=0
errtext="" : ordTransSession="" : ordAuthNumber="" : txn_id="" : affilID="" : ordCity="" : ordState="" : ordCountry="" : ordDiscountText="" : ordEmail="" : googleanalyticstrackorderinfo=""
if dateadjust="" then dateadjust=0
SESSION("couponapply")=empty
SESSION("giftcerts")=empty
SESSION("cpncode")=empty
paypalwaitipn=FALSE : showclickreload=FALSE
if debugmode then
	print "POST parameters<br />"
	for each objItem in request.form
		print "POST: " & objItem & " : " & request.form(objItem) & "<br />"
	next
	print "GET parameters<br />"
	for each objItem in request.querystring
		print "GET: " & objItem & " : " & request.querystring(objItem) & "<br />"
	next
end if
sub wait_paypal_ipn(ppordid,ordTransSession)
	print "<form id=""ppectform"" method=""post"">"
	for each objItem in request.form
		print whv(objItem,request.form(objItem))
	next
	print "</form>"
%>
<script>/* <![CDATA[ */
var totpptries=0;
var ajaxobj;
function checkipncallback(){
	if(ajaxobj.readyState==4){
		if(ajaxobj.responseText=='1'){
			document.getElementById("ppectform").submit();
		}else{
			totpptries++;
			if(totpptries<30)
				setTimeout('checkipnarrived()',1000);
			else
				document.getElementById("orderfail").innerHTML='<div style="margin:50px"><strong><%=jsescape(xxPPTOUT)%></strong></div>';
		}
	}
}
function checkipnarrived(){
	ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
	ajaxobj.onreadystatechange=checkipncallback;
	ajaxobj.open("GET","vsadmin/ajaxservice.asp?action=ipnarrived&oid=<%=ppordid & IIfVs(ordTransSession<>"","&sid=" & ordTransSession) %>",true);
	ajaxobj.send(null);
}
checkipnarrived();
/* ]]> */</script><%
end sub
sub order_failed()
	call order_failed_htmldisp(TRUE)
end sub
sub order_failed_htmldisp(dohtmldisplay)
	success=FALSE
%>	<div class="ectdiv">
		<div class="ectmessagescreen">
			<div id="orderfail" style="text-align:center"><%
		print xxThkErr
		if errtext<>"" then print "<div style=""margin:50px""><strong>" & IIfVr(dohtmldisplay,htmldisplay(errtext),errtext) & "</strong></div>"
		if paypalwaitipn then
			call wait_paypal_ipn(IIfVr(is_numeric(ordID),ordID,txn_id),ordTransSession)
			print "<div style=""text-align:center""><img style=""margin:30px"" src=""images/preloader.gif"" alt=""Loading"" /></div>"
		elseif showclickreload then
			print "<div style=""margin:50px;text-align:center""><input type=""button"" class=""ectbutton"" value=""" & xxClkRel & """ onclick=""window.location.reload()"" /></div>"
		end if
%>				<a class="ectlink" href="<%=storeUrl%>"><strong><%=xxCntShp%></strong></a>
			</div>
		</div>
	</div>
<%
end sub
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs2=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
if getget("action")="termpp" AND getget("paymentID")<>"" AND is_numeric(getget("ordid")) then
	ordID=getget("ordid")
	if is_numeric(ordID) then
		sSQL="SELECT ordStatus,ordPrivateStatus FROM orders WHERE ordID='" & escape_string(ordID) & "' AND ordStatus>=3 AND ordTransID='" & escape_string(getget("paymentID")) & "'"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			ordStatus=rs("ordStatus")
			errtext=rs("ordPrivateStatus")
		end if
		rs.close
		if ordStatus>=3 then
			call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
		else
			call order_failed()
		end if
	end if
elseif getget("mode")="authorize" AND getget("method")="29" AND getget("token-id")<>"" AND is_numeric(getget("ordernumber")) then ' NMI
	success=FALSE
	ordID=getget("ordernumber")
	if getpayprovdetx(29,data1,data2,data3,data4,data5,data6,ppflag1,ppflag2,ppflag3,ppbits,demomode,ppmethod) then
		tsessionid=""
		sXML="<?xml version=""1.0"" encoding=""UTF-8""?><complete-action>" & vrxmltag("api-key",data1) &  vrxmltag("token-id",getget("token-id")) & "</complete-action>"
		sSQL="SELECT ordID,ordStatus,ordSessionID FROM orders WHERE ordID=" & ordID
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			currstatus=rs("ordStatus")
			tsessionid=rs("ordSessionID")
		end if
		rs.close
		if currstatus>=3 AND tsessionid=getsessionid() then
			call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
		else
			xmlfnheaders=array(array("Content-Type","text/xml"))
			if callxmlfunction("https://secure.networkmerchants.com/api/v2/three-step",sXML,xmlres,"","Msxml2.ServerXMLHTTP",errormsg,FALSE) then
				set xmlDoc=Server.CreateObject("MSXML2.DOMDocument")
				xmlDoc.validateOnParse=FALSE
				xmlDoc.loadXML(xmlres)
				set nmiresultobj=xmlDoc.getElementsByTagName("result")
				resultcode=nmiresultobj.item(0).firstChild.nodeValue
				if resultcode="1" then
					set nmiresultobj=xmlDoc.getElementsByTagName("authorization-code")
					authcode=nmiresultobj.item(0).firstChild.nodeValue
					set nmiresultobj=xmlDoc.getElementsByTagName("transaction-id")
					transid=nmiresultobj.item(0).firstChild.nodeValue
					set nmiresultobj=xmlDoc.getElementsByTagName("avs-result")
					avsresult=nmiresultobj.item(0).firstChild.nodeValue
					success=TRUE
					sSQL="UPDATE cart SET cartCompleted=1 WHERE cartOrderID=" & escape_string(ordID)
					ect_query(sSQL)
					sSQL="UPDATE orders SET ordStatus=3,ordAuthStatus='',ordAuthNumber='" & escape_string(authcode) & "',ordTransID='" & escape_string(transid) & "',ordAVS='" & escape_string(avsresult) & "',ordCVV='' WHERE ordPayProvider=29 AND ordID=" & escape_string(ordID)
					ect_query(sSQL)
					call do_order_success(ordID,emailAddr,sendEmail,TRUE,TRUE,TRUE,TRUE)
				else
					set nmiformobj=xmlDoc.getElementsByTagName("result-text")
					call order_failed()
				end if
			else
				call order_failed()
			end if
		end if
	end if
elseif getpost("mode")="authorize" AND getpost("method")="28" AND is_numeric(getpost("ordernumber")) AND getpost("payment_method_nonce")<>"" AND getpost("txnid")<>"" then ' SquareUp
	ordID=getpost("ordernumber")
	btsessionid=getpost("sessionid")
	success=TRUE
	if getpayprovdetx(28,data1,data2,data3,data4,data5,data6,ppflag1,ppflag2,ppflag3,ppbits,demomode,ppmethod) then
		sSQL="SELECT ordID FROM orders WHERE ordID=" & ordID & " AND ordTransID='" & escape_string(getpost("txnid")) & "'"
		rs.open sSQL,cnn,0,1
		success=NOT rs.EOF
		rs.close
		if success then
			call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
		end if
	end if
elseif getget("pprov")="31" AND is_numeric(getget("ordid")) AND getget("sessionId")<>"" then ' Pay360
	if getpayprovdetails(31,data1,data2,data3,demomode,ppmethod) then
		success=FALSE
		sSQL="SELECT ordID,ordStatus,ordTransID,ordTransSession FROM orders WHERE ordID='" & escape_string(getget("ordid")) & "' AND ordPayProvider=31 AND ordTransSession='" & escape_string(getget("sessionId")) & "'"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			ordID=getget("ordid")
			ordTransSession=rs("ordTransSession")
			ordtransid=rs("ordTransID")
			ordstatus=rs("ordStatus")
			success=TRUE
		end if
		rs.close
		if success then
			url="https://api." & IIfVs(demomode,"mite.") & "pay360.com"
			xmlfnheaders=array(array("Content-Type","application/json"),array("Authorization", "Basic "&vrbase64_encrypt(data2&":"&data3)))
			if ordstatus>=3 then
				call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
			else
				if ordtransid="" then
					paypalwaitipn=TRUE
					xxThkErr=""
					ect_query("UPDATE cart SET cartCompleted=2 WHERE cartCompleted=0 AND cartOrderID='" & escape_string(ordID) & "'")
					ect_query("UPDATE orders SET ordAuthNumber='no ipn' WHERE ordAuthNumber='' AND ordPayProvider=31 AND ordID='" & escape_string(ordID) & "'")
					errtext=xxPPWIPN
					call order_failed()
				else
					if callxmlfunction(url & "/acceptor/rest/transactions/" & data1 & "/" & ordtransid,"",jres,"","Msxml2.ServerXMLHTTP",errormsg,FALSE) then
						transactionStatus=get_json_val(jres,"status","authResponse")
						total=get_json_val(jres,"amount","")
						transactionid=get_json_val(jres,"transactionId","")
						isdeferred=get_json_val(jres,"deferred","")
						ordavs=get_json_val(jres,"avsAddressCheck","") & " | " & get_json_val(jres,"avsPostcodeCheck","")
						ordcvv=get_json_val(jres,"cv2Check","")
						ordavs=replace(replace(ordavs,"FULL_MATCH","FM"),"MATCHED","M")
						ordcvv=replace(replace(ordcvv,"FULL_MATCH","FM"),"MATCHED","M")
						authcode=get_json_val(jres,"authCode","")
						ect_query("UPDATE orders SET ordTransID='" & escape_string(transactionid) & "' WHERE ordID='" & escape_string(ordID) & "' AND ordStatus<3")
						if transactionStatus="AUTHORISED" then
							ect_query("UPDATE cart SET cartCompleted=1 WHERE cartOrderID='" & escape_string(ordID) & "'")
							ect_query("UPDATE orders SET ordAVS='" & escape_string(trim(ordavs)) & "',ordCVV='" & escape_string(ordcvv) & "',ordStatus=3,ordAuthNumber='" & escape_string(authcode) & "',ordAuthStatus='" & escape_string(IIfVs(isdeferred="true","Pending: Capture")) & "' WHERE ordID='" & escape_string(ordID) & "' AND ordStatus<3")
							call do_order_success(ordID,emailAddr,sendEmail,TRUE,TRUE,TRUE,TRUE)
						else
							errtext=xxCCErro
							call order_failed()
						end if
					end if
				end if
			end if
		else
			errtext="Order Not Found."
			call order_failed()
		end if
	end if
elseif getpost("pprov")="21" AND is_numeric(getpost("ordernumber")) then ' Amazon Pay
	function calculateStringToSignV2()
		calculateStringToSignV2="POST" & vbLf & scripturl & vbLf & endpointpath & vbLf & amazonstr
	end function
	function amazonparam2(nam, val)
		amazonstr=amazonstr & IIfVs(amazonstr<>"","&") & nam & "=" & replace(rawurlencode(replaceaccents(val)),"%7E","~")
	end function
	ordID=getpost("ordernumber")
	if getpayprovdetails(21,data1,data2,data3,demomode,ppmethod) then
		success=TRUE
		alreadyprocessed=FALSE
		scripturl="mws-eu.amazonservices.com"
		if origCountryCode="US" then scripturl="mws.amazonservices.com"
		if origCountryCode="JP" then scripturl="mws.amazonservices.jp"
		endpointpath="/OffAmazonPayments" & IIfVs(demomode,"_Sandbox") & "/2013-01-01"
		endpoint="https://" & scripturl & endpointpath
		
		data2arr=split(data2,"&",2)
		if UBOUND(data2arr)>=0 then data2=data2arr(0)
		if UBOUND(data2arr)>0 then sellerid=data2arr(1)
		amazonstr=""

		itemtotal=0
		sSQL="SELECT ordAuthNumber,ordShipping,ordStateTax,ordCountryTax,ordHSTTax,ordHandling,ordTotal,ordDiscount,ordAuthNumber,ordTransID,ordEmail,ordStatus FROM orders WHERE ordPayProvider=21 AND ordTransID='" & escape_string(getpost("amzrefid")) & "' AND ordID=" & escape_string(ordID)
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			email=rs("ordEmail")
			itemtotal=(rs("ordShipping")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordHSTTax")+rs("ordTotal")+rs("ordHandling"))-rs("ordDiscount")
			ordStatus=rs("ordStatus")
			ordTransID=rs("ordTransID")
			alreadyprocessed=ordStatus>=3
		else
			success=FALSE
			errtext="The Order ID could not be found."
		end if
		rs.close
		
		if alreadyprocessed then
			call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
		elseif NOT success then
			call order_failed()
		else
			call amazonparam2("AWSAccessKeyId",data2)
			call amazonparam2("Action","SetOrderReferenceDetails")
			call amazonparam2("AmazonOrderReferenceId",getpost("amzrefid"))
			call amazonparam2("OrderReferenceAttributes.OrderTotal.Amount",FormatNumber(itemtotal,2,-1,0,0))
			call amazonparam2("OrderReferenceAttributes.OrderTotal.CurrencyCode",countryCurrency)
			call amazonparam2("OrderReferenceAttributes.SellerOrderAttributes.SellerOrderId",ordID)
			call amazonparam2("SellerId",sellerid)
			call amazonparam2("SignatureMethod","HmacSHA256")
			call amazonparam2("SignatureVersion",2)
			call amazonparam2("Timestamp",getutcdate(0))
			call amazonparam2("Version","2013-01-01")
			call amazonparam2("Signature",b64_hmac_sha256(data3,calculateStringToSignV2()))

			if NOT callxmlfunction(endpoint,amazonstr,res,"","WinHTTP.WinHTTPRequest.5.1",errtext,FALSE) then success=FALSE
			
			if success then
				amazonstr=""

				call amazonparam2("AWSAccessKeyId",data2)
				call amazonparam2("Action","ConfirmOrderReference")
				call amazonparam2("AmazonOrderReferenceId",getpost("amzrefid"))
				call amazonparam2("SellerId",sellerid)
				call amazonparam2("SignatureMethod","HmacSHA256")
				call amazonparam2("SignatureVersion",2)
				call amazonparam2("Timestamp",getutcdate(0))
				call amazonparam2("Version","2013-01-01")
				call amazonparam2("Signature",b64_hmac_sha256(data3,calculateStringToSignV2()))

				if NOT callxmlfunction(endpoint,amazonstr,res,"","WinHTTP.WinHTTPRequest.5.1",errtext,FALSE) then success=FALSE

				amazonstr=""
				
				call amazonparam2("AWSAccessKeyId",data2)
				call amazonparam2("Action","GetOrderReferenceDetails")
				call amazonparam2("AmazonOrderReferenceId",getpost("amzrefid"))
				call amazonparam2("SellerId",sellerid)
				call amazonparam2("SignatureMethod","HmacSHA256")
				call amazonparam2("SignatureVersion",2)
				call amazonparam2("Timestamp",getutcdate(0))
				call amazonparam2("Version","2013-01-01")
				call amazonparam2("Signature",b64_hmac_sha256(data3,calculateStringToSignV2()))
			end if
			if success then
				if callxmlfunction(endpoint,amazonstr,res,"","WinHTTP.WinHTTPRequest.5.1",errtext,FALSE) then
					ordEmail="" : ordName="" : ordAddress="" : ordAddress2="" : ordAddress3="" : ordCity="" : ordPhone=""
					set xmlDoc=Server.CreateObject("MSXML2.DOMDocument")
					xmlDoc.validateOnParse=FALSE
					xmlDoc.loadXML(res)
					Set nodeList=xmlDoc.getElementsByTagName("GetOrderReferenceDetailsResponse")
					Set nl=nodeList.Item(0)
					for ix1=0 to nl.childNodes.length - 1
						Set tx1=nl.childNodes.Item(ix1)
						if tx1.nodeName="Error" then
							for ix2=0 To tx1.childNodes.length - 1
								set tx2=tx1.childNodes.Item(ix2)
								if tx2.nodeName="Message" then
									errtext="Amazon Error: " & tx2.firstChild.nodeValue
									success=FALSE
								end if
							next
						elseif tx1.nodeName="GetOrderReferenceDetailsResult" then
							for ix2=0 To tx1.childNodes.length - 1
								set tx2=tx1.childNodes.Item(ix2)
								if tx2.nodeName="OrderReferenceDetails" then
									for ix3=0 To tx2.childNodes.length - 1
										set tx3=tx2.childNodes.Item(ix3)
										if tx3.nodeName="Constraints" then
											for ix4=0 To tx3.childNodes.length - 1
												set tx4=tx3.childNodes.Item(ix4)
												if tx4.nodeName="Constraint" then
													for ix5=0 To tx4.childNodes.length - 1
														set tx5=tx4.childNodes.Item(ix5)
														if tx5.nodeName="Description" then
															errtext=tx5.firstChild.nodeValue
															success=FALSE
														end if
													next
												end if
											next
										elseif tx3.nodeName="Destination" then
											for ix4=0 To tx3.childNodes.length - 1
												set tx4=tx3.childNodes.Item(ix4)
												if tx4.nodeName="PhysicalDestination" then
													for ix5=0 To tx4.childNodes.length - 1
														set tx5=tx4.childNodes.Item(ix5)
														if tx5.nodeName="Name" then
															ordName=tx5.firstChild.nodeValue
														elseif tx5.nodeName="AddressLine1" then
															ordAddress=tx5.firstChild.nodeValue
														elseif tx5.nodeName="AddressLine2" then
															ordAddress2=trim(tx5.firstChild.nodeValue&" "&ordAddress2)
														elseif tx5.nodeName="AddressLine3" then
															ordAddress2=trim(ordAddress2&" "&tx5.firstChild.nodeValue)
														elseif tx5.nodeName="City" then
															ordCity=tx5.firstChild.nodeValue
														elseif tx5.nodeName="StateOrRegion" then
															ordState=tx5.firstChild.nodeValue
														elseif tx5.nodeName="PostalCode" then
															ordZip=tx5.firstChild.nodeValue
														elseif tx5.nodeName="Phone" then
															ordPhone=tx5.firstChild.nodeValue
														elseif tx5.nodeName="Email" then
															ordEmail=tx5.firstChild.nodeValue
														end if
													next
												end if
											next
										elseif tx3.nodeName="Buyer" then
											for ix4=0 To tx3.childNodes.length - 1
												set tx4=tx3.childNodes.Item(ix4)
												if tx4.nodeName="Email" then ordEmail=tx4.firstChild.nodeValue
											next
										end if
									next
								end if
							next
						end if
					next
				else
					success=FALSE
				end if
				pendingreason=""
				authcode=""
				if success then
					success=FALSE
					amazonstr=""
					currdatetime=FormatDateTime(now(),0)
					call amazonparam2("AWSAccessKeyId",data2)
					call amazonparam2("Action","Authorize")
					call amazonparam2("AmazonOrderReferenceId",getpost("amzrefid"))
					call amazonparam2("AuthorizationAmount.Amount",FormatNumber(itemtotal,2,-1,0,0))
					call amazonparam2("AuthorizationAmount.CurrencyCode",countryCurrency)
					call amazonparam2("AuthorizationReferenceId",ordID&"_"&replace(replace(replace(currdatetime,"/",""),":","")," ",""))
					call amazonparam2("CaptureNow",IIfVr(ppmethod=1,"false","true"))
					call amazonparam2("SellerAuthorizationNote",IIfVr(ppmethod=1,"Authorization: ","Capture: ") & currdatetime)
					call amazonparam2("SellerId",sellerid)
					call amazonparam2("SignatureMethod","HmacSHA256")
					call amazonparam2("SignatureVersion",2)
					call amazonparam2("Timestamp",getutcdate(0))
					if ppmethod<>1 then call amazonparam2("TransactionTimeout","0")
					call amazonparam2("Version","2013-01-01")
					call amazonparam2("Signature",b64_hmac_sha256(data3,calculateStringToSignV2()))
					if callxmlfunction(endpoint,amazonstr,res,"","WinHTTP.WinHTTPRequest.5.1",errtext,FALSE) then
						set xmlDoc=Server.CreateObject("MSXML2.DOMDocument")
						xmlDoc.validateOnParse=FALSE
						xmlDoc.loadXML(res)
						set nodeList=xmlDoc.getElementsByTagName("AuthorizationStatus")
						if nodeList.length > 0 then
							set nl=nodeList.Item(0)
							for ix1=0 to nl.childNodes.length - 1
								set tx1=nl.childNodes.Item(ix1)
								if tx1.nodeName="State" then
									if (ppmethod=1 AND (tx1.firstChild.nodeValue="Pending" OR tx1.firstChild.nodeValue="Open")) OR (ppmethod<>1 AND tx1.firstChild.nodeValue="Closed") then
										success=TRUE
										if tx1.firstChild.nodeValue="Pending" then pendingreason="Capture"
									end if
								end if
							next
						else
							set nodeList=xmlDoc.getElementsByTagName("Error")
							if nodeList.length > 0 then
								set nl=nodeList.Item(0)
								for ix1=0 to nl.childNodes.length - 1
									set tx1=nl.childNodes.Item(ix1)
									if tx1.nodeName="Message" then errtext=tx1.firstChild.nodeValue
								next
							end if
						end if
						if success then
							set nodeList=xmlDoc.getElementsByTagName("AmazonAuthorizationId")
							if nodeList.length > 0 then
								authcode=nodeList.Item(0).firstChild.nodeValue
							else
								success=FALSE
							end if
						end if
					end if
				end if
				if success then
					if NOT alreadyprocessed then
						if pendingreason<>"" then pendingreason="Pending: " & pendingreason
						sSQL="UPDATE cart SET cartCompleted=1 WHERE cartOrderID=" & escape_string(ordID)
						ect_query(sSQL)
						sSQL="UPDATE orders SET ordStatus=3,ordAuthNumber='" & escape_string(authcode) & "',ordName='" & escape_string(ordName) & "',ordAddress='" & escape_string(ordAddress) & "',ordAddress2='" & escape_string(ordAddress2) & "',ordCity='" & escape_string(ordCity) & "',ordPhone='" & escape_string(ordPhone) & "',ordEmail='" & escape_string(ordEmail) & "',ordAuthStatus='" & escape_string(pendingreason) & "' WHERE ordPayProvider=21 AND ordID=" & escape_string(ordID)
						ect_query(sSQL)
						if autobillingtoshipping then ect_query("UPDATE orders SET ordShipName=ordName,ordShipAddress=ordAddress,ordShipAddress2=ordAddress2,ordShipCity=ordCity,ordShipState=ordState,ordShipCountry=ordCountry,ordShipPhone=ordPhone,ordShipZip=ordZip WHERE ordPayProvider=21 AND ordID=" & escape_string(ordID))
					end if
					call do_order_success(ordID,emailAddr,NOT alreadyprocessed,TRUE,NOT alreadyprocessed,NOT alreadyprocessed,NOT alreadyprocessed)
					SESSION("AmazonLogin")=""
					SESSION("AmazonLoginTimeout")=""
				else
					call order_failed()
				end if
			end if
		end if
	end if
elseif getget("method")="stripe" AND getget("sid")<>"" AND is_numeric(getget("soid")) then
	ordID=getget("soid")
	if SESSION("stripeid")<>getget("sid") OR NOT is_numeric(ordID) then
		errtext="Stripe Session Error."
		call order_failed()
	elseif getpayprovdetails(23,data1,data2,data3,demomode,ppmethod) then
		paymentintent=""
		sSQL="SELECT ordID,ordStatus,ordTransID FROM orders WHERE ordPayProvider=23 AND ordStatus>=3 AND ordID=" & ordID
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			success=TRUE
			paymentintent=rs("ordTransID")
		else
			success=FALSE
			errtext="The order has not yet been authorized. Please wait a few moments then refresh this browser window."
		end if
		rs.close
		if success then
			xmlfnheaders=array(array("User-Agent","Stripe/v1 RubyBindings/1.12.0"),array("Authorization","Bearer "&data1),array("Content-Type","application/x-www-form-urlencoded"))
			if callxmlfunction("https://api.stripe.com/v1/payment_intents/"&paymentintent,"",jres,"","Msxml2.ServerXMLHTTP", errtext, FALSE) then
				paymentstatus=get_json_val(jres,"status","payment_method_types")
				if (ppmethod<>1 AND paymentstatus="succeeded") OR (ppmethod=1 AND paymentstatus="requires_capture") then
					call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
				else
					errtext="Payment Status: " & paymentstatus
					success=FALSE
				end if
			else
				success=FALSE
			end if
		end if
		if NOT success then
			call order_failed()
		end if
	else
		errtext="Payment method not set."
		call order_failed()
	end if
elseif getpost("pprov")="23" AND getpost("stripeToken")<>"" AND getpost("stripeEmail")<>"" AND is_numeric(getpost("ordernumber")) then
	ordID=getpost("ordernumber")
	if getpayprovdetails(23,data1,data2,data3,demomode,ppmethod) then
		amount=0
		success=TRUE
		isstripedotcom=TRUE ' For callxmlfunction
		ordStatus=3
		ordTransID="xxxxx"
		alreadyprocessed=FALSE
		token=getpost("stripeToken")
		chargeid=""

		sSQL="SELECT ordAuthNumber,ordShipping,ordStateTax,ordCountryTax,ordHSTTax,ordHandling,ordTotal,ordDiscount,ordAuthNumber,ordTransID,ordEmail,ordStatus FROM orders WHERE ordPayProvider=23 AND ordID=" & ordID
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			email=rs("ordEmail")
			amount=(rs("ordShipping")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordHSTTax")+rs("ordTotal")+rs("ordHandling"))-rs("ordDiscount")
			ordStatus=rs("ordStatus")
			ordTransID=rs("ordTransID")
		else
			success=FALSE
			errtext="The Order ID could not be found."
		end if
		rs.close
		if ordStatus>=3 then
			if ordTransID=token then
				alreadyprocessed=TRUE
			else
				success=FALSE
				errtext="The Stripe.com Order Has Already Been Processed."
			end if
		elseif success then
			sXML="amount="&vsround(amount*100,0)&"&currency="&lcase(countryCurrency)&"&card="&token&"&capture="&IIfVr(ppmethod=1,"false","true")&"&description="&email&"&metadata[order_id]="&ordID
			xmlfnheaders=array(array("User-Agent","Stripe/v1 RubyBindings/1.12.0"),array("Authorization","Bearer "&data1),array("Content-Type","application/x-www-form-urlencoded"))
			success=callxmlfunction("https://api.stripe.com/v1/charges",sXML,xmlres,"","Msxml2.ServerXMLHTTP", errtext, FALSE)
			if success then
				idpos=instr(xmlres,"""id"":")
				if idpos>0 then
					startpos=instr(idpos+6,xmlres,"""")+1
					endpos=instr(startpos,xmlres,"""")
					chargeid=mid(xmlres,startpos,endpos-startpos)
				end if
			else
				idpos=instr(xmlres,"""message"":")
				if idpos>0 then
					startpos=instr(idpos+10,xmlres,"""")+1
					endpos=instr(startpos,xmlres,"""")
					errtext=mid(xmlres,startpos,endpos-startpos)
				end if
			end if
		end if
		if success then
			if NOT alreadyprocessed then
				sSQL="UPDATE cart SET cartCompleted=1 WHERE cartOrderID=" & ordID
				ect_query(sSQL)
				sSQL="UPDATE orders SET ordStatus=3,ordAuthStatus='',ordAuthNumber='" & escape_string(left(chargeid,48)) & "',ordTransID='" & escape_string(left(getpost("stripeToken"),48)) & "' WHERE ordPayProvider=23 AND ordID=" & ordID
				ect_query(sSQL)
			end if
			call do_order_success(ordID,emailAddr,NOT alreadyprocessed,TRUE,NOT alreadyprocessed,NOT alreadyprocessed,NOT alreadyprocessed)
		else
			call order_failed_htmldisp(FALSE)
		end if
	else
		errtext="Payment method not set."
		call order_failed()
	end if
elseif getget("method")="anethosted" AND is_numeric(getget("ordid")) then
	ordID=getget("ordid")
	if getpayprovdetails(3,data1,data2,data3,demomode,ppmethod) then
		fingerprint=UCASE(calcHMACSha512(data3,data1 & "^" & ordID & "^" & adminSecret & "^" & data2 & "^","TEXT","TEXT"))
		if getget("fp")=fingerprint then
			sSQL="SELECT ordID FROM orders WHERE ordPayProvider=3 AND ordStatus>=3 AND ordAuthNumber<>'' AND ordID=" & ordID
			rs.open sSQL,cnn,0,1
			paypalwaitipn=rs.EOF
			rs.close
			if paypalwaitipn then
				xxThkErr=""
				ect_query("UPDATE cart SET cartCompleted=2 WHERE cartCompleted=0 AND cartOrderID=" & ordID)
				ect_query("UPDATE orders SET ordAuthNumber='no ipn' WHERE ordAuthNumber='' AND ordPayProvider=3 AND ordID=" & ordID)
				errtext=xxPPWIPN
				call order_failed()
			else
				' errtext=$GLOBALS['xxPPTOUT']
				call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
			end if
		else
			errtext="Fingerprint mismatch"
			call order_failed()
		end if
	end if
elseif getget("ectprnm")="wpconfirm" AND is_numeric(getget("ordid")) AND is_numeric(getget("pprov")) then
	if NOT getpayprovdetails(getget("pprov"),data1,data2,data3,demomode,ppmethod) then
		errtext="Payment method not set."
		call order_failed()
	else
		ordID=getget("ordid")
		rethash=getget("rethash")
		sSQL="SELECT ordSessionID FROM orders WHERE ordID=" & escape_string(ordID)
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then sessionid=rs("ordSessionID") else sessionid="xxx"
		rs.close
		ourhash=ucase(calcmd5(ordID&"WPCONFHash"&getget("pprov")&sessionid&"1234"&adminSecret))
		if ourhash=rethash then
			call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
		else
			errtext="Hash values do not match"
			call order_failed()
		end if
	end if
elseif getget("PNREF")<>"" AND getget("SECURETOKEN")<>"" AND getget("SECURETOKENID")<>"" then ' PayPal Advanced
	txn_id=getget("PPREF")
	ordID=getget("INVOICE")
	print "<script>if(window!=top)top.location.href=location.href</script>"
	success=FALSE
	if getget("RESULT")="0" AND is_numeric(ordID) then
		if txn_id<>"" then
			sSQL="SELECT ordAuthNumber FROM orders WHERE ordPayProvider=22 AND ordStatus>=3 AND ordAuthNumber='" & escape_string(txn_id) & "' AND ordID=" & escape_string(ordID)
			rs.open sSQL,cnn,0,1
				if NOT rs.EOF then success=(rs("ordAuthNumber")<>"")
			rs.close
		end if
		if success then
			call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
		else
			sSQL="SELECT ordAuthNumber,ordShipping,ordStateTax,ordCountryTax,ordHSTTax,ordHandling,ordTotal,ordDiscount,ordAuthNumber,ordEmail FROM orders WHERE ordPayProvider=8 AND ordID=" & escape_string(ordID)
			rs.open sSQL,cnn,0,1			
			ispayflowtxn=FALSE
			if NOT rs.EOF then
				amount=formatnumber((rs("ordShipping")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordHSTTax")+rs("ordTotal")+rs("ordHandling"))-rs("ordDiscount"),2,-1,0,0)
				ispayflowtxn=TRUE
			end if
			rs.close
			if ispayflowtxn AND getpayprovdetails(8,data1,data2,data3,demomode,ppmethod) then
				vsdetails=Split(data1, "&")
				if UBOUND(vsdetails) > 0 then
					vs1=vsdetails(0)
					vs2=vsdetails(1)
					vs3=vsdetails(2)
					vs4=vsdetails(3)
				end if
				sXML="TRXTYPE=I&TENDER=C&PARTNER="&vs3&"&VENDOR="&vs2&"&USER="&vs1&"&PWD="&vs4&"&SECURETOKEN="&getget("SECURETOKEN")&"&SECURETOKENID="&getget("SECURETOKENID")&"&VERBOSITY=HIGH"
				success=callxmlfunction("https://" & IIfVr(demomode, "pilot-", "") & "payflowpro.paypal.com", sXML, curString, "", "WinHTTP.WinHTTPRequest.5.1", errormsg, FALSE)
				resparr=split(curString,"&")
				AUTHCODE="" : RESPMSG="" : AVSADDR="" : AVSZIP="" : CVV2MATCH="" : RESULT="" : TRANSTIME="" : AMT=""
				for each objItem in resparr
					itemarr=split(objItem,"=")
					if itemarr(0)="AUTHCODE" then AUTHCODE=itemarr(1)
					if itemarr(0)="RESPMSG" then errtext=itemarr(1)
					if itemarr(0)="AVSADDR " then AVSADDR =itemarr(1)
					if itemarr(0)="AVSZIP" then AVSZIP=itemarr(1)
					if itemarr(0)="CVV2MATCH" then CVV2MATCH=itemarr(1)
					if itemarr(0)="RESULT" then RESULT=itemarr(1)
					if itemarr(0)="TRANSTIME" then TRANSTIME=itemarr(1)
					if itemarr(0)="AMT" then AMT=itemarr(1)
				next
				if AUTHCODE="" then AUTHCODE=txn_id
				session.lcid=1033
				if isdate(TRANSTIME) then daysago=date()-datevalue(cdate(TRANSTIME)) else daysago=999
				session.lcid=savelcid
				if RESULT="0" AND AUTHCODE<>"" AND daysago<=1 AND abs(amount-AMT)<=0.01 then
					alreadysentemail=TRUE
					rs.open "SELECT ordStatus FROM orders WHERE ordID=" & ordID,cnn,0,1
					if NOT rs.EOF then alreadysentemail=rs("ordStatus")>=3
					rs.close
					if NOT alreadysentemail then
						ect_query("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
						ect_query("UPDATE orders SET ordStatus=3,ordAVS='"&escape_string(AVSADDR&AVSZIP)&"',ordCVV='"&escape_string(CVV2MATCH)&"',ordAuthNumber='"&escape_string(AUTHCODE)&"',ordTransID='"&escape_string(getget("PNREF"))&"' WHERE ordPayProvider=8 AND ordID="&ordID)
					end if
					call do_order_success(ordID,emailAddr,sendEmail AND NOT alreadysentemail,TRUE,NOT alreadysentemail,NOT alreadysentemail,NOT alreadysentemail)
				else
					call order_failed()
				end if
			else
				ect_query("UPDATE cart SET cartCompleted=2 WHERE cartCompleted=0 AND cartOrderID=" & ordID)
				ect_query("UPDATE orders SET ordAuthNumber='no ipn' WHERE ordAuthNumber='' AND ordPayProvider=22 AND ordID=" & ordID)
				xxThkErr=""
				if lcase(getrequest("PENDINGREASON"))="pending" then
					errtext=errtext&xxPPPend
				else
					sSQL="SELECT ordID FROM orders WHERE ordPayProvider=22 AND ordAuthNumber='no ipn' AND ordID=" & ordID
					rs.open sSQL,cnn,0,1
					paypalwaitipn=NOT rs.EOF
					rs.close
					if paypalwaitipn then errtext=xxPPWIPN else errtext=xxPPTOUT
				end if
				call order_failed_htmldisp(FALSE)
			end if
		end if
	else
		errtext=urldecode(getget("RESPMSG"))
		call order_failed()
	end if
elseif paypalhostedsolution AND getget("tx")<>"" then
	if NOT getpayprovdetails(18,data1,data2,data3,demomode,ppmethod) then
		errtext="Payment method not set."
		call order_failed
	else
		sXML="PWD=" & data2 & "&USER=" & data1 & IIfVr(data3<>"","&SIGNATURE="&data3,"") & "&METHOD=GetTransactionDetails&VERSION=84.0&TRANSACTIONID=" & getget("tx")
		if callxmlfunction("https://api-3t." & IIfVs(demomode,"sandbox.") & "paypal.com/nvp", sXML, sQuerystring, "", "Msxml2.ServerXMLHTTP", errtext, FALSE) then
			sParts=split(sQuerystring, "&")
			iParts=UBound(sParts) - 1
			pending_reason=""
			for i=0 to iParts
				aParts=split(sParts(i), "=", 2)
				sKey=aParts(0)
				sValue=aParts(1)
				select case sKey
				case "ACK"
					success=(sValue="Success")
				case "PAYMENTSTATUS"
					payment_status=sValue
				case "PENDINGREASON"
					pending_reason=sValue
				case "CUSTOM"
					ordID=replace(sValue,"'","")
				case "TRANSACTIONID"
					txn_id=replace(sValue,"'","")
				case "L_LONGMESSAGE0"
					errtext=urldecode(sValue)
				end select
			next
			if success then
				sSQL="SELECT ordAuthNumber FROM orders WHERE ordPayProvider=1 AND ordStatus>=3 AND ordAuthNumber='"&txn_id&"' AND ordID=" & ordID
				success=(txn_id<>"")
				rs.open sSQL,cnn,0,1
					if rs.EOF then success=FALSE
				rs.close
				if success then
					call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
				else
					ect_query("UPDATE cart SET cartCompleted=2 WHERE cartCompleted=0 AND cartOrderID="&ordID)
					ect_query("UPDATE orders SET ordAuthNumber='no ipn' WHERE ordAuthNumber='' AND ordPayProvider=1 AND ordID="&ordID)
					xxThkErr=""
					if lcase(payment_status)="pending" then
						errtext=errtext&xxPPPend
					else
						sSQL="SELECT ordID FROM orders WHERE ordPayProvider=1 AND ordAuthNumber='no ipn' AND ordID=" & ordID
						rs.open sSQL,cnn,0,1
						paypalwaitipn=NOT rs.EOF
						rs.close
						if paypalwaitipn then errtext=xxPPWIPN else errtext=xxPPTOUT
					end if
					call order_failed_htmldisp(FALSE)
				end if
			else
				call order_failed
			end if
		end if
	end if
' elseif getget("amt")<>"" AND getget("tx")<>"" AND getget("st")<>"" AND getget("cc")<>"" AND getget("cm")<>"" then
elseif getget("amt")<>"" AND getget("tx")<>"" AND getget("st")<>"" AND getget("cc")<>"" then
	ordID=0
	if NOT getpayprovdetails(1,data1,data2,data3,demomode,ppmethod) then
		errtext="Payment method not set."
		call order_failed
	elseif data2="" then
		errtext="Identity token for PayPal Payment Data Transfer (PDT) not set."
		call order_failed
	else
		tx=getget("tx")
		if instr(tx, ",")>0 then txarr=split(tx,",") : tx=trim(txarr(0))
		xmlfnheaders=array(array("Content-Type","application/x-www-form-urlencoded"),array("Host","www.paypal.com"),array("Connection","close"))
		if callxmlfunction("https://www." & IIfVs(demomode,"sandbox.") & "paypal.com/cgi-bin/webscr", "&cmd=_notify-synch&tx="&tx&"&at="&data2, sQuerystring, "", "Msxml2.ServerXMLHTTP", errtext, FALSE) then
			if mid(sQuerystring,1,7)="SUCCESS" then
				sQuerystring=mid(sQuerystring,9)
				sParts=Split(sQuerystring, vbLf)
				iParts=UBound(sParts) - 1
				pending_reason=""
				for i=0 to iParts
					aParts=split(sParts(i), "=", 2)
					sKey=aParts(0)
					sValue=aParts(1)
					select case sKey
					case "payment_status"
						payment_status=sValue
					case "pending_reason"
						pending_reason=sValue
					case "custom"
						ordID=replace(sValue,"'","")
					case "txn_id"
						txn_id=replace(sValue,"'","")
					end select
				next
				sSQL="SELECT ordAuthNumber FROM orders WHERE ordPayProvider=1 AND ordStatus>=3 AND ordAuthNumber='"&txn_id&"' AND ordID=" & ordID
				success=(txn_id<>"")
				rs.open sSQL,cnn,0,1
					if rs.EOF then success=FALSE
				rs.close
				if success then
					call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
				else
					ect_query("UPDATE cart SET cartCompleted=2 WHERE cartCompleted=0 AND cartOrderID="&ordID)
					ect_query("UPDATE orders SET ordAuthNumber='no ipn' WHERE ordAuthNumber='' AND ordPayProvider=1 AND ordID="&ordID)
					xxThkErr=""
					if lcase(payment_status)="pending" then
						errtext=errtext&xxPPPend
					else
						sSQL="SELECT ordID FROM orders WHERE ordPayProvider=1 AND ordAuthNumber='no ipn' AND ordID=" & ordID
						rs.open sSQL,cnn,0,1
						paypalwaitipn=NOT rs.EOF
						rs.close
						if paypalwaitipn then errtext=xxPPWIPN else errtext=xxPPTOUT
					end if
					call order_failed()
				end if
			else
				errtext=sQuerystring
				call order_failed
			end if
		else
			call order_failed
		end if
	end if
elseif getpost("custom")<>"" then ' PayPal
	ordID=replace(getpost("custom"),"'","")
	txn_id=replace(getpost("txn_id"),"'","")
	if NOT is_numeric(ordID) then ordID=0
	sSQL="SELECT ordAuthNumber FROM orders WHERE ordPayProvider=1 AND ordStatus>=3 AND ordAuthNumber='"&txn_id&"' AND ordID=" & ordID
	success=(txn_id<>"")
	rs.open sSQL,cnn,0,1
		if rs.EOF then success=FALSE
	rs.close
	if success then
		call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
	else
		ect_query("UPDATE cart SET cartCompleted=2 WHERE cartCompleted=0 AND cartOrderID="&ordID)
		ect_query("UPDATE orders SET ordAuthNumber='no ipn' WHERE ordAuthNumber='' AND ordPayProvider=1 AND ordID="&ordID)
		xxThkErr=""
		if getpost("payment_status")="pending" then
			errtext=errtext&xxPPPend
		else
			sSQL="SELECT ordID FROM orders WHERE ordPayProvider=1 AND ordAuthNumber='no ipn' AND ordID=" & ordID
			rs.open sSQL,cnn,0,1
			paypalwaitipn=NOT rs.EOF
			rs.close
			if paypalwaitipn then errtext=xxPPWIPN else errtext=xxPPTOUT
		end if
		call order_failed()
	end if
elseif getpost("method")="paypalexpress" AND getpost("token")<>"" then ' PayPal Express
	success=getpayprovdetails(19,username,data2pwd,data2hash,demomode,ppmethod)
	if username<>"" then username=trim(split(username,"/")(0))
	if data2pwd<>"" then data2pwd=urldecode(split(data2pwd,"&")(0))
	if data2pwd<>"" then data2pwd=trim(split(data2pwd,"/")(0))
	if data2hash<>"" then data2hash=trim(split(data2hash,"/")(0))
	if instr(username,"@AB@")<>0 then data2pwd="" : data2hash="AB"
	ordID=replace(getpost("ordernumber"),"'","")
	if NOT is_numeric(ordID) then ordID=0
	token=getpost("token")
	payerid=getpost("payerid")
	ordAuthNumber="" : status=""
	if demomode then sandbox=".sandbox" else sandbox=""
	sSQL="SELECT ordShipping,ordStateTax,ordCountryTax,ordHSTTax,ordHandling,ordTotal,ordDiscount,ordAuthNumber,ordEmail FROM orders WHERE ordID=" & ordID
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		session.LCID=1033
		amount=formatnumber((rs("ordShipping")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordHSTTax")+rs("ordTotal")+rs("ordHandling"))-rs("ordDiscount"),2,-1,0,0)
		session.LCID=saveLCID
		if rs("ordEmail")=getpost("email") then ordAuthNumber=trim(rs("ordAuthNumber")&"")
	else
		success=FALSE
	end if
	rs.close
	if success AND termsandconditions then
		if getpost("termsandconds")<>"1" then
			errtext="You must agree to our terms and conditions in order to complete your order. Please click below to go back and try again." & _
				"<br />&nbsp;<br />&nbsp;<br />&nbsp;<br /><input type=""button"" class=""ectbutton"" value=""" & xxGoBack & """ onclick=""history.go(-1)"" /><br />&nbsp;<br />&nbsp;<br />"
			success=FALSE
		end if
	end if
	if success then
		if ordAuthNumber="" then
			sXML=ppsoapheader(username, data2pwd, data2hash) & _
				"<soap:Body>" & _
				"  <DoExpressCheckoutPaymentReq xmlns=""urn:ebay:api:PayPalAPI"">" & _
				"    <DoExpressCheckoutPaymentRequest>" & _
				"      <Version xmlns=""urn:ebay:apis:eBLBaseComponents"">60.00</Version>" & _
				"      <DoExpressCheckoutPaymentRequestDetails xmlns=""urn:ebay:apis:eBLBaseComponents"">" & _
				"        <PaymentAction>" & IIfVr(ppmethod=1, "Authorization", "Sale") & "</PaymentAction>" & _
				"        <Token>" & token & "</Token><PayerID>" & payerid & "</PayerID>" & _
				"        <PaymentDetails>" & _
				"          <OrderTotal currencyID=""" & countryCurrency & """>" & amount & "</OrderTotal>" & _
				"          <ButtonSource>ecommercetemplates_Cart_EC_US</ButtonSource>" & _
				"    <NotifyURL>" & storeurl & "vsadmin/ppconfirm.asp</NotifyURL>" & _
				"        </PaymentDetails>" & _
				"      </DoExpressCheckoutPaymentRequestDetails>" & _
				"    </DoExpressCheckoutPaymentRequest>" & _
				"  </DoExpressCheckoutPaymentReq>" & _
				"</soap:Body></soap:Envelope>"
			if callxmlfunction("https://api" & IIfVs(data2hash<>"","-3t") & sandbox & ".paypal.com/2.0/", sXML, res, IIfVr(data2hash<>"","",username), "WinHTTP.WinHTTPRequest.5.1", vsRESPMSG, FALSE) then
				set xmlDoc=Server.CreateObject("MSXML2.DOMDocument")
				xmlDoc.validateOnParse=False
				xmlDoc.loadXML(res)
				Set nodeList=xmlDoc.getElementsByTagName("SOAP-ENV:Body")
				Set n=nodeList.Item(0)
				for j=0 to n.childNodes.length - 1
					Set e=n.childNodes.Item(i)
					if e.nodeName="DoExpressCheckoutPaymentResponse" then
						for k9=0 To e.childNodes.length - 1
							Set t=e.childNodes.Item(k9)
							if t.nodeName="Token" then
								if t.firstChild.nodeValue="Success" then success=TRUE
							elseif t.nodeName="DoExpressCheckoutPaymentResponseDetails" then
								set ff=t.childNodes
								for kk=0 to ff.length - 1
									set gg=ff.item(kk)
									if gg.nodeName="PaymentInfo" then
										set hh=gg.childNodes
										for ll=0 to hh.length - 1
											set ii=hh.item(ll)
											if ii.nodeName="PaymentStatus" then
												status=ii.firstChild.nodeValue
											elseif ii.nodeName="PendingReason" then
												pendingreason=ii.firstChild.nodeValue
											elseif ii.nodeName="TransactionID" then
												txn_id=ii.firstChild.nodeValue
											end if
										next
									end if
								next
							elseif t.nodeName="Errors" then
								set ff=t.childNodes
								for kk=0 to ff.length - 1
									set gg=ff.item(kk)
									if gg.nodeName="ShortMessage" then
										'errtext=gg.firstChild.nodeValue & "<br>" & errtext
									elseif gg.nodeName="LongMessage" then
										errtext=errtext & gg.firstChild.nodeValue
									elseif gg.nodeName="ErrorCode" then
										errcode=gg.firstChild.nodeValue
										errtext="(" & errcode & ") " & errtext
										if errcode="10486" OR errcode="10422" then
											response.redirect "https://www" & sandbox & ".paypal.com/webscr?cmd=_express-checkout&token=" & token
											print "<p align=""center"">" & xxAutFo & "</p>"
											print "<p align=""center"">" & xxForAut & " <a class=""ectlink"" href=""https://www" & sandbox & ".paypal.com/webscr?cmd=_express-checkout&token=" & token & """>" & xxClkHere & "</a></p>"
										end if
									end if
								next
							end if
						next
					end if
				next
			else
				success=FALSE
			end if
		else
			status="Refresh"
		end if
		if status="Completed" OR status="Pending" then
			if pendingreason="authorization" then pendingreason="Capture"
			if status="Pending" AND pendingreason<>"" then pendingreason="Pending: " & pendingreason else pendingreason=""
			ect_query("UPDATE cart SET cartCompleted=1 WHERE cartOrderID=" & replace(ordID,"'",""))
			ect_query("UPDATE orders SET ordStatus=3,ordAuthNumber='" & txn_id & "',ordAuthStatus='" & escape_string(pendingreason) & "',ordAddInfo='" & escape_string(getpost("ordAddInfo")) & "' WHERE ordPayProvider=19 AND ordID=" & replace(ordID,"'",""))
			call do_order_success(ordID,emailAddr,sendEmail,TRUE,TRUE,TRUE,TRUE)
		elseif status="Refresh" then
			call do_order_success(ordID,emailAddr,sendEmail,TRUE,FALSE,FALSE,FALSE)
		else
			call order_failed()
		end if
	else
		call order_failed_htmldisp(FALSE)
	end if
elseif getpost("txn_id")<>"" AND getpost("payer_id")<>"" AND getpost("payment_gross")<>"" then ' PayPal (Now not returning "custom" parameter
	txn_id=replace(getpost("txn_id"),"'","")
	sSQL="SELECT ordID,ordAuthNumber,ordStatus FROM orders WHERE ordDate>="&vsusdate(DateAdd("h",-36,Now()))&" AND ordPayProvider=1 AND ordAuthNumber='"&escape_string(txn_id)&"'"
	success=FALSE
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		success=rs("ordStatus")>=3
		ordID=rs("ordID")
	end if
	rs.close
	if success then
		call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
	else
		xxThkErr=""
		if getpost("payment_status")="pending" then
			errtext=errtext&xxPPPend
		else
			paypalwaitipn=TRUE
			errtext=xxPPWIPN
		end if
		call order_failed()
	end if
elseif getget("ncretval")<>"" AND getget("ncsessid")<>"" then ' NOCHEX
	ordID=trim(replace(getget("ncretval"),"'",""))
	ncsessid=trim(replace(getget("ncsessid"),"'",""))
	if NOT is_numeric(ordID) then ordID=0
	sSQL="SELECT ordAuthNumber FROM orders WHERE ordPayProvider=6 AND ordStatus>=3 AND ordSessionID='"&escape_string(ncsessid)&"' AND ordID=" & ordID
	success=TRUE
	rs.open sSQL,cnn,0,1
		if rs.EOF then success=FALSE
	rs.close
	if success then
		call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
	else
		ect_query("UPDATE cart SET cartCompleted=2 WHERE cartCompleted=0 AND cartOrderID="&ordID)
		ect_query("UPDATE orders SET ordAuthNumber='no apc' WHERE ordAuthNumber='' AND ordPayProvider=6 AND ordID="&ordID)
		errtext=xxNoCnf
		xxThkErr=""
		showclickreload=TRUE
		call order_failed
	end if
elseif getpost("xxpreauth")<>"" then
	ordID=replace(getpost("xxpreauth"),"'","")
	thesessionid=replace(getpost("thesessionid"),"'","")
	themethod=replace(getpost("xxpreauthmethod"),"'","")
	if NOT is_numeric(ordID) then ordID=0
	if NOT is_numeric(themethod) then themethod=0
	success=getpayprovdetails(themethod,data1,data2,data3,demomode,ppmethod)
	if success then
		success=FALSE
		sSQL="SELECT ordAuthNumber FROM orders WHERE ordSessionID='"&escape_string(thesessionid)&"' AND ordID=" & ordID
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then success=(trim(rs("ordAuthNumber")&"")<>"")
		rs.close
	end if
	if success then
		call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
	else
		call order_failed
	end if
elseif is_numeric(getget("sid")) AND is_numeric(getget("merchant_order_id")) AND is_numeric(getget("total")) AND getget("order_number")<>"" then ' 2Checkout Transaction
	if getget("credit_card_processed")="Y" then
		ordID=replace(getget("merchant_order_id"),"'","")
		if NOT is_numeric(ordID) then ordID=0
		success=getpayprovdetails(2,acctno,md5key,data3,demomode,ppmethod)
		keysmatch=TRUE
		if md5key<>"" then
			theirkey=getget("key")
			ourkey=trim(UCase(calcmd5(md5key&acctno&IIfVr(demomode,"1",getget("order_number"))&getget("total"))))
			if ourkey=theirkey then keysmatch=true else keysmatch=false
		end if
		if success AND keysmatch then
			sSQL="SELECT ordAuthNumber,((ordShipping+ordStateTax+ordCountryTax+ordHSTTax+ordTotal+ordHandling)-ordDiscount) AS ordGndTot FROM orders WHERE ordID=" & ordID
			rs.open sSQL,cnn,0,1
				if rs.EOF then success=FALSE else success=TRUE : ordGndTot=rs("ordGndTot")
			rs.close
			ect_query("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
			if ordGndTot<>cdbl(getget("total")) then ordAuthStatus="Pending: Total Paid " & getget("total") else ordAuthStatus=""
			ect_query("UPDATE orders SET ordStatus=3,ordAuthNumber='"&escape_string(getget("order_number"))&"',ordAuthStatus='"&ordAuthStatus&"' WHERE ordPayProvider=2 AND ordID="&ordID)
			call order_success(ordID,emailAddr,sendEmail)
		else
			call order_failed
		end if
	else
		call order_failed
	end if
elseif getpost("CUSTID")<>"" AND (getpost("AUTHCODE")<>"" OR getpost("PNREF")<>"") then ' PayFlow Link
	success=getpayprovdetails(8,data1,data2,data3,demomode,ppmethod)
	if success AND getpost("RESULT")="0" then
		ordID=replace(getpost("CUSTID"),"'","")
		if NOT is_numeric(ordID) then ordID=0
		ect_query("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
		ect_query("UPDATE orders SET ordStatus=3,ordAVS='"&escape_string(getpost("AVSDATA"))&"',ordCVV='"&escape_string(getpost("CSCMATCH"))&"',ordAuthNumber='"&escape_string(getpost("AUTHCODE"))&"' WHERE ordPayProvider=8 AND ordID="&ordID)
		call order_success(ordID,emailAddr,sendEmail)
	else
		errtext=getpost("RESPMSG")
		call order_failed
	end if
elseif getpost("emailorder")<>"" OR getpost("secondemailorder")<>"" then
	ordGndTot=1
	if emailorderstatus<>"" then ordStatus=emailorderstatus else ordStatus=3
	if getpost("emailorder")<>"" then
		ordID=replace(getpost("emailorder"),"'","")
		ppid=4
	else
		ordID=replace(getpost("secondemailorder"),"'","")
		ppid=17
	end if
	thesessionid=replace(getpost("thesessionid"),"'","")
	alreadysentemail=TRUE
	success=is_numeric(ordID)
	if success then
		rs.open "SELECT ordStatus,ordAuthNumber,((ordShipping+ordStateTax+ordCountryTax+ordHSTTax+ordTotal+ordHandling)-ordDiscount) AS ordGndTot FROM orders WHERE ordSessionID='"&escape_string(thesessionid)&"' AND ordID=" & ordID,cnn,0,1
		if rs.EOF then success=FALSE else alreadysentemail=rs("ordStatus")>=3  : ordGndTot=rs("ordGndTot")
		rs.close
	end if
	if success AND recaptchaenabled(256) then
		success=checkrecaptcha(errtext)
		if NOT success then ordID=0 : errtext="reCAPTCHA failure. If you inadvertently refreshed this page, your order has been logged in our system.<br />" & errtext
	end if
	if NOT alreadysentemail AND success then
		session.LCID=1033
		sSQL="SELECT payProvShow FROM payprovider WHERE (payProvEnabled=1 OR "&ordGndTot&"=0) AND payProvID="&ppid
		session.LCID=saveLCID
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			authnumber=rs("payProvShow")
			if ordGndTot=0 then ' Check if it was a gift cert
				sSQL="SELECT gcaGCID FROM giftcertsapplied WHERE gcaOrdID=" & ordID
				rs2.Open sSQL,cnn,0,1
				if NOT rs2.EOF then authnumber=xxGifCtc
				rs2.Close
			end if
			if authnumber="" then authnumber="Email"
		else
			success=FALSE
		end if
		rs.close
	end if
	if success AND alreadysentemail then
		call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
	elseif success then
		ect_query("UPDATE cart SET cartDateAdded=" & vsusdate(DateAdd("h",dateadjust,Now()))&",cartCompleted=1 WHERE cartCompleted<>1 AND cartOrderID="&ordID)
		ect_query("UPDATE orders SET ordStatus="&ordStatus&",ordAuthStatus='',ordAuthNumber='"&escape_string(left(authnumber,48))&"',ordDate="&vsusdatetime(DateAdd("h",dateadjust,Now()))&" WHERE ordAuthNumber='' AND (ordPayProvider="&ppid&" OR (ordTotal-ordDiscount)<=0) AND ordID="&ordID)
		call order_success(ordID,emailAddr,sendEmail)
	else
		call order_failed_htmldisp(FALSE)
	end if
elseif getget("OrderID")<>"" AND getget("TransRefNumber")<>"" then ' PSiGate
	sSQL="SELECT payProvID,payProvData1 FROM payprovider WHERE payProvEnabled=1 AND payProvID=11 OR payProvID=12"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then success=TRUE else success=FALSE
	rs.close
	if getget("Approved")<>"APPROVED" then success=FALSE
	if getget("CustomerRefNo") <> left(calcmd5(getget("OrderID")&":"&secretword),24) then success=FALSE
	if success then
		ordID=trim(replace(getget("OrderID"),"'",""))
		if NOT is_numeric(ordID) then ordID=0
		ect_query("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
		ect_query("UPDATE orders SET ordStatus=3,ordAuthStatus='',ordAVS='"&escape_string(getget("AVSResult")&"/"&getget("IPResult"))&"',ordCVV='"&escape_string(getget("CardIDResult"))&"',ordAuthNumber='"&escape_string(getget("CardAuthNumber"))&"',ordTransID='"&escape_string(getget("CardRefNumber"))&"' WHERE (ordPayProvider=11 OR ordPayProvider=12) AND ordID="&ordID)
		call order_success(ordID,emailAddr,sendEmail)
	else
		errtext=getget("ErrMsg")
		call order_failed
	end if
elseif getpost("ponumber")<>"" AND (getpost("approval_code")<>"" OR getpost("failReason")<>"") then ' Linkpoint
	ordID=replace(getpost("ponumber"),"'","")
	ordIDa=split(ordID,".")
	ordID=ordIDa(0)
	if is_numeric(ordID) AND getpayprovdetails(16,data1,data2,data3,demomode,ppmethod) then
		theauthcode=replace(getpost("approval_code"),"'","")
		thesuccess=lcase(getpost("status"))
		if (thesuccess="approved" OR thesuccess="submitted") AND theauthcode<>"" then
			autharr=split(theauthcode,":")
			if autharr(0)="Y" AND UBOUND(autharr) >= 3 then
				theauthcode=autharr(1)
				theavscode=autharr(2)
				sSQL="SELECT ordID FROM orders WHERE ordAuthNumber='"&left(theauthcode,6)&"' AND ordPayProvider=16 AND ordID=" & ordID
				rs.open sSQL,cnn,0,1
				orderexists=NOT rs.EOF
				rs.close
				if orderexists then
					call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
				else
					ect_query("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
					ect_query("UPDATE orders SET ordStatus=3,ordAuthStatus='',ordAVS='"&escape_string(left(theavscode,3))&"',ordCVV='"&escape_string(right(theavscode,1))&"',ordAuthNumber='"&left(theauthcode,6)&"',ordTransID='"&right(theauthcode,len(theauthcode)-6)&"' WHERE ordPayProvider=16 AND ordID="&ordID)
					call order_success(ordID,emailAddr,sendEmail)
				end if
			else
				errtext="Invalid auth code"
				call order_failed
			end if
		else
			errtext=getpost("failReason")
			errtextarr=split(errtext, ":")
			if IsArray(errtextarr) then
				if UBOUND(errtextarr)>0 then errtext=errtextarr(1)
			end if
			call order_failed
		end if
	else
		call order_failed
	end if
elseif getpost("oid")<>"" AND getpost("response_hash")<>"" AND getpost("txndate_processed")<>"" then ' Linkpoint 2.0
	ordID=getpost("oid")
	ordIDa=split(ordID,".")
	ordID=ordIDa(0)
	theauthcode=replace(getpost("approval_code"),"'","")
	lphash=getpost("response_hash")
	if is_numeric(ordID) AND getpayprovdetails(16,data1,data2,data3,demomode,ppmethod) then
		sSQL="SELECT ordPrivateStatus FROM orders WHERE ordID=" & ordID
		rs.open sSQL,cnn,0,1
		if rs.EOF then txndatetime="" else txndatetime=rs("ordPrivateStatus")
		rs.close
		str=data3 & theauthcode & getpost("chargetotal") & "840" & txndatetime & data1
		hex_str=""
		for i=1 to len(str)
			hex_str=hex_str + lcase(cstr(hex(asc(mid(str, i, 1)))))
		next
		ourhash=SHA256(hex_str)
		if ourhash<>lphash then
			errtext="Invalid Response Hash"
			call order_failed
		elseif ucase(getpost("status"))="APPROVED" OR ucase(getpost("status"))="SUBMITTED" then
			autharr=split(theauthcode,":")
			if autharr(0)="Y" AND UBOUND(autharr) >= 3 then
				theauthcode=autharr(1)
				theavscode=autharr(2)
				sSQL="SELECT ordID FROM orders WHERE ordAuthNumber='"&left(theauthcode,6)&"' AND ordPayProvider=16 AND ordID=" & ordID
				rs.open sSQL,cnn,0,1
				orderexists=NOT rs.EOF
				rs.close
				if orderexists then
					call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
				else
					ect_query("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
					ect_query("UPDATE orders SET ordStatus=3,ordAuthStatus='',ordAVS='"&escape_string(left(theavscode,3))&"',ordCVV='"&escape_string(right(theavscode,1))&"',ordAuthNumber='"&left(theauthcode,6)&"',ordTransID='"&right(theauthcode,len(theauthcode)-6)&"' WHERE ordPayProvider=16 AND ordID="&ordID)
					call order_success(ordID,emailAddr,sendEmail)
				end if
			else
				errtext="Invalid auth code"
				call order_failed
			end if
		else
			errtext=getpost("fail_reason")
			call order_failed
		end if
	else
		call order_failed
	end if
elseif left(getget("crypt"),1)="@" AND getpayprovdetails(24,data1,data2,data3,demomode,ppmethod) then ' SagePay
	function DecryptData(bytIn, bytPassword)
		Dim bytMessage(), bytKey(15), bytOut(), bytTemp(15), bytLast(15), lCount, lLength, lEncodedLength, bytLen(3), lPosition
		
		If Not IsInitialized(bytIn) Then Exit Function
		If Not IsInitialized(bytPassword) Then Exit Function
		
		lEncodedLength = UBound(bytIn) + 1
		If lEncodedLength Mod 16 <> 0 Then
			Exit Function
		End If
		
		For lCount = 0 To UBound(bytPassword)
			bytKey(lCount) = bytPassword(lCount)
			If lCount = 15 Then
				Exit For
			End If
		Next
		gentables
		gkey 4, 4, bytKey

		CopyBytesASP bytLast,0,bytPassword,0,16
		ReDim bytOut(lEncodedLength - 1)
		For lCount = 0 To lEncodedLength - 1 Step 16
			CopyBytesASP bytTemp, 0, bytIn, lCount, 16
			Decrypt bytTemp
			XORBlock bytTemp,bytLast
			CopyBytesASP bytLast, 0, bytIn, lCount, 16
			CopyBytesASP bytOut, lCount, bytTemp, 0, 16
		Next

		lLength = ubound(bytOut)
		ReDim bytMessage(lLength)
		CopyBytesASP bytMessage, 0, bytOut, 0, lLength+1
		DecryptData = bytMessage
	end function
	function AESDecrypt(sCypher, sPassword)
		Dim bytIn(), bytOut, bytPassword(), lCount, lLength,sTemp
		
		lLength = Len(sCypher)
		ReDim bytIn(lLength/2-1)
		for lCount = 0 To lLength/2-1
			bytIn(lCount) = CByte("&H" & Mid(sCypher,lCount*2+1,2))
		next
		
		lLength = Len(sPassword)
		ReDim bytPassword(lLength-1)
		for lCount = 1 To lLength
			bytPassword(lCount-1) = CByte(AscB(Mid(sPassword,lCount,1)))
		next

		bytOut = DecryptData(bytIn, bytPassword)
		lLength = UBound(bytOut) + 1 - bytOut(UBound(bytOut))
		sTemp = ""
		for lCount = 0 To lLength - 1
			sTemp = sTemp & Chr(bytOut(lCount))
		next
		AESDecrypt = sTemp
	end function
	public function getToken(thisString,thisToken)
		Dim Tokens, subString
		Tokens = Array("Status","StatusDetail","VendorTxCode","VPSTxId","TxAuthNo","AVSCV2","Amount","AddressResult","GiftAid")
		if instr(thisString,thisToken+"=")=0 then
			getToken=""
		else
			subString=mid(thisString,instr(thisString,thisToken)+len(thisToken)+1)
			i=0
			do while i<9
				if Tokens(i)<>thisToken then
					if instr(subString,"&"+Tokens(i))<>0 then 
						substring=left(substring,instr(subString,"&"+Tokens(i))-1)
					end if
				end if
				i = i +1
			loop
			getToken=subString
		end if
	end function
	Decoded=AESDecrypt(mid(getget("crypt"),2),data2)
	if Decoded<>"" then
		ordID=getToken(Decoded,"VendorTxCode")
		Status=getToken(Decoded,"Status")
		StatusDetail=getToken(Decoded,"StatusDetail")
		VPSTxId=getToken(Decoded,"VPSTxId")
		TxAuthNo=getToken(Decoded,"TxAuthNo")
		Amount=getToken(Decoded,"Amount")
		AVSCV2=getToken(Decoded,"AVSCV2")
		currorderstat=0
		if AVSCV2="ALL DATA MATCHED" OR AVSCV2="ALL DATA MATCH" OR AVSCV2="ALL MATCH" then
			ordAVS="Y"
			ordCVV="Y"
		elseif AVSCV2="SECURITY CODE MATCH ONLY" then
			ordAVS="N"
			ordCVV="Y"
		elseif AVSCV2="ADDRESS MATCH ONLY" then
			ordAVS="Y"
			ordCVV="N"
		elseif AVSCV2="NO DATA MATCHES" then
			ordAVS="N"
			ordCVV="N"
		elseif AVSCV2="DATA NOT CHECKED" then
			ordAVS="X"
			ordCVV="X"
		else
			ordAVS="E"
			ordCVV="E"
		end if
	end if
	if ordID<>"" AND (TxAuthNo<>"" OR Status="AUTHENTICATED") AND (Status="OK" OR Status="AUTHENTICATED") then
		theorderarray = split(ordID, "-")
		ordID = replace(theorderarray(0), "'", "")
		sSQL="SELECT ordStatus FROM orders WHERE ordPayProvider=24 AND ordID=" & ordID
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then currorderstat = int(rs("ordStatus"))
		rs.Close
		if currorderstat=2 then
			ordAuthStatus=""
			sSQL="UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID
			cnn.Execute(sSQL)
			if TxAuthNo="" then
				ordAuthStatus="Pending: Settlement"
				TxAuthNo="UNSETTLED"
			end if
			sSQL="UPDATE orders SET ordStatus=3,ordAVS='"&ordAVS&"',ordCVV='"&ordCVV&"',ordTransID='"&replace(VPSTxId, "'", "")&"',ordAuthNumber='"&replace(TxAuthNo, "'", "")&"',ordAuthStatus='"&ordAuthStatus&"' WHERE ordPayProvider=24 AND ordID="&ordID
			cnn.Execute(sSQL)
			Call order_success(ordID,emailAddr,sendEmail)
		else
			Call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
		end if
	else
		errtext=StatusDetail
		call order_failed
	end if
elseif(getget("AccessCode")<>"" AND getpayprovdetails(30,data1,data2,data3,demomode,ppmethod)) then ' eWay
	xmlfnheaders=array(array("Content-Type","application/json"),array("Authorization", "Basic " & vrbase64_encrypt(data1 & ":" & data2)))
	if callxmlfunction("https://api." & IIfVr(demomode,"sandbox.","") & "ewaypayments.com/AccessCode/"&getget("AccessCode"),"",jres,"","Msxml2.ServerXMLHTTP",errormsg,TRUE) then
		authcode=get_json_val(jres,"AuthorisationCode","")
		ordID=get_json_val(jres,"InvoiceReference","")
		transid=get_json_val(jres,"TransactionID","")
		ordAVS=get_json_val(jres,"Address","Verification")
		ordCVV=get_json_val(jres,"CVN","Verification")
		transstatus=get_json_val(jres,"TransactionStatus","")
		respmessage=get_json_val(jres,"ResponseMessage","")
		if is_numeric(ordID) AND transstatus then
			alreadysentemail=TRUE
			sSQL="SELECT ordStatus FROM orders WHERE ordID=" & ordID
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then alreadysentemail=rs("ordStatus")>=3
			rs.close
			if NOT alreadysentemail then
				sSQL="UPDATE cart SET cartCompleted=1 WHERE cartOrderID=" & ordID
				ect_query(sSQL)
				sSQL="UPDATE orders SET ordStatus=3,ordAVS='" & escape_string(ordAVS) & "',ordCVV='" & escape_string(ordCVV) & "',ordTransID='" & escape_string(transid) & "',ordAuthNumber='" & escape_string(authcode) & "',ordAuthStatus='' WHERE ordPayProvider=30 AND ordID=" & ordID
				ect_query(sSQL)
			end if
			call order_success(ordID,emailAddr,sendEmail)
		else
			errtext=xxCCErro
			call order_failed()
		end if
	else
		call order_failed()
	end if
elseif getget("OrdNo")<>"" AND getget("ErrMsg")<>"" then ' PSiGate Error Reporting
	errtext=getget("ErrMsg")
	call order_failed
elseif getget("method")="globalpayments" then ' Global Payments
	if getpayprovdetx(32,data1,data2,data3,data4,data5,data6,ppflag1,ppflag2,ppflag3,ppbits,demomode,ppmethod) then
		ordID=getget("ORDER_ID")
		AUTHCODE=getget("AUTHCODE")
		RESULT=getget("RESULT")
		hashstring=getget("TIMESTAMP")&"."&data1&"."&ordID&"."&RESULT&"."&getget("MESSAGE")&"."&getget("PASREF")&"."&AUTHCODE
		hashstring=hex_sha1(hex_sha1(hashstring)&"."&data2)
		ordIDarr=split(ordID,"-")
		ordID=ordIDarr(0)
		if NOT is_numeric(ordID) then
			errtext="Invalid Order ID"
			call order_failed()
		elseif hashstring<>getget("SHA1HASH") then
			errtext="Invalid Hash Value"
			call order_failed()
		else
			isauthorized=TRUE
			sSQL="SELECT ordStatus FROM orders WHERE ordID=" & ordID
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then isauthorized=rs("ordStatus")>=3
			rs.close
			if isauthorized then
				call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
			end if
		end if
	else
		errtext="Global Payments Not Enabled"
		call order_failed()
	end if
else
	if mobilebrowser AND request.querystring.count=0 AND request.form.count=0 AND getpayprovdetails(1,data1,data2,data3,demomode,ppmethod) then
		print "<div class=""ppmobilereturn0"" style=""text-align:center;margin:30px""><img src=""images/paypalacceptmark.gif"" alt=""PayPal"" /></div><div class=""ppmobilereturn1"" style=""text-align:center;margin:30px"">Your PayPal order details have now been received.</div>"
		print "<div class=""ppmobilereturn2"" style=""text-align:center;margin:30px"">For those that have paid by PayPal, please check your email inbox for your receipt and order details.</div>"
	else
		if getpayprovdetails(14,data1,data2,data3,demomode,ppmethod) then
%>
<!--#include file="customppreturn.asp"-->
<%		else
			errtext="No matching payment system"
			call order_failed
		end if
	end if
end if
if success AND (googleanalyticsinfo=TRUE OR usegoogleuniversal) AND is_numeric(ordID) then
	if ordID<>0 then
		session.LCID=1033
		' Order ID, Affiliation, Total, Tax, Shipping, City, State, Country
		googleanalyticstrackorderinfo="ga('ecommerce:addTransaction',{'id':'"&ordID&"','affiliation':'"&affilID&"','revenue':'"&ordTotal&"','shipping':'"&ordShipping&"','tax':'"&(ordStateTax+ordHSTTax+ordCountryTax)&"'});" & vbCrLf
		sSQL="SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity,"&getlangid("sectionName",256)&",pSKU FROM cart INNER JOIN (products INNER JOIN sections ON products.pSection=sections.sectionID) ON cart.cartProdID=products.pID WHERE cartOrderID="&ordID&" ORDER BY cartID"
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			sSQL="SELECT SUM(coPriceDiff) AS coPrDff FROM cartoptions WHERE coCartID="&rs("cartID")
			rs2.Open sSQL,cnn,0,1
			coPriceDiff=0
			if NOT rs2.EOF then
				if NOT IsNull(rs2("coPrDff")) then coPriceDiff=cDbl(rs2("coPrDff"))
			end if
			rs2.Close
			' Order ID, SKU, Product Name , Category, Price, Quantity
			googleanalyticstrackorderinfo=googleanalyticstrackorderinfo & "ga('ecommerce:addItem',{'id':'"&ordID&"','name':'"&jsescape(rs("cartProdName"))&"','sku':'"&jsescape(rs("cartProdID"))&"','category':'"&jsescape(rs(getlangid("sectionName",256)))&"','price':'"&(rs("cartProdPrice")+coPriceDiff)&"','quantity':'"&rs("cartQuantity")&"'});" & vbCrLf
			rs.MoveNext
		loop
		rs.close
		googleanalyticstrackorderinfo=googleanalyticstrackorderinfo & "ga('ecommerce:send');" & vbCrLf
		session.LCID=saveLCID
	end if
end if
cnn.Close
set rs=nothing
set rs2=nothing
set cnn=nothing
%>