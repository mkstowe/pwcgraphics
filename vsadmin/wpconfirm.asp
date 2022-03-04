<!--#include file="inc/md5.asp"-->
<!--#include file="db_conn_open.asp"-->
<!--#include file="inc/languagefile.asp"-->
<!--#include file="includes.asp"-->
<!--#include file="inc/incemail.asp"-->
<!--#include file="inc/incfunctions.asp"-->
<%	if wpconfirmpage="" then %>
<html>
<head>
<title>Thanks for shopping with us</title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=adminencoding%>">
<style type="text/css">
<!--
A:link{ COLOR:#FFFFFF; TEXT-DECORATION:none }
A:visited{ COLOR:#FFFFFF; TEXT-DECORATION:none }
A:active{ COLOR:#FFFFFF; TEXT-DECORATION:none }
A:hover{ COLOR:#f39000; TEXT-DECORATION:underline }
TD{ FONT-FAMILY:Verdana; FONT-SIZE:13px }
P{ FONT-FAMILY:Verdana; FONT-SIZE:13px }
-->
</style>
</head>
<%	end if
Dim rs,rs2,sSQL,orderText,custEmail,mailsystem,success,isworldpay,isauthnet,ordGrandTotal,ordID,ordAuthNumber
success=false
errtext=""
ordGrandTotal=0 : ordTotal=0 : ordStateTax=0 : ordHSTTax=0 : ordCountryTax=0 : ordShipping=0 : ordHandling=0 : ordDiscount=0
affilID="" : ordCity="" : ordState="" : ordCountry="" : ordDiscountText="" : ordEmail="" : emailtxt=""
SESSION("couponapply") = empty
SESSION("giftcerts") = empty
SESSION("cpncode") = empty
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
success=FALSE
worldpaycallbackerror=FALSE
isworldpay=FALSE
isauthnet=FALSE
isnetbanx=FALSE
issecpay=FALSE
wpconfreturl=""
function wpcheckalreadyprocessed(tordid)
	wpcheckalreadyprocessed=TRUE
	rs.open "SELECT ordStatus,ordAuthNumber FROM orders WHERE ordID=" & tordid,cnn,0,1
	if NOT rs.EOF then
		wpcheckalreadyprocessed=rs("ordStatus")>=3
		if wpcheckalreadyprocessed then ordAuthNumber=rs("ordAuthNumber")
	end if
	rs.close
end function
if trim(request.form("transStatus"))<>"" then ' WorldPay
	isworldpay=TRUE
	transstatus = trim(request.form("transStatus"))
	data2cbp = ""
	ordID = trim(replace(request.form("cartId"),"'",""))
	if getpayprovdetails(5,acctno,data2,data3,demomode,ppmethod) AND is_numeric(ordID) then
		data2arr = split(data2,"&",2)
		if UBOUND(data2arr) >= 0 then data2md5 = data2arr(0)
		if UBOUND(data2arr) > 0 then data2cbp = data2arr(1)
		if data2cbp <> "" then
			if data2cbp <> request.form("callbackPW") then
				transstatus=""
				errormsg = "Callback password incorrect"
				worldpaycallbackerror=TRUE
			end if
		end if
		if transstatus="Y" then
			avscode = trim(request.form("AVS"))
			if trim(request.form("wafMerchMessage"))<>"" then avscode = trim(request.form("wafMerchMessage")) & vbCrLf & avscode
			if NOT wpcheckalreadyprocessed(ordID) then
				cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
				cnn.Execute("UPDATE orders SET ordStatus=3,ordAVS='"&replace(avscode,"'","")&"',ordAuthNumber='"&replace(trim(request.form("transId")),"'","")&"' WHERE ordPayProvider=5 AND ordID="&ordID)
				call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
			end if
			success=TRUE
			sSQL = "SELECT ordSessionID FROM orders WHERE ordID=" & escape_string(ordID)
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then sessionid=rs("ordSessionID") else sessionid="notvalid"
			rs.close
			retprms="?ectprnm=wpconfirm&pprov=5&ordid="&ordID&"&rethash="&ucase(calcmd5(ordID&"WPCONFHash"&"5"&sessionid&"1234"&adminSecret))
			wpconfreturl=storeurlssl & "thanks.asp" & retprms
			print "<meta http-equiv=""refresh"" content=""0; URL=" & wpconfreturl & """ />"
		end if
	end if
elseif trim(request.form("x_response_code"))<>"" then ' Authorize.net
	ordID = trim(replace(request.form("x_invoice_num"),"'",""))
	if getpayprovdetails(3,data1,data2,data3,demomode,ppmethod) AND is_numeric(ordID) then
		isauthnet=TRUE
		validhash=FALSE
		if len(data3)>100 then
			fields="^" & getpost("x_trans_id") & "^" & _
				getpost("x_test_request") & "^" & _
				getpost("x_response_code") & "^" & _
				getpost("x_auth_code") & "^" & _
				getpost("x_cvv2_resp_code") & "^" & _
				getpost("x_cavv_response") & "^" & _
				getpost("x_avs_code") & "^" & _
				getpost("x_method") & "^" & _
				getpost("x_account_number") & "^" & _
				getpost("x_amount") & "^" & _
				getpost("x_company") & "^" & _
				getpost("x_first_name") & "^" & _
				getpost("x_last_name") & "^" & _
				getpost("x_address") & "^" & _
				getpost("x_city") & "^" & _
				getpost("x_state") & "^" & _
				getpost("x_zip") & "^" & _
				getpost("x_country") & "^" & _
				getpost("x_phone") & "^" & _
				getpost("x_fax") & "^" & _
				getpost("x_email") & "^" & _
				getpost("x_ship_to_company") & "^" & _
				getpost("x_ship_to_first_name") & "^" & _
				getpost("x_ship_to_last_name") & "^" & _
				getpost("x_ship_to_address") & "^" & _
				getpost("x_ship_to_city") & "^" & _
				getpost("x_ship_to_state") & "^" & _
				getpost("x_ship_to_zip") & "^" & _
				getpost("x_ship_to_country") & "^" & _
				getpost("x_invoice_num") & "^"
			hashstr=UCASE(calcHMACSha512(data3,fields,"TEXT","HEX"))
			validhash=(hashstr=ucase(getpost("x_SHA2_Hash")))
		else
			hashstr=ucase(calcmd5(trim(data3) & trim(data1) & request.form("x_trans_id") & FormatNumber(cdbl(request.form("x_amount")),2,-1,0,0)))
			validhash=(hashstr=ucase(getpost("x_MD5_Hash")))
		end if
		emailtxt = emailtxt & "MYHASH:" & hashstr & emlNl
		if trim(request.form("x_response_code"))="1" AND ordID<>"" AND validhash then
			vsAUTHCODE = trim(request.form("x_auth_code"))
			if vsAUTHCODE="" AND trim(request.form("x_method"))="ECHECK" then vsAUTHCODE="eCheck"
			if NOT wpcheckalreadyprocessed(ordID) then
				cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
				cnn.Execute("UPDATE orders SET ordStatus=3,ordAVS='"&replace(trim(request.form("x_avs_code")),"'","")&"',ordCVV='"&replace(trim(request.form("x_cvv2_resp_code")),"'","")&"',ordAuthNumber='"&replace(vsAUTHCODE,"'","")&"',ordTransID='"&replace(trim(request.form("x_trans_id")),"'","")&"' WHERE ordPayProvider=3 AND ordID="&ordID)
				call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
			end if
			success=TRUE
			sSQL = "SELECT ordSessionID FROM orders WHERE ordID=" & escape_string(ordID)
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then sessionid=rs("ordSessionID") else sessionid="notvalid"
			rs.close
			retprms="?ectprnm=wpconfirm&pprov=3&ordid="&ordID&"&rethash="&ucase(calcmd5(ordID&"WPCONFHash"&"3"&sessionid&"1234"&adminSecret))
			wpconfreturl=storeurlssl & "thanks.asp" & retprms
			print "<meta http-equiv=""refresh"" content=""0; URL=" & wpconfreturl & """ />"
		else
			if trim(data3)<>"" AND (hashstr<>ucase(request.form("x_MD5_Hash"))) then
				errormsg="Invalid Hash Value"
			else
				errormsg = request.form("x_response_code") & " (" & trim(request.form("x_response_reason_code")) & ") " & trim(request.form("x_response_reason_text"))
			end if
		end if
	end if
elseif trim(request("trans_id"))<>"" then ' Secpay / PayPoint
	if getpayprovdetails(9,data1,data2,data3,demomode,ppmethod) then
		issecpay=TRUE
		data2arr = split(data2,"&",2)
		if UBOUND(data2arr) >= 0 then data2md5 = data2arr(0)
		callbacksuccess=TRUE
		if trim(request("valid"))="true" AND trim(request("auth_code"))<>"" then
			ordID = trim(replace(request("trans_id"),"'",""))
			if trim(data2md5) <> "" then
				thehash = calcmd5("trans_id=" & ordID & "&amount=" & trim(request("amount")) & "&callback=" & storeurlssl & "vsadmin/" & IIfVr(wpconfirmpage="", "wpconfirm.asp", wpconfirmpage) & "&" & data2md5)
				if request("hash") <> thehash then callbacksuccess=FALSE
			end if
			if NOT callbacksuccess then
				errormsg = "Callback password incorrect"
			else
				if NOT wpcheckalreadyprocessed(ordID) then
					cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
					cnn.Execute("UPDATE orders SET ordStatus=3,ordAVS='" & replace(trim(request("cv2avs")),"'","") & "',ordAuthNumber='" & trim(replace(request("auth_code"),"'",""))&"' WHERE ordPayProvider=9 AND ordID="&ordID)
					call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
				else
					call do_order_success(ordID,emailAddr,sendEmail,FALSE,FALSE,FALSE,FALSE)
				end if
				success=TRUE
			end if
		else
			errormsg = trim(request("message"))
		end if
	end if
elseif trim(request.form("netbanx_reference"))<>"" then ' Netbanx
	if getpayprovdetails(15,data1,data2,data3,demomode,ppmethod) then
		isnetbanx=TRUE
		thereference = trim(request.form("netbanx_reference"))
		if trim(Request.ServerVariables("REMOTE_HOST"))<>"195.224.77.2" AND trim(Request.ServerVariables("REMOTE_HOST"))<>"80.65.254.6" then
			errormsg = "Error: This transaction does not appear to have been initiated by Netbanx"
		elseif thereference<>"0" AND trim(request.form("order_id"))<>"" then
			ordID = trim(replace(request.form("order_id"),"'",""))
			if NOT wpcheckalreadyprocessed(ordID) then
				cnn.Execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
				allchecks = "X"
				if trim(request.form("houseno_auth"))="Matched" then
					allchecks = "Y"
				elseif trim(request.form("houseno_auth"))="Not matched" then
					allchecks = "N"
				end if
				if trim(request.form("postcode_auth"))="Matched" then
					allchecks = allchecks & "Y"
				elseif trim(request.form("postcode_auth"))="Not matched" then
					allchecks = allchecks & "N"
				else
					allchecks = allchecks & "X"
				end if
				cvv = "X"
				if trim(request.form("CV2_auth"))="Matched" then
					cvv = "Y"
				elseif trim(request.form("CV2_auth"))="Not matched" then
					cvv = "N"
				end if
				cnn.Execute("UPDATE orders SET ordStatus=3,ordAVS='"&allchecks&"',ordCVV='"&cvv&"',ordAuthNumber='" & thereference &"' WHERE ordPayProvider=15 AND ordID="&ordID)
				call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
			else
				call do_order_success(ordID,emailAddr,sendEmail,FALSE,FALSE,FALSE,FALSE)
			end if
			success=TRUE
		else
			errormsg = "Transaction Declined"
		end if
	end if
end if
	if wpconfirmpage="" then
%>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#F39900">
  <tr>
    <td>
      <table width="100%" border="1" cellspacing="1" cellpadding="3">
        <tr> 
          <td rowspan="4" bgcolor="#333333">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td width="100%" bgcolor="#333333" align="center"><span style="color:#FFFFFF;font-family:Arial,Helvetica,sans-serif;font-weight:bold"><% response.write xxInAssc&"&nbsp;"
		if isworldpay then
			response.write "WorldPay"
		elseif isauthnet then
			response.write "Authorize.Net"
		elseif isnetbanx then
			response.write "Netbanx"
		elseif issecpay then
			response.write "SECPay"
		else
			response.write "<a href=""https://www.ecommercetemplates.com"">EcommerceTemplates.com</a>"
		end if %></span></td>
          <td rowspan="4" bgcolor="#333333">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
        </tr>
        <tr> 
          <td width="100%" bgcolor="#637BAD" align="center"><span style="color:#FFFFFF;font-family:Verdana,Helvetica,sans-serif;font-weight:bold;font-size:16px"><%=xxTnkStr%></span></td>
        </tr>
        <tr> 
          <td width="100%" align="center" bgcolor="#F5F5F5">
<%	end if ' wpconfirmpage
	if isworldpay then %>
<%		if worldpaycallbackerror then %>
			<p>&nbsp;</p>
			<p align="center"><span style="font-family:Verdana,Helvetica,sans-serif;font-weight:bold;font-size:12px"><%=xxTnkWit%> <WPDISPLAY ITEM=compName></span></p>
			<table width="100%" border="0" cellspacing="3" cellpadding="3" bgcolor="">
			  <tr> 
				<td width="100%" colspan="2" align="center"><%=xxThkErr%>
				<p>The error report returned by the server was:<br /><strong><%=errormsg%></strong></p>
				<a href="<%=storeUrl%>"><span style="color:#637BAD"><strong><%=xxCntShp%></strong></span></a><br />
				<p>&nbsp;</p>
				</td>
			  </tr>
			</table>
            <p><wpdisplay item="banner"></p>
			<p><span style="font-size:10px;font-weight:bold"><%=xxPlsNt1&" "&xxMerRef&" "&xxPlsNt2%></span></p>
<%		else %>
			<div style="text-align:center">You will now be forwarded to view your receipt.</div>
			<div style="text-align:center"><%=xxForAut&" <a href="""&wpconfreturl&""">"&xxClkHere&"</a>"%></div>
<%		end if %>
			<p>&nbsp;</p>
<%	elseif isauthnet AND success then %>
			<div style="text-align:center">You will now be forwarded to view your receipt.</div>
			<div style="text-align:center"><%=xxForAut&" <a href="""&wpconfreturl&""">"&xxClkHere&"</a>"%></div>
			<p>&nbsp;</p>
<%	elseif success then %>
		  <table border="0" cellspacing="0" cellpadding="0" width="98%" bgcolor="" align="center">
			<tr>
			  <td width="100%" align="center">
				<table width="80%" border="0" cellspacing="3" cellpadding="3" bgcolor="">
				  <tr> 
					<td width="100%" align="center"><%=xxThkYou%>
					</td>
				  </tr>
<%		if digidownloads=TRUE then
			response.write "</table>"
			noshowdigiordertext=TRUE
' To enable digital downloads, just add a "hash" back into the line below so it looks like this . . .
' <!--#include file="inc/digidownload.asp"-->
' If you apply an updater, you must repeat this step.
%>
<!--include file="inc/digidownload.asp"-->
<%			response.write "<table width=""80%"" border=""0"" cellspacing=""3"" cellpadding=""3"" bgcolor="""">"
		end if
%>
				  <tr> 
					<td width="100%"><%response.write Replace(orderText,vbCrLf,"<br />")%>
					</td>
				  </tr>
				  <tr> 
					<td width="100%" align="center"><br /><br />
					<%=xxRecEml%><br /><br />
					<a href="<%=storeUrl%>"><span style="color:#637BAD"><strong><%=xxCntShp%></strong></span></a><br />&nbsp;
					</td>
				  </tr>
				</table>
			  </td>
			</tr>
		  </table>
<%	else %>
		  <p>&nbsp;</p>
		  <table border="0" cellspacing="0" cellpadding="0" width="98%" bgcolor="" align="center">
			<tr>
			  <td width="100%">
				<table width="100%" border="0" cellspacing="3" cellpadding="3" bgcolor="">
				  <tr> 
					<td width="100%" colspan="2" align="center"><%=xxThkErr%>
					<p>The error report returned by the server was:<br /><strong><%=errormsg%></strong></p>
					<a href="<%=storeUrl%>"><span style="color:#637BAD"><strong><%=xxCntShp%></strong></span></a><br />
					<p>&nbsp;</p>
					</td>
				  </tr>
				</table>
			  </td>
			</tr>
		  </table>
<%	end if
googleanalyticstrackorderinfo=""
if googleanalyticsinfo=TRUE AND isnumeric(ordID) AND ordID<>"" AND NOT isworldpay AND NOT isauthnet then
	session.LCID = 1033
	' Order ID, Affiliation, Total, Tax, Shipping, City, State, Country
	googleanalyticstrackorderinfo = vbCrLf & IIfVr(usegoogleasync,"_gaq.push(['_addTrans',","pageTracker._addTrans(") & """" & ordID & """,""" & affilID & """,""" & ordTotal & """,""" & (ordStateTax+ordHSTTax+ordCountryTax) & """,""" & (ordShipping+ordHandling) & """,""" & IIfVr(usegoogleasync,replace(ordCity,"""","\""") & """,""","") & replace(ordState,"""","\""") & """,""" & replace(ordCountry,"""","\""") & """" & IIfVr(usegoogleasync,"]","") & ");" & vbCrLf
	
	sSQL = "SELECT cartProdID,cartProdName,cartProdPrice,cartQuantity,"&getlangid("sectionName",256)&",pSKU FROM cart INNER JOIN (products INNER JOIN sections ON products.pSection=sections.sectionID) ON cart.cartProdID=products.pID WHERE cartOrderID="&ordID&" ORDER BY cartID"
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		' Order ID, SKU, Product Name , Category, Price, Quantity
		googleanalyticstrackorderinfo = googleanalyticstrackorderinfo & IIfVr(usegoogleasync,"_gaq.push(['_addItem',","pageTracker._addItem(") & """" & ordID & """,""" & replace(rs("cartProdID"),"""","\""") & """,""" & replace(rs("cartProdName"),"""","\""") & """,""" & replace(rs(getlangid("sectionName",256)) & "","""","\""") & """,""" & rs("cartProdPrice") & """,""" & rs("cartQuantity") & """" & IIfVr(usegoogleasync,"]","") & ");" & vbCrLf
		rs.MoveNext
	loop
	rs.close
	googleanalyticstrackorderinfo = googleanalyticstrackorderinfo & IIfVr(usegoogleasync,"_gaq.push(['_trackTrans']);","pageTracker._trackTrans();") & vbCrLf
	session.LCID = saveLCID
end if
	if wpconfirmpage="" then %>
          </td>
        </tr>
        <tr> 
          <td width="100%" bgcolor="#333333" align="center"><span style="color:#FFFFFF;font-family:Verdana,Helvetica,sans-serif;font-weight:bold;font-size:12px"><a href="<%=storeUrl%>"><%=xxClkBck%></a></span></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%
if googleanalyticstrackorderinfo<>"" AND googleanalyticstrackid<>"" then
%>
<script>
var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>
<script>
try {
var pageTracker = _gat._getTracker("<%=googleanalyticstrackid%>");
pageTracker._trackPageview();
} catch(err) {}<%=googleanalyticstrackorderinfo%></script>
<%
end if %>
</body>
</html>
<%	end if ' wpconfirmpage
if debugmode=TRUE then
	if htmlemails=true then emlNl = "<br />" else emlNl=vbCrLf
	for each objItem In Request.Form
		emailtxt = emailtxt & objItem & " : " & Request.Form(objItem) & emlNl
	next
	Call DoSendEmailEO(emailAddr,emailAddr,"","wpconfirm.asp debug",emailtxt,emailObject,themailhost,theuser,thepass)
end if
cnn.Close
set rs = nothing
set rs2 = nothing
set cnn = nothing
%>
