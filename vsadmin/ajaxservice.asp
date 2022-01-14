<%
Response.Buffer=True
Response.Expires=60
Response.Expiresabsolute=Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl="no-cache"
'=========================================
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
%>
<!--#include file="inc/md5.asp"-->
<!--#include file="db_conn_open.asp"-->
<!--#include file="inc/languageadmin.asp"-->
<!--#include file="inc/languagefile.asp"-->
<!--#include file="includes.asp"-->
<%
Dim ordstatussubject(10),ordstatusemail(10)
on error resume next
if lcase(adminencoding)<>"utf-8" then response.codepage=65001
on error goto 0
response.charset="utf-8"
if getget("action")="imageupload" then server.scripttimeout=480
%>
<!--#include file="inc/incemail.asp"-->
<!--#include file="inc/incfunctions.asp"-->
<%
response.clear
emailtxt="" : WSP="" : OWSP=""
if debugmode then
	for each objitem in request.servervariables
		if (request.servervariables(objitem))<>"" then emailtxt=emailtxt&"SERVER:"&objitem & " : " & request.servervariables(objitem) & emlNl
	next
	for each objitem in request.form
		emailtxt=emailtxt&"FORM:"&objitem & " : " & request.form(objitem) & emlNl
	next
	for each objitem in request.querystring
		if (request.querystring(objitem))<>"" then emailtxt=emailtxt&"QUERYSTR:"&objitem & " : " & request.querystring(objitem) & emlNl
	next
end if
function getcsvline()
	getcsvline=""
	do while csvcurrpos <= csvlen
		tmpchar=mid(csvfile, csvcurrpos, 1)
		csvcurrpos=csvcurrpos+1
		if tmpchar=vbCr OR tmpchar=vbLf then exit do else getcsvline=getcsvline&tmpchar
	loop
	do while csvcurrpos <= csvlen
		tmpchar=mid(csvfile, csvcurrpos, 1)
		if tmpchar=vbCr OR tmpchar=vbLf then csvcurrpos=csvcurrpos+1 else exit do
	loop
end function
if storesessionvalue="" then storesessionvalue="virtualstore"
if NOT ((getget("SS-UserName")<>"" AND getget("SS-Password")<>"") OR getget("action")="pay360" OR getget("action")="globalpayments" OR getget("action")="squareup" OR getget("action")="imageupload" OR getget("action")="autosearch" OR getget("action")="appstatus" OR getget("action")="screlated" OR getget("action")="executeppsale" OR getget("action")="createppsale" OR getget("action")="termsandconditions" OR getget("action")="logoutaccount" OR getget("action")="loginaccount" OR getget("action")="createaccount" OR request.querystring("action")="ipnarrived" OR request.querystring("action")="notifystock" OR request.querystring("action")="clord" OR request.querystring("action")="applycert" OR request.querystring("action")="centinel" OR request.querystring("action")="centinel2") then
	if NOT disallowlogin then
<!--#include file="inc/incloginfunctions.asp"-->
	end if
	if SESSION("loggedon")<>storesessionvalue OR disallowlogin then response.redirect "login.asp"
end if
set rs=Server.CreateObject("ADODB.RecordSet")
set rs2=Server.CreateObject("ADODB.RecordSet")
set rs3=server.createobject("ADODB.RecordSet")
set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
b64pad="="
if dateadjust="" then dateadjust=0
thedate=DateAdd("h",dateadjust,Now())
function jsenc(tstr)
	jsenc=tstr
end function
if SESSION("clientLoginLevel")<>"" then minloglevel=SESSION("clientLoginLevel") else minloglevel=0
function rsbtostr(rsbinary)
	set rsbts=createobject("ADODB.Recordset")
	lenbinary=lenb(rsbinary)
	if lenbinary>0 Then
		rsbts.fields.append "mBinary", 201, lenbinary
		rsbts.open
		rsbts.addnew
		rsbts("mBinary").appendchunk rsbinary 
		rsbts.update
		rsbtostr=rsbts("mBinary")
	else  
		rsbtostr=""
	end If
end function
function saveimageupload(ordid)
	imagesperitem=1
	maximages=0
	maxfilesize=2*1024*1024
	filetypes=".bmp,.gif,.jpe,.jpeg,.jpg,.pdf,.png,.tif,.tiff"
	if uploaditemsperorder<>"" then imagesperitem=uploaditemsperorder
	if uploadmaxfilesize<>"" then maxfilesize=uploadmaxfilesize
	if uploadfiletypes<>"" then filetypes=uploadfiletypes
	sSQL="SELECT SUM(cartQuantity) AS sumquant FROM cart INNER JOIN products ON cart.cartProdID=products.pID WHERE cartOrderID=" & escape_string(ordid) & " AND products.pUpload<>0"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then maximages=rs("sumquant")*imagesperitem
	rs.close
	extension=lcase(fs.getextensionname(getpost("filename")))
	if len(extension)<3 OR instr(lcase(filetypes),"."&lcase(extension))=0 then
		saveimageupload="ILLEGALEXTENSION"
		exit function
	end if
	imgnum=0
	sSQL="SELECT COUNT(*) AS countups FROM imageuploads WHERE upOrderID=" & escape_string(ordid)
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then 
		if rs("countups")>maximages then
			saveimageupload="MAXIMAGES"
			exit function
		end if
	end if
	rs.close
	sSQL="SELECT uploadDir FROM admin WHERE adminID=1"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then imageuploaddir=trim(rs("uploadDir")&"")
	rs.close
	freeimageslot=FALSE
	do while imgnum<100 AND NOT freeimageslot
		imagefilename="i" & zeropadint(getpost("orderid"),6) & "_" & zeropadint(imgnum,2) & "." & extension
		freeimageslot=NOT fs.fileexists(imageuploaddir & "\" & imagefilename)
		imgnum=imgnum+1
	loop
	if imageuploaddir="" then
		saveimageupload="NOUPLOADDIR"
		exit function
	elseif freeimageslot then
		on error resume next
		err.number=0
		set oFile=fs.CreateTextFile(imageuploaddir & "\" & imagefilename, TRUE)
		errnum=err.number
		on error goto 0
		if errnum=0 then
			imgsrcarr=split(getpost("imgsrc"),"base64,",2)
			imgsrc=Base64Decode(imgsrcarr(1))
			if lenb(imgsrc)>maxfilesize then
				saveimageupload="MAXFILESIZE"
				exit function
			else
				comments=left(strip_tags2(getpost("comments")),txtcollen)
				oFile.Write rsbtostr(imgsrc)
				oFile.Close
				sSQL="INSERT INTO imageuploads (upOrderID,upFilename,upComments) VALUES (" & escape_string(ordid) & ",'" & escape_string(imagefilename) & "','" & escape_string(comments) & "')"
				cnn.execute(sSQL)
				saveimageupload="SUCCESS"
			end if
		else
			saveimageupload="NOOPENFILE"
		end if
	else
		saveimageupload="MAXIMAGES"
	end if
end function
if getget("SS-UserName")<>"" AND getget("SS-Password")<>"" then ' Ship Station
	retval=""
	username = getget("SS-UserName")
	password = getget("SS-Password")
	sSQL="SELECT adminShipID FROM adminshipping WHERE shipStationUser='" & escape_string(username) & "' AND shipStationPass='" & escape_string(password) & "'"
	rs.open sSQL,cnn,0,1
	loginsuccess=NOT rs.EOF
	rs.close
	if NOT loginsuccess then
		retval="Login Error"
	elseif getget("action")="shipnotify" then
		if is_numeric(getget("order_number")) then
			sSQL="SELECT orderstatussubject,orderstatussubject2,orderstatussubject3,orderstatusemail,orderstatusemail2,orderstatusemail3 FROM emailmessages WHERE emailID=1"
			rs.open sSQL,cnn,0,1
			ordstatussubject(1)=trim(rs("orderstatussubject")&"")
			ordstatusemail(1)=rs("orderstatusemail")&""
			ordstatussubject(2)=trim(rs("orderstatussubject2")&"")
			ordstatusemail(2)=rs("orderstatusemail2")&""
			ordstatussubject(3)=trim(rs("orderstatussubject3")&"")
			ordstatusemail(3)=rs("orderstatusemail3")&""
			rs.close

			shipcarrier=""
			if shipstationshippedstatus="" then shipstationshippedstatus=6
			if getget("carrier")="USPS" then shipcarrier=3
			if getget("carrier")="UPS" then shipcarrier=4
			if getget("carrier")="CanadaPost" then shipcarrier=6
			if getget("carrier")="FedEx" then shipcarrier=7
			if getget("carrier")="DHL" OR getget("carrier")="DHLGlobalMail" OR getget("carrier")="DHLCanada" then shipcarrier=9
			if getget("carrier")="AustraliaPost" then shipcarrier=10
			
			tracking_number=getget("tracking_number")
			sSQL="SELECT ordTrackNum FROM orders WHERE ordID=" & escape_string(getget("order_number"))
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if trim(rs("ordTrackNum"))<>"" then
					if instr(rs("ordTrackNum"),tracking_number)=0 then
						tracking_number=rs("ordTrackNum") & "," & tracking_number
					else
						tracking_number=rs("ordTrackNum")
					end if
				end if
			end if
			rs.close
			
			sSQL="UPDATE orders SET ordTrackNum='" & escape_string(tracking_number) & "'"
			if shipcarrier<>"" then sSQL=sSQL&",ordShipCarrier='" & escape_string(shipcarrier) & "'"
			sSQL=sSQL&" WHERE ordID='" & escape_string(getget("order_number")) & "'"
			ect_query(sSQL)
			call updateorderstatus(getget("order_number"), shipstationshippedstatus, TRUE)
		else
			retval="Illegal Order Number"
		end if
	elseif getget("action")="export" AND isdate(getget("start_date")) AND isdate(getget("end_date")) then
        set oShell = CreateObject("WScript.Shell")
        utcoffsetminutes = oShell.RegRead("HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias")
		if is_numeric(utcoffsetminutes) then utcoffset=utcoffsetminutes/60 else utcoffset=0
		startdate=cdate(getget("start_date"))
		enddate=cdate(getget("end_date"))
		startdate = DateAdd("h",dateadjust-utcoffset,startdate)
		enddate = DateAdd("h",dateadjust-utcoffset,enddate)

		Dim payproviders(50)
		sSQL="SELECT payProvID,payProvName FROM payprovider ORDER BY payProvID"
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			payproviders(rs("payProvID"))=rs("payProvName")
			rs.movenext
		loop
		rs.close

		if (adminUnits AND 3)=1 then weightunits="Pounds" else weightunits="Grams"
		sSQL="SELECT ordID,ordName,ordLastName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordExtra1,ordExtra2,ordShipExtra1,ordShipExtra2,ordCheckoutExtra1,ordCheckoutExtra2,ordShipName,ordShipLastName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,ordShipPhone,ordPayProvider,ordAuthNumber,ordTotal,ordDate,ordStateTax,ordCountryTax,ordHSTTax,ordShipping,ordHandling,ordShipType,ordDiscount,ordAffiliate,ordDiscountText,ordStatus,ordStatusDate,statPrivate,ordAddInfo,ordPrivateStatus FROM orders INNER JOIN orderstatus ON orders.ordStatus=orderstatus.statID WHERE ordDate BETWEEN " & vsusdatetime(startdate) & " AND " & vsusdatetime(enddate)
		
		' sSQL2="UPDATE orders SET ordAddInfo='" & escape_string(getget("start_date")&" <br>"&getget("end_date")&" <br>"&formatdatetime(startdate,0)&" <br>"&formatdatetime(enddate,0)) & "' WHERE ordID='3747'"
		'sSQL2="UPDATE orders SET ordAddInfo='" & escape_string(sSQL) & "' WHERE ordID='3747'"
		'cnn.execute(sSQL2)

		rs.open sSQL,cnn,0,1
		retval="<?xml version=""1.0"" encoding=""utf-8""?>" & vbLf & _
			"<Orders pages=""1"">"
		do while NOT rs.EOF
			hasshipaddress=(trim(rs("ordShipAddress"))<>"")
			sSQL="SELECT countryCode FROM countries WHERE countryName='" & escape_string(IIfVr(hasshipaddress,rs("ordShipCountry"),rs("ordCountry"))) & "'"
			rs2.open sSQL,cnn,0,1
			countryCode=""
			if NOT rs2.EOF then countryCode=rs2("countryCode")
			rs2.close
			retval=retval&"<Order>" & _
			"<OrderID><![CDATA[" & rs("ordID") & "]]></OrderID>" & _
			"<OrderNumber><![CDATA[" & rs("ordID") & "]]></OrderNumber>" & _
			"<OrderDate>" & formatdatetime(dateadd("h",utcoffset,rs("ordDate")),0) & "</OrderDate>" & _
			"<OrderStatus><![CDATA[" & left(rs("statPrivate")&"",50) & "]]></OrderStatus>" & _
			"<LastModified>" & formatdatetime(dateadd("h",utcoffset,rs("ordStatusDate")),0) & "</LastModified>" & _
			"<ShippingMethod><![CDATA[" & left(rs("ordShipType")&"",100) & "]]></ShippingMethod>" & _
			"<PaymentMethod><![CDATA[" & left(payproviders(rs("ordPayProvider")&""),50) & "]]></PaymentMethod>" & _
			"<OrderTotal>" & rs("ordTotal") & "</OrderTotal>" & _
			"<TaxAmount>" & (rs("ordStateTax")+rs("ordCountryTax")+rs("ordHSTTax")) & "</TaxAmount>" & _
			"<ShippingAmount>" & (rs("ordShipping")+rs("ordHandling")) & "</ShippingAmount>" & _
			"<CustomerNotes><![CDATA[" & left(rs("ordAddInfo")&"",1000) & "]]></CustomerNotes>" & _
			"<InternalNotes><![CDATA[" & left(rs("ordPrivateStatus")&"",1000) & "]]></InternalNotes>" & vbLf & _
			"<Customer>" & _
				"<CustomerCode><![CDATA[" & rs("ordEmail") & "]]></CustomerCode>" & _
				"<BillTo>" & _
					"<Name><![CDATA[" & left(trim(rs("ordName") & " " & rs("ordLastName")),50) & "]]></Name>" & _
					"<Company><![CDATA[" & left(rs("ordExtra1")&"",50) & "]]></Company>" & _
					"<Phone><![CDATA[" & left(rs("ordPhone")&"",50) & "]]></Phone>" & _
					"<Email><![CDATA[" & left(rs("ordEmail")&"",50) & "]]></Email>" & _
				"</BillTo>" & vbLf & _
				"<ShipTo>" & _
					"<Name><![CDATA[" & left(IIfVr(hasshipaddress,trim(rs("ordShipName") & " " & rs("ordShipLastName")),trim(rs("ordName") & " " & rs("ordLastName"))),50) & "]]></Name>" & _
					"<Company><![CDATA[" & left(IIfVr(hasshipaddress,rs("ordShipExtra1"),rs("ordExtra1"))&"",50) & "]]></Company>" & _
					"<Address1><![CDATA[" & left(IIfVr(hasshipaddress,rs("ordShipAddress"),rs("ordAddress"))&"",50) & "]]></Address1>" & _
					"<Address2><![CDATA[" & left(IIfVr(hasshipaddress,rs("ordShipAddress2"),rs("ordAddress2"))&"",50) & "]]></Address2>" & _
					"<City><![CDATA[" & left(IIfVr(hasshipaddress,rs("ordShipCity"),rs("ordCity"))&"",50) & "]]></City>" & _
					"<State><![CDATA[" & left(getstateorabbreviation(IIfVr(hasshipaddress,rs("ordShipState"),rs("ordState"))),50) & "]]></State>" & _
					"<PostalCode><![CDATA[" & left(IIfVr(hasshipaddress,rs("ordShipZip"),rs("ordZip"))&"",50) & "]]></PostalCode>" & _
					"<Country><![CDATA[" & countryCode & "]]></Country>" & _
					"<Phone><![CDATA[" & left(IIfVr(hasshipaddress,rs("ordShipPhone"),rs("ordPhone"))&"",50) & "]]></Phone>" & _
				"</ShipTo>" & vbLf & _
			"</Customer>" & vbLf & _
			"<Items>"
			sSQL="SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity,pWeight FROM cart INNER JOIN products on cart.cartProdID=products.pID WHERE cartOrderID=" & rs("ordID") & " ORDER BY cartID"
			rs2.open sSQL,cnn,0,1
			do while NOT rs2.EOF
				optionpricediff=0
				optionretval=""
				sSQL="SELECT coOptGroup,coCartOption,coPriceDiff,coWeightDiff FROM cartoptions WHERE coCartID=" & rs2("cartID")
				rs3.open sSQL,cnn,0,1
				if NOT rs3.EOF then
					optionretval="<Options>"
					do while NOT rs3.EOF
						optionretval=optionretval&"<Option>" & _
							"<Name><![CDATA[" & left(rs3("coOptGroup"),100) & "]]></Name>" & _
							"<Value><![CDATA[" & left(rs3("coCartOption"),100) & "]]></Value>" & _
							"<Weight>" & IIfVr(weightunits="Grams",rs3("coWeightDiff")*1000,rs3("coWeightDiff")) & "</Weight>" & _
							"</Option>" & vbLf
						optionpricediff=optionpricediff+rs3("coPriceDiff")
						rs3.movenext
					loop
					optionretval=optionretval&"</Options>"
				end if
				rs3.close
				retval=retval&"<Item>" & _
					"<SKU><![CDATA[" & left(rs2("cartProdID"),50) & "]]></SKU>" & _
					"<Name><![CDATA[" & left(rs2("cartProdName"),200) & "]]></Name>" & _
					"<Weight>" & IIfVr(weightunits="Grams",rs2("pWeight")*1000,rs2("pWeight")) & "</Weight>" & _
					"<WeightUnits>" & weightunits & "</WeightUnits>" & _
					"<Quantity>" & rs2("cartQuantity") & "</Quantity>" & _
					"<UnitPrice>" & (rs2("cartProdPrice")+optionpricediff) & "</UnitPrice>" & optionretval & "</Item>" & vbLf
				rs2.movenext
			loop
			rs2.close
			if rs("ordDiscount")>0 then
				retval=retval&"<Item>" & _
					"<SKU></SKU>" & _
					"<Name><![CDATA[Discounts]]></Name>" & _
					"<Quantity>1</Quantity>" & _
					"<UnitPrice>-" & rs("ordDiscount") & "</UnitPrice>" & _
					"<Adjustment>true</Adjustment>" & _
					"</Item>" & vbLf
			end if
			retval=retval&"</Items>" & _
				"</Order>"
			rs.movenext
		loop
		retval=retval&"</Orders>"
		rs.close
	else
		retval="Illegal Action"
	end if
	print retval
elseif getget("action")="pay360" then
	ordID=getget("orderid")
	if is_numeric(ordID) AND getpayprovdetails(31,data1,data2,data3,demomode,ppmethod) then
		sSQL="SELECT ordShipping,ordStateTax,ordCountryTax,ordHSTTax,ordTotal,ordHandling,ordDiscount,ordEmail,ordName,ordLastName,ordAddress,ordAddress2,ordCity,ordState,ordCountry,ordZip,ordPhone FROM orders WHERE ordPayProvider=31 AND ordStatus=2 AND ordID=" & ordID
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			sSQL="SELECT countryID,countryCode,countryCode3,loadStates FROM countries WHERE countryName='" & escape_string(rs("ordCountry")) & "'"
			rs2.open sSQL,cnn,0,1
			if NOT rs2.EOF then countryCode3=rs2("countryCode3")
			rs2.close
			grandtotal=(rs("ordShipping")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordHSTTax")+rs("ordTotal")+rs("ordHandling"))-rs("ordDiscount")
			signature=sha256(ordID & adminSecret& "pay360")
			url="https://api." & IIfVs(demomode,"mite.") & "pay360.com"
			post="{""session"":{""returnUrl"":{""url"":""" & storeurlssl & "thanks.asp?pprov=31&ordid=" & ordID & """}," & _
				"""transactionNotification"":{""url"":""" & storeurlssl & "vsadmin/stripewebhook.asp?pprov=31&signature=" & signature & """,""format"":""REST_JSON""}" & _
				"},""transaction"":{""merchantReference"":""" & ordID & """,""money"":{""amount"":{""fixed"":" & grandtotal & "},""currency"":""USD""},""deferred"":" & IIfVr(ppmethod=1,"true","false") & "}," & _
				"""customer"":{""registered"":false,""details"":{" & _
				"""name"":" & json_encode(trim(rs("ordName")&" "&rs("ordLastName"))) & ",""address"":{""line1"":" & json_encode(rs("ordAddress")) & ",""line2"":" & json_encode(rs("ordAddress2")) & ",""city"":" & json_encode(rs("ordCity")) & ",""region"":" & json_encode(rs("ordState")) & ",""postcode"":" & json_encode(rs("ordZip")) & ",""countryCode"":" & json_encode(countryCode3) & "}," & _
				"""telephone"":" & json_encode(rs("ordPhone")) & ",""emailAddress"":" & json_encode(rs("ordEmail")) & ",""ipAddress"":" & json_encode(REMOTE_ADDR) & "}}}"
			xmlfnheaders=array(array("Content-Type","application/json"),array("Authorization", "Basic "&vrbase64_encrypt(data2&":"&data3)))
			if callxmlfunction(url & "/hosted/rest/sessions/" & data1 & "/payments",post,jres,"","Msxml2.ServerXMLHTTP",errormsg,FALSE) then
				thestatus=get_json_val(jres,"status","")
				if thestatus<>"SUCCESS" then
					success=FALSE
					print "0:" & get_json_val(jres,"reasonMessage","")
				else
					redirecturl=get_json_val(jres,"redirectUrl","")
					pay360sessid=get_json_val(jres,"sessionId","")
					if redirecturl<>"" AND pay360sessid<>"" then
						sSQL="UPDATE orders SET ordTransID='',ordTransSession='" & escape_string(pay360sessid) & "' WHERE ordStatus<3 AND ordID=" & ordID
						ect_query(sSQL)
						print "1:" & redirecturl
					else
						print "0:Redirect URL not returned"
					end if
				end if
			else
				print "0:" & errormsg
			end if
		end if
		rs.close
	end if
elseif getget("action")="globalpayments" then %>
<style>
input.ectbutton,button.ectbutton{
background:#006ABA;
color:#FFF;
padding:6px 12px;
border:0;
border-radius:4px;
font-family:FontAwesome,sans-serif;
cursor:pointer;
font-weight:normal;
-webkit-appearance:none;
}
input.ectbutton:hover,button.ectbutton:hover{
background:#DDD;
color:#000;
}
</style>
<%
	errtext=""
	ordID=getpost("ORDER_ID")
	AUTHCODE=getpost("AUTHCODE")
	RESULT=getpost("RESULT")
	CVV2MATCH=getpost("CVNRESULT")
	AVSADDR=getpost("AVSADDRESSRESULT")
	AVSZIP=getpost("AVSPOSTCODERESULT")
	if getpayprovdetx(32,data1,data2,data3,data4,data5,data6,ppflag1,ppflag2,ppflag3,ppbits,demomode,ppmethod) then
		hashstring=getpost("TIMESTAMP")&"."&data1&"."&ordID&"."&RESULT&"."&getpost("MESSAGE")&"."&getpost("PASREF")&"."&AUTHCODE
		hashstring=hex_sha1(hex_sha1(hashstring)&"."&data2)
		if hashstring<>getpost("SHA1HASH") then
			errtext="Invalid Hash Value"
		elseif RESULT<>"00" then
			errtext=getpost("MESSAGE")
		else %>
<script>
function globalpaymentscont(){
	window.top.location='<%=storeurlssl%>thanks<%=extension%>?method=globalpayments&TIMESTAMP=<%=getpost("TIMESTAMP")%>&ORDER_ID=<%=ordID%>&RESULT=<%=getpost("RESULT")%>&MESSAGE=<%=urlencode(getpost("MESSAGE"))%>&PASREF=<%=getpost("PASREF")%>&AUTHCODE=<%=getpost("AUTHCODE")%>&SHA1HASH=<%=getpost("SHA1HASH")%>';
}
</script>
<%			ordIDarr=split(ordID,"-")
			ordID=ordIDarr(0)
			if NOT is_numeric(ordID) then
				errtext="Invalid Order ID"
			else
				alreadysentemail=TRUE
				sSQL="SELECT ordStatus FROM orders WHERE ordID=" & ordID
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then alreadysentemail=rs("ordStatus")>=3
				rs.close
				if NOT alreadysentemail then
					ect_query("UPDATE cart SET cartCompleted=1 WHERE cartOrderID=" & ordID)
					ect_query("UPDATE orders SET ordStatus=3,ordAVS='" & escape_string(AVSADDR & AVSZIP) & "',ordCVV='" & escape_string(CVV2MATCH) & "',ordAuthNumber='" & escape_string(AUTHCODE) & "',ordTransID='" & escape_string(getpost("PASREF")) & "' WHERE ordPayProvider=32 AND ordID=" & ordID)
					call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
				end if
				print "<div style=""text-align:center;font-family:sans-serif;margin-top:50px"">"
				print "<div style=""font-size:30px"">Global Payments</div>"
				print "<div style=""margin:50px 0;font-weight:bold"">Your transaction has been processed successfully.</div>"
				print "<div>" & imageorbutton("","Please Click Here To View Your Receipt","gpaymentscont","globalpaymentscont()",TRUE) & "</div>"
				print "</div>"
			end if
		end if
	else
		errtext="Global Payments Not Enabled"
	end if
	if errtext<>"" then %>
<script>
function globalpaymentscont(){
	window.top.location='<%=storeurlssl%>cart<%=extension%>?mode=checkout';
}
</script>
	<div style="text-align:center;font-family:sans-serif;margin-top:50px"><%
		print "<div style=""font-size:30px;margin:50px 0"">Global Payments</div>"
		print xxThkErr
		print "<div style=""margin:50px 0;font-weight:bold"">" & errtext & "</div>"
		print "<div>" & imageorbutton("","Please Click Here To Go Back And Try Again","gpaymentsfail","globalpaymentscont()",TRUE) & "</div>"
%>	</div>
<%
	end if
elseif getget("action")="squareup" then
	ordGrandTotal=0 : ordTotal=0 : ordStateTax=0 : ordHSTTax=0 : ordCountryTax=0 : ordShipping=0 : ordHandling=0 : ordDiscount=0
	affilID="" : ordCity="" : ordState="" : ordCountry="" : ordDiscountText="" : ordEmail=""
	ordID=getpost("ordernumber")
	btsessionid=getpost("sessionid")
	success=TRUE
	callxmlfunctionstatus=0
	if getpayprovdetx(28,data1,data2,data3,data4,data5,data6,ppflag1,ppflag2,ppflag3,ppbits,demomode,ppmethod) AND is_numeric(ordID) then
		grandtotal=0
		sJSON=""
		sSQL="SELECT ordShipping,ordStateTax,ordCountryTax,ordHSTTax,ordTotal,ordHandling,ordDiscount,ordEmail,ordName,ordLastName,ordAddress,ordAddress2,ordCity,ordState,ordCountry,ordZip,ordPhone FROM orders WHERE ordID=" & ordID
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			sSQL="SELECT countryCode FROM countries WHERE countryName='" & escape_string(rs("ordCountry")) & "'"
			rs2.open sSQL,cnn,0,1
			if NOT rs2.EOF then countryCode=rs2("countryCode")
			rs2.close
			grandtotal=(rs("ordShipping")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordHSTTax")+rs("ordTotal")+rs("ordHandling"))-rs("ordDiscount")
			firstname=rs("ordName") : lastname=rs("ordLastName")
			if NOT usefirstlastname then call splitname(rs("ordName"),firstname,lastname)
			sJSON="{""idempotency_key"": """ & uniqid() & """," & _
				"""amount_money"": {""amount"": " & int(grandtotal*100) & ",""currency"": """ & countryCurrency & """}," & _
				"""source_id"": " & json_encode(getpost("nonce")) & ","
			if data3<>"" then sJSON=sJSON&"""location_id"":" & json_encode(data3) & ",""token"":" & json_encode(getpost("verification")) & ","
			sJSON=sJSON&"""autocomplete"": " & IIfVr(ppmethod=1,"false","true") & "," & _
				"""billing_address"": {""first_name"": " & json_encode(firstname) & ", ""last_name"": " & json_encode(lastname) & ", ""address_line_1"": " & json_encode(rs("ordAddress")) & ", ""address_line_2"": " & json_encode(rs("ordAddress2")) & ", ""administrative_district_level_1"": " & json_encode(rs("ordState")) & ", ""country"": " & json_encode(countryCode) & ", ""postal_code"": " & json_encode(rs("ordZip")) & "}," & _
				"""buyer_email_address"": " & json_encode(rs("ordEmail")) & "," & _
				"""reference_id"": """ & ordID & """," & _
				"""note"": """ & ordID & """" & _
			"}"
		end if
		rs.close
		' print sJSON & "\n"
		xmlfnheaders=array(array("Authorization","Bearer "&data2),array("Content-Type","application/json"))
		if callxmlfunction("https://connect.squareup" & IIfVs(demomode,"sandbox") & ".com/v2/payments",sJSON,jres,"","Msxml2.ServerXMLHTTP",errormsg,FALSE) then
			txn_id=""
			status=get_json_val(jres,"status","")
			if status="COMPLETED" OR (status="APPROVED" AND ppmethod=1) then
				ordcvv=IIfVr(get_json_val(jres,"cvv_status","")="CVV_ACCEPTED",1,0)
				ordavs=IIfVr(get_json_val(jres,"avs_status","")="AVS_ACCEPTED",1,0)
				pendingreason=IIfVs(ppmethod=1,"Pending: Capture")
				txn_id=get_json_val(jres,"id","")
				order_id=get_json_val(jres,"order_id","")
				cnn.execute("UPDATE cart SET cartCompleted=1 WHERE cartOrderID=" & ordID)
				cnn.execute("UPDATE orders SET ordStatus=3,ordAuthNumber='" & escape_string(order_id) & "',ordAuthStatus='" & escape_string(pendingreason) & "',ordTransID='" & escape_string(txn_id) & "',ordAVS='" & escape_string(ordavs) & "',ordCVV='" & escape_string(ordcvv) & "' WHERE ordPayProvider=28 AND ordID=" & ordID)
				call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
				print "SUCCESS:" & txn_id
			else
				print "FAILURE:" & get_json_val(jres,"detail","")
			end if
			' print "\n".$jres;
		else
			if callxmlfunctionstatus<>0 then
				print "FAILURE:" & xxCCErro
			else
				print "FAILURE:" & errormsg
			end if
		end if
	end if
elseif getget("action")="imageget" then
	imagefilename=""
	imageuploaddir=""
	if is_numeric(getget("id")) then
		sSQL="SELECT upFilename FROM imageuploads WHERE upID=" & escape_string(getget("id"))
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then imagefilename=rs("upFilename")
		rs.close
	end if
	sSQL="SELECT uploadDir FROM admin WHERE adminID=1"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then imageuploaddir=rs("uploadDir")
	rs.close
	hasfso=TRUE
	err.number=0
	on error resume next
	set fs=server.createobject("Scripting.FileSystemObject")
	if err.number<>0 then hasfso=FALSE
	on error goto 0
	if NOT hasfso then
		print "NOFSOBJECT"
	elseif imageuploaddir<>"" AND imagefilename<>"" then
		'ob_clean();
		fullfilename=imageuploaddir & "\" & imagefilename
		extension=lcase(fs.getextensionname(imagefilename))
		sendextension="jpg"
		if extension="bmp" then
			sendextension="bmp"
		elseif extension="gif" then
			sendextension="gif"
		elseif extension="tif" OR extension="tiff" then
			sendextension="tif"
		end if
		if getget("preview")="true" then print sendextension&":"
		set objstream=Server.CreateObject("ADODB.Stream")
		objstream.type=1
		objstream.open
		objstream.loadfromfile fullfilename
		buffer=objstream.Read()
		if getget("preview")="true" then
			buffer=Base64Encode(buffer)
			response.write(buffer)
		else
			response.addheader "Content-disposition", "filename=" & imagefilename
			response.contenttype="application/octet-stream"
			response.binarywrite(buffer)
		end if
		objstream.close
		set objstream=nothing
	end if
elseif getget("action")="imageupload" then
	if is_numeric(getpost("orderid")) then
		ordSessionID="none"
		sSQL="SELECT ordSessionID FROM orders WHERE ordID=" & escape_string(getpost("orderid"))
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			ordSessionID=rs("ordSessionID")
		end if
		rs.close
		uploadcheck=hex_sha1("imageupload^" & getpost("orderid") & "^" & adminSecret & "^fromect" & ordSessionID)
		if uploadcheck<>getpost("check") then
			print "ILLEGAL"
		else
			hasfso=TRUE
			err.number=0
			on error resume next
			set fs=server.createobject("Scripting.FileSystemObject")
			if err.number<>0 then hasfso=FALSE
			on error goto 0
			if hasfso then
				print saveimageupload(getpost("orderid"))
			else
				print "NOFSOBJECT"
			end if
		end if
	else
		print "ILLEGAL"
	end if
elseif request.querystring("action")="autosearch" then
	if noautosearch<>TRUE then
		rc=0
		if maxautosearch="" then maxautosearch=14
		listtext=replace(getpost("listtext"),"[","[[]")
		listcat=getpost("listcat")
		secids=""
		if is_numeric(listcat) then
			secids=getsectionids(listcat, FALSE)
		end if
		for index=0 to 1
			if index=1 then
				if rc=0 AND len(listtext)>3 AND lcase(right(listtext,1))="s" then listtext=left(listtext,len(listtext)-1) else exit for
			end if
			sSQL="SELECT"&IIfVr(mysqlserver<>true," TOP "&maxautosearch,"")&" pID,"&getlangid("pName",1)&" FROM products INNER JOIN sections ON products.pSection=sections.sectionID WHERE sectionDisabled<="&minloglevel&" AND pDisplay<>0 AND " & IIfVs(ectsiteid<>"", "pSiteID=" & ectsiteid & " AND ") & IIfVs(nosearcharticles,"pSchemaType=0 AND ")
			sSQL=sSQL & "(" & IIfVs(NOT nosearchprodid,"pID LIKE '"&escape_string(listtext)&"%' OR ") & "pSKU LIKE '"&escape_string(listtext)&"%' OR "&getlangid("pName",1)&" LIKE '"&escape_string(listtext)&"%') "
			if secids<>"" then sSQL=sSQL & "AND pSection IN (" & secids & ") "
			sSQL=sSQL & "ORDER BY "&getlangid("pName",1)
			if mysqlserver=TRUE then sSQL=sSQL & " LIMIT 0,"&maxautosearch
			rs.open sSQL,cnn,0,1
			do while NOT rs.EOF
				print rawurlencode(rs(getlangid("pName",1)))&"&"
				rc=rc+1
				rs.movenext
			loop
			rs.close
			if rc < maxautosearch-5 then
				sSQL="SELECT"&IIfVr(mysqlserver<>true," TOP "&maxautosearch-rc,"")&" pID,"&getlangid("pName",1)&" FROM products INNER JOIN sections ON products.pSection=sections.sectionID WHERE sectionDisabled<="&minloglevel&" AND pDisplay<>0 AND " & IIfVs(ectsiteid<>"", "pSiteID=" & ectsiteid & " AND ") & IIfVs(nosearcharticles,"pSchemaType=0 AND ")
				sSQL=sSQL & "(" & IIfVs(NOT nosearchprodid,"pID LIKE '%"&escape_string(listtext)&"%' OR ") & "pSKU LIKE '%"&escape_string(listtext)&"%' OR "&getlangid("pName",1)&" LIKE '%"&escape_string(listtext)&"%') AND NOT (pID LIKE '"&escape_string(listtext)&"%' OR pSKU LIKE '"&escape_string(listtext)&"%' OR "&getlangid("pName",1)&" LIKE '"&escape_string(listtext)&"%') "
				if secids<>"" then sSQL=sSQL & "AND pSection IN (" & secids & ") "
				sSQL=sSQL & "ORDER BY "&getlangid("pName",1)
				if mysqlserver=TRUE then sSQL=sSQL & " LIMIT 0,"&maxautosearch-rc
				rs.open sSQL,cnn,0,1
				do while NOT rs.EOF
					print rawurlencode(rs(getlangid("pName",1)))&"&"
					rc=rc+1
					rs.movenext
				loop
				rs.close
			end if
		next
	end if
elseif getget("action")="appstatus" then
	newdeviceid=replace(replace(getpost("token")," ",""),"'","")
	if disallowlogin then
		print "LOGINDISABLED"
	elseif getpost("subact")="unregister" then
		sSQL="DELETE FROM devicenotifications WHERE dnID='"&escape_string(newdeviceid)&"'"
		cnn.execute(sSQL)
		print "UNREGISTERED"
	else
		dofloodcontrol=FALSE
		haspermissions=FALSE
		if getpost("subact")="register" then
			cnn.execute("DELETE FROM ajaxfloodcontrol WHERE afcAction=5 AND afcDate<" & vsusdatetime(dateadd("s",-5,now())))
			sSQL="SELECT afcID FROM ajaxfloodcontrol WHERE afcAction=5 AND (afcIP='" & escape_string(REMOTE_ADDR) & "' OR afcSession='" & escape_string(session.sessionid) & "')"
			rs.open sSQL,cnn,0,1
			dofloodcontrol=NOT rs.EOF
			rs.close
		end if
		hasgoodlogin=FALSE
		if NOT dofloodcontrol then
			if getpost("pw")<>"" then hashedpw=dohashpw(getpost("pw")) else hashedpw="            "
			sSQL="SELECT adminID FROM admin WHERE adminUser='"&escape_string(getpost("id"))&"' AND adminPassword='"&escape_string(hashedpw)&"'"
			rs.open sSQL,cnn,0,1
			hasgoodlogin=NOT rs.EOF
			if hasgoodlogin then haspermissions=TRUE
			rs.close
		end if
		if NOT hasgoodlogin AND NOT dofloodcontrol then
			sSQL="SELECT adminloginid,adminloginpermissions FROM adminlogin WHERE adminloginname='"&escape_string(getpost("id"))&"' AND adminloginpassword='"&escape_string(hashedpw)&"'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				hasgoodlogin=TRUE
				if mid(rs("adminloginpermissions"),15,1)="X" then haspermissions=TRUE
			end if
			rs.close
		end if
		if dofloodcontrol then
			print "FLOODCONTROL"
		else
			sSQL="SELECT dnID FROM devicenotifications WHERE dnID='"&escape_string(newdeviceid)&"'"
			rs.open sSQL,cnn,0,1
			hasdeviceid=NOT rs.EOF
			rs.close
			if newdeviceid="" OR NOT hasgoodlogin OR (getpost("subact")<>"register" AND NOT hasdeviceid) then
				print "INVALIDLOGIN"
			elseif NOT haspermissions then
				print "NOPERMISSIONS"
			elseif getpost("subact")="clearsalecount" then
				cnn.execute("UPDATE devicenotifications SET dnLastUpdated=" & vsusdatetime(DateAdd("h",dateadjust,Now())) & " WHERE dnID='"&escape_string(newdeviceid)&"'")
				print "DONECLEAR"
			elseif getpost("subact")="register" then
				if hasdeviceid then
					print "ALREADYREGISTERED"
				else				
					sSQL="INSERT INTO devicenotifications (dnID,dnLastUpdated) VALUES ('"&escape_string(newdeviceid)&"'," & vsusdatetime(DateAdd("h",dateadjust,Now())) & ")"
					cnn.execute(sSQL)
					print "GOODCONNECT"
				end if
			elseif getpost("subact")="getdashboard" then
				if homeordersstatus<>"" then ordersstatus=homeordersstatus else ordersstatus="3,4,5,6,7,8,9,10,11,12,13,14,15,16,17"
				' This month / This year = Orders + Total
				sSQL = " SELECT COUNT(*) AS totalorders, SUM(ordTotal-ordDiscount) AS totalvalue FROM orders WHERE ordStatus IN (" & ordersstatus & ") AND ordDate BETWEEN " & vsusdate(dateserial(year(date()),month(date()),1))&" AND " & vsusdate(date()+1)
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					if NOT isnull(rs("totalvalue")) then print rs("totalorders") & "&" & jsurlencode(FormatEuroCurrency(rs("totalvalue"))) else print "0&0"
				end if
				rs.close
				' This month / Last year = Orders + Total
				sSQL = " SELECT COUNT(*) AS totalorders, SUM(ordTotal-ordDiscount) AS totalvalue FROM orders WHERE ordStatus IN (" & ordersstatus & ") AND ordDate BETWEEN " & vsusdate(dateserial(year(date())-1,month(date()),1))&" AND " & vsusdate(dateserial(year(date())-1,month(date())+1,1))
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					if NOT isnull(rs("totalvalue")) then print "&" & rs("totalorders") & "&" & jsurlencode(FormatEuroCurrency(rs("totalvalue"))) else print "&0&0"
				end if
				rs.close

				' Last month / This year = Orders + Total
				sSQL = " SELECT COUNT(*) AS totalorders, SUM(ordTotal-ordDiscount) AS totalvalue FROM orders WHERE ordStatus IN (" & ordersstatus & ") AND ordDate BETWEEN " & vsusdate(dateserial(year(date()),month(date())-1,1))&" AND " & vsusdate(dateserial(year(date()),month(date()),1))
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					if NOT isnull(rs("totalvalue")) then print "&" & rs("totalorders") & "&" & jsurlencode(FormatEuroCurrency(rs("totalvalue"))) else print "&0&0"
				end if
				rs.close
				' Last month / Last year = Orders + Total
				sSQL = " SELECT COUNT(*) AS totalorders, SUM(ordTotal-ordDiscount) AS totalvalue FROM orders WHERE ordStatus IN (" & ordersstatus & ") AND ordDate BETWEEN " & vsusdate(dateserial(year(date())-1,month(date())-1,1))&" AND " & vsusdate(dateserial(year(date())-1,month(date()),1))
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					if NOT isnull(rs("totalvalue")) then print "&" & rs("totalorders") & "&" & jsurlencode(FormatEuroCurrency(rs("totalvalue"))) else print "&0&0"
				end if
				rs.close

				' Jan 1 to Now / This year = Orders + Total
				sSQL = " SELECT COUNT(*) AS totalorders, SUM(ordTotal-ordDiscount) AS totalvalue FROM orders WHERE ordStatus IN (" & ordersstatus & ") AND ordDate BETWEEN " & vsusdate(dateserial(year(date()),1,1))&" AND " & vsusdate(date()+1)
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					if NOT isnull(rs("totalvalue")) then print "&" & rs("totalorders") & "&" & jsurlencode(FormatEuroCurrency(rs("totalvalue"))) else print "&0&0"
				end if
				rs.close
				' Jan 1 to Now / Last year = Orders + Total
				sSQL = " SELECT COUNT(*) AS totalorders, SUM(ordTotal-ordDiscount) AS totalvalue FROM orders WHERE ordStatus IN (" & ordersstatus & ") AND ordDate BETWEEN " & vsusdate(dateserial(year(date())-1,1,1))&" AND " & vsusdate(dateserial(year(date())-1,month(date()),day(date())))
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					if NOT isnull(rs("totalvalue")) then print "&" & rs("totalorders") & "&" & jsurlencode(FormatEuroCurrency(rs("totalvalue"))) else print "&0&0"
				end if
				rs.close

				' 12 Months / This year = Orders + Total
				sSQL = " SELECT COUNT(*) AS totalorders, SUM(ordTotal-ordDiscount) AS totalvalue FROM orders WHERE ordStatus IN (" & ordersstatus & ") AND ordDate BETWEEN " & vsusdate(dateserial(year(date())-1,month(date()),day(date())))&" AND " & vsusdate(date()+1)
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					if NOT isnull(rs("totalvalue")) then print "&" & rs("totalorders") & "&" & jsurlencode(FormatEuroCurrency(rs("totalvalue"))) else print "&0&0"
				end if
				rs.close
				' 12 Months / Last year = Orders + Total
				sSQL = " SELECT COUNT(*) AS totalorders, SUM(ordTotal-ordDiscount) AS totalvalue FROM orders WHERE ordStatus IN (" & ordersstatus & ") AND ordDate BETWEEN " & vsusdate(dateserial(year(date())-2,month(date()),day(date())))&" AND " & vsusdate(dateserial(year(date())-1,month(date()),day(date())))
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					if NOT isnull(rs("totalvalue")) then print "&" & rs("totalorders") & "&" & jsurlencode(FormatEuroCurrency(rs("totalvalue"))) else print "&0&0"
				end if
				rs.close

				' All time = Orders + Total
				sSQL = " SELECT COUNT(*) AS totalorders, SUM(ordTotal-ordDiscount) AS totalvalue FROM orders WHERE ordStatus IN (" & ordersstatus & ")"
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					if NOT isnull(rs("totalvalue")) then print "&" & rs("totalorders") & "&" & jsurlencode(FormatEuroCurrency(rs("totalvalue"))) else print "&0&0"
				end if
				rs.close
				
				cnn.execute("UPDATE devicenotifications SET dnLastUpdated=" & vsusdatetime(DateAdd("h",dateadjust,Now())) & " WHERE dnID='"&escape_string(newdeviceid)&"'")
			elseif getpost("subact")="alltimetotals" then
				sSQL="SELECT COUNT(*) AS thecnt FROM orders WHERE ordStatus>=2 AND ordDate>="&vsusdate(DateAdd("h",dateadjust,Now()))
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					if NOT isnull(rs("thecnt")) then neworders=rs("thecnt")
				end if
				rs.close
				newratings=0
				sSQL="SELECT COUNT(*) AS thecnt FROM ratings WHERE rtApproved=0"
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					if NOT isnull(rs("thecnt")) then newratings=rs("thecnt")
				end if
				rs.close
				newaccounts=0
				sSQL="SELECT COUNT(*) AS thecnt FROM customerlogin WHERE clDateCreated>"&vsusdate(Date()-7)
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					if NOT isnull(rs("thecnt")) then newaccounts=rs("thecnt")
				end if
				rs.close
				newmaillist=0
				sSQL="SELECT COUNT(*) AS thecnt FROM mailinglist WHERE mlConfirmDate>"&vsusdate(Date()-7)
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					if NOT isnull(rs("thecnt")) then newmaillist=rs("thecnt")
				end if
				rs.close
				newaffiliate=0
				sSQL="SELECT COUNT(*) AS thecnt FROM affiliates WHERE affilDate>"&vsusdate(Date()-7)
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					if NOT isnull(rs("thecnt")) then newaffiliate=rs("thecnt")
				end if
				rs.close
				newgiftcert=0
				sSQL="SELECT COUNT(*) AS thecnt FROM giftcertificate WHERE gcDateCreated>"&vsusdate(Date()-7)
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					if NOT isnull(rs("thecnt")) then newgiftcert=rs("thecnt")
				end if
				rs.close
				newstocknotify=0
				if notifybackinstock then
					sSQL="SELECT COUNT(*) AS thecnt FROM notifyinstock"
					rs.open sSQL,cnn,0,1
					if NOT rs.EOF then
						if NOT isnull(rs("thecnt")) then newstocknotify=rs("thecnt")
					end if
					rs.close
				end if
				if SESSION("loginid")=0 then
					sSQL="SELECT COUNT(*) AS thecnt FROM auditlog WHERE eventSuccess=0"
					rs.open sSQL,cnn,0,1
					if NOT rs.EOF then
						if NOT isnull(rs("thecnt")) then newlogevents=rs("thecnt")
					end if
					rs.close
				end if
				print neworders & "&" & newratings & "&" & newaccounts & "&" & newmaillist & "&" & newaffiliate & "&" & newgiftcert & "&" & newstocknotify & "&" & newlogevents
				sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 20 ")&"ordID,ordName,ordLastName,ordShipName,ordShipLastName,ordReferer,ordDate,ordAddress,ordCity,ordState,ordTotal,ordAddInfo,ordStatus,statPrivate FROM orders,orderstatus WHERE statID=ordStatus AND ordStatus>1 ORDER BY ordID DESC"&IIfVs(mysqlserver=TRUE," LIMIT 0,20")
				rs.open sSQL,cnn,0,1
				do while NOT rs.EOF
					print "&"&rs("ordID")&"|"
					print jsurlencode(trim(rs("ordName")&" "&rs("ordLastName")))&"|"
					if trim(rs("ordShipName")&"")<>"" AND lcase(rs("ordShipName")&"")<>lcase(rs("ordName")&"") then
						print jsurlencode(trim(rs("ordShipName")&" "&rs("ordShipLastName")))
					end if
					print "|"&jsurlencode(formatdatetime(rs("ordDate"),2)&" "&formatdatetime(rs("ordDate"),4))&"|"&jsurlencode(rs("statPrivate"))&"|"&jsurlencode(FormatEuroCurrency(rs("ordTotal")))
					rs.movenext
				loop
				rs.close

				cnn.execute("UPDATE devicenotifications SET dnLastUpdated=" & vsusdatetime(DateAdd("h",dateadjust,Now())) & " WHERE dnID='"&escape_string(newdeviceid)&"'")
			end if
		end if
	end if
elseif getget("action")="screlated" then
	if screlatedlayout="" then screlatedlayout="productimage,productid,productname,price,description"
	customlayoutarray=split(lcase(replace(screlatedlayout," ","")),",")
	WSP=""
	get_wholesaleprice_sql()
	sSQL="SELECT pId,pSKU,pSection,"&getlangid("pName",1)&","&WSP&"pPrice,pStaticPage,pStaticURL,pDateAdded,pExemptions,pOrder,pTax,pStockByOpts,pNumRatings,pTotRating,"&getlangid("pDescription",2)&" FROM products INNER JOIN relatedprods ON products.pId=relatedprods.rpRelProdID WHERE " & IIfVs(ectsiteid<>"", "pSiteID=" & ectsiteid & " AND ") & "pDisplay<>0 AND rpProdID='"&escape_string(getpost("prodid"))&"'"
	if relatedproductsbothways=TRUE then sSQL=sSQL & " UNION SELECT pId,pSKU,pSection,"&getlangid("pName",1)&","&WSP&"pPrice,pStaticPage,pStaticURL,pDateAdded,pExemptions,pOrder,pTax,pStockByOpts,pNumRatings,pTotRating,"&getlangid("pDescription",2)&" FROM products INNER JOIN relatedprods ON products.pId=relatedprods.rpProdID WHERE " & IIfVs(ectsiteid<>"", "pSiteID=" & ectsiteid & " AND ") & "pDisplay<>0 AND rpRelProdID='"&escape_string(getpost("prodid"))&"'"
	if sSortBy<>"" then sSQL=sSQL&" ORDER BY " & sSortBy & IIfVs(isdesc," DESC")
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then
		if softcartrelatedheader<>"" then print "<h1 class=""scrheader"">" & softcartrelatedheader & "</h1>"
		do while NOT rs.EOF
			rpsmallimage=""
			sSQL="SELECT imageSrc,imageType FROM productimages WHERE imageProduct='" & rs("pId") & "' AND (imageType=0 OR imageType=1) AND imageNumber=0 ORDER BY imageType"
			rs2.open sSQL,cnn,0,1
			if NOT rs2.EOF then
				rpsmallimage=rs2("imageSrc")
			end if
			rs2.close
			thedetailslink=getdetailsurl(rs("pId"),rs("pStaticPage"),rs(getlangid("pName",1)),trim(rs("pStaticURL")&""),"","")
			startlink="<a class=""ectlink"" href="""&htmlspecials(thedetailslink)&""">"
			endlink="</a>"
			if perproducttaxrate=TRUE AND NOT IsNull(rs("pTax")) then thetax=rs("pTax") else thetax=countryTaxRate
			if instr(screlatedlayout,"price")>0 then
				sSQL="SELECT poOptionGroup,optType,optFlags,optTxtMaxLen,optAcceptChars,0 AS isDepOpt FROM prodoptions INNER JOIN optiongroup ON optiongroup.optGrpID=prodoptions.poOptionGroup WHERE poProdID='"&escape_string(rs("pID"))&"' AND NOT (poProdID='"&escape_string(giftcertificateid)&"' OR poProdID='"&escape_string(donationid)&"') ORDER BY poID"
				rs2.open sSQL,cnn,0,1
				if rs2.EOF then prodoptions="" else prodoptions=rs2.getrows
				rs2.close
			end if
			print "<div class=""scrproduct"">"
			for each layoutoption in customlayoutarray
				if layoutoption="productimage" then
					print "<div class=""scrimage"">" & IIfVs(rpsmallimage<>"",startlink & "<img class=""scrimage"" src=""" & rpsmallimage & """ style=""border:0"" alt="""&replace(strip_tags2(rs(getlangid("pName",1))),"""","&quot;")&""" />" & endlink) & "</div>" & vbLf
				elseif layoutoption="productname" then
					print "<div class=""scrprodname"">" & startlink & rs(getlangid("pName",1)) & endlink & "</div>" & vbLf
				elseif layoutoption="productid" then
					print "<div class=""scrprodid"">" & rs("pId") & "</div>" & vbLf
				elseif layoutoption="price" then
					totprice=rs("pPrice")
					if isarray(prodoptions) then
						optdiff=getoptionspricediff(thetax)
						totprice=totprice+optdiff
					end if
					print "<div class=""scrprodprice"">"
						print IIfVr(totprice=0 AND pricezeromessage<>"",pricezeromessage,FormatEuroCurrency(IIfVr(showtaxinclusive=2 AND (rs("pExemptions") AND 2)<>2, totprice+(totprice*thetax/100.0), totprice)))
						if showtaxinclusive=1 AND (rs("pExemptions") AND 2)<>2 AND totprice>0 then print " <span class=""inctax scrinctax"">" & replace(ssIncTax,"%s",FormatEuroCurrency(totprice+(totprice*thetax/100.0))) & "</span>"
					print "</div>" & vbLf
				elseif layoutoption="description" then
					shortdesc=rs(getlangid("pDescription",2))&""
					print "<div class=""scrproddescription"">" & shortdesc & "</div>" & vbLf
				elseif layoutoption="reviewstars" then
					if rs("pNumRatings")>0 then print showproductreviews(2,"scrating") else print softcartnoratings
				else
					print "UNKNOWN LAYOUT OPTION:"&layoutoption&"<br />"
				end if
			next
			print "</div>"
			rs.movenext
		loop
	end if
	rs.Close
elseif getget("action")="executeppsale" then
	if is_numeric(getget("ordid")) AND getpost("paymentID")<>"" then
		if getpayprovdetails(27,xmlfnuser,xmlfnpassword,data3,demomode,ppmethod) then
			ordID=getget("ordid")
			requestxx="{""payer_id"": """ & getpost("payerID") & """}"
			xmlfnheaders=array(array("Content-Type","application/json"),array("PayPal-Partner-Attribution-Id","ecommercetemplates_Cart_EC_US"))
			call callxmlfunction("https://api." & IIfVs(demomode,"sandbox.") & "paypal.com/v1/payments/payment/" & getpost("paymentID") & "/execute",requestxx,res,"","Msxml2.ServerXMLHTTP",errormsg,FALSE)
			relrespos=instr(res,"""related_resources""")
			if relrespos>0 then
				idpos=instr(relrespos,res,"""id""")
				if idpos>0 then
					delim1=instr(idpos+4,res,"""")
					delim2=instr(delim1+1,res,"""")
					tid=mid(res,delim1+1,delim2-delim1-1)
				end if
			end if
			if tid<>"" then
				sSQL="SELECT ordID FROM orders WHERE ordID='" & escape_string(ordID) & "' AND ordStatus<3 AND ordTransID='" & escape_string(getpost("paymentID")) & "'"
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					sSQL="UPDATE cart SET cartCompleted=1 WHERE cartOrderID='" & escape_string(ordID) & "'"
					cnn.execute(sSQL)
					sSQL="UPDATE orders SET ordStatus=3,ordAuthStatus='',ordAuthNumber='" & escape_string(left(tid,48)) & "',ordTransID='" & escape_string(left(getpost("paymentID"),48)) & "' WHERE ordPayProvider IN (27) AND ordID='" & escape_string(ordID) & "'"
					cnn.execute(sSQL)
					call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
				end if
				rs.close
			else
				messagepos=instr(res,"""message""")
				if messagepos>0 then
					delim1=instr(messagepos+4,res,"""")
					delim2=instr(delim1+1,res,"""")
					message=mid(res,delim1+1,delim2-delim1-1)
					sSQL="UPDATE orders SET ordPrivateStatus='" & escape_string(left(message,255)) & "' WHERE ordPayProvider IN (27) AND ordID='" & escape_string(ordID) & "' AND ordTransID='" & escape_string(getpost("paymentID")) & "'"
					ect_query(sSQL)
				end if
			end if
		end if
	end if
elseif getget("action")="createppsale" then
	if is_numeric(getget("ordid")) then
		if getpayprovdetails(27,xmlfnuser,xmlfnpassword,data3,demomode,ppmethod) then
			grandtotal=0
			sSQL="SELECT ordEmail,ordName,ordLastName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordPhone,ordShipping,ordStateTax,ordCountryTax,ordHSTTax,ordTotal,ordHandling,ordDiscount FROM orders WHERE ordPayProvider=27 AND ordStatus<3 AND ordID=" & escape_string(getget("ordid"))
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				grandtotal=(rs("ordShipping")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordHSTTax")+rs("ordTotal")+rs("ordHandling"))-rs("ordDiscount")
				methodname="sale"
				if ppmethod=1 then methodname="authorize"
				if ppmethod=2 then methodname="order"
				if usefirstlastname then
					first_name=rs("ordName")
					last_name=rs("ordLastName")
				else
					call splitname(rs("ordName"),first_name,last_name)
				end if
				sSQL="SELECT countryCode FROM countries WHERE countryName='" & escape_string(rs("ordCountry")) & "' OR countryName2='" & escape_string(rs("ordCountry")) & "' OR countryName3='" & escape_string(rs("ordCountry")) & "'"
				rs2.open sSQL,cnn,0,1
				if NOT rs2.EOF then countryCode=rs2("countryCode")
				rs2.close
				ordState=rs("ordState")
				if countryCode="US" OR countryCode="CA" then
					sSQL="SELECT stateAbbrev FROM states WHERE (stateCountryID=1 OR stateCountryID=2) AND stateName='" & escape_string(ordState) & "' OR stateName2='" & escape_string(ordState) & "' OR stateName3='" & escape_string(ordState) & "' OR stateAbbrev='" & escape_string(ordState) & "'"
					rs2.open sSQL,cnn,0,1
					if NOT rs2.EOF then ordState=rs2("stateAbbrev")
					rs2.close
				end if
				requestxx="{" & _
				"	""intent"":""" & methodname & """," & _
				"	""redirect_urls"":{" & _
				"		""return_url"":""" & storeurlssl & "cart" & extension & """," & _
				"		""cancel_url"":""" & storeurlssl & "cart" & extension & """" & _
				"	}," & _
				"	""payer"":{" & _
				"		""payment_method"":""paypal""," & _
				"		""payer_info"":{" & _
				"			""email"":" & json_encode(rs("ordEmail")) & "," & _
				"			""first_name"":" & json_encode(first_name) & "," & _
				"			""last_name"":" & json_encode(last_name) & "," & _
				"			""billing_address"":{" & _
				"				""line1"":" & json_encode(rs("ordAddress")) & "," & _
				"				""line2"":" & json_encode(rs("ordAddress2")) & "," & _
				"				""city"":" & json_encode(rs("ordCity")) & "," & _
				"				""postal_code"":" & json_encode(rs("ordZip")) & "," & _
				"				""phone"":" & json_encode(rs("ordPhone")) & "," & _
				"				""state"":" & json_encode(ordState) & "," & _
				"				""country_code"":" & json_encode(countryCode) & _
				"			}" & _
				"		}" & _
				"	}," & _
				"	""transactions"":[{" & _
				"		""amount"":{" & _
				"		""total"":""" & formatnumber(grandtotal,2,-1,0,0) & """," & _
				"		""currency"":""" & countryCurrency & """" & _
				"		}," & _
				"		""custom"":""" & getget("ordid") & """" & _
				"	}]" & _
				"}"
				' A certificate is required to complete client authentication
				' https://www.paypal-community.com/t5/REST-APIs/Classic-ASP-Attempting-to-generate-OAUTH/td-p/1379272
				xmlfnheaders=array(array("Content-Type","application/json"),array("PayPal-Partner-Attribution-Id","ecommercetemplates_Cart_EC_US"),array("Authorization", "Basic " & vrbase64_encrypt(xmlfnuser&":"&xmlfnpassword)))
				if callxmlfunction("https://api." & IIfVs(demomode,"sandbox.") & "paypal.com/v1/payments/payment/",requestxx,res,"","Msxml2.ServerXMLHTTP",errormsg,FALSE) then
					idpos=instr(res,"""id""")
					if idpos>0 then
						delim1=instr(idpos+4,res,"""")
						delim2=instr(delim1+1,res,"""")
						tid=mid(res,delim1+1,delim2-delim1-1)
					end if
					sSQL="UPDATE orders SET ordTransID='" & escape_string(tid) & "' WHERE ordID=" & escape_string(getget("ordid"))
					cnn.execute(sSQL)
				end if
				print res
			else
				print "Order not found"
			end if
			rs.close
		end if
	end if
elseif getget("action")="termsandconditions" then
	print "<div class=""termsandconds"">"
		print "<div style=""padding:3px;float:right;text-align:right"" class=""tncclose no-print""><a href=""#"" onclick=""return closetandc()""><img src=""images/close.gif"" style=""border:0"" alt=""" & xxClsWin & """ /></a></div>"
		print "<div id=""ecttermsandconds"" class=""termsandcondsinner"" style=""background-color:#FFF"">"
		sSQL="SELECT contentData FROM contentregions WHERE contentName='termsandconditions'"
		rs.open sSQL,cnn,0,1
		if rs.EOF then
			print "<div style=""margin:100px;text-align:center"">You can create a terms and conditions page by going to the Admin Content Regions page and clicking &quot;New Content Region&quot; and creating a new content region with &quot;termsandconditions&quot; as the Region Name.</div>"
		else
			print rs("contentData")
		end if
		rs.close
		print "</div>"
		print "<div class=""termsandcondsbutton"" style=""text-align:center;padding:20px"">"
			print "<input type=""button"" class=""no-print"" value=""" & htmlspecials(xxTncYes) & """ style=""margin:20px"" onclick=""document.getElementById('ecttnccheckbox').checked=true;closetandc()"" />"
			print "<input type=""button"" class=""no-print"" value=""" & htmlspecials(xxTncPrn) & """ style=""margin:20px"" onclick=""doprintterms()"" />"
			print "<input type=""button"" class=""no-print"" value=""" & htmlspecials(xxTncNo) & """ style=""margin:20px"" onclick=""document.getElementById('ecttnccheckbox').checked=false;closetandc()"" />"
		print "</div>"
	print "</div>"
elseif getget("action")="shipcarrierupdate" AND is_numeric(getget("updfield")) then
	sSQL="UPDATE orders SET ordShipCarrier=" & getget("updfield") & " WHERE ordID=" & getget("ordid")
	cnn.execute(sSQL)
	print "SUCCESS:shipcarrier"
elseif getget("action")="tracknumupdate" then
	sSQL="UPDATE orders SET ordTrackNum='" & escape_string(getget("updfield")) & "' WHERE ordID=" & getget("ordid")
	cnn.execute(sSQL)
	print "SUCCESS:ordTrackNum"
elseif getget("action")="invoiceupdate" then
	sSQL="UPDATE orders SET ordInvoice='" & escape_string(getget("updfield")) & "' WHERE ordID=" & getget("ordid")
	cnn.execute(sSQL)
	print "SUCCESS:ordInvoice"
elseif getget("action")="logoutaccount" then
	thesessionid=replace(getsessionid(),"'","")
	SESSION("clientID")=empty : SESSION("clientUser")=empty : SESSION("clientActions")=empty : SESSION("clientLoginLevel")=empty : SESSION("clientPercentDiscount")=empty
	sSQL="DELETE FROM tmplogin WHERE tmploginid='" & escape_string(thesessionid) & "'"
	ect_query(sSQL)
	call setacookie("WRITECLL","",-7)
	call setacookie("WRITECLP","",-7)
	print "DONELOGOUT"
elseif getget("action")="loginaccount" OR getget("action")="createaccount" then
	loginsuccess=FALSE
	thesessionid=replace(getsessionid(),"'","")
	if getget("action")="loginaccount" then
		afcaction=IIfVr(request.servervariables("HTTPS")="on",2,1)
		cnn.execute("DELETE FROM ajaxfloodcontrol WHERE afcAction="&afcaction&" AND afcDate<" & vsusdatetime(dateadd("s",-5,now())))
		sSQL="SELECT afcID FROM ajaxfloodcontrol WHERE afcAction="&afcaction&" AND (afcIP='" & escape_string(REMOTE_ADDR) & "' OR afcSession='" & escape_string(session.sessionid) & "')"
		rs.open sSQL,cnn,0,1
		dofloodcontrol=NOT rs.EOF
		rs.close
		if dofloodcontrol then
			print "ERROR=1<div>"&xxFloCon&". "&xxMuWaLo&".</div><div>"&xxPlTrAg&" <span id=""fclitspan"">"&xxInSecs&"</span>.</div>"
		elseif getget("lc")="" OR getget("lc")<>sha256(adminSecret & "ect admin login" & session.sessionid) then
			print "ERROR=3Hash check failed."
		else
			clientEmail=cleanupemail(getpost("email"))
			clientPW=dohashpw(getpost("pw"))
			sSQL="SELECT clID,clUserName,clActions,clLoginLevel,clPercentDiscount FROM customerlogin WHERE (clEmail<>'' AND clEmail='"&escape_string(clientEmail)&"' AND clPW='"&escape_string(clientPW)&"') OR (clEmail='' AND clUserName='"&escape_string(clientEmail)&"' AND clPW='"&escape_string(clientPW)&"')"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				SESSION("clientID")=rs("clID")
				SESSION("clientUser")=rs("clUserName")
				SESSION("clientActions")=rs("clActions")
				SESSION("clientLoginLevel")=rs("clLoginLevel")
				SESSION("clientPercentDiscount")=(100.0-cdbl(rs("clPercentDiscount")))/100.0
				call setacookie("WRITECLL",clientEmail,IIfVr(getpost("licook")="ON",365,0))
				call setacookie("WRITECLP",clientPW,IIfVr(getpost("licook")="ON",365,0))
				print "SUCCESS:"&rs("clLoginLevel")
				loginsuccess=TRUE
			else
				print "ERROR=2" & xxNoLogD
			end if
			rs.close
			sSQL="INSERT INTO ajaxfloodcontrol (afcAction,afcIP,afcSession,afcDate) VALUES ("&afcaction&",'" & escape_string(REMOTE_ADDR) & "','" & escape_string(session.sessionid) & "'," & vsusdatetime(now()) & ")"
			ect_query(sSQL)
		end if
	elseif getget("action")="createaccount" then
		loginsuccess=TRUE
		if recaptchaenabled(8) then loginsuccess=checkrecaptcha(loginerror)
		if loginsuccess then
			loginerror=xxEmExi
			cnn.execute("DELETE FROM ajaxfloodcontrol WHERE afcAction=0 AND afcDate<" & vsusdatetime(dateadd("n",-30,now())))
			sSQL="SELECT afcID FROM ajaxfloodcontrol WHERE afcAction=0 AND (afcIP='" & escape_string(REMOTE_ADDR) & "' OR afcSession='" & escape_string(session.sessionid) & "')"
			rs.open sSQL,cnn,0,1
			dofloodcontrol=NOT rs.EOF
			rs.close
		end if
		if NOT loginsuccess then
			print "ERROR=" & loginerror
		elseif dofloodcontrol then
			loginsuccess=FALSE
			print "ERROR=Flood Control. You have already created an account in the last 30 minutes"
		else
			clientemail=getpost("email")
			clientpw=dohashpw(getpost("pw"))
			if NOT allowclientregistration then
				loginsuccess=FALSE
				loginerror="Client Registration is Disabled"
			else
				if getpost("fullname")<>"" AND clientpw<>"" AND instr(clientemail,"@")>0 AND instr(clientemail,".")>0 then
					sSQL="SELECT clID FROM customerlogin WHERE clEmail='"&escape_string(clientemail)&"'"
					rs.open sSQL,cnn,0,1
					loginsuccess=rs.EOF
					rs.close
					if NOT loginsuccess then loginerror="An account with that email address already exists."
				else
					loginsuccess=FALSE
					loginerror="Invalid login details"
				end if
			end if
			if loginsuccess AND (instr(getpost("fullname"),"<")>0 OR instr(getpost("fullname"),">")>0) then
				loginsuccess=FALSE
				loginerror="Invalid Characters in Login Name"
			end if
			if loginsuccess then
				if defaultcustomerloginlevel="" then defaultcustomerloginlevel=0
				if defaultcustomerloginactions="" then defaultcustomerloginactions=0
				if defaultcustomerlogindiscount="" then defaultcustomerlogindiscount=0 else defaultcustomerloginactions=(int(defaultcustomerloginactions) OR 16)
				if sqlserver=TRUE then
					sSQL="INSERT INTO customerlogin (clUserName,clEmail,clPw,clDateCreated,clLoginLevel,clActions,clPercentDiscount,clientCustom1,clientCustom2) VALUES ('"&escape_string(getpost("fullname"))&"','"&escape_string(clientemail)&"','"&escape_string(clientpw)&"'," & vsusdate(DateAdd("h",dateadjust,Now())) & "," & defaultcustomerloginlevel & "," & defaultcustomerloginactions & "," & defaultcustomerlogindiscount & ",'"&escape_string(strip_tags2(getpost("extraclientfield1")))&"','"&escape_string(strip_tags2(getpost("extraclientfield2")))&"')"
					ect_query(sSQL)
					rs.open "SELECT @@IDENTITY AS lstIns",cnn,0,1
					SESSION("clientID")=int(cstr(rs("lstIns")))
					rs.close
				else
					rs.open "customerlogin",cnn,1,3,&H0002
					rs.AddNew
					rs.Fields("clUserName")		= getpost("fullname")
					rs.Fields("clEmail")		= clientemail
					rs.Fields("clPw")			= clientpw
					rs.Fields("clDateCreated")	= DateAdd("h",dateadjust,Now())
					rs.Fields("clLoginLevel")	= defaultcustomerloginlevel
					rs.Fields("clActions")		= defaultcustomerloginactions
					rs.Fields("clPercentDiscount")	= defaultcustomerlogindiscount
					rs.Fields("clientCustom1")	= getpost("extraclientfield1")
					rs.Fields("clientCustom2")	= getpost("extraclientfield2")
					rs.Update
					if mysqlserver=TRUE then
						rs.close
						rs.open "SELECT LAST_INSERT_ID() AS lstIns",cnn,0,1
						SESSION("clientID")=rs("lstIns")
					else
						SESSION("clientID")=rs.Fields("clID")
					end if
					rs.close
				end if
				if getpost("allowemail")="1" then call addtomailinglist(clientemail,getpost("fullname"))
				if (adminEmailConfirm AND 4)=4 then
					emailmessage="There has been a new customer signup at your store: " & emlNl & _
						"Email: " & clientemail & emlNl & _
						"Name: " & getpost("fullname") & emlNl
					call DoSendEmailEO(emailAddr,emailAddr,clientemail,"New Customer Signup",emailmessage,emailObject,themailhost,theuser,thepass)
				end if
				SESSION("clientUser")=getpost("fullname")
				SESSION("clientActions")=defaultcustomerloginactions
				SESSION("clientLoginLevel")=defaultcustomerloginlevel
				SESSION("clientPercentDiscount")=(100.0-cdbl(defaultcustomerlogindiscount))/100.0
				sSQL="INSERT INTO ajaxfloodcontrol (afcAction,afcIP,afcSession,afcDate) VALUES (0,'" & escape_string(REMOTE_ADDR) & "','" & session.sessionid & "'," & vsusdatetime(now()) & ")"
				ect_query(sSQL)
				print "SUCCESS"
			else
				print "ERROR=" & loginerror
			end if
		end if
	end if
	if loginsuccess then
		get_wholesaleprice_sql()
		sSQL="SELECT ordID FROM orders WHERE ordStatus>1 AND ordAuthNumber='' AND " & getordersessionsql()
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then ordID=rs("ordID") else ordID=""
		rs.close
		if ordID<>"" then
			release_stock(ordID)
			ect_query("UPDATE cart SET cartSessionID='"&replace(getsessionid(),"'","")&"',cartClientID="&IIfVr(SESSION("clientID")<>"",SESSION("clientID"),0)&" WHERE cartCompleted=0 AND cartOrderID=" & ordID)
			ect_query("UPDATE orders SET ordAuthStatus='MODWARNOPEN',ordShipType='MODWARNOPEN' WHERE ordID=" & ordID)
		end if
		sSQL="SELECT cartID,cartProdID FROM cart WHERE cartCompleted=0 AND cartClientID="&replace(SESSION("clientID"),"'","")
		rs.open sSQL,cnn,0,1
			if NOT rs.EOF then cartarr=rs.getRows else cartarr=""
		rs.close
		if isarray(cartarr) then
			for index=0 to UBOUND(cartarr, 2)
				hasoptions=TRUE
				sSQL="SELECT cartID,cartQuantity FROM cart WHERE cartClientID=0 AND cartCompleted=0 AND cartSessionID='"&escape_string(thesessionid)&"' AND cartProdID='" & escape_string(cartarr(1,index)) & "'"
				rs.open sSQL,cnn,0,1
					if NOT rs.EOF then thecartid=rs("cartID") : thequant=rs("cartQuantity") else thecartid=""
				rs.close
				if thecartid<>"" then ' check options
					sSQL="SELECT coOptID,coCartOption FROM cartoptions WHERE coCartID=" & cartarr(0, index)
					rs.open sSQL,cnn,0,1
						if NOT rs.EOF then optarr1=rs.getRows else optarr1=""
					rs.close
					sSQL="SELECT coOptID,coCartOption FROM cartoptions WHERE coCartID=" & thecartid
					rs.open sSQL,cnn,0,1
						if NOT rs.EOF then optarr2=rs.getRows else optarr2=""
					rs.close
					if (isarray(optarr1) AND NOT isarray(optarr2)) OR (NOT isarray(optarr1) AND isarray(optarr2)) then hasoptions=FALSE
					if isarray(optarr1) AND isarray(optarr2) then
						if UBOUND(optarr1,2)<>UBOUND(optarr2,2) then hasoptions=FALSE
						if hasoptions then
							for index2=0 to UBOUND(optarr1,2)
								hasthisoption=FALSE
								for index3=0 to UBOUND(optarr2,2)
									if optarr1(0,index2)=optarr2(0,index3) AND optarr1(1,index2)=optarr2(1,index3) then hasthisoption=TRUE
								next
								if NOT hasthisoption then hasoptions=FALSE
							next
						end if
					end if
				end if
				if thecartid<>"" AND hasoptions then
					ect_query("DELETE FROM cartoptions WHERE coCartID="&cartarr(0,index))
					ect_query("DELETE FROM cart WHERE cartID="&cartarr(0,index))
				end if
			next
		end if
		sSQL="UPDATE cart SET cartClientID="&replace(SESSION("clientID"),"'","")&" WHERE cartClientID=0 AND cartCompleted=0 AND cartSessionID='"&escape_string(thesessionid)&"'"
		ect_query(sSQL)
		sSQL="UPDATE cart SET cartSessionID='"&escape_string(thesessionid)&"' WHERE cartCompleted=0 AND cartClientID="&replace(SESSION("clientID"),"'","")
		ect_query(sSQL)
		sSQL="SELECT cartID,cartProdID,cartProdPrice,pID,"&WSP&"pPrice FROM cart LEFT JOIN products ON cart.cartProdId=products.pID WHERE cartClientID="&replace(SESSION("clientID"),"'","")&" AND cartCompleted=0 AND cartProdID<>'"&giftcertificateid&"' AND cartProdID<>'"&donationid&"' AND cartProdID<>'"&giftwrappingid&"'"
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			if isnull(rs("pID")) then
				cartchanged=TRUE
				ect_query("DELETE FROM cartoptions WHERE coCartID="&rs("cartID"))
				ect_query("DELETE FROM cart WHERE cartID="&rs("cartID"))
			else
				newprice=checkpricebreaks(rs("cartProdID"),rs("pPrice"))
				if rs("cartProdPrice")<>newprice then cartchanged=TRUE ' recalculate wholesale price plus quant discounts
				if mysqlserver=TRUE then
					sSQL="SELECT coID,coPriceDiff,"&OWSP&"optPriceDiff,optFlags FROM cart INNER JOIN cartoptions ON cart.cartID=cartoptions.coCartID INNER JOIN options ON cartoptions.coOptID=options.optID INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optType IN (-4,-2,-1,1,2,4) AND cartID="&rs("cartID")
				else
					sSQL="SELECT coID,coPriceDiff,"&OWSP&"optPriceDiff,optFlags FROM cart INNER JOIN (cartoptions INNER JOIN (options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID) ON cartoptions.coOptID=options.optID) ON cart.cartID=cartoptions.coCartID WHERE optType IN (-4,-2,-1,1,2,4) AND cartID="&rs("cartID")
				end if
				rs2.open sSQL,cnn,0,1
				do while NOT rs2.EOF
					sSQL="UPDATE cartoptions SET coPriceDiff="&IIfVr((rs2("optFlags") AND 1)=0, rs2("optPriceDiff"), vsround((rs2("optPriceDiff") * newprice)/100.0, 2))&" WHERE coID="&rs2("coID")
					ect_query(sSQL)
					rs2.movenext
				loop
				rs2.close
			end if
			rs.movenext
		loop
		rs.close
		SESSION("xsshipping")=empty : SESSION("discounts")=empty : SESSION("xscountrytax")=empty : SESSION("tofreeshipamount")=empty : SESSION("tofreeshipquant")=empty
	end if
elseif getget("action")="dashboard" then
	sSQL="SELECT COUNT(*) AS thecnt FROM orders WHERE ordStatus>=2 AND ordDate>="&vsusdate(DateAdd("h",dateadjust,Now()))
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		if NOT isnull(rs("thecnt")) then neworders=rs("thecnt")
	end if
	rs.close
	newratings=0
	sSQL="SELECT COUNT(*) AS thecnt FROM ratings WHERE rtApproved=0"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		if NOT isnull(rs("thecnt")) then newratings=rs("thecnt")
	end if
	rs.close
	newaccounts=0
	sSQL="SELECT COUNT(*) AS thecnt FROM customerlogin WHERE clDateCreated>"&vsusdate(Date()-7)
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		if NOT isnull(rs("thecnt")) then newaccounts=rs("thecnt")
	end if
	rs.close
	newmaillist=0
	sSQL="SELECT COUNT(*) AS thecnt FROM mailinglist WHERE mlConfirmDate>"&vsusdate(Date()-7)
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		if NOT isnull(rs("thecnt")) then newmaillist=rs("thecnt")
	end if
	rs.close
	newaffiliate=0
	sSQL="SELECT COUNT(*) AS thecnt FROM affiliates WHERE affilDate>"&vsusdate(Date()-7)
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		if NOT isnull(rs("thecnt")) then newaffiliate=rs("thecnt")
	end if
	rs.close
	newgiftcert=0
	sSQL="SELECT COUNT(*) AS thecnt FROM giftcertificate WHERE gcDateCreated>"&vsusdate(Date()-7)
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		if NOT isnull(rs("thecnt")) then newgiftcert=rs("thecnt")
	end if
	rs.close
	newstocknotify=0
	if notifybackinstock then
		sSQL="SELECT COUNT(*) AS thecnt FROM notifyinstock"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			if NOT isnull(rs("thecnt")) then newstocknotify=rs("thecnt")
		end if
		rs.close
	end if
	if SESSION("loginid")=0 then
		sSQL="SELECT COUNT(*) AS thecnt FROM auditlog WHERE eventSuccess=0"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			if NOT isnull(rs("thecnt")) then newlogevents=rs("thecnt")
		end if
		rs.close
	end if
	print neworders & "&" & newratings & "&" & newaccounts & "&" & newmaillist & "&" & newaffiliate & "&" & newgiftcert & "&" & newstocknotify & "&" & newlogevents
	sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 5 ")&"ordID,ordName,ordLastName,ordShipName,ordShipLastName,ordReferer,ordDate,ordAddress,ordCity,ordState,ordTotal,ordAddInfo,ordStatus,statPrivate FROM orders,orderstatus WHERE statID=ordStatus AND ordStatus>1 ORDER BY ordID DESC"&IIfVs(mysqlserver=TRUE," LIMIT 0,5")
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		print "&"&rs("ordID")&"|"
		print jsurlencode(trim(rs("ordName")&" "&rs("ordLastName")))&"|"
		if trim(rs("ordShipName")&"")<>"" AND lcase(rs("ordShipName")&"")<>lcase(rs("ordName")&"") then
			print jsurlencode(trim(rs("ordShipName")&" "&rs("ordShipLastName")))
		end if
		print "|"&jsurlencode(formatdatetime(rs("ordDate")))&"|"&jsurlencode(rs("statPrivate"))&"|"&jsurlencode(FormatEuroCurrency(rs("ordTotal")))
		rs.movenext
	loop
	rs.close
elseif getget("action")="ipnarrived" AND getget("oid")<>"" then
	oid=getget("oid")
	if getget("sid")<>"" then
		sSQL="SELECT ordID FROM orders WHERE ordTransID<>'' AND ordTransSession='" & escape_string(getget("sid")) & "' AND ordID='" & escape_string(oid) & "'"
	elseif is_numeric(getget("oid")) then
		sSQL="SELECT ordID FROM orders WHERE ordAuthNumber<>'no ipn' AND ordID=" & oid
	else
		sSQL="SELECT ordID FROM orders WHERE ordDate>="&vsusdate(DateAdd("h",-36,Now()))&" AND ordAuthNumber='" & escape_string(oid) & "'"
	end if
	rs.open sSQL,cnn,0,1
	if rs.EOF then print "0" else print "1"
	rs.close
elseif request.querystring("action")="notifystock" then
	oid=request.querystring("oid")
	legalrequest=TRUE
	if NOT is_numeric(oid) then legalrequest=FALSE
	sSQL="SELECT pId FROM products WHERE pId='"&escape_string(request.querystring("pid"))&"'"
	rs.open sSQL,cnn,0,1
	if rs.EOF then legalrequest=FALSE
	rs.close
	sSQL="SELECT pId FROM products WHERE pId='"&escape_string(request.querystring("tpid"))&"'"
	rs.open sSQL,cnn,0,1
	if rs.EOF then legalrequest=FALSE
	rs.close
	ect_query("DELETE FROM notifyinstock WHERE nsDate<"&vsusdate(Date()-365))
	ect_query("DELETE FROM notifyinstock WHERE nsTriggerProdID='"&escape_string(request.querystring("tpid"))&"' AND nsEmail='"&escape_string(request.querystring("email"))&"'")
	if legalrequest then ect_query("INSERT INTO notifyinstock (nsProdID,nsTriggerProdID,nsOptID,nsEmail,nsDate) VALUES ('"&escape_string(request.querystring("pid"))&"','"&escape_string(request.querystring("tpid"))&"',"&oid&",'"&escape_string(request.querystring("email"))&"',"&vsusdate(Date())&")")
	print "SUCCESS"
elseif request.querystring("action")="clord" then
	if closeorderimmediately=TRUE then
		thesessionid=SESSION("sessionid")
		sSQL="SELECT ordID FROM orders WHERE ordStatus>1 AND ordAuthNumber='' AND " & getordersessionsql()
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then orderid=rs("ordID") else orderid=""
		rs.close
		if orderid<>"" then
			ect_query("UPDATE cart SET cartCompleted=2 WHERE cartCompleted=0 AND cartOrderID="&orderid)
			ect_query("UPDATE orders SET ordAuthNumber='CHECK MANUALLY' WHERE ordID="&orderid)
		end if
	end if
elseif request.querystring("action")="centinel2" then
	if debugmode then
		emailtxt=emailtxt&"-----VALUES-----"&emlNl
		emailtxt=emailtxt&"cardinal_method: "&SESSION("cardinal_method")&emlNl
		emailtxt=emailtxt&"cardinal_ordernum: "&SESSION("cardinal_ordernum")&emlNl
		emailtxt=emailtxt&"cardinal_sessionid: "&SESSION("cardinal_sessionid")&emlNl
		call DoSendEmailEO(emailAddr,emailAddr,"","ajaxservice.asp debug",emailtxt,emailObject,themailhost,theuser,thepass)
	end if
	signatureverify=""
	SESSION("ErrorDesc")=""
	SESSION("PAResStatus")=""
	sXML="<CardinalMPI>" & _
		addtag("Version","1.7") & _
		addtag("MsgType","cmpi_authenticate") & _
		addtag("ProcessorId",cardinalprocessor) & _
		addtag("MerchantId",cardinalmerchant) & _
		addtag("TransactionPwd",cardinalpwd) & _
		addtag("TransactionType","C") & _
		addtag("TransactionId",SESSION("cardinal_transaction")) & _
		addtag("OrderId",SESSION("cardinal_orderid")) & _
		addtag("PAResPayload",request.form("PaRes")) & "</CardinalMPI>"
	theurl="https://"&IIfVr(SESSION("cardinal_method")="7" OR SESSION("cardinal_method")="18","paypal","centinel400")&".cardinalcommerce.com/maps/txns.asp"
	if cardinaltestmode then theurl="https://centineltest.cardinalcommerce.com/maps/txns.asp"
	if cardinalurl<>"" then theurl=cardinalurl
	if callxmlfunction(theurl, "cmpi_msg=" & Server.URLEncode(sXML), res, "", "WinHTTP.WinHTTPRequest.5.1", vsRESPMSG, 12) then
		set xmlDoc=Server.CreateObject("MSXML2.DOMDocument")
		xmlDoc.validateOnParse=False
		xmlDoc.loadXML(res)
		set oNodeList=xmlDoc.documentElement.childNodes
		for i=0 To oNodeList.length - 1
			Set Item=oNodeList.item(i)
			if Item.nodeName="PAResStatus" then SESSION("PAResStatus")=Item.Text
			if Item.nodeName="SignatureVerification" then signatureverify=Item.Text
			if Item.nodeName="Cavv" then SESSION("Cavv")=Item.Text
			if Item.nodeName="EciFlag" then SESSION("EciFlag")=Item.Text
			if Item.nodeName="Xid" then SESSION("Xid")=Item.Text
			if Item.nodeName="ErrorDesc" then SESSION("ErrorDesc")=Item.Text
		next
	end if
	if signatureverify<>"Y" OR SESSION("PAResStatus")="N" then SESSION("centinelok")="N" else SESSION("centinelok")="Y"
%>
<html><head>
<script>
function onLoadHandler(){document.frmLaunchACS.submit();}
</script>
</head>
<body onload="onLoadHandler();">
<center>
<form name="frmLaunchACS" method="post" action="<%=storeurlssl%>cart<%=extension%>" target="_top">
<noscript>
<p>&nbsp;</p>
<h1>Processing your transaction</h1>
<%="<p>"&xxNoJS&"</p><p>"&xxMstClk&"</p>"%><p>&nbsp;</p>
<input type="submit" value="<%=xxSubmt%>"></center>
</noscript>
<input type="hidden" name="mode" value="authorize" />
<input type="hidden" name="method" value="<%=SESSION("cardinal_method")%>" />
<input type="hidden" name="ordernumber" value="<%=SESSION("cardinal_ordernum")%>" />
<input type="hidden" name="sessionid" value="<%=SESSION("cardinal_sessionid")%>" />
</form></center></body></html>
<%
elseif request.querystring("action")="centinel" then
	if debugmode then
		emailtxt=emailtxt&"-----VALUES-----"&emlNl
		emailtxt=emailtxt&"PaReq: "&SESSION("cardinal_pareq")&emlNl
		emailtxt=emailtxt&"TermUrl: "&storeurlssl&"vsadmin/ajaxservice.asp?action=centinel2"&emlNl
		call DoSendEmailEO(emailAddr,emailAddr,"","ajaxservice.asp debug",emailtxt,emailObject,themailhost,theuser,thepass)
	end if %><html><head>
<script>
function onLoadHandler(){document.frmLaunchACS.submit();}
</script>
</head>
<body onload="onLoadHandler();">
<center>
<form name="frmLaunchACS" method="post" action="<%=replace(request.querystring("url"),"""","")%>">
<noscript>
<p>&nbsp;</p>
<h1>Processing your transaction</h1>
<%="<p>"&xxNoJS&"</p><p>"&xxMstClk&"</p>"%><p>&nbsp;</p>
<input type="submit" value="<%=xxSubmt%>"></center>
</noscript>
<input type="hidden" name="PaReq" value="<%=replace(SESSION("cardinal_pareq"),"""","&quot;")%>">
<input type="hidden" name="TermUrl" value="<%=storeurlssl&"vsadmin/ajaxservice.asp?action=centinel2"%>">
<input type="hidden" name="MD" value="">
</form></center></body></html>
<%
elseif request.querystring("action")="applycert" then
	if SESSION("lastcertapplied")<>"" then lastapplied=timer() - SESSION("lastcertapplied") else lastapplied=1000
	SESSION("lastcertapplied")=timer()
	cpncode=request.querystring("cpncode")
	if lastapplied < 3 AND request.querystring("act")<>"delete" then
		print IIfVs(getget("stg1")="1","fail&")&xxFldCnt
	elseif cpncode<>"" then
		gotcpncode=FALSE
		if getget("stg1")<>"1" then print "<div style=""display:table""><div>"
		if request.querystring("act")="delete" then
			gotcpncode=TRUE
			SESSION("giftcerts")=replace(SESSION("giftcerts"), trim(replace(cpncode,"'","")) & " ", "")
			SESSION("cpncode")=replace(SESSION("cpncode"), trim(replace(cpncode,"'","")) & " ", "")
			print IIfVs(getget("stg1")="1","success&")&xxCpGcDl
		else
			sSQL="SELECT gcID FROM giftcertificate WHERE gcRemaining>0 AND gcAuthorized<>0 AND gcID='" & escape_string(replace(cpncode,"'","")) & "'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if instr(SESSION("giftcerts"), rs("gcID") & " ")=0 then
					SESSION("giftcerts")=SESSION("giftcerts") & rs("gcID") & " "
					print IIfVs(getget("stg1")="1","success&")&xxGcApld
				else
					print IIfVs(getget("stg1")="1","fail&")&xxGcAlAp
				end if
				gotcpncode=TRUE
			end if
			rs.close
		end if
		if NOT gotcpncode then
			sSQL="SELECT cpnID,cpnNumber FROM coupons WHERE cpnIsCoupon=1 AND cpnNumAvail>0 AND cpnEndDate>="&vsusdate(Date()) & " AND cpnNumber='" & trim(replace(cpncode,"'","")) & "'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if instr(SESSION("cpncode"), rs("cpnNumber") & " ")=0 then
					if trim(SESSION("cpncode"))="" then
						SESSION("cpncode")=SESSION("cpncode") & rs("cpnNumber") & " "
						print IIfVs(getget("stg1")="1","success&")&xxCpnApd
					else
						print IIfVs(getget("stg1")="1","fail&")&xxOnOnCp
					end if
				else
					print IIfVs(getget("stg1")="1","fail&")&xxCpAlAp
				end if
			else
				print IIfVs(getget("stg1")="1","fail&")&xxGCCNoF
			end if
			rs.close
		end if
		if getget("stg1")<>"1" then print "</div>"
		if (SESSION("giftcerts")<>"" OR SESSION("cpncode")<>"") AND getget("stg1")<>"1" then
			gcarr=split(trim(SESSION("giftcerts")), " ")
			for index=0 to UBOUND(gcarr)
				print "<div style=""display:table-row""><div style=""display:table-cell"">" & xxAppGC & "</div><div style=""display:table-cell"">" & gcarr(index) & "</div><div style=""display:table-cell"">(<a href=""#"" onclick=""return removecert('"&gcarr(index)&"')"">"&xxRemove&"</a>)</div></div>"
			next
			cpnarr=split(trim(SESSION("cpncode")), " ")
			for index=0 to UBOUND(cpnarr)
				print "<div style=""display:table-row""><div style=""display:table-cell"">" & xxApdCpn & "</div><div style=""display:table-cell"">" & cpnarr(index) & "</div><div style=""display:table-cell"">(<a href=""#"" onclick=""return removecert('"&cpnarr(index)&"')"">"&xxRemove&"</a>)</div></div>"
			next
		end if
		if getget("stg1")<>"1" then print "</div>"
	end if
elseif request.querystring("action")="checkupdates" then
	if proxyserver<>"" then
		set objHttp=Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
		objHttp.setProxy 2, proxyserver
	else
		set objHttp=Server.CreateObject("Msxml2.ServerXMLHTTP")
	end if
	objHttp.setTimeouts 10000, 10000, 10000, 10000
	objHttp.open "GET", "https://www.ecommercetemplates.com/updaterversions.asp?versions=true&format=ASP&plusver="&server.urlencode(request.querystring("storever")), false
	objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	err.number=0
	errnum=0
	on error resume next
	objHttp.Send ""
	errnum=err.number
	if errnum<>0 OR objHttp.status<>200 then
		recommendedversion=yyCdNoCo
		securityrelease="false"
		shouldupdate="true"
		xxhaserror="true"
	else
		recommendedversion=objHttp.responseXML.getElementsByTagName("recommendedversion").Item(0).firstChild.nodeValue
		securityrelease=(objHttp.responseXML.getElementsByTagName("securityrelease").Item(0).firstChild.nodeValue="true")
		shouldupdate=(objHttp.responseXML.getElementsByTagName("shouldupdate").Item(0).firstChild.nodeValue="true")
		xxhaserror="false"
		sSQL="UPDATE admin SET updLastCheck="&vsusdate(Date())&",updRecommended='"&recommendedversion&"',updSecurity="&IIfVr(securityrelease,"1","0")&",updShouldUpd="&IIfVr(shouldupdate,"1","0")&" WHERE adminID=1"
		ect_query(sSQL)
	end if
	on error goto 0
	Response.ContentType="text/xml"
	print "<?xml version=""1.0""?><updaterresults><recommendedversion>"&recommendedversion&"</recommendedversion><securityupdate>"&IIfVr(securityrelease,"true","false")&"</securityupdate><shouldupdate>"&IIfVr(shouldupdate,"true","false")&"</shouldupdate><haserror>"&xxhaserror&"</haserror></updaterresults>"
elseif request.querystring("action")="dazzleupd" then
	' print request.querystring("ordid") & " :: " & request.querystring("trackno")
	newtrackingnum=trim(request.querystring("trackno"))
	iordid=request.querystring("ordid")
	emailstat=request.querystring("emstatus")
	ordstatus=request.querystring("ordstatus")
	if is_numeric(iordid) AND newtrackingnum<>"" then
		iordid=int(iordid)
		rs.open "SELECT ordStatus,ordAuthNumber,ordEmail,ordDate,"&getlangid("statPublic",64)&",ordName,ordLastName,ordTrackNum,ordPayProvider,ordLang,ordClientID,loyaltyPoints,ordTotal,ordDiscount,pointsRedeemed,ordStatusInfo FROM orders INNER JOIN orderstatus ON orders.ordStatus=orderstatus.statID WHERE ordID="&iordid,cnn,0,1
		if NOT rs.EOF then
			oldordstatus=rs("ordStatus")
			ordauthno=rs("ordAuthNumber")
			ordemail=rs("ordEmail")
			orddate=rs("ordDate")
			oldstattext=rs(getlangid("statPublic",64))&""
			ordstatinfo=rs("ordStatusInfo")&""
			if htmlemails=TRUE then ordstatinfo=replace(ordstatinfo, vbCrLf, "<br />")
			ordername=trim(rs("ordName")&" "&rs("ordLastName"))
			trackingnum=trim(rs("ordTrackNum")&"")
			payprovider=rs("ordPayProvider")
			languageid=rs("ordLang")+1
			ordClientID=rs("ordClientID")
			loyaltypointtotal=rs("loyaltyPoints")
			ordTotal=rs("ordTotal")
			ordDiscount=rs("ordDiscount")
			pointsredeemed=rs("pointsRedeemed")
		end if
		rs.close

		if instr(trackingnum,newtrackingnum)>0 then
			iordid="" ' Already set
		elseif trackingnum<>"" then
			ect_query("UPDATE orders SET ordTrackNum='" & escape_string(trackingnum & "," & newtrackingnum) & "',ordStatusDate=" & vsusdatetime(DateAdd("h",dateadjust,Now())) & " WHERE ordID=" & iordid)
			trackingnum=trackingnum & "," & newtrackingnum
		else
			ect_query("UPDATE orders SET ordTrackNum='" & escape_string(newtrackingnum) & "',ordStatusDate=" & vsusdatetime(DateAdd("h",dateadjust,Now())) & " WHERE ordID=" & iordid)
			trackingnum=newtrackingnum
		end if

		if iordid<>"" AND ordstatus<>"" then
			ordstatus=int(ordstatus)
			' if oldordstatus<>Int(ordstatus) then
				if emailstat="1" AND ordstatus<>1 then
					sSQL="SELECT orderstatussubject,orderstatussubject2,orderstatussubject3,orderstatusemail,orderstatusemail2,orderstatusemail3 FROM emailmessages WHERE emailID=1"
					rs.open sSQL,cnn,0,1
					ordstatussubject(1)=trim(rs("orderstatussubject")&"")
					ordstatusemail(1)=rs("orderstatusemail")&""
					ordstatussubject(2)=trim(rs("orderstatussubject2")&"")
					ordstatusemail(2)=rs("orderstatusemail2")&""
					ordstatussubject(3)=trim(rs("orderstatussubject3")&"")
					ordstatusemail(3)=rs("orderstatusemail3")&""
					rs.close

					rs.open "SELECT "&getlangid("statPublic",64)&",emailstatus FROM orderstatus WHERE statID=" & ordstatus,cnn,0,1
					if NOT rs.EOF then
						newstattext=rs(getlangid("statPublic",64))&""
						emailstatus=cint(rs("emailstatus"))<>0
					else
						emailstatus=FALSE
					end if
					rs.close
					if getget("noemail")="true" then emailstatus=FALSE

					if (adminlangsettings AND 4096)=0 then languageid=1
					if ordstatussubject(languageid)<>"" then emailsubject=ordstatussubject(languageid) else emailsubject="Order status updated"
					ose=ordstatusemail(languageid)
					for index=0 to 18
						replaceone=TRUE
						do while replaceone
							ose=replaceemailtxt(ose, "%statusid" & index & "%", IIfVr(index=ordstatus,"%ectpreserve%",""), replaceone)
						loop
					next
					ose=replace(ose, "%orderid%", iordid)
					ose=replace(ose, "%orderdate%", FormatDateTime(orddate, 1) & " " & FormatDateTime(orddate, 4))
					ose=replace(ose, "%oldstatus%", oldstattext)
					ose=replace(ose, "%newstatus%", newstattext)
					ose=replace(ose, "%date%", FormatDateTime(DateAdd("h",dateadjust,Now()), 1) & " " & FormatDateTime(DateAdd("h",dateadjust,Now()), 4))
					ose=replace(ose, "%ordername%", ordername)
					ose=replaceemailtxt(ose, "%statusinfo%", ordstatinfo, replaceone)
					tracknumarr=split(trackingnum,",")
					for index=0 to UBOUND(tracknumarr)
						ose=replaceemailtxt(ose, "%trackingnum%", tracknumarr(index), replaceone)
					next
					do while instr(ose, "%trackingnum%")>0
						ose=replaceemailtxt(ose, "%trackingnum%", "", replaceone)
					loop
					reviewlinks=""
					norepeatlinks=""
					if instr(ose, "%reviewlinks%") > 0 then
						sSQL="SELECT cartProdID,cartOrigProdID FROM cart WHERE cartOrderID="&iordid
						rs2.open sSQL,cnn,0,1
						do while NOT rs2.EOF
							if trim(rs2("cartOrigProdID")&"")<>"" then cartprodid=rs2("cartOrigProdID") else cartprodid=rs2("cartProdID")
							if instr(norepeatlinks,",'"&cartprodid&"'")=0 then
								norepeatlinks=norepeatlinks&",'"&cartprodid&"'"
								sSQL="SELECT pID,"&getlangid("pName",1)&",pStaticPage,pStaticURL,pDisplay FROM products WHERE pDisplay<>0 AND pID='"&escape_string(cartprodid)&"'"
								rs.open sSQL,cnn,0,1
								if NOT rs.EOF then
									thelink=storeurl & getdetailsurl(rs("pID"),rs("pStaticPage"),rs(getlangid("pName",1)),trim(rs("pStaticURL")&""),"review=true","")
									if htmlemails=TRUE then thelink="<a href=""" & thelink & """>" & thelink & "</a>"
									reviewlinks=reviewlinks & thelink & emlNl
								end if
								rs.close
							end if
							rs2.MoveNext
						loop
						rs2.close
					end if
					ose=replaceemailtxt(ose, "%reviewlinks%", reviewlinks, replaceone)
					ose=replace(ose, "%nl%", emlNl)
					ose=replace(ose, "<br />", emlNl)
					if emailstatus then Call DoSendEmailEO(ordemail,emailAddr,"",emailsubject,ose,emailObject,themailhost,theuser,thepass)
				end if
			' end if
			if oldordstatus<>Int(ordstatus) then ect_query("UPDATE orders SET ordStatus=" & ordstatus & ",ordStatusDate=" & vsusdatetime(DateAdd("h",dateadjust,Now())) & " WHERE ordID=" & iordid)
		end if
		print "SUCCESS|"&request.querystring("rowid")
	else
		print "ERROR"
	end if
elseif request.querystring("action")="dazzle" then
	isdazzlefile=FALSE
	isworldshipfile=FALSE
	csvcurrpos=1
	addressindex=-1
	trackingindex=-1
	csvfile=trim(request.form("dazzletext"))
	csvlen=len(csvfile)
	csvline=lcase(replace(getcsvline(),"""",""))
	columnarr=split(csvline,vbTab)
	fileseparator=vbTab
	if isarray(columnarr) then
		for index=0 to UBOUND(columnarr)
			if columnarr(index)="address" then addressindex=index
			if columnarr(index)="tracking_id" then trackingindex=index
		next
		if addressindex<>-1 AND trackingindex<>-1 then isdazzlefile=TRUE
	end if
	if NOT isdazzlefile then
		columnarr=split(csvline,",")
		if isarray(columnarr) then
			for index=0 to UBOUND(columnarr)
				if columnarr(index)="shiptocompanyorname" then addressindex=index
				if columnarr(index)="packagetrackingnumber" then trackingindex=index
				if columnarr(index)="shiptoemailaddress" then emailindex=index
			next
			if addressindex<>-1 AND trackingindex<>-1 AND emailindex<>-1 then isworldshipfile=TRUE
			fileseparator=","
		end if
	end if
	if isdazzlefile OR isworldshipfile then
		do while csvcurrpos<csvlen
			csvline=trim(getcsvline())
			orderid=0
			if csvline<>"" then
				columnarr=split(csvline,fileseparator)
				if isarray(columnarr) then
					theaddress=trim(columnarr(addressindex))
					thetracknum=trim(replace(columnarr(trackingindex),"""",""))
					if isworldshipfile then theemail=trim(replace(columnarr(emailindex),"""",""))
					if (isdazzlefile AND theaddress<>"") OR (isworldshipfile AND theemail<>"") then
						addressarr=split(theaddress,",")
						print "==DAZZLELINE==" & thetracknum & "==ORIGADD==" & theaddress
						if isarray(addressarr) OR isworldshipfile then
							sSQL="SELECT TOP 5 ordID,ordStatus,ordName,ordLastName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry FROM orders WHERE ordStatus>=3 AND"
							if isworldshipfile then
								sSQL=sSQL & " ordEmail='"&escape_string(theemail)&"' AND ordEmail<>''"
							else
								if usefirstlastname then
									call splitname(addressarr(0),tfirstn,tlastn)
									sSQL=sSQL & " (((ordName='"&escape_string(trim(tfirstn))&"' AND ordLastName='"&escape_string(trim(tlastn))&"') OR (ordShipName='"&escape_string(trim(tfirstn))&"' AND ordShipLastName='"&escape_string(trim(tlastn))&"' AND (ordShipName<>'' OR ordShipLastName<>'')))"
									call splitname(addressarr(1),tfirstn,tlastn)
									sSQL=sSQL & " OR ((ordName='"&escape_string(trim(tfirstn))&"' AND ordLastName='"&escape_string(trim(tlastn))&"') OR (ordShipName='"&escape_string(trim(tfirstn))&"' AND ordShipLastName='"&escape_string(trim(tlastn))&"' AND (ordShipName<>'' OR ordShipLastName<>''))))"
								else
									sSQL=sSQL & " (ordName='"&escape_string(trim(addressarr(0)))&"' OR ordName='"&escape_string(trim(addressarr(1)))&"' OR ((ordShipName='"&escape_string(trim(addressarr(0)))&"' OR ordShipName='"&escape_string(trim(addressarr(1)))&"') AND ordShipName<>''))"
								end if
							end if
							sSQL=sSQL & " ORDER BY ordID DESC"
							rs.open sSQL,cnn,0,1
							do while NOT rs.EOF
								print "==MATCHLINE==" & rs("ordID") & "|" & rs("ordStatus") & "==FULLADD==" & trim(rs("ordName")&" "&rs("ordLastName")) & ", " & rs("ordAddress") & ", "
								if trim(rs("ordAddress2"))<>"" then print rs("ordAddress2") & ", "
								print rs("ordCity") & ", " & rs("ordState") & ", " & rs("ordZip")
								rs.movenext
							loop
							rs.close
						end if
					end if
				end if
			end if
		loop
	else
		print "ERRORFILEFORMAT"
	end if
elseif request.querystring("action")="getlist" then
	rc=0
	if maxadminlookup="" then maxadminlookup=50
	if request.querystring("listtype")<>"adddets" then print request.querystring("objid")&"==LISTOBJ=="
	listtext=replace(request.form("listtext"),"[","[[]")
	if request.querystring("listtype")="adddets" then
		actarr=split(request.form("listtext"),"|")
		if actarr(0)="0" then
			print actarr(1) & "==LISTELM=="
			sSQL="SELECT clID,clUserName FROM customerlogin WHERE clID="&actarr(1)
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if usefirstlastname then
					ordName=trim(rs("clUserName")&"")
					ordLastName=""
					if instr(ordName," ")>0 then
						namearr=split(ordName," ",2)
						ordName=namearr(0)
						ordLastName=namearr(1)
					end if
					print jsenc(ordName) & "==LISTELM==" & jsenc(ordLastName)
				else
					print jsenc(rs("clUserName")&"") & "==LISTELM=="
				end if
			end if
			rs.close
		elseif actarr(0)="1" OR actarr(0)="2" then
			if actarr(0)="1" then
				sSQL="SELECT addCustID,addName,addLastName,addAddress,addAddress2,addCity,addState,addZip,addCountry,addPhone,addExtra1,addExtra2 FROM address WHERE addID="&actarr(1)
			else
				sSQL="SELECT 0 AS addCustID,ordName AS addName,ordLastName AS addLastName,ordAddress AS addAddress,ordAddress2 AS addAddress2,ordCity AS addCity,ordState AS addState,ordZip AS addZip,ordCountry AS addCountry,ordPhone AS addPhone,ordExtra1 AS addExtra1,ordExtra2 AS addExtra2 FROM orders WHERE ordID="&actarr(1)
			end if
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				print rs("addCustID") & "==LISTELM==" & jsenc(rs("addName")&"") & "==LISTELM==" & jsenc(rs("addLastName")&"") & "==LISTELM=="
				print jsenc(rs("addAddress")&"") & "==LISTELM==" & jsenc(rs("addAddress2")&"") & "==LISTELM=="
				print jsenc(rs("addCity")&"") & "==LISTELM==" & jsenc(rs("addState")&"") & "==LISTELM=="
				print jsenc(rs("addZip")&"") & "==LISTELM==" & jsenc(rs("addCountry")&"") & "==LISTELM=="
				print jsenc(rs("addPhone")&"") & "==LISTELM=="
				print jsenc(rs("addExtra1")&"") & "==LISTELM==" & jsenc(rs("addExtra2")&"")
			end if
			rs.close
		end if
	elseif request.querystring("listtype")="email" then
		' noaddress=0 addresstable=1 orderstable=2 | clid or addid or ordid
		gotresults=false
		if rc<10 then
			sSQL="SELECT"&IIfVr(mysqlserver<>true," TOP 20","")&" clID,clEmail,clUserName,clActions,clPercentDiscount,addName,addLastName,addAddress,addAddress2,addID,addCity,addState,addZip,addCountry FROM customerlogin INNER JOIN address ON customerlogin.clID=address.addCustID WHERE "
			sSQL=sSQL & "clEmail LIKE '"&escape_string(listtext)&"%' "
			sSQL=sSQL & "ORDER BY clEmail"
			if mysqlserver=TRUE then sSQL=sSQL & " LIMIT 0,20"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF AND gotresults then print jsenc("----------------")&"==LISTELM=="&jsenc("----------------")&"==LISTOBJ==" : gotresults=FALSE
			do while NOT rs.EOF
				theaddress=rs("clEmail")&" / "&trim(rs("addName")&" "&rs("addLastName"))&" - "&rs("addAddress")
				if trim(rs("addAddress2")&"")<>"" then theaddress=theaddress & ", " & rs("addAddress2")
				theaddress=theaddress & ", " & rs("addState") & ", " & rs("addZip") & ", " & rs("addCountry")
				thecode="1|"&rs("addID")
				print jsenc(rs("clEmail"))&"==LISTELM=="&jsenc(theaddress)&"==LISTELM=="&thecode&"==LISTELM=="&IIfVr((rs("clActions") AND 8)=8,1,0)&"==LISTELM=="&IIfVr((rs("clActions") AND 16)=16,1,0)&"==LISTELM=="&rs("clPercentDiscount")&"==LISTOBJ=="
				rc=rc+1
				gotresults=TRUE
				rs.movenext
			loop
			rs.close
		end if
		if rc<20 then
			sSQL="SELECT DISTINCT"&IIfVr(mysqlserver<>true," TOP 20","")&" MAX(ordID) AS ordID,ordEmail,ordName,ordLastName,ordAddress,ordAddress2,ordState,ordZip,ordCountry FROM orders WHERE "
			sSQL=sSQL & "ordEmail LIKE '"&escape_string(listtext)&"%' "
			sSQL=sSQL & "GROUP BY ordEmail,ordName,ordLastName,ordAddress,ordAddress2,ordState,ordZip,ordCountry ORDER BY ordEmail"
			if mysqlserver=TRUE then sSQL=sSQL & " LIMIT 0,20"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF AND gotresults then print jsenc("----------------")&"==LISTELM=="&jsenc("----------------")&"==LISTOBJ==" : gotresults=FALSE
			do while NOT rs.EOF
				theaddress=rs("ordEmail")&" / "&trim(rs("ordName")&" "&rs("ordLastName"))&" - "&rs("ordAddress")
				if trim(rs("ordAddress2")&"")<>"" then theaddress=theaddress & ", " & rs("ordAddress2")
				theaddress=theaddress & ", " & rs("ordState") & ", " & rs("ordZip") & ", " & rs("ordCountry")
				thecode="2|"&rs("ordID")
				print jsenc(rs("ordEmail"))&"==LISTELM=="&jsenc(theaddress)&"==LISTELM=="&thecode&"==LISTOBJ=="
				rc=rc+1
				gotresults=TRUE
				rs.movenext
			loop
			rs.close
		end if
		if rc<40 then
			sSQL="SELECT"&IIfVr(mysqlserver<>true," TOP 10","")&" clID,clEmail,clUserName FROM customerlogin WHERE "
			sSQL=sSQL & "clEmail LIKE '"&escape_string(listtext)&"%' "
			sSQL=sSQL & "ORDER BY clEmail"
			if mysqlserver=TRUE then sSQL=sSQL & " LIMIT 0,10"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF AND gotresults then print jsenc("----------------")&"==LISTELM=="&jsenc("----------------")&"==LISTOBJ==" : gotresults=FALSE
			do while NOT rs.EOF
				theaddress=rs("clEmail")&" / "&rs("clUserName")
				thecode="0|"&rs("clID")
				print jsenc(rs("clEmail"))&"==LISTELM=="&jsenc(theaddress)&"==LISTELM=="&thecode&"==LISTOBJ=="
				rc=rc+1
				gotresults=TRUE
				rs.movenext
			loop
			rs.close
		end if
	elseif request.querystring("listtype")="prodid" OR request.querystring("listtype")="prodname" then
		sSQL="SELECT"&IIfVr(mysqlserver<>true," TOP "&maxadminlookup,"")&" pID,pName FROM products WHERE "
		if request.querystring("listtype")="prodname" then
			sSQL=sSQL & "pName LIKE '"&escape_string(listtext)&"%' ORDER BY pName"
		else
			sSQL=sSQL & "pID LIKE '"&escape_string(listtext)&"%' OR pSKU LIKE '"&escape_string(listtext)&"%' ORDER BY pID"
		end if
		if mysqlserver=TRUE then sSQL=sSQL & " LIMIT 0,"&maxadminlookup
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			print jsenc(rs("pID"))&"==LISTELM=="&jsenc(rs("pName"))&"==LISTOBJ=="
			rc=rc+1
			rs.movenext
		loop
		rs.close
		if maxadminlookup-30 > rc then
			sSQL="SELECT"&IIfVr(mysqlserver<>true," TOP "&maxadminlookup,"")&" pID,pName FROM products WHERE "
			if request.querystring("listtype")="prodname" then
				sSQL=sSQL & "pName LIKE '%"&escape_string(listtext)&"%' AND NOT (pName LIKE '"&escape_string(listtext)&"%') ORDER BY pName"
			else
				sSQL=sSQL & "pID LIKE '%"&escape_string(listtext)&"%' AND NOT (pID LIKE '"&escape_string(listtext)&"%') ORDER BY pID"
			end if
			if mysqlserver=TRUE then sSQL=sSQL & " LIMIT 0,"&maxadminlookup
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF AND rc<>0 then print jsenc("----------------")&"==LISTELM=="&jsenc("----------------")&"==LISTOBJ=="
			do while NOT rs.EOF
				print jsenc(rs("pID"))&"==LISTELM=="&jsenc(rs("pName"))&"==LISTOBJ=="
				rc=rc+1
				rs.movenext
			loop
			rs.close
		end if
	end if
elseif request.querystring("action")="updateoptions" then
	session.LCID=1033
	byoptions=FALSE
	optstockjs=""
	id=request.querystring("index")
	productid=request.form("productid")
	haswsp=request.querystring("wsp")="1"
	percdisc=request.querystring("perc")
	if haswsp then
		WSP="pWholesalePrice AS "
		TWSP="pWholesalePrice"
		if wholesaleoptionpricediff=TRUE then OWSP="optWholesalePriceDiff AS "
	end if
	if is_numeric(percdisc) then
		percdisc=(100.0-cdbl(percdisc))/100.0
		WSP=percdisc & "*"&IIfVr(haswsp,"pWholesalePrice","pPrice")&" AS "
		TWSP=percdisc & "*"&IIfVr(haswsp,"pWholesalePrice","pPrice")
		OWSP=percdisc & "*"&IIfVr(haswsp AND wholesaleoptionpricediff,"optWholesalePriceDiff","optPriceDiff")&" AS "
	end if
	sSQL="SELECT "&getlangid("pName",1)&","&WSP&"pPrice,pStockByOpts,pInStock,pExemptions FROM products WHERE pID='"&escape_string(productid)&"'"
	rs.open sSQL,cnn,0,1
	if rs.EOF then
		prodname="Not Found: " & productid
		prodprice=0
		prodstock="''"
		prodexemptions=0
	else
		prodname=rs(getlangid("pName",1))&""
		prodprice=vsround(rs("pPrice"),2)
		if rs("pStockByOPts")<>0 then prodstock="'bo'" : byoptions=TRUE else prodstock=IIfVr(isnull(rs("pInStock")),"0",rs("pInStock"))
		prodexemptions=rs("pExemptions")
	end if
	rs.close
	opttext=""
	sSQL="SELECT poOptionGroup,optType,optFlags FROM prodoptions INNER JOIN optiongroup ON optiongroup.optGrpID=prodoptions.poOptionGroup WHERE poProdID='"&escape_string(productid)&"' ORDER BY poID"
	prodoptions=""
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		prodoptions=rs.getrows
	else
		opttext=opttext& "-"
	end if
	rs.close
	jstext=""
	if IsArray(prodoptions) then
		opttext=opttext& "<div class=""optionstable"">"
		for rowcounter=0 to UBOUND(prodoptions,2)
			index=0
			sSQL="SELECT optID,"&getlangid("optName",32)&","&getlangid("optGrpName",16)&","&OWSP&"optPriceDiff,optType,optFlags,optStock,optTxtCharge FROM options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optGroup="&prodoptions(0,rowcounter)&" ORDER BY optID"
			rs2.Open sSQL,cnn,0,1
			if NOT rs2.EOF then
				opttext=opttext & "<div class=""optionstablerow""><div class=""optionstableleft"">"&rs2(getlangid("optGrpName",16)) & ":</div><div>"
				if abs(Int(rs2("optType")))=3 then
					opttext=opttext& "<input type='hidden' name='optn"&id&"_"&rs2("optID")&"' value='"&rs2("optID")&"' />"
					if rs2("optTxtCharge")<>0 then jstext=jstext&"opttxtcharge["&rs2("optID")&"]="&rs2("optTxtCharge")&";"
					opttext=opttext& "<textarea wrap='virtual' name='voptn"&id&"_"&rs2("optID")&"' id='voptn"&id&"_"&rs2("optID")&"' cols='30' rows='3'>"
					opttext=opttext& rs2(getlangid("optName",32))&"</textarea>"
				else
					opttext=opttext& "<select class=""prodoption"" onchange=""dorecalc(true)"" name='optn"&id&"_"&rowcounter&"' id='optn"&id&"_"&rowcounter&"' size='1'>"
					opttext=opttext& "<option value=''>"&xxPlsSel&"</option>"
					do while not rs2.EOF
						opttext=opttext& "<option value='"&rs2("optID")&"|"&IIfVr((rs2("optFlags") AND 1)=1,(prodprice*rs2("optPriceDiff"))/100.0,rs2("optPriceDiff"))&"'>"&rs2(getlangid("optName",32))
						if cdbl(rs2("optPriceDiff"))<>0 then
							opttext=opttext& " "
							if cdbl(rs2("optPriceDiff")) > 0 then opttext=opttext& "+"
							if (rs2("optFlags") AND 1)=1 then
								opttext=opttext& FormatNumber((prodprice*rs2("optPriceDiff"))/100.0,2)
							else
								opttext=opttext& FormatNumber(rs2("optPriceDiff"),2)
							end if
						end if
						opttext=opttext& "</option>"&vbCrLf
						if byoptions then
							optstockjs=optstockjs & "if(typeof(stock['oid_"&rs2("optID")&"'])==""undefined""){"
							optstockjs=optstockjs & "stock['oid_"&rs2("optID")&"']="&rs2("optStock")&";"
							optstockjs=optstockjs & "}"
						end if
						rs2.MoveNext
					loop
					opttext=opttext& "</select>"
				end if
				opttext=opttext& "</div></div>"
			end if
			rs2.Close
		next
		opttext=opttext& "</div>"
	end if
	jstext=jstext&"document.getElementById('prodname"&id&"').value='"&replace(prodname,"'","\'")&"';"
	jstext=jstext&"document.getElementById('price"&id&"').value='"&prodprice&"';"
	jstext=jstext&"document.getElementById('stateexempt"&id&"').value='"&IIfVr((prodexemptions AND 1)=1,"true","false")&"';"
	jstext=jstext&"document.getElementById('countryexempt"&id&"').value='"&IIfVr((prodexemptions AND 2)=2,"true","false")&"';"
	jstext=jstext&"document.getElementById('optdiffspan"&id&"').value=0;"
	jstext=jstext&"if(typeof(stock['pid_"&replace(productid,"'","\'")&"'])==""undefined""){"
	jstext=jstext&"stock['pid_"&replace(productid,"'","")&"']="&prodstock&";"
	jstext=jstext&optstockjs&"}"
	print id&"==LISTELM=="&jsenc(opttext)&"==LISTELM=="&jsenc(jstext)
elseif request.querystring("processor")="AuthNET" then
	ordID=replace(getget("gid"),"'","")
	payprov="" : transid=""
	success=FALSE
	if is_numeric(ordID) then
		sSQL="SELECT ordTransID,ordPayProvider FROM orders WHERE ordID=" & ordID
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			transid=rs("ordTransID")
			payprov=rs("ordPayProvider")
		end if
		rs.close
	end if
	if is_numeric(ordID) AND is_numeric(payprov) then
		if getpayprovdetails(payprov,data1,data2,data3,demomode,ppmethod) then
			if secretword<>"" then
				data1=upsdecode(data1, secretword)
				data2=upsdecode(data2, secretword)
			end if
			if getget("act")="charge" then
				anaction="priorAuthCaptureTransaction"
			elseif getget("act")="refund" then
				anaction="refundTransaction"
			elseif getget("act")="void" then
				anaction="voidTransaction"
			elseif getget("act")="reauth" then
				anaction="authOnlyContinueTransaction"
			end if
			if anaction<>"" then
				sjson="{""createTransactionRequest"":{""merchantAuthentication"":{""name"":" & json_encode(data1) & ",""transactionKey"":" & json_encode(data2) & "}," & _
					"""transactionRequest"":{""transactionType"":""" & anaction & """," & IIfVs(anaction<>"voidTransaction","""amount"":""" & getpost("amount") & """,") & """refTransId"":""" & transid & """}}}"
				success=callxmlfunction("https://api" & IIfVs(demomode,"test") & ".authorize.net/xml/v1/request.api",sjson,jres,"","Msxml2.ServerXMLHTTP",vsRESPMSG,FALSE)
				if NOT success then print errormsg
			end if
			if success then
				resultcode=get_json_val(jres,"resultCode","messages")
				if resultcode="Ok" then
					if is_numeric(getpost("capstatus")) then ect_query("UPDATE orders SET ordStatus=" & escape_string(getpost("capstatus")) & ",ordStatusDate=" & vsusdatetime(DateAdd("h",dateadjust,now())) & " WHERE ordID=" & escape_string(ordID) & " AND ordStatus<>" & escape_string(getpost("capstatus")))
					print get_json_val(jres,"description","messages")
				else
					errtxt=get_json_val(jres,"errorText","")
					if errtxt<>"" then print errtxt&"<br />"
					print get_json_val(jres,"text","messages")
				end if
			end if
		end if
	end if
elseif request.querystring("processor")="PayPalCO" then
	if getpayprovdetails(27,xmlfnuser,xmlfnpassword,data3,demomode,ppmethod) then
		ordID=replace(getget("gid"),"'","")
		if is_numeric(ordID) then
			transid=""
			sSQL="SELECT ordID,ordAuthNumber FROM orders WHERE ordID='" & escape_string(ordID) & "'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then transid=rs("ordAuthNumber")
			rs.close
			if getget("act")="charge" then
				requestxx="{""amount"":{""currency"": """ & countryCurrency & """,""total"": """ & getpost("amount") & """}" & IIfVs(getpost("additionalcapture")="1",",""is_final_capture"": true") & "}"
				xmlfnheaders=array(array("Content-Type: application/json"))
				call callcurlfunction("https://api." & IIfVs(demomode,"sandbox.") & "paypal.com/v1/payments/orders/" & transid & "/capture",requestxx,res,"",errormsg,FALSE)
			elseif getget("act")="void" then
				requestxx="{}"
				xmlfnheaders=array(array("Content-Type: application/json"))
				call callcurlfunction("https://api." & IIfVs(demomode,"sandbox.") & "paypal.com/v1/payments/authorization/" & transid & "/void",requestxx,res,"",errormsg,FALSE)
			elseif getget("act")="reauth" then
				requestxx="{""amount"":{""currency"": """ & countryCurrency & """,""total"": """ & getpost("amount") & """}}"
				xmlfnheaders=array(array("Content-Type: application/json"))
				call callcurlfunction("https://api." & IIfVs(demomode,"sandbox.") & "paypal.com/v1/payments/authorization/" & transid & "/reauthorize",requestxx,res,"",errormsg,FALSE)
			end if
			tstate=""
			idpos=instr(res,"""state""")
			if idpos>0 then
				delim1=instr(idpos+4,res,"""")
				delim2=instr(delim1+1,res,"""")
				tstate=mid(res,delim1+1,delim2-delim1-1)
			end if
			tmessage=""
			idpos=instr(res,"""message""")
			if idpos>0 then
				delim1=instr(idpos+4,res,"""")
				delim2=instr(delim1+1,res,"""")
				tmessage=mid(res,delim1+1,delim2-delim1-1)
			end if
			if tstate="completed" then
				ect_query("UPDATE orders SET ordStatus='" & escape_string(getpost("capstatus")) & "',ordStatusDate=" & vsusdatetime(DateAdd("h",dateadjust,now())) & " WHERE ordID='" & escape_string(ordID) & "' AND ordStatus<>'" & escape_string(getpost("capstatus")) & "'")
				print "Completed"
			else
				print tmessage
			end if
		end if
	end if
elseif request.querystring("processor")="PayPal" then
	ordID=replace(request.querystring("gid"),"'","")
	sSQL="SELECT ordPayProvider,ordAuthNumber,ordTransID,ordShipping,ordStateTax,ordCountryTax,ordHSTTax,ordTotal,ordHandling,ordDiscount FROM orders WHERE ordID=" & ordID
	rs.open sSQL,cnn,0,1
	authcode=rs("ordAuthNumber")
	transid=rs("ordTransID")
	grandtotal=(rs("ordShipping")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordHSTTax")+rs("ordTotal")+rs("ordHandling"))-rs("ordDiscount")
	rs.close
	success=getpayprovdetails(19,username,data2pwd,data2hash,demomode,ppmethod)
	if data2pwd<>"" then data2pwd=urldecode(split(data2pwd,"&")(0))
	if instr(username,"/")>0 then
		data1arr=split(username,"/")
		if grandtotal<12 then username=trim(data1arr(1)) else username=trim(data1arr(0))
		data1arr=split(data2pwd,"/")
		if grandtotal<12 AND ubound(data1arr)>0 then data2pwd=trim(data1arr(1)) else data2pwd=trim(data1arr(0))
		data1arr=split(data2hash,"/")
		if grandtotal<12 AND ubound(data1arr)>0 then data2hash=trim(data1arr(1)) else data2hash=trim(data1arr(0))
	end if
	if demomode then sandbox=".sandbox" else sandbox=""
	if NOT success then
		print "username / pw not set for express checkout"
	else
		if request.querystring("act")="charge" then
			sXML=ppsoapheader(username, data2pwd, data2hash) & "<soap:Body><DoCaptureReq xmlns=""urn:ebay:api:PayPalAPI"">" & _
				"<DoCaptureRequest xmlns=""urn:ebay:api:PayPalAPI"">" & _
					"<Version xmlns=""urn:ebay:apis:eBLBaseComponents"" xsi:type=""xsd:string"">1.0</Version>" & _
					"<AuthorizationID>" & authcode & "</AuthorizationID>" & _
					"<Amount currencyID=""" & countryCurrency & """ xsi:type=""cc:BasicAmountType"">" & request.form("amount") & "</Amount>" & _
					"<CompleteType>" & IIfVr(request.form("additionalcapture")="1", "NotComplete", "Complete") & "</CompleteType>" & _
					"<Note>" & request.form("comments") & "</Note>" & _
				"</DoCaptureRequest></DoCaptureReq></soap:Body></soap:Envelope>"
		elseif request.querystring("act")="void" then
			sXML=ppsoapheader(username, data2pwd, data2hash) & "<soap:Body><DoVoidReq xmlns=""urn:ebay:api:PayPalAPI"">" & _
				"<DoVoidRequest xmlns=""urn:ebay:api:PayPalAPI"">" & _
					"<Version xmlns=""urn:ebay:apis:eBLBaseComponents"" xsi:type=""xsd:string"">1.0</Version>" & _
					"<AuthorizationID>" & authcode & "</AuthorizationID>" & _
					"<Note>" & request.form("comments") & "</Note>" & _
				"</DoVoidRequest></DoVoidReq></soap:Body></soap:Envelope>"
		elseif request.querystring("act")="reauth" then
			sXML=ppsoapheader(username, data2pwd, data2hash) & "<soap:Body><DoReauthorizationReq xmlns=""urn:ebay:api:PayPalAPI"">" & _
				"<DoReauthorizationRequest xmlns=""urn:ebay:api:PayPalAPI"">" & _
					"<Version xmlns=""urn:ebay:apis:eBLBaseComponents"" xsi:type=""xsd:string"">1.0</Version>" & _
					"<AuthorizationID>" & authcode & "</AuthorizationID>" & _
					"<Amount currencyID=""" & countryCurrency & """ xsi:type=""cc:BasicAmountType"">" & request.form("amount") & "</Amount>" & _
					"<Note>" & request.form("comments") & "</Note>" & _
				"</DoReauthorizationRequest></DoReauthorizationReq></soap:Body></soap:Envelope>"
		end if
		if callxmlfunction("https://api" & IIfVr(data2hash<>"", "-3t", "") & sandbox & ".paypal.com/2.0/", sXML, res, IIfVr(data2hash<>"","",username), "WinHTTP.WinHTTPRequest.5.1", errormsg, FALSE) then
			success=FALSE:vsERRCODE="":vsRESPMSG="":vsTRANSID=""
			set xmlDoc=Server.CreateObject("MSXML2.DOMDocument")
			xmlDoc.validateOnParse=False
			xmlDoc.loadXML (res)
			Set nodeList=xmlDoc.getElementsByTagName("SOAP-ENV:Body")
			Set n=nodeList.Item(0)
			for j=0 to n.childNodes.length - 1
				Set e=n.childNodes.Item(i)
				if e.nodeName="DoCaptureResponse" OR e.nodeName="DoVoidResponse" OR e.nodeName="DoReauthorizationResponse" then
					for k9=0 To e.childNodes.length - 1
						Set t=e.childNodes.Item(k9)
						if t.nodeName="Ack" then
							if t.firstChild.nodeValue="Success" OR t.firstChild.nodeValue="SuccessWithWarning" then
								success=TRUE
								vsRESPMSG="Success"
							end if
						elseif t.nodeName="Errors" then
							themsg="" : thecode="" : iswarning=FALSE
							set ff=t.childNodes
							for kk=0 to ff.length - 1
								set gg=ff.item(kk)
								if gg.nodeName="LongMessage" then
									themsg=gg.firstChild.nodeValue
								elseif gg.nodeName="ErrorCode" then
									vsERRCODE=gg.firstChild.nodeValue
								elseif gg.nodeName="SeverityCode" then
									if gg.hasChildNodes then iswarning=(gg.firstChild.nodeValue="Warning")
								end if
							next
							if NOT iswarning then
								vsRESPMSG=themsg & "<br />" & vsRESPMSG
							end if
						elseif t.nodeName="DoCaptureResponseDetails" then
							set ff=t.getElementsByTagName("TransactionID")
							if ff.length > 0 then
								vsTRANSID=ff.Item(0).firstChild.nodeValue
							end if
						end if
					next
				end if
			next
			if success then
				if request.querystring("act")="charge" then
					ect_query("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&ordID)
					ect_query("UPDATE orders SET ordTransID='"&vsTRANSID&"' WHERE ordID="&ordID)
					ect_query("UPDATE orders SET ordStatus=3,ordStatusDate=" & vsusdatetime(thedate) & " WHERE ordStatus<3 AND ordID="&ordID)
				elseif request.querystring("act")="void" then
					ect_query("UPDATE orders SET ordTransID='void' WHERE ordID="&ordID)
					ect_query("UPDATE orders SET ordStatus=0,ordStatusDate=" & vsusdatetime(thedate) & " WHERE ordID="&ordID)
				end if
				print vsRESPMSG
			elseif vsERRCODE<>"" then
				print vsRESPMSG & " (" & vsERRCODE & ")"
			end if
		else
			print errormsg
		end if
	end if
elseif getget("processor")="Amazon" AND is_numeric(getget("gid")) then
	amazonstr=""
	isstripedotcom=TRUE
	function calculateStringToSignV2()
		calculateStringToSignV2="POST" & vbLf & scripturl & vbLf & endpointpath & vbLf & amazonstr
	end function
	function amazonparam2(nam, val)
		amazonstr=amazonstr & IIfVs(amazonstr<>"","&") & nam & "=" & replace(rawurlencode(replaceaccents(val)),"%7E","~")
	end function
	ordID=replace(getget("gid"),"'","")
	sSQL="SELECT ordPayProvider,ordAuthNumber,payProvData1,payProvData2,payProvData3,payProvDemo FROM orders INNER JOIN payprovider ON orders.ordPayProvider=payprovider.payProvID WHERE payProvEnabled=1 AND ordPayProvider=21 AND ordID=" & ordID
	rs.open sSQL,cnn,0,1
	authcode=rs("ordAuthNumber")
	data1=rs("payProvData1")
	data2=rs("payProvData2")
	data3=rs("payProvData3")
	demomode=rs("payProvDemo")
	rs.close
	scripturl="mws-eu.amazonservices.com"
	if origCountryCode="US" then scripturl="mws.amazonservices.com"
	if origCountryCode="JP" then scripturl="mws.amazonservices.jp"
	endpointpath="/OffAmazonPayments" & IIfVs(demomode,"_Sandbox") & "/2013-01-01"
	endpoint="https://" & scripturl & endpointpath
	if getget("act")="settle" then
		data2arr=split(data2,"&",2)
		if UBOUND(data2arr)>=0 then data2=data2arr(0)
		if UBOUND(data2arr)>0 then sellerid=data2arr(1)
		currdatetime=FormatDateTime(now(),0)
		call amazonparam2("AWSAccessKeyId",data2)
		call amazonparam2("Action","Capture")
		call amazonparam2("AmazonAuthorizationId",authcode)
		call amazonparam2("CaptureAmount.Amount",FormatNumber(getget("amount"),2,-1,0,0))
		call amazonparam2("CaptureAmount.CurrencyCode",countryCurrency)
		call amazonparam2("CaptureReferenceId",ordID&"_"&replace(replace(replace(currdatetime,"/",""),":","")," ",""))
		call amazonparam2("SellerCaptureNote","Capture: " & currdatetime)
		call amazonparam2("SellerId",sellerid)
		call amazonparam2("SignatureMethod","HmacSHA256")
		call amazonparam2("SignatureVersion",2)
		call amazonparam2("Timestamp",getutcdate(0))
		call amazonparam2("Version","2013-01-01")
		call amazonparam2("Signature",b64_hmac_sha256(data3,calculateStringToSignV2()))

		captureamount=0
		currentstate=""
		success=callxmlfunction(endpoint,amazonstr,res,"","WinHTTP.WinHTTPRequest.5.1",errormsg,FALSE)

		if res="" then
			print "<span style=""color:#FF0000"">" & "Error, couldn't update order " & ordID & "<br />" & errormsg & "</span><br/>"
		else
			set xmlDoc=Server.CreateObject("MSXML2.DOMDocument")
			xmlDoc.validateOnParse=False
			xmlDoc.loadXML(res)
			if NOT success then
				set obj2=xmlDoc.getElementsByTagName("Error")
				if obj2.length > 0 then
					set obj3=obj2.item(0).getElementsByTagName("Message")
					if obj3.length>0 then print "<span style=""color:#FF0000"">" & obj3.item(0).firstChild.nodeValue & "</span>" else print "<span style=""color:#FF0000"">" & errormsg & "</span>"
				else
					print "<span style=""color:#FF0000"">" & errormsg & "</span>"
				end if
			else
				set obj2=xmlDoc.getElementsByTagName("CaptureDetails")
				if obj2.length > 0 then
					for ix1=0 to obj2.item(0).childNodes.length - 1
						set tx1=obj2.item(0).childNodes.Item(ix1)
						if tx1.nodeName="CaptureAmount" then
							set obj3=tx1.getElementsByTagName("Amount")
							if obj3.length>0 then captureamount=obj3.item(0).firstChild.nodeValue
						elseif tx1.nodeName="CaptureStatus" then
							set obj3=tx1.getElementsByTagName("State")
							if obj3.length>0 then currentstate=obj3.item(0).firstChild.nodeValue
						end if
					next
					if currentstate="Completed" then
						ect_query("UPDATE orders SET ordAuthStatus='',ordStatusDate=" & vsusdatetime(thedate) & " WHERE ordPayProvider=21 AND ordID="&ordID)
					end if
					print "Captured: " & captureamount & ". The status of this capture is: " & currentstate
				end if
			end if
		end if
	end if
end if
cnn.Close
set cnn=nothing
set rs3=nothing
set rs2=nothing
set rs=nothing
%>