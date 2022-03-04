<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
Dim sSQL,rs,success,cnn,errmsg,rowcounter,startfont,endfont,smonth,allorders,addcomma,delOptions,ordAddInfo,ordstatussubject(10),ordstatusemail(10)
if storesessionvalue="" then storesessionvalue="virtualstore"
netnav=FALSE
if htmlemails=TRUE then emlNl="<br />" else emlNl=vbCrLf
if instr(Request.ServerVariables("HTTP_USER_AGENT"), "Gecko") > 0 then netnav=TRUE
if dateadjust="" then dateadjust=0
thedate=DateAdd("h",dateadjust,Now())
thedate=DateSerial(year(thedate),month(thedate),day(thedate))
themask=cStr(DateSerial(2003,12,11))
themask=replace(themask,"2003","yyyy")
themask=replace(themask,"12","mm")
themask=replace(themask,"11","dd")
if getget("doedit")="true" OR getget("id")="new" then doedit=TRUE else doedit=FALSE
isinvoice=(getget("invoice")="true")
if maxordersperpage="" then maxordersperpage=250
iNumOfPages=0
if NOT is_numeric(getget("pg")) then CurPage=1 else CurPage=vrmax(1, int(getget("pg")))
function trimoldcartitems(cartitemsdel)
	if dateadjust="" then dateadjust=0
	thetocdate=DateAdd("h",dateadjust,Now())
	sSQL="SELECT adminDelUncompleted,adminClearCart FROM admin WHERE adminID=1"
	rs.open sSQL,cnn,0,1
	delAfter=rs("adminDelUncompleted")
	delSavedCartAfter=rs("adminClearCart")
	rs.close
	if delAfter<>0 then
		sSQL="SELECT "&IIfVs(mysqlserver<>true,"TOP 1000 ")&"ordID FROM orders WHERE ordAuthNumber='' AND ordDate<" & vsusdate(thetocdate-delAfter) & " AND ordStatus=2"&IIfVs(mysqlserver=true," LIMIT 0,1000")
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			release_stock(rs("ordID"))
			ect_query("UPDATE cart SET cartOrderID=0 WHERE cartOrderID="&rs("ordID"))
			ect_query("DELETE FROM orders WHERE ordID="&rs("ordID"))
			rs.MoveNext
		loop
		rs.close
	end if
	sSQL="SELECT "&IIfVs(mysqlserver<>true,"TOP 1000 ")&"cartID,listOwner FROM cart LEFT JOIN customerlists ON cart.cartListID=customerlists.listID WHERE cartCompleted=0 AND cartOrderID=0 AND "
	sSQL=sSQL & "((cartClientID=0 AND cartDateAdded<" & vsusdatetime(cartitemsdel) & ") "
	if delSavedCartAfter<>0 then sSQL=sSQL & "OR (cartDateAdded<" & vsusdate(thetocdate-delSavedCartAfter) & ") "
	sSQL=sSQL & ")"&IIfVs(mysqlserver=true," LIMIT 0,1000")
	rs.open sSQL,cnn,0,1
	rowcounter=0
	if NOT rs.EOF then
		delOptions="" : addcomma=""
		do while NOT rs.EOF
			if NOT isnull(rs("listOwner")) then
				ect_query("UPDATE cart SET cartCompleted=3,cartClientID="&rs("listOwner")&" WHERE cartID="&rs("cartID"))
			else
				delOptions=delOptions & addcomma & rs("cartID")
				addcomma=","
			end if
			rowcounter=rowcounter+1
			rs.MoveNext
		loop
		if delOptions<>"" then ect_query("DELETE FROM cartoptions WHERE coCartID IN ("&delOptions&")")
		if delOptions<>"" then ect_query("DELETE FROM cart WHERE cartID IN ("&delOptions&")")
	end if
	rs.close
	trimoldcartitems=IIfVr(rowcounter>950,"","1")
end function
function editspecial(data,col,size,special)
	if doedit then editspecial="<input type=""text"" id="""&col&""" name="""&col&""" value=""" & htmlspecialsucode(data) & """ size="""&size&""" "&special&" />" else editspecial=htmldisplay(data&"")
end function
function editfunc(data,col,size)
	if doedit then editfunc="<input type=""text"" id="""&col&""" name="""&col&""" value=""" & htmlspecialsucode(data) & """ size="""&size&""" />" else editfunc=htmldisplay(data&"")
end function
function editnumeric(data,col,size,evt)
	if doedit then editnumeric="<input type=""text"" id="""&col&""" name="""&col&""" value="""&replace(FormatNumber(strip_tags2(data&""),2),",","")&""" size="""&size&""" "&evt&"/>" else editnumeric=FormatEuroCurrency(data)
end function
function getNumericField(fldname)
	fldval=getpost(fldname)
	if NOT is_numeric(fldval) then getNumericField=0 else getNumericField=cdbl(fldval)
end function
function decodehtmlentities(thestr)
	thestr=replace(thestr, "&quot;", """")
	thestr=replace(thestr, "&nbsp;", " ")
	decodehtmlentities=thestr
end function
sub writesearchparams()
	call writehiddenvar("fromdate", SESSION("fromdate"))
	call writehiddenvar("todate", SESSION("todate"))
	call writehiddenvar("notstatus", SESSION("notstatus"))
	call writehiddenvar("notsearchfield", SESSION("notsearchfield"))
	call writehiddenvar("searchtext", SESSION("searchtext"))
	call writehiddenvar("ordStatus", SESSION("ordStatus"))
	call writehiddenvar("ordstate", SESSION("ordstate"))
	call writehiddenvar("ordcountry", SESSION("ordcountry"))
	call writehiddenvar("payprovider", SESSION("payprovider"))
	call writehiddenvar("ordersearchfield", request.cookies("ordersearchfield"))
end sub
function showgetoptionsselect(oid)
	showgetoptionsselect="<div style=""position:absolute""><select id="""&oid&""" size=""15"" " & _
		"style=""display:none;position:absolute;min-width:280px;top:0px;left:0px;"" " & _
		"onblur=""this.style.display='none'"" " & _
		"onchange=""comboselect_onchange(this)"" " & _
		"onclick=""comboselect_onclick(this)"" " & _
		"onkeyup=""comboselect_onkeyup(event.keyCode,this)"">" & _
		"<option value="""">Populating...</option>" & _
		"</select></div>"
end function
sub getdates()
	if fromdate<>"" then
		hasfromdate=TRUE
		if is_numeric(fromdate) AND instr(fromdate, ".")=0 then
			thefromdate=(thedate-fromdate)
		else
			if isdate(fromdate) then
				thefromdate=datevalue(fromdate)
			else
				success=FALSE
				errmsg=yyDatInv & " - " & fromdate
				thefromdate=thedate
			end if
		end if
	else
		thefromdate=thedate
	end if
	if todate<>"" then
		hastodate=TRUE
		if is_numeric(todate) AND instr(todate, ".")=0 then
			thetodate=(thedate-todate)
		else
			if isdate(todate) then
				thetodate=datevalue(todate)
			else
				success=FALSE
				errmsg=yyDatInv & " - " & todate
				thetodate=thedate
			end if
		end if
	else
		thetodate=thedate
	end if
	if hasfromdate AND hastodate then
		if thefromdate > thetodate then
			tmpdate=thetodate
			thetodate=thefromdate
			thefromdate=tmpdate
		end if
	end if
end sub
sub getordsearchwheresql()
	dim searchtext,ordersearchfield,ordstatus,ordstate,ordcountry,payprovider
	editablefield=request.cookies("editablefield")
	searchtext=escape_string(getrequest("searchtext"))
	ordersearchfield=getrequest("ordersearchfield")
	if ordersearchfield<>"" then
		response.cookies("ordersearchfield")=ordersearchfield
		response.cookies("ordersearchfield").Expires=Date()+365
	end if
	if request.servervariables("HTTPS")="on" then response.cookies("ordersearchfield").secure=TRUE
	ordstatus=getrequest("ordStatus")
	ordstate=getrequest("ordstate")
	ordcountry=getrequest("ordcountry")
	payprovider=getrequest("payprovider")
	if ordersearchfield="product" AND searchtext<>"" then whereSQL=whereSQL & " INNER JOIN cart ON orders.ordID=cart.cartOrderID "
	if (ordersearchfield="ordid" OR ordersearchfield="") AND searchtext<>"" AND is_numeric(searchtext) then
		whereSQL=whereSQL & " WHERE ordID=" & searchtext & " "
	else
		if editablefield="abandoned" then
			if getpost("abandonedstatus")="recovered" then
				whereSQL=whereSQL & " WHERE ordStatus>2 AND ordAuthNumber<>''"
			else
				whereSQL=whereSQL & " WHERE ordStatus=2 AND ordAuthNumber=''"
			end if
		else
			if ordstatus<>"" then whereSQL=whereSQL & " WHERE " & IIfVr(getrequest("notstatus")="ON","NOT ","") & "(ordStatus IN (" & ordstatus & "))" else whereSQL=whereSQL & " WHERE ordStatus<>1"
		end if
		if ordstate<>"" then whereSQL=whereSQL & " AND " & IIfVr(getrequest("notsearchfield")="ON","NOT ","") & "(ordState IN ('" & replace(replace(escape_string(ordstate),", ",","),",","','") & "'))"
		if ordcountry<>"" then whereSQL=whereSQL & " AND " & IIfVr(getrequest("notsearchfield")="ON","NOT ","") & "(ordCountry IN ('" & replace(replace(escape_string(ordcountry),", ",","),",","','") & "'))"
		if payprovider<>"" then whereSQL=whereSQL & " AND " & IIfVr(getrequest("notsearchfield")="ON","NOT ","") & "(ordPayProvider IN ("&payprovider&")) "
		if hasfromdate then
			whereSQL=whereSQL & " AND ordDate BETWEEN " & vsusdate(thefromdate) & " AND " & vsusdate(IIfVr(hastodate, thetodate+1, thefromdate+1))
		elseif searchtext="" AND ordstatus="" AND ordstate="" AND ordcountry="" AND payprovider="" then
			whereSQL=whereSQL & " AND ordDate BETWEEN " & vsusdate(thedate) & " AND " & vsusdate(thedate+1)
		end if
		if searchtext<>"" then
			if ordersearchfield="ordid" OR ordersearchfield="" OR ordersearchfield="name" then
				if usefirstlastname then
					call splitfirstlastname(searchtext,firstname,lastname)
					if lastname="" then
						namesql="(ordName LIKE '%"&firstname&"%' OR ordLastName LIKE '%"&firstname&"%')"
					else
						namesql="(ordName LIKE '%"&firstname&"%' AND ordLastName LIKE '%"&lastname&"%')"
					end if
				else
					namesql="ordName LIKE '%"&searchtext&"%'"
				end if
			end if
			if ordersearchfield="ordid" OR ordersearchfield="" then
				whereSQL=whereSQL & " AND (ordEmail LIKE '%" & searchtext & "%' OR "&namesql&")"
			elseif ordersearchfield="email" then
				whereSQL=whereSQL & " AND ordEmail LIKE '%"&searchtext&"%'"
			elseif ordersearchfield="authcode" then
				whereSQL=whereSQL & " AND (ordAuthNumber LIKE '%"&searchtext&"%' OR ordTransID LIKE '%"&searchtext&"%')"
			elseif ordersearchfield="name" then
				whereSQL=whereSQL & " AND " & namesql
			elseif ordersearchfield="product" then
				whereSQL=whereSQL & " AND (cartProdID LIKE '%"&searchtext&"%' OR cartProdName LIKE '%"&searchtext&"%')"
			elseif ordersearchfield="address" then
				whereSQL=whereSQL & " AND (ordAddress LIKE '%"&searchtext&"%' OR ordAddress2 LIKE '%"&searchtext&"%' OR ordCity LIKE '%"&searchtext&"%' OR ordState LIKE '%"&searchtext&"%' OR ordShipAddress LIKE '%"&searchtext&"%' OR ordShipAddress2 LIKE '%"&searchtext&"%' OR ordShipCity LIKE '%"&searchtext&"%' OR ordShipState LIKE '%"&searchtext&"%')"
			elseif ordersearchfield="phone" then
				whereSQL=whereSQL & " AND ordPhone LIKE '%"&searchtext&"%'"
			elseif ordersearchfield="zip" then
				whereSQL=whereSQL & " AND ordZip LIKE '%"&searchtext&"%'"
			elseif ordersearchfield="invoice" then
				whereSQL=whereSQL & " AND ordInvoice LIKE '%"&searchtext&"%'"
			elseif ordersearchfield="tracknum" then
				whereSQL=whereSQL & " AND ordTrackNum LIKE '%"&searchtext&"%'"
			elseif ordersearchfield="affiliate" then
				whereSQL=whereSQL & " AND ordAffiliate='"&searchtext&"'"
			elseif ordersearchfield="extra1" then
				whereSQL=whereSQL & " AND ordExtra1 LIKE '%"&searchtext&"%'"
			elseif ordersearchfield="extra2" then
				whereSQL=whereSQL & " AND ordExtra2 LIKE '%"&searchtext&"%'"
			elseif ordersearchfield="checkout1" then
				whereSQL=whereSQL & " AND ordCheckoutExtra1 LIKE '%"&searchtext&"%'"
			elseif ordersearchfield="checkout2" then
				whereSQL=whereSQL & " AND ordCheckoutExtra2 LIKE '%"&searchtext&"%'"
			end if
		end if
	end if
end sub
set rs=Server.CreateObject("ADODB.RecordSet")
set rs2=Server.CreateObject("ADODB.RecordSet")
set rs3=Server.CreateObject("ADODB.RecordSet")
set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
success=TRUE
alreadygotadmin=getadminsettings()
homecountrytaxrate=countryTaxRate
if getpost("act")="abandoned" OR getpost("act")="oneabandoned" then
	if (getpost("act")="abandoned" AND getpost("ids")<>"") OR (getpost("act")="oneabandoned" AND getpost("id")<>"") then
		dim abandonedcartsubject(3),abandonedcartemail(3)
		replaceone=FALSE
		sSQL="SELECT abandonedcartsubject,abandonedcartsubject2,abandonedcartsubject3,abandonedcartemail,abandonedcartemail2,abandonedcartemail3 FROM emailmessages WHERE emailID=1"
		rs.open sSQL,cnn,0,1
		abandonedcartsubject(1)=trim(rs("abandonedcartsubject")&"")
		abandonedcartemail(1)=rs("abandonedcartemail")&""
		abandonedcartsubject(2)=trim(rs("abandonedcartsubject2")&"")
		abandonedcartemail(2)=rs("abandonedcartemail2")&""
		abandonedcartsubject(3)=trim(rs("abandonedcartsubject3")&"")
		abandonedcartemail(3)=rs("abandonedcartemail3")&""
		rs.close
		if getpost("act")="abandoned" then idlist=split(getpost("ids"),",") else idlist=split(getpost("id"),",")
		for each theid in idlist
			ordername=""
			sSQL="SELECT COUNT(*) AS tcount FROM abandonedcartemail WHERE aceOrderID=" & theid
			rs.open sSQL,cnn,0,1
			if rs.EOF then tcount=1 else tcount=rs("tcount")+1
			rs.close
			rs.open "SELECT ordStatus,ordAuthNumber,ordEmail,ordDate,ordName,ordLastName,ordLang FROM orders WHERE ordID="&theid,cnn,0,1
			if NOT rs.EOF then
				languageid=rs("ordLang")+1
				ordername=trim(rs("ordName")&" "&rs("ordLastName"))
				ordemail=rs("ordEmail")
				orddate=rs("ordDate")
			end if
			rs.close
			if (adminlangsettings AND 4096)=0 then languageid=1
			if abandonedcartsubject(languageid)<>"" then emailsubject=abandonedcartsubject(languageid)
			ose=abandonedcartemail(languageid)
			for uoindex=1 to 3
				ose=replaceemailtxt(ose,"%email" & uoindex & "%", IIfVr(uoindex=tcount,"%ectpreserve%",""),replaceone)
			next
			ose=replace(ose, "%ordername%", ordername)
			ose=replace(ose, "%orderdate%", FormatDateTime(orddate, 1) & " " & FormatDateTime(orddate, 4))
			acekey=calcmd5("ECT Abandoned Cart"&trim(theid)&":"&ordemail&":"&adminSecret)
			abandonedcartid=storeurl&"cart"&extension&"?acartid="&trim(theid)&"&acarthash="&acekey
			ose=replace(ose, "%abandonedcartid%", abandonedcartid)
			ose=replace(ose, "%nl%", emlNl)
			ose=replace(ose, "<br />", emlNl)
			if emailsubject<>"" AND ose<>"" then
				sSQL="INSERT INTO abandonedcartemail (aceOrderID,aceDateSent,aceKey) VALUES (" & theid & "," & vsusdate(date()) & ",'" & escape_string(acekey) & "')"
				cnn.execute(sSQL)
				call DoSendEmailEO(ordemail,emailAddr,"",emailsubject,ose,emailObject,themailhost,theuser,thepass)
			else
				print "BLANK EMAIL: NOT SENT<br />"
			end if
		next
	end if
	yyUpdSuc="Sending Emails..."
elseif getpost("updatestatus")="1" OR getpost("act")="status" then
	sSQL="SELECT orderstatussubject,orderstatussubject2,orderstatussubject3,orderstatusemail,orderstatusemail2,orderstatusemail3 FROM emailmessages WHERE emailID=1"
	rs.open sSQL,cnn,0,1
	ordstatussubject(1)=trim(rs("orderstatussubject")&"")
	ordstatusemail(1)=rs("orderstatusemail")&""
	ordstatussubject(2)=trim(rs("orderstatussubject2")&"")
	ordstatusemail(2)=rs("orderstatusemail2")&""
	ordstatussubject(3)=trim(rs("orderstatussubject3")&"")
	ordstatusemail(3)=rs("orderstatusemail3")&""
	rs.close
end if
if getpost("updatestatus")="1" AND is_numeric(getpost("orderid")) then
	ect_query("UPDATE orders SET ordStatusInfo='"&escape_string(getpost("ordStatusInfo"))&"',ordPrivateStatus='"&escape_string(getpost("ordPrivateStatus"))&"' WHERE ordID="&getpost("orderid"))
	ect_query("UPDATE orders set ordAuthNumber='" & escape_string(yyManAut) & "' WHERE ordStatus>=3 AND ordAuthNumber='' AND ordID=" & getpost("orderid"))
	ect_query("UPDATE cart SET cartCompleted=1 WHERE cartOrderID=" & getpost("orderid"))
	call updateorderstatus(getpost("orderid"), int(getpost("ordStatus")), getpost("emailstat")="1")
elseif is_numeric(getget("id")) AND getget("id")<>"multi" AND getget("id")<>"new" then
	if getpost("delccdets")<>"" then
		sSQL="UPDATE orders SET ordCNum='' WHERE ordID="&getget("id")
		ect_query(sSQL)
	end if
else
	delccafter=0
	if delccafter<>0 then ect_query("UPDATE orders SET ordCNum='' WHERE ordDate<" & vsusdate(thedate-delccafter))
	if persistentcart="" then persistentcart=3
	if SESSION("hasdeletedoldcart")<>"1" then SESSION("hasdeletedoldcart")=trimoldcartitems(thedate - persistentcart)
end if
sSQL="SELECT statID,statPrivate FROM orderstatus WHERE statPrivate<>'' ORDER BY statID"
rs.open sSQL,cnn,0,1
	allstatus=rs.GetRows
rs.close
if getpost("act")="delnotification" then
	sSQL="DELETE FROM devicenotifications WHERE dnID='" & escape_string(getpost("dnid")) & "'"
	cnn.execute(sSQL)
	print "<div style=""text-align:center;padding:50px"">Device Notification Deleted</div>"
	print "<meta http-equiv=""refresh"" content=""2; URL=adminorders.asp"" />"
elseif getget("act")="dnotif" then
	' Device Notifications
	sSQL="SELECT dnID FROM devicenotifications ORDER BY dnID"
	rs2.open sSQL,cnn,0,1
	if NOT rs2.EOF then
%>
<script>
function removenotification(dnid){
	if(confirm("Are you sure you want to remove this device notification?")){
		document.getElementById('dnid').value=dnid;
		document.getElementById('dnform').submit();
	}
}
</script>
<form method="post" action="adminorders.asp" id="dnform">
<input type="hidden" name="act" value="delnotification" />
<input type="hidden" name="dnid" id="dnid" value="" />
<h3 class="round_top half_top"><%="Device Notifications"%></h3>
<table class="admin-table-b keeptable">
	<tr>
		<th colspan="2">The following devices are registered to receive notifications for store sales</th>
	</tr>
<%		do while NOT rs2.EOF %>
	<tr>
		<td><%=left(rs2("dnID"),20)&"..."%></td>
		<td><input type="button" value="Remove Device Registration" onclick="removenotification('<%=jsescape(rs2("dnID"))%>')" /></td>
	</tr>
<%			rs2.movenext
		loop %>
</table>
</form>
<%
	end if
	rs2.close
elseif getpost("updatestatus")="1" OR getpost("act")="abandoned" OR getpost("act")="oneabandoned" then %>
		<form id="searchparamsform" method="post" action="adminorders.asp">
<%			writesearchparams() %>
            <div style="text-align:center;padding:30px;font-weight:bold"><%=yyUpdSuc%></div>
			<div style="text-align:center;padding:30px"><%=yyNowFrd%></div>
            <div style="text-align:center;padding:30px"><%=yyNoAuto%></div>
			<div style="text-align:center;padding-bottom:30px"><input type="submit" value="<%=yyClkHer%>"></div>
		</form>
<script>
setTimeout('document.getElementById("searchparamsform").submit()', 500);
</script>
<%
elseif getpost("doedit")="true" then
	session.LCID=1033
	OWSP=""
	orderid=getpost("orderid")
	ordstatus=int(getpost("ordStatus"))
	oldordstatus=0
	if orderid<>"new" then
		sSQL="SELECT ordSessionID,ordClientID,ordAuthStatus,ordShipType,loyaltyPoints,ordStatus FROM orders WHERE ordID=" & orderid
		rs.open sSQL,cnn,0,1
		thesessionid=rs("ordSessionID")
		thecustomerid=rs("ordClientID")
		loyaltypointtotal=rs("loyaltyPoints")
		oldordstatus=rs("ordStatus")
		ordAuthStatus=rs("ordAuthStatus")
		ordShipType=rs("ordShipType")
		rs.close
		if oldordstatus>=2 AND getpost("updatestock")="ON" then release_stock(orderid)
		if thecustomerid<>0 AND loyaltypoints<>"" AND loyaltypointtotal<>0 AND oldordstatus>=3 then ect_query("UPDATE customerlogin SET loyaltyPoints=loyaltyPoints-" & loyaltypointtotal & " WHERE clID=" & thecustomerid)
		if ordAuthStatus="MODWARNOPEN" OR isnull(ordAuthStatus) then ect_query("UPDATE orders SET ordAuthStatus='' WHERE ordID=" & orderid)
		if ordShipType="MODWARNOPEN" OR isnull(ordShipType) then ect_query("UPDATE orders SET ordShipType='' WHERE ordID=" & orderid)
	end if
	ordComLoc=0
	if getpost("commercialloc")="Y" then ordComLoc=1
	if getpost("wantinsurance")="Y" then ordComLoc=ordComLoc+2
	if getpost("saturdaydelivery")="Y" then ordComLoc=ordComLoc+4
	if getpost("signaturerelease")="Y" then ordComLoc=ordComLoc+8
	if getpost("insidedelivery")="Y" then ordComLoc=ordComLoc+16
	discounttext=replace(getpost("discounttext"),vbCrLf,"<br />")
	discounttext=replace(discounttext,vbCr,"<br />")
	discounttext=replace(discounttext,vbLf,"<br />")
	orddate=DateAdd("h",dateadjust,Now())
	if getpost("doeditdate")="1" then
		neworddate=getpost("editorddate") & " " & getpost("editordtime")
		if isdate(neworddate) then orddate=datevalue(neworddate)+timevalue(neworddate)
	end if
	ordcountry="" : ordshipcountry=""
	if is_numeric(getpost("country")) then
		sSQL="SELECT countryName FROM countries WHERE countryID=" & escape_string(getpost("country"))
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then ordcountry=rs("countryName")
		rs.close
	end if
	if is_numeric(getpost("scountry")) then
		sSQL="SELECT countryName FROM countries WHERE countryID=" & escape_string(getpost("scountry"))
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then ordshipcountry=rs("countryName")
		rs.close
	end if
	if sqlserver=TRUE then ' {
		ordauthnumber=getpost("ordAuthNumber")
		if ordstatus>2 AND ordauthnumber="" then ordauthnumber="manual auth"
		if orderid="new" then
			thesessionid="A1"
			thecustomerid=0
			sSQL="INSERT INTO orders (ordSessionID,ordClientID,ordName,ordLastName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordShipName,ordShipLastName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,ordShipPhone,ordPayProvider,ordAuthNumber,ordTransID,ordShipping,ordStateTax,ordCountryTax,ordHSTTax,ordLang,ordHandling,ordShipType,ordShipCarrier,loyaltyPoints,ordTotal,ordDate,ordStatusInfo,ordPrivateStatus,ordStatus,ordAuthStatus,ordStatusDate,ordComLoc,ordIP,ordAffiliate,ordExtra1,ordExtra2,ordShipExtra1,ordShipExtra2,ordCheckoutExtra1,ordCheckoutExtra2,ordDiscount,ordDiscountText,ordInvoice,ordAddInfo) VALUES (" & _
				"'" & thesessionid & "'," & _
				"'" & thecustomerid & "'," & _
				"'" & escape_string(getpost("name")) & "'," & _
				"'" & escape_string(getpost("lastname")) & "'," & _
				"'" & escape_string(getpost("address")) & "'," & _
				"'" & escape_string(getpost("address2")) & "'," & _
				"'" & escape_string(getpost("city")) & "'," & _
				"'" & escape_string(getpost("state")) & "'," & _
				"'" & escape_string(getpost("zip")) & "'," & _
				"'" & escape_string(ordcountry) & "'," & _
				"'" & escape_string(getpost("email")) & "'," & _
				"'" & escape_string(getpost("phone")) & "'," & _
				"'" & escape_string(getpost("sname")) & "'," & _
				"'" & escape_string(getpost("slastname")) & "'," & _
				"'" & escape_string(getpost("saddress")) & "'," & _
				"'" & escape_string(getpost("saddress2")) & "'," & _
				"'" & escape_string(getpost("scity")) & "'," & _
				"'" & escape_string(getpost("sstate")) & "'," & _
				"'" & escape_string(getpost("szip")) & "'," & _
				"'" & escape_string(ordshipcountry) & "'," & _
				"'" & escape_string(getpost("sphone")) & "'," & _
				"4," & _
				"'" & escape_string(getpost("ordAuthNumber")) & "'," & _
				"'" & escape_string(getpost("ordTransID")) & "'," & _
				getNumericField("ordShipping") & "," & _
				getNumericField("ordStateTax") & "," & _
				getNumericField("ordCountryTax") & "," & _
				IIfVr(origCountryID=2,getNumericField("ordHSTTax"),0) & "," & _
				IIfVr(is_numeric(getpost("ordlang")),getpost("ordlang"),0) & "," & _
				getNumericField("ordHandling") & "," & _
				"'" & escape_string(getpost("shipmethod")) & "'," & _
				"'" & escape_string(getpost("shipcarrier")) & "'," & _
				"'" & getNumericField("loyaltyPoints") & "'," & _
				"'" & escape_string(getpost("ordtotal")) & "'," & _
				vsusdatetime(orddate) & "," & _
				"'" & escape_string(getpost("ordStatusInfo")) & "'," & _
				"'" & escape_string(getpost("ordPrivateStatus")) & "'," & _
				getNumericField("ordstatus") & ",''," & _
				vsusdatetime(DateAdd("h",dateadjust,Now())) & "," & _
				"'" & ordComLoc & "'," & _
				"'" & escape_string(getpost("ipaddress")) & "'," & _
				"'" & getpost("PARTNER") & "'," & _
				"'" & escape_string(getpost("extra1")) & "'," & _
				"'" & escape_string(getpost("extra2")) & "'," & _
				"'" & escape_string(getpost("shipextra1")) & "'," & _
				"'" & escape_string(getpost("shipextra2")) & "'," & _
				"'" & escape_string(getpost("checkoutextra1")) & "'," & _
				"'" & escape_string(getpost("checkoutextra2")) & "'," & _
				getNumericField("ordDiscount") & "," & _
				"'" & escape_string(discounttext) & "'," & _
				"'" & escape_string(getpost("ordInvoice")) & "'," & _
				"'" & escape_string(getpost("ordAddInfo")) & "')"
			ect_query(sSQL)
			rs.open "SELECT @@IDENTITY AS lstIns",cnn,0,1
			orderid=int(cstr(rs("lstIns")))
			rs.close
		else
			sSQL="UPDATE orders SET " & _
				"ordName='" & escape_string(getpost("name")) & "'," & _
				"ordLastName='" & escape_string(getpost("lastname")) & "'," & _
				"ordAddress='" & escape_string(getpost("address")) & "'," & _
				"ordAddress2='" & escape_string(getpost("address2")) & "'," & _
				"ordCity='" & escape_string(getpost("city")) & "'," & _
				"ordState='" & escape_string(getpost("state")) & "'," & _
				"ordZip='" & escape_string(getpost("zip")) & "'," & _
				"ordCountry='" & escape_string(ordcountry) & "'," & _
				"ordEmail='" & escape_string(getpost("email")) & "'," & _
				"ordPhone='" & escape_string(getpost("phone")) & "'," & _
				"ordShipName='" & escape_string(getpost("sname")) & "'," & _
				"ordShipLastName='" & escape_string(getpost("slastname")) & "'," & _
				"ordShipAddress='" & escape_string(getpost("saddress")) & "'," & _
				"ordShipAddress2='" & escape_string(getpost("saddress2")) & "'," & _
				"ordShipCity='" & escape_string(getpost("scity")) & "'," & _
				"ordShipState='" & escape_string(getpost("sstate")) & "'," & _
				"ordShipZip='" & escape_string(getpost("szip")) & "'," & _
				"ordShipCountry='" & escape_string(ordshipcountry) & "'," & _
				"ordShipPhone='" & escape_string(getpost("sphone")) & "'," & _
				"ordShipType='" & escape_string(getpost("shipmethod")) & "'," & _
				"ordShipCarrier='" & escape_string(getpost("shipcarrier")) & "'," & _
				"ordIP='" & escape_string(getpost("ipaddress")) & "'," & _
				"ordComLoc=" & ordComLoc & "," & _
				"ordAffiliate='" & getpost("PARTNER") & "'," & _
				"ordAddInfo='" & escape_string(getpost("ordAddInfo")) & "'," & _
				"ordStatusInfo='" & escape_string(getpost("ordStatusInfo")) & "'," & _
				"ordPrivateStatus='" & escape_string(getpost("ordPrivateStatus")) & "'," & _
				"ordStatus=" & getNumericField("ordstatus") & ","
			if getpost("doeditdate")="1" then sSQL=sSQL & "ordDate=" & vsusdatetime(orddate) & ","
			sSQL=sSQL & "ordTrackNum='" & escape_string(getpost("ordTrackNum")) & "'," & _
				"ordDiscountText='" & escape_string(discounttext) & "'," & _
				"ordInvoice='" & escape_string(getpost("ordInvoice")) & "'," & _
				"ordExtra1='" & escape_string(getpost("extra1")) & "'," & _
				"ordExtra2='" & escape_string(getpost("extra2")) & "'," & _
				"ordShipExtra1='" & escape_string(getpost("shipextra1")) & "'," & _
				"ordShipExtra2='" & escape_string(getpost("shipextra2")) & "'," & _
				"ordCheckoutExtra1='" & escape_string(getpost("checkoutextra1")) & "'," & _
				"ordCheckoutExtra2='" & escape_string(getpost("checkoutextra2")) & "'," & _
				"ordShipping=" & getNumericField("ordShipping") & "," & _
				"ordStateTax=" & getNumericField("ordStateTax") & "," & _
				"ordCountryTax=" & getNumericField("ordCountryTax") & "," & _
				IIfVs(origCountryID=2,"ordHSTTax=" & getNumericField("ordHSTTax") & ",") & _
				"ordLang=" & IIfVr(is_numeric(getpost("ordlang")),getpost("ordlang"),0) & "," & _
				"ordDiscount=" & getNumericField("ordDiscount") & "," & _
				"ordHandling=" & getNumericField("ordHandling") & "," & _
				"ordAuthNumber='" & escape_string(ordauthnumber) & "'," & _
				"ordTransID='" & escape_string(getpost("ordTransID")) & "'," & _
				"loyaltyPoints='" & getNumericField("loyaltyPoints") & "'," & _
				"ordTotal=" & getNumericField("ordtotal") & " WHERE ordID=" & getNumericField("orderid")
			ect_query(sSQL)
		end if
	else ' }{ sqlserver=TRUE
		if orderid="new" then
			rs.open "orders",cnn,1,3,&H0002
			rs.AddNew
			thesessionid="A1"
			thecustomerid=0
			rs.Fields("ordDate")		= orddate
			rs.Fields("ordStatusDate")	= DateAdd("h",dateadjust,Now())
			rs.Fields("ordPayProvider")	= 4
			rs.Fields("ordAuthStatus")	= ""
			rs.Fields("ordSessionID")	= thesessionid
		else
			if mysqlserver then rs.CursorLocation=3
			rs.open "SELECT * FROM orders WHERE ordID="&orderid,cnn,1,3,&H0001
		end if
		if is_numeric(getpost("custid")) then
			rs.Fields("ordClientID")=getpost("custid")
			thecustomerid=getpost("custid")
		end if
		rs.Fields("ordLang")		= IIfVr(is_numeric(getpost("ordlang")),getpost("ordlang"),0)
		rs.Fields("ordName")		= getpost("name")
		rs.Fields("ordLastName")	= getpost("lastname")
		rs.Fields("ordAddress")		= getpost("address")
		rs.Fields("ordAddress2")	= getpost("address2")
		rs.Fields("ordCity")		= getpost("city")
		rs.Fields("ordState")		= getpost("state")
		rs.Fields("ordZip")			= getpost("zip")
		rs.Fields("ordCountry")		= ordcountry
		rs.Fields("ordEmail")		= getpost("email")
		rs.Fields("ordPhone")		= getpost("phone")
		rs.Fields("ordShipName")	= getpost("sname")
		rs.Fields("ordShipLastName")= getpost("slastname")
		rs.Fields("ordShipAddress")	= getpost("saddress")
		rs.Fields("ordShipAddress2")= getpost("saddress2")
		rs.Fields("ordShipCity")	= getpost("scity")
		rs.Fields("ordShipState")	= getpost("sstate")
		rs.Fields("ordShipZip")		= getpost("szip")
		rs.Fields("ordShipCountry")	= ordshipcountry
		rs.Fields("ordShipPhone")	= getpost("sphone")
		rs.Fields("ordShipType")	= getpost("shipmethod")
		rs.Fields("ordShipCarrier")	= getpost("shipcarrier")
		rs.Fields("ordIP")			= getpost("ipaddress")
		rs.Fields("ordComLoc")		= ordComLoc
		rs.Fields("ordAffiliate")	= getpost("PARTNER")
		rs.Fields("ordAddInfo")		= getpost("ordAddInfo")
		rs.Fields("ordStatusInfo")	= getpost("ordStatusInfo")
		rs.Fields("ordPrivateStatus")	= getpost("ordPrivateStatus")
		rs.Fields("ordStatus")		= getpost("ordStatus")
		rs.Fields("ordTrackNum")	= getpost("ordTrackNum")
		rs.Fields("ordDiscountText")= discounttext
		rs.Fields("ordInvoice")		= getpost("ordInvoice")
		rs.Fields("ordExtra1")		= getpost("extra1")
		rs.Fields("ordExtra2")		= getpost("extra2")
		rs.Fields("ordShipExtra1")	= getpost("shipextra1")
		rs.Fields("ordShipExtra2")	= getpost("shipextra2")
		rs.Fields("ordCheckoutExtra1")	= getpost("checkoutextra1")
		rs.Fields("ordCheckoutExtra2")	= getpost("checkoutextra2")
		rs.Fields("ordShipping")	= getNumericField("ordShipping")
		if origCountryID=2 then rs.Fields("ordHSTTax")=getNumericField("ordHSTTax")
		rs.Fields("ordStateTax")	= getNumericField("ordStateTax")
		rs.Fields("ordCountryTax")	= getNumericField("ordCountryTax")
		rs.Fields("ordDiscount")	= getNumericField("ordDiscount")
		rs.Fields("ordHandling")	= getNumericField("ordHandling")
		ordauthnumber=getpost("ordAuthNumber")
		if ordstatus>2 AND ordauthnumber="" then ordauthnumber="manual auth"
		rs.Fields("ordAuthNumber")	= ordauthnumber
		rs.Fields("ordTransID")		= getpost("ordTransID")
		rs.Fields("ordTotal")		= getNumericField("ordtotal")
		rs.Fields("loyaltyPoints")	= getNumericField("loyaltyPoints")
		rs.Update
		if orderid="new" then
			if mysqlserver=TRUE then
				rs.close
				rs.open "SELECT LAST_INSERT_ID() AS lstIns",cnn,0,1
				orderid=rs("lstIns")
			else
				orderid=rs.Fields("ordID")
			end if
		end if
		rs.close
	end if ' }
	if ordstatus>2 then ect_query("UPDATE giftcertificate SET gcAuthorized=1 WHERE gcOrderID=" & orderid)
	if ordstatus<=2 then ect_query("UPDATE giftcertificate SET gcAuthorized=0 WHERE gcOrderID=" & orderid)
	if thecustomerid<>0 AND loyaltypoints<>"" AND ordstatus>=3 then ect_query("UPDATE customerlogin SET loyaltyPoints=loyaltyPoints+" & getNumericField("loyaltyPoints") & " WHERE clID=" & thecustomerid)
	Dim forminorder()
	redim forminorder(100)
	formitemcnt=0
	for jj=1 to request.form.Count
		for each objElem in request.form
			if request.form(objElem) is request.form(jj) then
				if Left(objElem,6)="prodid" OR Left(objElem,4)="optn" then
					forminorder(formitemcnt)=objElem
					formitemcnt=formitemcnt + 1
					forminorderubound=UBOUND(forminorder)
					if formitemcnt > forminorderubound then redim preserve forminorder(forminorderubound+100)
				end if
				exit for
			end if
		next
	next
	for jj=0 to formitemcnt-1
		objForm=forminorder(jj)
		'print objForm & " : " & getpost(objForm) & "<br />"
		if Left(objForm,6)="prodid" then
			idno=trim(right(objForm, Len(objForm)-6))
			cartid=getpost("cartid"&idno)
			prodid=getpost("prodid"&idno)
			quant=getpost("quant"&idno)
			if NOT is_numeric(quant) then quant=1
			theprice=getpost("price"&idno)
			if NOT is_numeric(theprice) then theprice=0
			prodname=getpost("prodname"&idno)
			delitem=getpost("del_"&idno)
			sSQL="SELECT pWeight FROM products WHERE pID='"&escape_string(prodid)&"'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then thepweight=rs("pWeight") else thepweight=0
			rs.close
			if delitem="yes" OR (cartid<>"" AND trim(prodid)="") then
				ect_query("DELETE FROM cart WHERE cartID=" & cartid)
				ect_query("DELETE FROM cartoptions WHERE coCartID=" & cartid)
				cartid=""
			elseif cartid<>"" then
				Session.LCID=1033
				sSQL="UPDATE cart SET cartProdID='"&escape_string(prodid)&"',cartProdPrice="&theprice&",cartProdName='"&escape_string(prodname)&"',cartQuantity="&quant&",cartCompleted=0 WHERE cartID="&cartid
				ect_query(sSQL)
				Session.LCID=saveLCID
				ect_query("DELETE FROM cartoptions WHERE coCartID=" & cartid)
			else
				if sqlserver=TRUE then
					Session.LCID=1033
					sSQL="INSERT INTO cart (cartSessionID,cartClientID,cartProdID,cartOrigProdID,cartQuantity,cartCompleted,cartProdName,cartProdPrice,cartOrderID,cartDateAdded) VALUES (" & _
					"'" & escape_string(thesessionid) & "'," & thecustomerid & ",'" & escape_string(prodid) & "',''," & quant & ",0,'" & escape_string(prodname) & "'," & vsround(theprice,2) & "," & orderid & "," & vsusdate(DateAdd("h",dateadjust,Now())) & ")"
					ect_query(sSQL)
					rs.open "SELECT @@IDENTITY AS lstIns",cnn,0,1
					cartid=int(cstr(rs("lstIns")))
					rs.close
					Session.LCID=saveLCID
				else
					rs.open "cart",cnn,1,3,&H0002
					rs.AddNew
					rs.Fields("cartSessionID")		= thesessionid
					rs.Fields("cartClientID")		= thecustomerid
					rs.Fields("cartProdID")			= prodid
					rs.Fields("cartQuantity")		= quant
					rs.Fields("cartCompleted")		= 0
					rs.Fields("cartProdName")		= prodname
					rs.Fields("cartProdPrice")		= theprice
					rs.Fields("cartOrderID")		= orderid
					rs.Fields("cartDateAdded")		= DateAdd("h",dateadjust,Now())
					rs.Update
					if mysqlserver=TRUE then
						rs.close
						rs.open "SELECT LAST_INSERT_ID() AS lstIns",cnn,0,1
						cartid=rs("lstIns")
					else
						cartid=rs.Fields("cartID")
					end if
					rs.close
				end if
			end if
			if cartid<>"" then
				if ordstatus<>2 then ect_query("UPDATE cart SET cartCompleted=1 WHERE cartID="&cartid)
				optprefix="optn"&idno&"_"
				prefixlen=len(optprefix)
				for kk=0 to formitemcnt-1
					objForm=forminorder(kk)
					if left(objForm,prefixlen)=optprefix AND getpost(objForm)<>"" then
						optidarr=split(getpost(objForm),"|")
						optid=optidarr(0)
						if getpost("v"&objForm)="" then
							sSQL="SELECT optID,"&getlangid("optGrpName",16)&","&getlangid("optName",32)&","&OWSP&"optPriceDiff,optWeightDiff,optType,optFlags FROM options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optID="&Replace(optid,"'","")
							rs.open sSQL,cnn,0,1
							if abs(rs("optType"))<>3 then
								sSQL="INSERT INTO cartoptions (coCartID,coOptID,coOptGroup,coCartOption,coPriceDiff,coWeightDiff) VALUES ("&cartID&","&rs("optID")&",'"&escape_string(rs(getlangid("optGrpName",16)))&"','"&escape_string(rs(getlangid("optName",32)))&"',"
								sSQL=sSQL & optidarr(1) & ","
								if (rs("optFlags") AND 2)=0 then sSQL=sSQL & rs("optWeightDiff") & ")" else sSQL=sSQL & ((thepweight*rs("optWeightDiff"))/100.0) & ")"
							else
								sSQL="INSERT INTO cartoptions (coCartID,coOptID,coOptGroup,coCartOption,coPriceDiff,coWeightDiff) VALUES ("&cartID&","&rs("optID")&",'"&escape_string(rs(getlangid("optGrpName",16)))&"','',0,0)"
							end if
							rs.close
							ect_query(sSQL)
						else
							sSQL="SELECT optID,"&getlangid("optGrpName",16)&","&getlangid("optName",32)&",optTxtCharge,optMultiply,optAcceptChars FROM options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optID="&replace(optid,"'","")
							rs.open sSQL,cnn,0,1
							if NOT rs.EOF then
								theopttoadd=getpost("v"&objForm)
								optPriceDiff=IIfVr(rs("optTxtCharge")<0 AND theopttoadd<>"",abs(rs("optTxtCharge")),rs("optTxtCharge")*len(theopttoadd))
								optmultiply=0
								if rs("optMultiply")<>0 then
									if is_numeric(theopttoadd) then optmultiply=cdbl(theopttoadd) else theopttoadd="#NAN"
								end if
								sSQL="INSERT INTO cartoptions (coCartID,coOptID,coOptGroup,coCartOption,coPriceDiff,coWeightDiff,coMultiply) VALUES ("&cartID&","&rs("optID")&",'"&escape_string(rs(getlangid("optGrpName",16)))&"','"&escape_string(getpost("v"&objForm))&"',"&optPriceDiff&",0," & IIfVr(rs("optMultiply")<>0,1,0) & ")"
								ect_query(sSQL)
							end if
							rs.close
						end if
					end if
				next
			end if
		end if
	next
	if ordstatus>=2 AND getpost("updatestock")="ON" then stock_subtract(orderid)
	session.LCID=saveLCID %>
		<form id="searchparamsform" method="post" action="adminorders.asp<% if getget("view")="true" then print "?id=" & getpost("orderid") %>">
<%			writesearchparams()
			if getpost("orderid")<>"new" then call writehiddenvar("ctrlmod", 2) %>
			<div style="text-align:center;padding:30px;font-weight:bold"><%=yyUpdSuc%></div>
			<div style="text-align:center;padding:30px"><%=yyNowFrd%></div>
            <div style="text-align:center;padding:30px"><%=yyNoAuto%></div>
			<div style="text-align:center;padding-bottom:30px"><input type="submit" value="<%=yyClkHer%>"></div>
		</form>
<script>
setTimeout('document.getElementById("searchparamsform").submit()', 500);
</script>
<%
elseif getget("id")<>"" then
	if getget("id")="new" then
		idlist=array("0")
	elseif getget("id")="multi" then
		idlist=split(getpost("ids"), ",")
	else
		idlist=split(getget("id"), ",")
	end if
	numids=UBOUND(idlist)
	numorders=0
	if getget("id")<>"multi" AND getget("id")<>"new" AND is_numeric(getget("id")) then
		print "<form method=""post"" action=""adminorders.asp"" id=""researchform"">"
		call writehiddenvar("fromdate", getpost("fromdate"))
		call writehiddenvar("todate", getpost("todate"))
		call writehiddenvar("notstatus", getpost("notstatus"))
		call writehiddenvar("notsearchfield", getpost("notsearchfield"))
		call writehiddenvar("searchtext", getpost("searchtext"))
		call writehiddenvar("ordStatus", getpost("ordStatus"))
		call writehiddenvar("ordstate", getpost("ordstate"))
		call writehiddenvar("ordcountry", getpost("ordcountry"))
		call writehiddenvar("payprovider", getpost("payprovider"))
		hastodate=FALSE : hasfromdate=FALSE
		todate=getrequest("todate") : fromdate=getrequest("fromdate")
		thetodate=thedate : thefromdate=thedate
		call getdates()
		origsearchtext=htmlspecials(getrequest("searchtext"))
		sSQL="SELECT "&IIfVs(mysqlserver<>true,"TOP 1 ")&"ordID FROM (orders INNER JOIN payprovider ON payprovider.payProvID=orders.ordPayProvider)"
		whereSQL="" : namesql=""
		call getordsearchwheresql()
		sSQL=sSQL&whereSQL&" AND ordID<"&getget("id")&" ORDER BY ordID DESC"&IIfVs(mysqlserver=true," LIMIT 0,1")
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then previousbysearch=rs("ordID") else previousbysearch=""
		rs.close
		sSQL="SELECT "&IIfVs(mysqlserver<>true,"TOP 1 ")&"ordID FROM (orders INNER JOIN payprovider ON payprovider.payProvID=orders.ordPayProvider)"
		whereSQL="" : namesql=""
		call getordsearchwheresql()
		sSQL=sSQL&whereSQL&" AND ordID>"&getget("id")&" ORDER BY ordID ASC"&IIfVs(mysqlserver=true," LIMIT 0,1")
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then nextbysearch=rs("ordID") else nextbysearch=""
		rs.close
		print "</form>"
	end if
	for each theid in idlist
		noloyaltypoints=FALSE : orderstatetaxexempt=FALSE : ordercountrytaxexempt=FALSE
		numids=numids-1
		allorders=""
		statetaxrate=0
		countrytaxrate=0
		hsttaxrate=0
		countryorder=0
		if getget("id")="new" then
			ordStatus=3 : ordAuthStatus="" : ordStatusDate=Now() : ordID="" : ordName="" : ordLastName="" : ordAddress="" : ordAddress2="" : ordCity="" : ordState="" : ordZip="" : ordCountry="" : ordEmail="" : ordPhone="" : ordShipName="" : ordShipLastName="" : ordShipAddress="" : ordShipAddress2="" : ordShipCity="" : ordShipState="" : ordShipZip="" : ordShipCountry="" : ordShipPhone="" : ordPayProvider=0 : ordAuthNumber="manual auth" : ordTransID="" : ordTotal=0 : ordDate=now() : ordStateTax=0 : ordCountryTax=0 : ordShipping=0 : ordShipType="" : ordShipCarrier=0 : ordIP=left(request.servervariables("REMOTE_ADDR"), 48) : ordAffiliate="" : ordDiscount=0 : ordDiscountText="" : ordHandling=0 : ordComLoc=0 : ordExtra1="" : ordExtra2="" : ordShipExtra1="" : ordShipExtra2="" : ordCheckoutExtra1="" : ordCheckoutExtra2="" : ordHSTTax=0 : ordTrackNum="" : ordInvoice="" : ordClientID=0 : ordReferer="" : ordUserAgent="" : ordAddInfo=""
		else
			if NOT is_numeric(theid) then theid=-1
			if viewordersort="" then viewordersort="cartID"
			if isprinter AND packingslipsort<>"" then viewordersort=packingslipsort
			if isinvoice AND invoicesort<>"" then viewordersort=invoicesort
			sSQL="SELECT cartProdId,cartProdName,cartProdPrice,cartQuantity,cartID,pStockByOpts,pExemptions,cartGiftWrap,cartGiftMessage,pWeight,cartOrigProdID FROM cart LEFT JOIN products on cart.cartProdID=products.pId WHERE cartOrderID="&theid&" ORDER BY " & viewordersort
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				allorders=rs.getrows
				numorders=UBOUND(allorders, 2)+1
			end if
			rs.close
			sSQL="SELECT ordID,ordStatus,ordAuthStatus,ordStatusDate,ordName,ordLastName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordShipName,ordShipLastName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,ordShipPhone,ordPayProvider,ordAuthNumber,ordTransID,ordTotal,ordDate,ordStateTax,ordCountryTax,ordShipping,ordShipType,ordShipCarrier,ordIP,ordAffiliate,ordDiscount,ordDiscountText,ordHandling,ordComLoc,ordExtra1,ordExtra2,ordShipExtra1,ordShipExtra2,ordCheckoutExtra1,ordCheckoutExtra2,ordHSTTax,ordTrackNum,ordInvoice,ordClientID,ordReferer,ordUserAgent,ordQuerystr,loyaltyPoints,ordLang,ordAddInfo FROM orders INNER JOIN payprovider ON payprovider.payProvID=orders.ordPayProvider WHERE ordID="&theid
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				ordID=rs("ordID") : ordStatus=rs("ordStatus") : ordAuthStatus=rs("ordAuthStatus") : ordStatusDate=rs("ordStatusDate") : ordName=rs("ordName") : ordLastName=trim(rs("ordLastName")&"") : ordAddress=rs("ordAddress") : ordAddress2=rs("ordAddress2") : ordCity=rs("ordCity") : ordState=rs("ordState") : ordZip=rs("ordZip") : ordCountry=rs("ordCountry") : ordEmail=rs("ordEmail") : ordPhone=rs("ordPhone") : ordShipName=rs("ordShipName") : ordShipLastName=rs("ordShipLastName") : ordShipAddress=rs("ordShipAddress") : ordShipAddress2=rs("ordShipAddress2") : ordShipCity=rs("ordShipCity") : ordShipState=rs("ordShipState") : ordShipZip=rs("ordShipZip") : ordShipCountry=rs("ordShipCountry") : ordShipPhone=rs("ordShipPhone") : ordPayProvider=rs("ordPayProvider") : ordAuthNumber=rs("ordAuthNumber") : ordTransID=rs("ordTransID") : ordTotal=rs("ordTotal") : ordDate=rs("ordDate") : ordStateTax=rs("ordStateTax") : ordCountryTax=rs("ordCountryTax") : ordShipping=rs("ordShipping") : ordShipType=rs("ordShipType") : ordShipCarrier=rs("ordShipCarrier") : ordIP=rs("ordIP") : ordAffiliate=rs("ordAffiliate") : ordDiscount=rs("ordDiscount") : ordDiscountText=rs("ordDiscountText") : ordHandling=rs("ordHandling") : ordComLoc=rs("ordComLoc") : ordExtra1=rs("ordExtra1") : ordExtra2=rs("ordExtra2") : ordShipExtra1=rs("ordShipExtra1") : ordShipExtra2=rs("ordShipExtra2") : ordCheckoutExtra1=rs("ordCheckoutExtra1") : ordCheckoutExtra2=rs("ordCheckoutExtra2") : ordHSTTax=rs("ordHSTTax") : ordTrackNum=rs("ordTrackNum") : ordInvoice=rs("ordInvoice") : ordClientID=rs("ordClientID") : ordReferer=rs("ordReferer") : ordUserAgent=trim(rs("ordUserAgent")&"") : ordQuerystr=rs("ordQuerystr") : loyaltypointtotal=rs("loyaltyPoints") : ordlang=rs("ordLang") : ordAddInfo=trim(rs("ordAddInfo")&"")
			else
				ordID="Invalid Order ID" : ordStatus=0 : ordAuthStatus="" : ordStatusDate=Date() : ordName="&nbsp;" : ordLastName="" : ordAddress="" : ordAddress2="" : ordCity="" : ordState="" : ordZip="" : ordCountry="" : ordEmail="" : ordPhone="" : ordShipName="" : ordShipLastName="" : ordShipAddress="" : ordShipAddress2="" : ordShipCity="" : ordShipState="" : ordShipZip="" : ordShipCountry="" : ordShipPhone="" : ordPayProvider=0 : ordAuthNumber="" : ordTransID="" : ordTotal=0 : ordDate=dateserial(1970,1,1) : ordStateTax=0 : ordCountryTax=0 : ordShipping=0 : ordShipType="" : ordShipCarrier=0 : ordIP="" : ordAffiliate="" : ordDiscount=0 : ordDiscountText="" : ordHandling=0 : ordComLoc=0 : ordExtra1="" : ordExtra2="" : ordShipExtra1="" : ordShipExtra2="" : ordCheckoutExtra1="" : ordCheckoutExtra2="" : ordHSTTax=0 : ordTrackNum="" : ordInvoice="" : ordClientID=0 : ordReferer="" : ordUserAgent="" : ordQuerystr="" : loyaltypointtotal="" : ordlang=0 : ordAddInfo=""
			end if
			rs.close
			if ordClientID<>0 then
				sSQL="SELECT clActions FROM customerlogin WHERE clID=" & ordClientID
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					if (rs("clActions") AND 1)=1 then orderstatetaxexempt=TRUE
					if (rs("clActions") AND 2)=2 then ordercountrytaxexempt=TRUE
					if loyaltypointsnowholesale AND (rs("clActions") AND 8)=8 then noloyaltypoints=TRUE
					if loyaltypointsnopercentdiscount AND (rs("clActions") AND 16)=16 then noloyaltypoints=TRUE
				end if
				rs.close
			end if
			if getget("id")<>"multi" then
				sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 1 ")&"ordID FROM orders WHERE ordID<"&theid&" ORDER BY ordID DESC"&IIfVs(mysqlserver=TRUE," LIMIT 0,1")
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then previousid=rs("ordID")
				rs.close
				sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 1 ")&"ordID FROM orders WHERE ordID>"&theid&" ORDER BY ordID"&IIfVs(mysqlserver=TRUE," LIMIT 0,1")
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then nextid=rs("ordID")
				rs.close
			end if
		end if
		if doedit then
			Session.LCID=1033
			print "<form method=""post"" name=""editform"" id=""editform"" action=""adminorders.asp" & IIfVs(getget("view")="true","?view=true") & """ onsubmit=""return confirmedit()""><input type=""hidden"" name=""orderid"" value="""&getget("id")&""" /><input type=""hidden"" name=""doedit"" value=""true"" />"
			currSymbolHTML=""
			currPostAmount=FALSE
			currDecimalSep="."
			currThousandsSep=""
		end if
		if NOT isprinter then
%>
<script>
/* <![CDATA[ */
var newwin="";
var plinecnt=0;
var numcartitems=<%=numorders%>;
function popupaddress(isship){
	var tntry=document.getElementById('ord'+isship+'country').innerHTML;
	document.getElementById('addresstextarea').value='';
	if(document.getElementById('ord'+isship+'extra1')&&document.getElementById('ord'+isship+'extra1').innerHTML!='')document.getElementById('addresstextarea').value+=document.getElementById('ord'+isship+'extra1').innerHTML + "\r\n";
	document.getElementById('addresstextarea').value+=document.getElementById('ord'+isship+'name').innerHTML + "\r\n" +
				document.getElementById('ord'+isship+'address').innerHTML + "\r\n";
	if(document.getElementById('ord'+isship+'address2')&&document.getElementById('ord'+isship+'address2').innerHTML!='')document.getElementById('addresstextarea').value+=document.getElementById('ord'+isship+'address2').innerHTML + "\r\n";
	if(document.getElementById('ord'+isship+'city').innerHTML!='')document.getElementById('addresstextarea').value+=document.getElementById('ord'+isship+'city').innerHTML + ", ";
	document.getElementById('addresstextarea').value+=document.getElementById('ord'+isship+'state').innerHTML + " " + document.getElementById('ord'+isship+'zip').innerHTML + ((tntry!='USA'&&tntry!='United States of America'&&tntry!='United States')<% if origCountryID<>1 then print "||true"%>?"\r\n" + tntry:'');
	document.getElementById('addressdiv').style.display='block';
	document.getElementById('addresstextarea').select();
}
function openemailpopup(id){
<%	if trim(ordEmail)="" OR instr(trim(ordEmail&""),"@")=0 then %>
	alert("There is no valid email set.");
<%	else %>
	popupWin=window.open('popupemail.asp?'+id,'emailpopup','menubar=no, scrollbars=no, width=300, height=250, directories=no,location=no,resizable=yes,status=no,toolbar=no')
<%	end if %>
}
function uaajaxcallback(){
	if(ajaxobj.readyState==4){
		var restxt=ajaxobj.responseText;
		resarr=restxt.split('==LISTELM==');
		if(resarr.length>0){
			document.getElementById("custid").value=resarr[0];
			document.getElementById("name").value=resarr[1];
<%	if usefirstlastname then print "document.getElementById('lastname').value=resarr[2];" & vbCrLf %>
		}
		if(resarr.length>5){
			document.getElementById("address").value=resarr[3];
<%	if useaddressline2 then print "document.getElementById('address2').value=resarr[4];" & vbCrLf %>
			document.getElementById("city").value=resarr[5];
			document.getElementById("state").value=resarr[6];
			document.getElementById("zip").value=resarr[7];
			cntry=document.getElementById("country");
			cntxt=resarr[8];
			for(index=0; index<cntry.length; index++){
				if(cntry.options[index].text==cntxt||cntry.options[index].value==cntxt){
					cntry.selectedIndex=index;
				}
			}
			document.getElementById("phone").value=resarr[9];
<%	if trim(extraorderfield1)<>"" then print "document.getElementById('extra1').value=resarr[10];" & vbCrLf
	if trim(extraorderfield2)<>"" then print "document.getElementById('extra2').value=resarr[11];" & vbCrLf %>
			setstatetax();
			setcountrytax();
		}
	}
}
function updateaddress(id){
	document.getElementById('percdisc').value=(adiscnts[id][1]=='1'?adiscnts[id][2]:'');
	document.getElementById('wholesaledisc').checked=(adiscnts[id][0]=='1');
	ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
	ajaxobj.onreadystatechange=uaajaxcallback;
	ajaxobj.open("POST", "ajaxservice.asp?action=getlist&listtype=adddets", true);
	ajaxobj.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
	ajaxobj.send('listtext='+adds[id]);
}
function upajaxcallback(){
	if(ajaxobj.readyState==4){
		var restxt=ajaxobj.responseText.replace(/^\s+|\s+$/g,"");
		resarr=restxt.split('==LISTELM==');
		document.getElementById('optionsspan'+resarr[0]).innerHTML=resarr[1];
		try{eval(resarr[2]);}catch(err){document.getElementById('optionsspan'+resarr[0]).innerHTML='javascript error'}
	}
}
function updateoptions(id){
	prodid=document.getElementById('prodid'+id).value;
	if(prodid != ''){
		ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
		ajaxobj.onreadystatechange=upajaxcallback;
		ajaxobj.open("POST", 'ajaxservice.asp?action=updateoptions&index='+id+'&wsp='+(document.getElementById('wholesaledisc').checked?'1':'0')+'&perc='+document.getElementById('percdisc').value.replace(/[^0-9\.]/g, ''), true);
		ajaxobj.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
		ajaxobj.send('productid='+prodid);
	}
	return(false);
}
function extraproduct(plusminus){
	var thetable=document.getElementById('producttable');
if(plusminus=='+'){
	numcartitems++;
	var node=document.createElement("DIV");
	node.className='ectorderrow';
	thetable.appendChild(node);
	
	var subnode=document.createElement("DIV");
	subnode.innerHTML='<div class="smallscreenquant"><input id="smallquant'+(plinecnt+1000)+'" size="3" onkeyup="document.getElementById(\'quant'+(plinecnt+1000)+'\').value=this.value" onchange="dorecalc(true)" type="text" /> x </div>' +
		'<input type="button" value="..." onclick="updateoptions('+(plinecnt+1000)+')" />&nbsp;<input name="prodid'+(plinecnt+1000)+'" size="18" id="prodid'+(plinecnt+1000)+'" AUTOCOMPLETE="off" onkeydown="return combokey(this,event)" onkeyup="combochange(this,event)" /><input type="hidden" id="stateexempt'+(plinecnt+1000)+'" value="false" /><input type="hidden" id="countryexempt'+(plinecnt+1000)+'" value="false" />' +
		'<%=jsescapel(showgetoptionsselect("xxxx"))%>'.replace(/xxxx/,'selectprodid'+(plinecnt+1000));
	node.appendChild(subnode);
	
	var subnode=document.createElement("DIV");
	subnode.innerHTML='<input type="text" id="prodname'+(plinecnt+1000)+'" name="prodname'+(plinecnt+1000)+'" size="24" AUTOCOMPLETE="off" onkeydown="return combokey(this,event)" onkeyup="combochange(this,event)" />' +
		'<%=jsescapel(showgetoptionsselect("xxxx"))%>'.replace(/xxxx/,'selectprodname'+(plinecnt+1000));
	node.appendChild(subnode);

	var subnode=document.createElement("DIV");
	subnode.innerHTML='<div id="optionsspan'+(plinecnt+1000)+'"><div style="text-align:center">-</div></div>';
	node.appendChild(subnode);

	var subnode=document.createElement("DIV");
	subnode.innerHTML='<div class="largescreenquant"><input type="text" id="quant'+(plinecnt+1000)+'" name="quant'+(plinecnt+1000)+'" size="5" onkeyup=""document.getElementById(\'smallquant'+(plinecnt+1000)+'\').value=this.value"" value="1" /></div>';
	node.appendChild(subnode);

	var subnode=document.createElement("DIV");
	subnode.innerHTML='<input type="text" class="orderprice" id="price'+(plinecnt+1000)+'" name="price'+(plinecnt+1000)+'" value="0" size="7" /><br /><input type="hidden" id="optdiffspan'+(plinecnt+1000)+'" value="0" />';
	node.appendChild(subnode);

	var subnode=document.createElement("DIV");
	subnode.innerHTML='&nbsp;';
	node.appendChild(subnode);

	plinecnt++;
}else{
	if(plinecnt>0){
	thetable.removeChild(thetable.lastChild);
	plinecnt--;
	numcartitems--;
	}
}
}
function confirmedit(){
<%	if useStockManagement then %>
var stockwarn="The following items do not have sufficient stock\n\n";
var outstock=false;
var oostock=new Array();
var oostockqnt=new Array();
var inputs=document.forms['editform'].getElementsByTagName("input");
for(ceindex=0;ceindex<inputs.length;ceindex++){
	var thename=inputs[ceindex].name;
	if(thename.substr(0,5)=="quant"){
		var theid=thename.substr(5);
		delbutton=document.getElementById("del_"+theid);
		if(delbutton==null)
			isdeleted=false;
		else
			isdeleted=delbutton.checked;
		if(! isdeleted){
			var pid=document.getElementById("prodid"+theid).value;
			var stocklevel=stock['pid_' + pid];
			var quant=document.getElementById("quant"+theid).value
			if(typeof(stocklevel)=="undefined"){
				// Do nothing, pid not defined.
			}else if(stocklevel=="bo"){ // By Options
				for(var ii in document.forms.editform){
					var opttext="optn"+theid+"_";
					if(ii.substr(0,opttext.length)==opttext){
						theitem=document.getElementById(ii);
						if(document.getElementById('v'+ii)==null){
							thevalue=theitem[theitem.selectedIndex].value.split('|')[0];
							stocklevel=stock['oid_'+thevalue];
							if(typeof(oostockqnt['oid_'+thevalue])=="undefined")
								oostockqnt['oid_'+thevalue]=parseInt(quant);
							else
								oostockqnt['oid_'+thevalue]+=parseInt(quant);
							if(parseInt(stocklevel)<oostockqnt['oid_'+thevalue]){
								oostock['oid_'+thevalue]=document.getElementById("prodname"+theid).value + " (" + theitem[theitem.selectedIndex].text + ") : Required " + oostockqnt['oid_'+thevalue] + " available ";
							}
						}
					}
				}
			}else{
				if(typeof(oostockqnt['pid_' + pid])=="undefined")
					oostockqnt['pid_' + pid]=parseInt(quant);
				else
					oostockqnt['pid_' + pid]+=parseInt(quant);
				if(parseInt(stocklevel)<oostockqnt['pid_' + pid]){
					oostock['pid_' + pid]=document.getElementById("prodname"+theid).value + ": Required " + oostockqnt['pid_' + pid] + " available ";
				}
			}
		}
	}
}
for(var i in oostock){
	outstock=true;
	stockwarn+=oostock[i] + stock[i] + "\n";
}
if(outstock){
	if(! confirm(stockwarn+"\nPress \"OK\" to submit changes or cancel to adjust quantities\n"))
		return(false);
}
<%	end if %>
if(confirm("<%=jscheck(yyChkRec)%>"))
	return(true);
return(false);
}
function calcshipping(){
	var txturl='shipservice.asp?';
	var editformelems=document.getElementById('editform').elements;
	for(var iix=0; iix<editformelems.length;iix++){
		var ii=editformelems[iix].name;
		if(ii.substr(0,6)=='prodid'){
			var theid=ii.substr(6);
			txturl+=ii+"="+editformelems[iix].value+"&";
			var thequant=parseInt(document.getElementById("quant"+theid).value);
			if(isNaN(thequant)) thequant=0;
			txturl+="quant"+theid+"="+thequant+"&";
			for(var iix2=0; iix2<editformelems.length;iix2++){
				var ii2=editformelems[iix2].name;
				var opttext="optn"+theid+"_";
				if(ii2.substr(0,opttext.length)==opttext){
					theitem=document.getElementById(ii2);
					if(document.getElementById('v'+ii2)==null){
						thevalue=theitem[theitem.selectedIndex].value;
						txturl+="optn"+theid+"_"+iix2+"="+thevalue.split('|')[0]+"&";
					}
				}
			}
		}
	}
	var isship=(document.getElementById('sstate').value!=''&&document.getElementById('szip').value!=''?'s':'');
	var shipstate=encodeURIComponent(document.getElementById(isship+'state').value);
	var shipzip=encodeURIComponent(document.getElementById(isship+'zip').value);
	var shipcountry=encodeURIComponent(document.getElementById(isship+'country').value);
	var comloc=document.getElementById('commercialloc')[document.getElementById('commercialloc').selectedIndex].value;
	var popupWin=window.open(txturl+"action=admincalc&destzip="+shipzip+"&sc="+shipcountry+"&sta="+shipstate+"&cl="+comloc,'calcshipping','menubar=no, scrollbars=yes, width=500, height=400, directories=no,location=no,resizable=yes,status=no,toolbar=no');
}
var opttxtcharge=[];
function dorecalc(onlytotal){
var thetotal=0,totoptdiff=0,statetaxabletotal=0,countrytaxabletotal=0;
for(var zz=0; zz < document.forms.editform.length; zz++){
var iq=document.forms.editform[zz].name;
if(iq.substr(0,5)=="quant"){
	theid=iq.substr(5);
	totopts=0;
	delbutton=document.getElementById("del_"+theid);
	if(delbutton==null)
		isdeleted=false;
	else
		isdeleted=delbutton.checked;
	if(! isdeleted){
		var editformelems=document.getElementById('editform').elements;
        for(var iix=0; iix<editformelems.length;iix++){
			var ii=editformelems[iix].name;
			var opttext="optn"+theid+"_";
			if(ii.substr(0,opttext.length)==opttext){
				theitem=document.getElementById(ii);
				if(document.getElementById('v'+ii)==null){
					thevalue=theitem[theitem.selectedIndex].value;
					if(thevalue.indexOf('|')>0){
						totopts+=parseFloat(thevalue.substr(thevalue.indexOf('|')+1));
					}
				}else{
					optid=parseInt(ii.substr(opttext.length));
					if(opttxtcharge[optid]){
						if(opttxtcharge[optid]>0){
							totopts+=opttxtcharge[optid]*document.getElementById('v'+ii).value.length;
						}else if(document.getElementById('v'+ii).value.length>0){
							totopts+=Math.abs(opttxtcharge[optid]);
						}
					}
				}
			}
		}
		thequant=parseInt(document.getElementById(iq).value);
		if(isNaN(thequant)) thequant=0;
		theprice=parseFloat(document.getElementById("price"+theid).value);
		if(isNaN(theprice)) theprice=0;
		document.getElementById("optdiffspan"+theid).value=totopts;
		optdiff=parseFloat(document.getElementById("optdiffspan"+theid).value);
		if(isNaN(optdiff)) optdiff=0;
		thetotal+=thequant * (theprice + optdiff);
		if(document.getElementById('orderstatetaxexempt').checked&&(!document.getElementById("stateexempt"+theid)||document.getElementById("stateexempt"+theid).value!='true'))
			statetaxabletotal+=thequant * (theprice + optdiff);
		if(document.getElementById('ordercountrytaxexempt').checked&&(!document.getElementById("countryexempt"+theid)||document.getElementById("countryexempt"+theid).value!='true'))
			countrytaxabletotal+=thequant * (theprice + optdiff);
		totoptdiff+=thequant * optdiff;
	}
}
}
document.getElementById("optdiffspan").innerHTML=totoptdiff.toFixed(2);
document.getElementById("ordtotal").value=thetotal.toFixed(2);
if(onlytotal==true) return;<%
if origCountryID=2 then print vbCrLf & "var ssa=getshipstateabbrev();" %>
statetaxrate=parseFloat(document.getElementById("staterate").value);
if(isNaN(statetaxrate)) statetaxrate=0;
var homecountrytaxrate=<%=homecountrytaxrate%>;
countrytaxrate=parseFloat(document.getElementById("countryrate").value);
if(isNaN(countrytaxrate)) countrytaxrate=0;
discount=parseFloat(document.getElementById("ordDiscount").value);
if(isNaN(discount)){
	discount=0;
	document.getElementById("ordDiscount").value=0;
}
statetaxtotal=(statetaxrate * Math.max(statetaxabletotal-discount,0)) / 100.0;
<%	if showtaxinclusive=3 then %>
countrytaxtotal=Math.round((countrytaxabletotal*100) / ((100+homecountrytaxrate)/homecountrytaxrate))/100.0;
thetotal-=countrytaxtotal;
document.getElementById("ordtotal").value=thetotal.toFixed(2);
if(countrytaxrate!=homecountrytaxrate&&homecountrytaxrate!=0)
	if(countrytaxrate!=0) countrytaxtotal=countrytaxtotal*(countrytaxrate/homecountrytaxrate); else countrytaxtotal=0;
<%	else %>
countrytaxtotal=(countrytaxrate * Math.max(countrytaxabletotal-discount,0)) / 100.0;
<%	end if %>
shipping=parseFloat(document.getElementById("ordShipping").value);
if(isNaN(shipping)){
	shipping=0;
	document.getElementById("ordShipping").value=0;
}
handling=parseFloat(document.getElementById("ordHandling").value);
if(isNaN(handling)){
	handling=0;
	document.getElementById("ordHandling").value=0;
}
<%	if taxShipping=2 then %>
statetaxtotal+=(statetaxrate * shipping) / 100.0;
countrytaxtotal+=(countrytaxrate * shipping) / 100.0;
<%	end if
	if taxHandling=2 then %>
statetaxtotal+=(statetaxrate * handling) / 100.0;
countrytaxtotal+=(countrytaxrate * handling) / 100.0;
<%	end if %>
var hsttax=0;
<%	if origCountryID=2 then %>
	if(getshipcountry()=='canada'){
		if(ssa=="NB" || ssa=="NF" || ssa=="NS" || ssa=="ON" || ssa=="PE"){
			hsttax=statetaxtotal+countrytaxtotal;
			statetaxtotal=0;
			countrytaxtotal=0;
		}
	}
	document.getElementById("ordHSTTax").value=hsttax.toFixed(2);
<%	end if %>
statetaxtotal=roundNumber(statetaxtotal,2);
countrytaxtotal=roundNumber(countrytaxtotal,2);
document.getElementById("ordStateTax").value=statetaxtotal.toFixed(2);
document.getElementById("ordCountryTax").value=countrytaxtotal.toFixed(2);
grandtotal=(thetotal + shipping + handling + statetaxtotal + countrytaxtotal + hsttax) - discount;
document.getElementById("grandtotalspan").innerHTML=grandtotal.toFixed(2);
<%	if loyaltypoints<>"" then %>
	document.getElementById("loyaltyPoints").value=Math.round((thetotal.toFixed(2)-discount)*<%=IIfVr(noloyaltypoints,0,loyaltypoints)%>);
<%	end if %>
}
function roundNumber(num, dec){
	var result=Math.round(Math.round(num * Math.pow(10, dec+1) ) / 10) / Math.pow(10,dec);
	return result;
}
function ppajaxcallback(){
	if(ajaxobj.readyState==4){
		document.getElementById("googleupdatespan").innerHTML=ajaxobj.responseText;
	}
}
function updategoogleorder(theprocessor,theact,ordid){
	if(theact=='settle') tmsg='Capture the amount of: '+document.getElementById("txamount").value; else tmsg='Inform '+theprocessor+' of change to order id '+ordid+'?';
	if(confirm(tmsg)){
		document.getElementById("googleupdatespan").innerHTML='';
		ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
		ajaxobj.onreadystatechange=ppajaxcallback;
		extraparams='';
		if(theact=='ship'){
			shipcar=document.getElementById("shipcarrier");
			if(shipcar!= null){
				trackno=document.getElementById("ordTrackNum").value
				if(trackno!='' && confirm('Include tracking and carrier info?')){
					extraparams='&carrier='+(shipcar.options[shipcar.selectedIndex].value)+'&trackno='+document.getElementById("ordTrackNum").value;
				}
			}
		}
		if(document.getElementById("txamount")){
			extraparams+='&amount='+document.getElementById("txamount").value;
		}
		document.getElementById("googleupdatespan").innerHTML='Connecting...';
		ajaxobj.open("GET", "ajaxservice.asp?processor="+theprocessor+"&gid="+ordid+"&act="+theact+extraparams, true);
		ajaxobj.send(null);
	}
}
function updatepaypalorder(theprocessor,ordid){
	if(confirm('Inform '+(theprocessor=='AuthNET'?'Authorize.net':'PayPal')+' of change to order id ' + ordid + "?")){
		document.getElementById("googleupdatespan").innerHTML='';
		ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
		ajaxobj.onreadystatechange=ppajaxcallback;
		var additionalcapture=document.getElementById("additionalcapture")?document.getElementById("additionalcapture")[document.getElementById("additionalcapture").selectedIndex].value:0;
		var theact=document.getElementById("paypalaction")[document.getElementById("paypalaction").selectedIndex].value;
		document.getElementById("googleupdatespan").innerHTML='Connecting...';
		postdata="additionalcapture=" + additionalcapture + "&amount=" + encodeURIComponent(document.getElementById("captureamount").value) + (document.getElementById("buyernote")?"&comments=" + encodeURIComponent(document.getElementById("buyernote").value):'');
		if(document.getElementById('capstatus')) postdata+="&capstatus=" + document.getElementById('capstatus')[document.getElementById("capstatus").selectedIndex].value;
		ajaxobj.open("POST", "ajaxservice.asp?processor="+theprocessor+"&gid="+ordid+"&act="+theact, true);
		ajaxobj.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
		ajaxobj.send(postdata);
	}
}
function setpaypalelements(){
	var theact=document.getElementById("paypalaction")[document.getElementById("paypalaction").selectedIndex].value;
	if(theact=='void'){
		document.getElementById("captureamount").disabled=true;
		document.getElementById("additionalcapture").disabled=true;
	}else if(theact=='reauth'){
		document.getElementById("captureamount").disabled=false;
		document.getElementById("additionalcapture").disabled=true;
	}else{
		document.getElementById("captureamount").disabled=false;
		document.getElementById("additionalcapture").disabled=false;
	}
}
function copybillingtoshipping(){
<%	if Trim(extraorderfield1)<>"" then print "document.getElementById('shipextra1').value=document.getElementById('extra1').value;" %>
	document.getElementById("sname").value=document.getElementById("name").value;
<%	if usefirstlastname then print "document.getElementById('slastname').value=document.getElementById('lastname').value;" %>
	document.getElementById("saddress").value=document.getElementById("address").value;
<%	if useaddressline2=TRUE then print "document.getElementById('saddress2').value=document.getElementById('address2').value;" %>
	document.getElementById("scity").value=document.getElementById("city").value;
	document.getElementById("sstate").value=document.getElementById("state").value;
	document.getElementById("szip").value=document.getElementById("zip").value;
	document.getElementById("scountry").selectedIndex=document.getElementById("country").selectedIndex;
	document.getElementById("sphone").value=document.getElementById("phone").value;
<%	if Trim(extraorderfield2)<>"" then print "document.getElementById('shipextra2').value=document.getElementById('extra2').value;" %>
}
<%			if doedit then %>
function setCookie(c_name,value,expiredays){
	var exdate=new Date();
	exdate.setDate(exdate.getDate()+expiredays);
	document.cookie=c_name+ "=" +escape(value)+((expiredays==null) ? "" : ";expires="+exdate.toGMTString());
}
var adds=[];
var opensels=[];
var adiscnts=[];
document.getElementById('main').onclick=function(){
	for(var ii=0; ii<opensels.length; ii++)
		document.getElementById(opensels[ii]).style.display='none';
};
function addopensel(id){
	for(var ii=0; ii<opensels.length; ii++)
		if(id==opensels[ii]) return;
	opensels.push(id);
}
function plajaxcallback(){
	if(ajaxobj.readyState==4){
		var resarr=ajaxobj.responseText.replace(/^\s+|\s+$/g,"").split('==LISTOBJ==');
		var index,isname=false;
		oSelect=document.getElementById(resarr[0]);
		var act=resarr[0].replace(/\d/g,'');
		for(index=0; index<resarr.length-2; index++){
			var splitelem=resarr[index+1].split('==LISTELM==');
			var val1=splitelem[0];
			var val2=splitelem[1];
			var haswsdisc=0,hasperdisc=0,perdisc=0;
			if(splitelem.length>=2) adds[index]=splitelem[2];
			if(splitelem.length>=5) haswsdisc=splitelem[3];
			if(splitelem.length>=5) hasperdisc=splitelem[4];
			if(splitelem.length>=5) perdisc=splitelem[5];
			adiscnts[index]=new Array(haswsdisc,hasperdisc,perdisc);
			if(index<oSelect.length)
				var y=oSelect.options[index];
			else
				var y=document.createElement('option');
			if(act=='selectprodname'){
				y.text=val2;
				y.value=val1;
			}else if(act=='selectemail'){
				y.text=val2;
				y.title=val2;
				y.value=val1;
			}else{
				y.text=val1;
				y.value=val1;
			}
			if(y.text=='----------------') y.disabled=true; else y.disabled=false;
			if(index>=oSelect.length){
				try{oSelect.add(y, null);} // FF etc
				catch(ex){oSelect.add(y);} // IE
			}
		}
		if(oSelect){
			for(var ii=oSelect.length;ii>=index;ii--){
				oSelect.remove(ii);
			}
		}
	}
}
var gsid;
var gltyp;
var gtxt;
var tmrid;
function populatelist(){
	var objid=gsid;
	var listtype=gltyp;
	var stext=gtxt;
	ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
	ajaxobj.onreadystatechange=plajaxcallback;
	ajaxobj.open("POST", "ajaxservice.asp?action=getlist&objid="+objid+"&listtype="+listtype, true);
	ajaxobj.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
	ajaxobj.send('listtext='+stext);
}
function combochange(oText,e){
	if(document.getElementById("autocomplete").checked==false)
		return;
	keyCode=e.keyCode;
	if(keyCode<32&&keyCode!=8)return true;
	oSelect=document.getElementById('select'+oText.id);
	addopensel(oSelect.id);
	oSelect.style.display='';
	toFind=oText.value.toLowerCase();
	gsid=oSelect.id;
	gltyp=oText.id.replace(/\d/g,'');
	gtxt=toFind;
	clearTimeout(tmrid);
	tmrid=setTimeout("populatelist()",800);
}
function combokey(oText,e){
	if(document.getElementById("autocomplete").checked==false)
		return
	oSelect=document.getElementById('select'+oText.id);
	keyCode=e.keyCode;
	if(keyCode==40 || keyCode==38){ // Up / down arrows
		addopensel(oSelect.id);
		oSelect.style.display='';
		oSelect.focus();
		comboselect_onchange(oSelect);
	}
	else if(keyCode==13){
		oSelect.style.display='none';
		oText.focus();
		updateoptions(oText.id.replace(/prodid|prodname/,''));
		return getvalsfromserver(oSelect);
	}
	return true;
}
function getvalsfromserver(oSelect){
	var act=oSelect.id.replace(/\d/g,'');
	oText=document.getElementById(oSelect.id.replace('select',''));
	if(oSelect.selectedIndex != -1){
		if(act=='selectprodname'){
			oText.value=oSelect.options[oSelect.selectedIndex].text;
			document.getElementById(oText.id.replace('prodname','prodid')).value=oSelect.options[oSelect.selectedIndex].value;
		}else
			oText.value=oSelect.options[oSelect.selectedIndex].value;
		oSelect.style.display='none';
		oText.focus();
		if(act=='selectemail')
			updateaddress(oSelect.selectedIndex);
		else
			updateoptions(oText.id.replace(/prodid|prodname/,''));
	}
	return false;
}
function comboselect_onclick(oSelect){
	return(getvalsfromserver(oSelect));
}
function comboselect_onchange(oSelect){
	oText=document.getElementById(oSelect.id.replace('select',''));
	if(oSelect.selectedIndex != -1){
		if(oText.id.indexOf('prodname')!=-1)
			oText.value=oSelect.options[oSelect.selectedIndex].text;
		else
			oText.value=oSelect.options[oSelect.selectedIndex].value;
	}
}
function comboselect_onkeyup(keyCode,oSelect){
	if(keyCode==13){
		getvalsfromserver(oSelect);
	}
	return(false);
}
var countrytaxrates=[];
var statetaxrates=[];
var stateabbrevs=[];
<%	sSQL="SELECT stateName,stateAbbrev,stateTax,stateCountryID FROM states WHERE stateTax<>0"
	rs2.Open sSQL,cnn,0,1
	do while not rs2.EOF
		print "statetaxrates["""&lcase(rs2("stateName"))&"""]="&rs2("stateTax")&";"&vbCrLf
		if rs2("stateCountryID")=1 OR rs2("stateCountryID")=2 then
			print "statetaxrates["""&lcase(rs2("stateAbbrev"))&"""]="&rs2("stateTax")&";"&vbCrLf
			print "stateabbrevs["""&lcase(rs2("stateName"))&"""]="""&rs2("stateAbbrev")&""";"&vbCrLf
		end if
		rs2.MoveNext
	loop
	rs2.Close
	sSQL="SELECT countryID,countryName,countryTax FROM countries WHERE countryTax<>0"
	rs2.Open sSQL,cnn,0,1
	do while not rs2.EOF
		print "countrytaxrates["&lcase(rs2("countryID"))&"]="&rs2("countryTax")&";"&vbCrLf
		rs2.MoveNext
	loop
	rs2.Close %>
function setstatetax(){
	var addans='';
	if(document.getElementById('saddress').value!='') addans='s';
	var rgnname=document.getElementById(addans+'state').value.toLowerCase();
	if(statetaxrates[rgnname]) statetaxrate=parseFloat(statetaxrates[rgnname]); else statetaxrate=0;
	document.getElementById("staterate").value=statetaxrate;
}
function setcountrytax(){
	var addans='';
	if(document.getElementById('saddress').value!='') addans='s';
	var tobj=document.getElementById(addans+'country');
	var rgnid=tobj.options[tobj.selectedIndex].value.toLowerCase();
	if(countrytaxrates[rgnid]) countrytaxrate=parseFloat(countrytaxrates[rgnid]); else countrytaxrate=0;
	document.getElementById("countryrate").value=countrytaxrate;
}
function getshipstateabbrev(){
	var addans='';
	if(document.getElementById('saddress').value!='') addans='s';
	var rgnname=document.getElementById(addans+'state').value.toLowerCase();
	if(stateabbrevs[rgnname]){
		document.getElementById('staterate').value=statetaxrates[rgnname];
		return(stateabbrevs[rgnname]);
	}else
		return document.getElementById(addans+'state').value;
}
function getshipcountry(){
	var addans='';
	if(document.getElementById('saddress').value!='') addans='s';
	var tobj=document.getElementById(addans+'country');
	return(tobj.options[tobj.selectedIndex].value.toLowerCase());
}
<%			else ' NOT doedit
%>
function dosavefieldcb(){
	if(ajaxobj.readyState==4){
		if(ajaxobj.responseText.substr(0,7)=='SUCCESS'){
			document.getElementById(ajaxobj.responseText.substr(8)).style.borderColor='green';
		}else
			alert('Error updating');
	}
}
function dosavefield(tact,ordid,newval){
	ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
	ajaxobj.onreadystatechange=dosavefieldcb;
	ajaxobj.open("GET","ajaxservice.asp?action="+tact+"&ordid="+ordid+"&updfield="+encodeURIComponent(newval),true);
	ajaxobj.send(null);
}
var invisflashing=false,tnumisflashing=false;
var invintervalid=0,tnumintervalid=0;
var invflipflop=0,tnumflipflop=0;
function fliptext(tact){
	if(tact=='invoiceupdate'){
		document.getElementById(tact+'t').innerHTML=invflipflop==0?'Press Enter To Save:':'<%=jscheck(yyInvNum)%>:';
		invflipflop=Math.abs(invflipflop-1);
	}else if(tact=='tracknumupdate'){
		document.getElementById(tact+'t').innerHTML=tnumflipflop==0?'Press Enter To Save:':'<%=jscheck(yyTraNum)%>:';
		tnumflipflop=Math.abs(tnumflipflop-1);
	}
}
function startflash(tact){
	if(tact=='invoiceupdate'&&!invisflashing){
		fliptext('invoiceupdate');
		invisflashing=true;
		invintervalid=setInterval("fliptext('invoiceupdate')",2000);
	}else if(tact=='tracknumupdate'&&!tnumisflashing){
		fliptext('tracknumupdate');
		tnumisflashing=true;
		tnumintervalid=setInterval("fliptext('tracknumupdate')",2000);
	}
}
function stopflash(tact){
	if(tact=='invoiceupdate'){
		clearInterval(invintervalid);
		document.getElementById(tact+'t').innerHTML='<%=jscheck(yyInvNum)%>:';
		invflipflop=0;
		invisflashing=false;
	}else if(tact=='tracknumupdate'){
		clearInterval(tnumintervalid);
		document.getElementById(tact+'t').innerHTML='<%=jscheck(yyTraNum)%>:';
		tnumflipflop=0;
		tnumisflashing=false;
	}
}
function checksavefield(tact,evt,tobj,ordid){
	var code=(evt.keyCode ? evt.keyCode : evt.which);
	if(code==13){
		stopflash(tact);
		document.getElementById(tact+'t').className=document.getElementById(tact+'t').className.replace(/ ectred/g,'');
		dosavefield(tact,ordid,tobj.value);
	}else if(code==8||code==32||(code>=46&&code<=90)){
		tobj.style.borderColor='red';
		document.getElementById(tact+'t').className+=' ectred';
		tobj.title='Press Enter To Save';
		startflash(tact);
	}
}
function checkstatusupdfield(evt,tobj){
	var code = (evt.keyCode ? evt.keyCode : evt.which);
	if(code==8||code==13||code==32||(code>=46&&code<=90)){
		document.getElementById('statusupdcontainer').style.display=''
	}
}
<%			end if %>
/* ]]> */
</script>
<%			if NOT doedit then %>
<div id="addressdiv" onclick="this.style.display='none'" style="display:none;position:absolute;width:100%;height:2000px;background-image:url(adminimages/opaquepixel.png);top:0px;left:0px;text-align:center;z-index:10000;">
<br /><br /><br /><br /><br /><br /><br /><br />
<textarea id="addresstextarea" rows="10" cols="40" onclick="return false"></textarea>
</div>
<%			end if
		end if ' NOT isprinter
		sSQL="SELECT packingslipuseinvoice FROM admin WHERE adminID=1"
		rs.open sSQL,cnn,0,1
		packingslipuseinvoice=rs("packingslipuseinvoice")
		rs.close
		sSQL="SELECT invoiceheader,invoiceaddress,invoicefooter"&IIfVs(packingslipuseinvoice=0,",packingslipheader,packingslipaddress,packingslipfooter")&" FROM emailmessages WHERE emailID=1"
		rs.open sSQL,cnn,0,1
		invoiceheader=rs("invoiceheader")
		invoiceaddress=rs("invoiceaddress")
		invoicefooter=rs("invoicefooter")
		if packingslipuseinvoice then
			packingslipheader=invoiceheader
			packingslipaddress=invoiceaddress
			packingslipfooter=invoicefooter
		else
			packingslipheader=rs("packingslipheader")
			packingslipaddress=rs("packingslipaddress")
			packingslipfooter=rs("packingslipfooter")
		end if
		rs.close
%>
<script>
/* <![CDATA[ */
function researchformgo(tid,switchinvoice){
	var currinvoice='<%=IIfVs(getget("printer")="true","&printer=true")&IIfVs(getget("invoice")="true","&invoice=true")%>';
	var newinvoice='';
	if(switchinvoice==0)
		newinvoice=currinvoice;
	else if(switchinvoice==1||switchinvoice==2)
		newinvoice=switchinvoice==2?'&printer=true':'&invoice=true';
	else if(switchinvoice==5)
		newinvoice='&doedit=true&view=true';
	document.getElementById('researchform').action='adminorders.asp'+(switchinvoice!=3?'?id='+tid+newinvoice:'');
	document.getElementById('researchform').submit();
}
/* ]]> */
</script>
<script src="popcalendar.js"></script>
<div class="orderdetails"<% if numids>=0 then print " style=""page-break-after:always"""%>>
<%		extraclass="order"
		if getget("invoice")="true" then extraclass="invoice"
		if getget("printer")="true" then extraclass="packingslip"
		if doedit then extraclass="edit"
		if isprinter AND isempty(packingslipheader) then packingslipheader=invoiceheader
		if isinvoice AND invoiceheader<>"" then %>
		<div class="orderheader invoiceheader"><%=invoiceheader%></div>
<%		elseif isprinter AND packingslipheader<>"" then %>
		<div class="orderheader packslipheader"><%=packingslipheader%></div>
<%		end if %>
		<div class="buttons">
			<div class="buttonleft"><%
		if doedit then
			print "&nbsp;<input type=""checkbox"" value=""ON"" name=""autocomplete"" id=""autocomplete"" onclick=""setCookie('ectautocomp',this.checked?1:0,600)"" "&IIfVr(request.cookies("ectautocomp")="1","checked=""checked"" ","")&"/> <strong>"&yyUsAuCo&"</strong>"
		elseif getget("id")<>"multi" AND getget("id")<>"new" then
			print "<input "&IIfVs(previousid="","disabled=""disabled"" ")&"class=""no-print"" style=""width:110px;margin:2px"" type=""button"" value=""&laquo; "&yyPrev&""" onclick=""researchformgo('"&previousid&"',0)"" />"
			print "<input "&IIfVs(previousbysearch="","disabled=""disabled"" ")&"class=""no-print"" style=""width:110px;margin:2px"" type=""button"" value="""&"Prev. By Search"&""" onclick=""researchformgo('"&previousbysearch&"',0)"" /><br />"
			print "<input "&IIfVs(NOT(getget("printer")="true" OR getget("invoice")="true"),"disabled=""disabled"" ")&"class=""no-print"" style=""width:110px;margin:2px"" type=""button"" value="""&"Print"&""" onclick=""window.print()"" />"
			print "<input class=""no-print"" style=""width:110px;margin:2px"" type=""button"" value="""&"Orders"&""" onclick=""researchformgo('"&ordID&"',3)"" />"
		end if %></div><div class="idanddate"><%
		if NOT doedit then print "<div class=""no-print""><input type=""button"" style=""width:130px;margin:2px"" value=""Edit Order: " & ordID & """ onclick=""researchformgo('"&ordID&"',5)"" /></div>"
		print "<div class=""orderid"">" & xxOrdNum & " " & IIfVr(getget("id")="new","("&yyNewOrd&")",ordID)
		if doedit AND adminlanguages>0 then
			print " - Language ID: <select size=""1"" name=""ordlang"">"
			for index=0 to adminlanguages
				print "<option value="""&index&""""&IIfVs(index=ordlang," selected=""selected""")&">"&(index+1)&"</option>"
			next
			print "</select>"
		end if
		thedatetime=FormatDateTime(ordDate, 1) & IIfVs(NOT isinvoice," " & FormatDateTime(ordDate, 4))
		print "</div><div class=""orderdate"">"
		if doedit then
			print "<input type=""button"" value=""" & thedatetime & """ onclick=""document.getElementById('doeditdate').value='1';this.style.display='none';document.getElementById('editorddatediv').style.display='inline';popUpCalendar(this, document.getElementById('editorddate'), '" & themask & "', 0)"" />"
			print "<div id=""editorddatediv"" style=""position:relative;display:none"">"
				print "<input type=""text"" style=""width:120px"" id=""editorddate"" name=""editorddate"" value="""& FormatDateTime(ordDate, 2) & """ onfocus=""popUpCalendar(this, document.getElementById('editorddate'), '" & themask & "', 0)"" />"
				print "<input type=""text"" style=""width:60px"" name=""editordtime"" value="""& FormatDateTime(ordDate, 4) & """ />"
				call writehiddenidvar("doeditdate","")
			print "</div>"
		else
			print thedatetime
		end if
		print "</div>" %>
			</div><div class="buttonright">
<%		if getget("id")<>"multi" AND getget("id")<>"new" AND NOT doedit then
			print "<input "&IIfVs(nextbysearch="","disabled=""disabled"" ")&"class=""no-print"" style=""width:110px;margin:2px"" type=""button"" value="""&"Next By Search"&""" onclick=""researchformgo('"&nextbysearch&"',0)"" />"
			print "<input "&IIfVs(nextid="","disabled=""disabled"" ")&"class=""no-print"" style=""width:110px;margin:2px"" type=""button"" value="""&yyNext&" &raquo;"" onclick=""researchformgo('"&nextid&"',0)"" /><br />"
			print "<input class=""no-print"" style=""width:110px;margin:2px"" type=""button"" value="""&IIfVr(getget("invoice")="true","Details","Invoice")&""" onclick=""researchformgo('"&ordID&"',"&IIfVr(getget("invoice")="true",4,1)&")"" />"
			print "<input class=""no-print"" style=""width:110px;margin:2px"" type=""button"" value="""&IIfVr(getget("printer")="true","Details","Packing Slip")&""" onclick=""researchformgo('"&ordID&"',"&IIfVr(getget("printer")="true",4,2)&")"" />"
		end if
%>			</div>
		</div>
<%		if isprinter AND isempty(packingslipaddress) then packingslipaddress=invoiceaddress
		if isinvoice AND invoiceaddress<>"" then %>
		<div class="orderaddress invoiceaddress"><%=invoiceaddress%></div>
<%		elseif isprinter AND packingslipaddress<>"" then %>
		<div class="orderaddress packslipaddress"><%=packingslipaddress%></div>
<%		end if %>
		<input type="hidden" name="custid" id="custid" value="" />

			<div class="addresses">
				<div class="billing address">
					<div class="billing colheading"><h2><%=yyBilDet%>.</h2></div>
<%		if trim(extraorderfield1)<>"" AND (NOT isprinter OR (trim(ordExtra1&"")<>"" AND NOT extraorderfield1noprint)) then %>
					<div class="billing container container1">
					  <div class="extra1 left"><%=extraorderfield1 %>:</div>
					  <div class="extra1 right" id="ordextra1"><%=editfunc(ordExtra1,"extra1",25)%></div>
					</div>
<%		end if %>
					<div class="billing container container2">
					  <div class="ordname left"><%=yyName%>:</div>
					  <div class="ordname right" id="ordname"><% if usefirstlastname OR ordLastName<>"" then print editfunc(ordName,"name",11)&" "&editfunc(ordLastName,"lastname",11) else print editfunc(ordName,"name",25)%></div>
					</div>
					<div class="billing container container3">
					  <div class="ordaddress left"><%=xxAddress%>:</div>
					  <div class="ordaddress right" id="ordaddress"><%=editfunc(ordAddress,"address",25)%></div>
					</div>
<%		if (doedit AND useaddressline2) OR trim(ordAddress2&"")<>"" then %>
					<div class="billing container container4">
					  <div class="ordaddress2 left"><%=IIfVr(isprinter,"&nbsp;",xxAddress2&":")%></div>
					  <div class="ordaddress2 right" id="ordaddress2"><%=editfunc(ordAddress2,"address2",25)%></div>
					</div>
<%		end if
		if isprinter then %>
					<div class="billing container container5">
					  <div class="ordcity left">&nbsp;</div>
					  <div class="ordcity right"><%=ordCity&IIfVr(trim(ordCity&"")<>"" AND trim(ordState&"")<>"", ", ", "")&ordState%></div>
					</div>
<%		else %>
					<div class="billing container container6">
					  <div class="ordcity left"><%=xxCity%>:</div>
					  <div class="ordcity right" id="ordcity"><%=editfunc(ordCity,"city",25)%></div>
					</div>
					<div class="billing container container7">
					  <div class="ordstate left"><%=xxAllSta%>:</div>
					  <div class="ordstate right" id="ordstate"><%=editspecial(ordState,"state",25,"onblur=""setstatetax()""")%></div>
					</div>
<%		end if %>
					<div class="billing container container8">
					  <div class="ordzip left"><%=IIfVr(isprinter,"&nbsp;",xxZip&":")%></div>
					  <div class="ordzip right" id="ordzip"><%=editfunc(ordZip,"zip",15)%></div>
					</div>
					<div class="billing container container9">
					  <div class="ordcountry left"><%=IIfVr(isprinter,"&nbsp;",xxCountry&":")%></div>
					  <div class="ordcountry right" id="ordcountry"><%
		if doedit then
			foundmatch=FALSE
			countryid=0
			loadstates=-1
			print "<select name=""country"" id=""country"" size=""1"" onchange=""setcountrytax()"">"
			sSQL="SELECT countryID,countryName,countryTax,loadStates FROM countries ORDER BY countryOrder DESC, countryName"
			rs2.Open sSQL,cnn,0,1
			do while not rs2.EOF
				print "<option value=""" & rs2("countryID") & """"
				if ordCountry=rs2("countryName") OR (getget("id")="new" AND NOT foundmatch) then
					print " selected=""selected"""
					foundmatch=TRUE
					countryid=rs2("countryID")
					countrytaxrate=rs2("countryTax")
					loadstates=rs2("loadStates")
				end if
				print ">"&rs2("countryName")&"</option>"&vbCrLf
				rs2.MoveNext
			loop
			rs2.Close
			if NOT foundmatch then print "<option value=""" & htmlspecials(ordCountry) & """ selected=""selected"">"&ordCountry&"</option>"&vbCrLf
			print "</select>"
			if loadstates>0 then
				sSQL="SELECT stateTax FROM states WHERE stateCountryID="&countryid&" AND (stateName='"&escape_string(ordState)&"' OR stateAbbrev='"&escape_string(ordState)&"')"
				rs2.Open sSQL,cnn,0,1
				if NOT rs2.EOF then statetaxrate=rs2("stateTax")
				rs2.Close
			end if
		else
			shipcountrytext=IIfVr(ordShipCountry<>"", ordShipCountry, ordCountry)
			sSQL="SELECT countryID,countryTax FROM countries WHERE countryEnabled<>0 AND (countryName='" & escape_string(shipcountrytext) & "' OR countryName2='" & escape_string(shipcountrytext) & "' OR countryName3='" & escape_string(shipcountrytext) & "')"
			rs2.Open sSQL,cnn,0,1
			if NOT rs2.EOF then countrytaxrate=rs2("countryTax")
			rs2.close
			print ordCountry
		end if %></div>
					</div>
					<div class="billing container container10">
					  <div class="ordphone left"><%=xxPhone%>:</div>
					  <div class="ordphone right"><%=editfunc(ordPhone,"phone",25)%></div>
					</div>
<%		if trim(extraorderfield2)<>"" AND (NOT isprinter OR (trim(ordExtra2&"")<>"" AND NOT extraorderfield2noprint)) then %>
					<div class="billing container container11">
					  <div class="ordextra2 left"><% print extraorderfield2 %>:</div>
					  <div class="ordextra2 right" id="ordextra2"><%=editfunc(ordExtra2,"extra2",25)%></div>
					</div>
<%		end if
		if NOT (isprinter OR doedit) then %>
					<div class="clipboard container container12">
					  <div class="left"></div>
					  <div class="right"><input type="button" value="Copy to Clipboard" onclick="popupaddress('')" /></div>
					</div>
<%		end if %>
				</div><%
		if trim(ordShipName&"")<>"" OR trim(ordShipAddress&"")<>"" OR trim(ordShipCity&"")<>"" OR trim(ordShipExtra1&"")<>"" OR doedit then
				%><div class="ship shipaddress">
					<div class="ship colheading"><h2><%=xxShpDet%>.
<%			if doedit then print " &raquo; <a href=""#"" onclick=""copybillingtoshipping(); return(false);"">"&yyCopBil&"</a>"%></h2></div>
<%			if trim(extraorderfield1)<>"" AND (NOT isprinter OR (trim(ordShipExtra1&"")<>"" AND NOT extraorderfield1noprint)) then %>
					<div class="ship container container1">
					  <div class="ordshipextra1 left"><%=extraorderfield1 %>:</div>
					  <div class="ordshipextra1 right" id="ordsextra1"><%=editfunc(ordShipExtra1,"shipextra1",25)%></div>
					</div>
<%			end if %>
					<div class="ship container container2">
					  <div class="ordshipname left"><%=yyName%>:</div>
					  <div class="ordshipname right" id="ordsname"><% if usefirstlastname then print editfunc(ordShipName,"sname",11)&" "&editfunc(ordShipLastName,"slastname",11) else print editfunc(ordShipName,"sname",25)%></div>
					</div>
					<div class="ship container container3">
					  <div class="ordshipaddress left"><%=xxAddress%>:</div>
					  <div class="ordshipaddress right" id="ordsaddress"><%=editspecial(ordShipAddress,"saddress",25,"onblur=""setstatetax();setcountrytax();""")%></div>
					</div>
<%			if (doedit AND useaddressline2) OR trim(ordShipAddress2&"")<>"" then %>
					<div class="ship container container4">
					  <div class="ordshipaddress2 left"><%=IIfVr(isprinter, "&nbsp;", xxAddress2&":")%></div>
					  <div class="ordshipaddress2 right" id="ordsaddress2"><%=editfunc(ordShipAddress2,"saddress2",25)%></div>
					</div>
<%			end if
			if isprinter then %>
					<div class="ship container container5">
					  <div class="ordshipcity left">&nbsp;</div>
					  <div class="ordshipcity right"><%=ordShipCity&IIfVr(trim(ordShipCity&"")<>"" AND trim(ordShipState&"")<>"", ", ", "")&ordShipState%></div>
					</div>
<%			else %>
					<div class="ship container container6">
					  <div class="ordshipcity left"><%=xxCity%>:</div>
					  <div class="ordshipcity right" id="ordscity"><%=editfunc(ordShipCity,"scity",25)%></div>
					</div>
					<div class="ship container container7">
					  <div class="ordshipstate left"><%=xxAllSta%>:</div>
					  <div class="ordshipstate right" id="ordsstate"><%=editspecial(ordShipState,"sstate",25,"onblur=""setstatetax()""")%></div>
					</div>
<%			end if %>
					<div class="ship container container8">
					  <div class="ordshipzip left"><%=IIfVr(isprinter, "&nbsp;", xxZip&":")%></div>
					  <div class="ordshipzip right" id="ordszip"><%=editfunc(ordShipZip,"szip",15)%></div>
					</div>
					<div class="ship container container9">
					  <div class="ordshipcountry left"><%=IIfVr(isprinter, "&nbsp;", xxCountry&":")%></div>
					  <div class="ordshipcountry right" id="ordscountry"><%
			if doedit then
				if trim(ordShipName&"")<>"" OR trim(ordShipAddress&"")<>"" then usingshipcountry=TRUE else usingshipcountry=FALSE
				foundmatch=(getget("id")="new")
				countryid=0
				loadstates=-1
				print "<select name=""scountry"" id=""scountry"" size=""1"" onchange=""setcountrytax()"">"
				sSQL="SELECT countryID,countryName,countryTax,loadStates FROM countries ORDER BY countryOrder DESC, countryName"
				rs2.Open sSQL,cnn,0,1
				do while not rs2.EOF
					print "<option value=""" & rs2("countryID") & """"
					if ordShipCountry=rs2("countryName") then
						print " selected=""selected"""
						foundmatch=TRUE
						countryid=rs2("countryID")
						if usingshipcountry then countrytaxrate=rs2("countryTax")
						loadstates=rs2("loadStates")
					end if
					print ">"&rs2("countryName")&"</option>"&vbCrLf
					rs2.MoveNext
				loop
				rs2.Close
				if NOT foundmatch then print "<option value=""" & htmlspecials(ordShipCountry) & """ selected=""selected"">"&ordShipCountry&"</option>"&vbCrLf
				print "</select>"
				if loadstates>0 AND usingshipcountry then
					sSQL="SELECT stateTax FROM states WHERE stateCountryID="&countryid&" AND (stateName='"&escape_string(ordShipState&"")&"' OR stateAbbrev='"&escape_string(ordShipState)&"')"
					rs2.Open sSQL,cnn,0,1
					if NOT rs2.EOF then statetaxrate=rs2("stateTax")
					rs2.Close
				end if
			else
				print ordShipCountry
			end if %></div>
					</div>
					<div class="ship container container10">
					  <div class="ordshipphone left"><%=xxPhone%>:</div>
					  <div class="ordshipphone right"><%=editfunc(ordShipPhone,"sphone",25)%></div>
					</div>
<%			if trim(extraorderfield2)<>"" AND (NOT isprinter OR (trim(ordShipExtra2&"")<>"" AND NOT extraorderfield2noprint)) then %>
					<div class="ship container container11">
					  <div class="ordshipextra2 left"><% print extraorderfield2 %>:</div>
					  <div class="ordshipextra2 right" id="ordsextra2"><%=editfunc(ordShipExtra2,"shipextra2",25)%></div>
					</div>
<%			end if
			if NOT (isprinter OR doedit) then %>
					<div class="clipboard container container12">
					  <div class="left"></div>
					  <div class="right"><input type="button" value="Copy to Clipboard" onclick="popupaddress('s')" /></div>
					</div>
<%			end if %>
				</div>
<%		end if %>
			</div><div class="adddetails <%=extraclass%>adddetails">
					<div class="colheading"><h2><%=yyAddDet%>.</h2></div>
					<div class="container container1">
					  <div class="ordemail left"><%=xxEmail%>:</div>
					  <div class="ordemail right"><%
		if isprinter OR doedit then print editspecial(ordEmail,"email",35,"AUTOCOMPLETE=""off"" onkeydown=""return combokey(this,event)"" onkeyup=""combochange(this,event)""") else print "<a href=""mailto:" & htmlspecials(ordEmail) & """>"&htmlspecials(ordEmail)&"</a>"
		if doedit then print showgetoptionsselect("selectemail") %></div>
					</div>
<%		if trim(extracheckoutfield1)<>"" AND NOT (isprinter AND extracheckoutfield1noprint) then
			checkoutfield1=extracheckoutfield1
			checkoutfield2=editfunc(ordCheckoutExtra1,"checkoutextra1",25)
%>					<div class="container container2">
					  <div class="ordcheckoutextra1 left"><% if extracheckoutfield1reverse then print checkoutfield2 else print checkoutfield1 & ":" %></div>
					  <div class="ordcheckoutextra1 right"><% if extracheckoutfield1reverse then print checkoutfield1 else print checkoutfield2 %></div>
					</div>
<%		end if
		if trim(extracheckoutfield2)<>"" AND NOT (isprinter AND extracheckoutfield2noprint) then
			checkoutfield1=extracheckoutfield2
			checkoutfield2=editfunc(ordCheckoutExtra2,"checkoutextra2",25)
%>					<div class="container container3">
					  <div class="ordcheckoutextra2 left"><% if extracheckoutfield2reverse then print checkoutfield2 else print checkoutfield1 & ":" %></div>
					  <div class="ordcheckoutextra2 right"><% if extracheckoutfield2reverse then print checkoutfield1 else print checkoutfield2 %></div>
					</div>
<%		end if
		if ordAddInfo<>"" OR doedit then %>
					<div class="container container4">
					  <div class="ordaddinfo left"><%=replace(xxAddInf, "  ", "&nbsp;&nbsp;")%>:</div>
					  <div class="ordaddinfo right"><%
			if doedit then
				print "<textarea name=""ordAddInfo"" cols=""50"" rows=""4"">" & strip_tags2(ordAddInfo&"") & "</textarea>"
			else
				print replace(strip_tags2(ordAddInfo&""),vbNewLine,"<br />")
			end if %></div>
					</div>
<%		end if
		if NOT isprinter then %>
					<div class="container container5">
					  <div class="ordip left"><%=yyIPAdd%>:</div>
					  <div class="ordip right"><% if doedit then print editfunc(ordIP,"ipaddress",15) else print "<a href=""http://www.infosniper.net/index.php?lang=1&ip_address="&urlencode(ordIP&"")&""" target=""_blank"">"&htmlspecials(ordIP&"")&"</a>"%></div>
					</div>
					<div class="container container6">
					  <div class="ordaffiliate left"><%=yyAffili%>:</div>
					  <div class="ordaffiliate right"><%=editfunc(ordAffiliate,"PARTNER",15)%></div>
					</div>
<%		end if
		if (trim(ordDiscountText)<>"" AND (NOT isprinter OR isinvoice)) OR doedit then %>
					<div class="container container7">
					  <div class="orddiscounttext left"><%=xxAppDs%>:</div>
					  <div class="orddiscounttext right"><% if doedit then print "<textarea name=""discounttext"" cols=""50"" rows=""2"">" & replace(replace(ordDiscountText&"","<br />",vbNewLine),"<","&lt;") & "</textarea>" else print replace(htmlspecials(replace(ordDiscountText,"<br />",vbCrLf)),vbCrLf,"<br />") %></div>
					</div>
<%		end if
		if NOT isprinter then
		sSQL="SELECT gcaGCID,gcaAmount FROM giftcertsapplied WHERE gcaOrdID="&theid
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			print "<div class=""container container8""><div class=""left"">" & yyCerNum & "</div><div class=""right"">" & rs("gcaGCID") & " " & FormatEuroCurrency(rs("gcaAmount")) & " " & "<a href=""admingiftcert.asp?id=" & rs("gcaGCID") & """>" & yyClkVw & "</a></div></div>"
			rs.MoveNext
		loop
		rs.close
		end if
		ordAuthNumber=trim(ordAuthNumber&"")
		ordTransID=trim(ordTransID&"")
		if NOT isprinter AND (ordAuthNumber<>"" OR ordTransID<>"" OR doedit) then %>
					<div class="container container9">
					  <div class="ordauthnumber left"><%=yyAutCod%>:</div>
					  <div class="ordauthnumber right"><%=editfunc(ordAuthNumber,"ordAuthNumber",25) %></div>
					</div>
					<div class="container container10">
					  <div class="ordtransid left"><%=yyTranID%>:</div>
					  <div class="ordtransid right"><%=editfunc(ordTransID,"ordTransID",25) %></div>
					</div>
<%		end if
		if (isinvoice AND ordInvoice<>"") OR NOT isprinter then %>
					<div class="container container11">
					  <div class="ordinvoice left" id="invoiceupdatet"><%=yyInvNum%>:</div>
					  <div class="ordinvoice right"><%
						if isinvoice then
							print editfunc(ordInvoice,"ordInvoice",15)
						else %>
	<input type="text" name="ordInvoice" id="ordInvoice" size="25" value="<%=htmlspecials(ordInvoice)%>" <%=IIfVs(NOT (isprinter OR doedit),"onkeyup=""checksavefield('invoiceupdate',event,this,"&getget("id")&")"" ")%>/>
<%						end if %></div>
					</div>
<%		end if
		if NOT isprinter then
			if loyaltypoints<>"" then %>
					<div class="container container12">
					  <div class="loyaltypoints left"><%=xxLoyPoi%>:</div>
					  <div class="loyaltypoints right"><%=editfunc(loyaltypointtotal,"loyaltyPoints",10) %></div>
					</div>
<%			end if
			if ordReferer<>"" then %>
					<div class="container container13">
					  <div class="ordreferer left">Referer:</div>
					  <div class="ordreferer right"><input type="text" name="ordreferer" value="<%=replace(ordReferer & IIfVr(ordQuerystr<>"", "?" & ordQuerystr, ""), """", "&quot;")%>" size="80" /></div>
					</div>
<%			end if
			if ordUserAgent<>"" then %>
					<div class="container container15">
					  <div class="orduseragent left">User Agent:</div>
					  <div class="orduseragent right"><%=htmlspecials(ordUserAgent)%></div>
					</div><%
			end if
			if trim(ordAuthNumber&"")<>"" AND NOT doedit then %>
					<div class="container container14">
					  <div class="reemail left"></div>
					  <div class="reemail right"><input type="button" value="Resend Email" onclick="javascript:openemailpopup('id=<%=ordID%>')" /></div>
					</div>
<%			end if %>
				</div><%
			if getget("id")<>"new" AND is_numeric(getget("id")) then
				rs2.Open "SELECT ordStatusInfo FROM orders WHERE ordID="&getget("id"),cnn,0,1
				if NOT rs2.EOF then ordStatusInfo=rs2("ordStatusInfo")
				rs2.Close
				rs2.Open "SELECT ordPrivateStatus FROM orders WHERE ordID="&getget("id"),cnn,0,1
				if NOT rs2.EOF then ordPrivateStatus=rs2("ordPrivateStatus")
				rs2.Close
			end if
				%><div class="orderstatus <%=extraclass%>orderstatus">
					<div class="colheading"><h2><%=yyOrdSta%>.</h2></div>
<%			if NOT doedit then print "<form method=""post"" action=""adminorders.asp""><input type=""hidden"" name=""updatestatus"" value=""1"" /><input type=""hidden"" name=""orderid"" value="""&getget("id")&""" />"
				isauthorized=TRUE
				if ordAuthStatus="MODWARNOPEN" OR ordShipType="MODWARNOPEN" then
					isauthorized=FALSE %>
					<div class="container containerMW">
					  <div class="modwarning left ectred"><%=yyWarni%>:</div>
					  <div class="modwarning right"><%=yyMoWarn%></div>
					</div>
<%				end if %>
					<div class="container container1">
					  <div class="updatestatus left"><%=yyOrdSta%>:</div>
					  <div class="updatestatus right"><select name="ordStatus" size="1" onchange="document.getElementById('statusupdcontainer').style.display=''" style="width:190px"><%
				for index=0 to UBOUND(allstatus,2)
					if NOT isauthorized AND allstatus(0,index)>2 then exit for
					print "<option value=""" & allstatus(0,index) & """"
					if ordStatus=allstatus(0,index) then print " selected=""selected""" & ">" & allstatus(1,index) & " " & FormatDateTime(ordStatusDate, 2) & " " & FormatDateTime(ordStatusDate, 4) & "</option>" else print ">" & allstatus(1,index) & "</option>"
				next %></select></div>
					</div>
					<div class="container container2">
					  <div class="ordstatusinfo left"><%=yyStaInf%>:</div>
					  <div class="ordstatusinfo right">
						<textarea name="ordStatusInfo" id="ordStatusInfo" cols="50" rows="4" onkeyup="checkstatusupdfield(event,this)"><%=htmlspecials(ordStatusInfo&"")%></textarea>
					  </div>
					</div>
					<div class="container container3">
					  <div class="ordprivatestatus left"><%=yyPriSta%>:</div>
					  <div class="ordprivatestatus right"><textarea name="ordPrivateStatus" cols="50" rows="4" onkeyup="checkstatusupdfield(event,this)"><%=htmlspecials(ordPrivateStatus&"")%></textarea></div>
					</div>
<%			if NOT doedit then %>
					<div class="container container4">
					  <div class="emailstat left"><input type="checkbox" name="emailstat" value="1" <% if getpost("emailstat")="1" OR alwaysemailstatus=TRUE then print "checked"%> /></div>
					  <div class="emailstat right"><%=yyEStat%>
						<div style="padding-top:6px;display:none" id="statusupdcontainer"><input type="submit" value="<%=yyUpdate%>" /></div>
					  </div>
					</div>
<%				print "</form>"
			end if
		end if %>
				</div><%
			if ordShipCarrier<>0 OR ordShipType<>"" OR doedit then
				%><div class="shipping <%=extraclass%>shipping">
					<div class="colheading"><h2><%=yyShip%>.</h2></div>
					<div class="container container1">
					  <div class="ordshiptype left"><%=xxShpMet%>:</div>
					  <div class="ordshiptype right"><%
				if isprinter then
					print IIfVr(ordShipType="MODWARNOPEN",yyMoWarn,ordShipType)
				else %>
							<select name="shipcarrier" id="shipcarrier" size="1" onchange="dosavefield('shipcarrierupdate',<%=getget("id")%>,this[this.selectedIndex].value)">
							<option value="<%=ordShipCarrier%>"><%=yyOther%></option>
							<option value="3"<%if Int(ordShipCarrier)=3 then print " selected=""selected"""%>>USPS</option>
							<option value="4"<%if Int(ordShipCarrier)=4 then print " selected=""selected"""%>>UPS</option>
							<option value="6"<%if Int(ordShipCarrier)=6 then print " selected=""selected"""%>>CanPos</option>
							<option value="7"<%if Int(ordShipCarrier)=7 then print " selected=""selected"""%>>FedEx</option>
							<option value="8"<%if Int(ordShipCarrier)=8 then print " selected=""selected"""%>>FedEx SmartPost</option>
							<option value="9"<%if Int(ordShipCarrier)=9 then print " selected=""selected"""%>>DHL</option>
							<option value="10"<%if Int(ordShipCarrier)=10 then print " selected=""selected"""%>>Australia Post</option>
							</select> <%
				end if %></div>
					</div>
<%				isauthorized=NOT (ordAuthStatus="MODWARNOPEN" OR ordShipType="MODWARNOPEN")
				if (ordShipType<>"" OR doedit) AND NOT isprinter then %>
					<div class="container container2">
					  <div class="left<% if NOT isauthorized then print " ectred"%>"><% if NOT isauthorized then print yyWarni&":" else print "&nbsp;"%></div>
					  <div class="right"><% if isauthorized then print editfunc(ordShipType,"shipmethod",25) else print yyMoWarn %></div>
					</div>
<%				end if
				if NOT isprinter then %>
					<div class="container container3">
					  <div class="ordtracknum left" id="tracknumupdatet"><%=yyTraNum%>:</div>
					  <div class="ordtracknum right"><input type="text" name="ordTrackNum" id="ordTrackNum" size="30" value="<%=htmlspecials(ordTrackNum&"")%>" <%=IIfVs(NOT (isprinter OR doedit),"onkeyup=""checksavefield('tracknumupdate',event,this,"&getget("id")&")"" ")%>/></div>
					</div>
<%				end if %>
					<div class="container container4">
					  <div class="left"><% if doedit then print xxCLoc & ":"%></div>
					  <div class="right"><%	if doedit then
												print "<select name=""commercialloc"" id=""commercialloc"" size=""1"">"
												print "<option value=""N"">"&yyNo&"</option>"
												print "<option value=""Y"""&IIfVr((ordComLoc AND 1)=1," selected=""selected""","")&">"&yyYes&"</option>"
												print "</select>"
											end if %></div>
					</div>
<%				if doedit then %>
					<div class="container container5">
					  <div class="left"><%=xxShpIns%>:</div>
					  <div class="right"><%	print "<select name=""wantinsurance"" size=""1"">"
											print "<option value=""N"">"&yyNo&"</option>"
											print "<option value=""Y"""&IIfVr((ordComLoc AND 2)=2," selected=""selected""","")&">"&yyYes&"</option>"
											print "</select>" %></div>
					</div>
					<div class="container container6">
					  <div class="left"><%=xxSatDe2%>:</div>
					  <div class="right"><%	print "<select name=""saturdaydelivery"" size=""1"">"
											print "<option value=""N"">"&yyNo&"</option>"
											print "<option value=""Y"""&IIfVr((ordComLoc AND 4)=4," selected=""selected""","")&">"&yyYes&"</option>"
											print "</select>" %></div>
					</div>
					<div class="container container7">
					  <div class="left"><%=xxSigRe2%>:</div>
					  <div class="right"><%	print "<select name=""signaturerelease"" size=""1"">"
											print "<option value=""N"">"&yyNo&"</option>"
											print "<option value=""Y"""&IIfVr((ordComLoc AND 8)=8," selected=""selected""","")&">"&yyYes&"</option>"
											print "</select>" %></div>
					</div>
					<div class="container container8">
					  <div class="left"><%=xxInsDe2%>:</div>
					  <div class="right"><%	print "<select name=""insidedelivery"" size=""1"">"
											print "<option value=""N"">"&yyNo&"</option>"
											print "<option value=""Y"""&IIfVr((ordComLoc AND 16)=16," selected=""selected""","")&">"&yyYes&"</option>"
											print "</select>" %></div>
					</div>
<%				elseif ordComLoc>0 OR forceinsuranceselection then
					if isprinter then thestyle="" else thestyle=" style=""color:#FF0000"""
					shipopts="Shipping options:"
					if (ordComLoc AND 1)=1 then print "<div class=""container container9""><div class=""left"">"&shipopts&"</div><div class=""right""" & thestyle&">"&xxCerCLo&"</div></div>" : shipopts=""
					if ((ordComLoc AND 2)=2) OR forceinsuranceselection then print "<div class=""container container10""><div class=""left"">"&shipopts&"</div><div class=""right""" & thestyle&">"&IIfVr((ordComLoc AND 2)=2,xxShiInI,xxNoWtIn)&"</div></div>" : shipopts=""
					if (ordComLoc AND 4)=4 then print "<div class=""container container11""><div class=""left"">"&shipopts&"</div><div class=""right""" & thestyle&">"&xxSatDeR&"</div></div>" : shipopts=""
					if (ordComLoc AND 8)=8 then print "<div class=""container container12""><div class=""left"">"&shipopts&"</div><div class=""right""" & thestyle&">"&xxSigRe2&"</div></div>" : shipopts=""
					if (ordComLoc AND 16)=16 then print "<div class=""container container13""><div class=""left"">"&shipopts&"</div><div class=""right""" & thestyle&">"&xxInsDe2&"</div></div>" : shipopts=""
				end if %>
				</div><%
			end if

		if NOT isprinter AND ordAuthNumber<>"" AND NOT doedit then
			if ordPayProvider=21 then
				%><div class="authcapture">
					<div class="colheading"><h2>Amazon Capture.</h2></div>
					<div class="container container1">
					  <div class="left">Status:</div>
					  <div class="right ectred" id="googleupdatespan"></div>
					</div>
					<div class="container container2">
					  <div class="left">Capture Amount:</div>
					  <div class="right"><input type="text" name="txamount" id="txamount" size="5" value="<%=FormatNumber((ordTotal+ordStateTax+ordCountryTax+ordShipping+ordHSTTax+ordHandling)-ordDiscount, 2)%>" /></div>
					</div>
					<div class="container container3">
					  <div class="left">&nbsp;</div>
					  <div class="right"><input type="button" value="Capture Order" onclick="updategoogleorder('Amazon','settle',<%=ordID%>)" /></div>
					</div>
				</div><%
				elseif ordPayProvider=1 OR ordPayProvider=18 OR ordPayProvider=19 then
				%><div class="authcapture">
					<div class="colheading"><h2>PayPal Authorization / Capture.</h2></div>
					<div class="container container4">
					  <div class="left">Status:</div>
					  <div class="right ectred" id="googleupdatespan"></div>
					</div>
					<div class="container container5">
					  <div class="left">Capture Amount:</div>
					  <div class="right"><input type="text" name="captureamount" id="captureamount" size="10" value="<%=FormatNumber((ordTotal+ordStateTax+ordCountryTax+ordShipping+ordHSTTax+ordHandling)-ordDiscount, 2)%>" />
					  <select name="additionalcapture" id="additionalcapture" size="1"><option value="0">Close Authorization</option><option value="1">Leave Open for Additional Capture</option></select>
					  </div>
					</div>
					<div class="container container6">
					  <div class="left">Note to buyer:</div>
					  <div class="right"><textarea name="buyernote" id="buyernote" cols="50" rows="4"></textarea></div>
					</div>
					<div class="container container7">
					  <div class="left">Action:</div>
					  <div class="right">
						<select name="paypalaction" id="paypalaction" size="1" onchange="setpaypalelements()"><option value="charge">Capture</option><option value="void">Void</option><option value="reauth">Reauthorization</option></select>
					  </div>
					</div>
					<div class="container container8">
					  <div class="left">&nbsp;</div>
					  <div class="right"><input type="button" value="Inform PayPal" onclick="updatepaypalorder('PayPal',<%=ordID%>)" /></div>
					</div>
				</div><%
			elseif ordPayProvider=3 OR ordPayProvider=13 OR ordPayProvider=27 then
				isauthnet=ordPayProvider<>27
				%><div class="authcapture">
					<div class="colheading"><h2><%=IIfVr(isauthnet,"Authorize.net","PayPal Checkout")%> Authorization / Capture.</h2></div>
					<div class="container container4">
					  <div class="left">Status:</div>
					  <div class="right ectred" id="googleupdatespan"></div>
					</div>
					<div class="container container5">
					  <div class="left">Capture Amount:</div>
					  <div class="right"><input type="text" name="captureamount" id="captureamount" size="10" value="<%=FormatNumber((ordTotal+ordStateTax+ordCountryTax+ordShipping+ordHSTTax+ordHandling)-ordDiscount, 2)%>" />
<%				if NOT isauthnet then %>
						<select name="additionalcapture" id="additionalcapture" size="1"><option value="0">Close Authorization</option><option value="1">Leave Open for Additional Capture</option></select>
<%				end if %>
					  </div>
					</div>
					<div class="container container6">
					  <div class="left">Change Status To:</div>
					  <div class="right"><select id="capstatus" size="1" style="width:190px"><%
				for index=0 to UBOUND(allstatus,2)
					if NOT isauthorized AND allstatus(0,index)>2 then exit for
					print "<option value=""" & allstatus(0,index) & """"
					if ordStatus=allstatus(0,index) then print " selected=""selected""" & ">" & allstatus(1,index) & " " & FormatDateTime(ordStatusDate, 2) & " " & FormatDateTime(ordStatusDate, 4) & "</option>" else print ">" & allstatus(1,index) & "</option>"
				next %></select></div>
					</div>
					<div class="container container7">
					  <div class="left">Action:</div>
					  <div class="right">
						<select name="paypalaction" id="paypalaction" size="1" onchange="setpaypalelements()" style="width:190px">
							<option value="charge">Capture</option>
							<option value="void">Void</option>
<%				if NOT isauthnet then %>
							<option value="reauth">Reauthorization</option>
<%				end if %>
						</select>
					  </div>
					</div>
					<div class="container container8">
					  <div class="left">&nbsp;</div>
					  <div class="right"><input type="button" value="Inform <%=IIfVr(isauthnet,"Authorize.net","PayPal")%>" onclick="updatepaypalorder('<%=IIfVr(isauthnet,"AuthNET","PayPalCO")%>',<%=ordID%>)" /></div>
					</div>
				</div><%
			end if
			sSQL="SELECT upID,upComments,upFilename FROM imageuploads WHERE upOrderID=" & ordID
			rs2.open sSQL,cnn,0,1
			if NOT rs2.EOF then
				%><div class="imgupload">
<script language="javascript">
    function getuploadedimage(tid,ispreview){
        ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
        ajaxobj.open("GET", "ajaxservice.asp?action=imageget"+(ispreview?"&preview=true":"")+"&id="+tid,true); // Make sure file is in same server
        ajaxobj.send(null);
		ajaxobj.onreadystatechange=function(){
			if(ajaxobj.readyState==4){
				var image=document.getElementById("previewimupload");
				image.src="data:image/"+ajaxobj.responseText.substr(0,3)+";base64," + ajaxobj.responseText.substr(4);
				document.getElementById("uploadpreviewdiv").style.display='';
				document.getElementById("closeprevimg").style.display='';
			}
		};
    }
</script>
					<div class="colheading"><img src="adminimages/close.gif" id="closeprevimg" alt="Close Preview" style="float:right;display:none" onclick="document.getElementById('uploadpreviewdiv').style.display='none';this.style.display='none'" /><h2>Image Uploads.</h2>
						<div id="uploadpreviewdiv" style="text-align:center;display:none"><img style="max-width:90%" src="" id="previewimupload" /></div>
						<div id="uploadedimagestable" class="ecttable">
							<div class="ecttablerow"><div>Filename</div><div>Comments</div><div>&nbsp;</div><div>&nbsp;</div></div>
<%
					do while NOT rs2.EOF
						print "<div class=""ecttablerow""><div>" & htmldisplay(rs2("upFilename")) & "</div><div>" & htmldisplay(rs2("upComments")) & "</div><div><input type=""button"" value=""Preview"" onclick=""getuploadedimage(" & rs2("upID") & ",true)"" /></div><div><input type=""button"" value=""Download"" onclick=""document.location='ajaxservice.asp?action=imageget&id=" & rs2("upID") & "'"" /></div></div>" & vbCrLf
						rs2.movenext
					loop
%>						</div>
					</div>
				</div><%
			end if
			rs2.close
		end if
		
		if NOT isprinter AND NOT doedit then
			' if Int(ordPayProvider)=10 then
			if FALSE then
			%><div class="authcapture">
<%				if request.servervariables("HTTPS")<>"on" AND (Request.ServerVariables("SERVER_PORT_SECURE") <> "1") AND nochecksslserver<>TRUE then %>
					<div class="container" style="color:#FF0000;font-weight:bold">You do not appear to be viewing this page on a secure (https) connection. Credit card information cannot be shown.</div>
<%				else
					sSQL="SELECT ordCNum FROM orders WHERE ordID="&getget("id")
					rs2.Open sSQL,cnn,0,1
					if NOT rs2.EOF then ordCNum=rs2("ordCNum")
					rs2.Close
					if encryptmethod="aspencrypt" OR encryptmethod="" then %>
<OBJECT classid="CLSID:F9463571-87CB-4A90-A1AC-2284B7F5AF4E" 
	codeBase="https://www.ecommercetemplates.com/aspencrypt.dll" 
	id="XEncrypt">
</OBJECT>
<%					end if
					if ordCNum<>"" then
						if encryptmethod="none" then
							cnumarr=split(ordCNum, "&")
						elseif encryptmethod="aspencrypt" OR encryptmethod="" then
%>
<SCRIPT LANGUAGE="VBScript">
function URLDecodeHex(match, hex_digits, pos, source)
	URLDecodeHex=chr("&H" & hex_digits)
end function
function URLDecode(decstr)
	set re=new RegExp
	decstr=Replace(decstr, "+", " ")
	re.Pattern="%([0-9a-fA-F]{2})"
	re.Global=True
	URLDecode=re.Replace(decstr, GetRef("URLDecodeHex"))
end function
	' Set Context=XEncrypt.OpenContextEx("Microsoft Enhanced Cryptographic Provider v1.0", "mycontainer", False)
	Set Context=XEncrypt.OpenContext("mycontainer", False)
	Set Msg=Context.CreateMessage(True) ' use 3DES
	on error resume next
		err.number=0
		cnum=Msg.DecryptText("<%=Replace(ordCNum,vbNewLine,"")%>", "")
		If err.number=0 then
			cnumarr=split(cnum, "&")
		else
			Document.Write err.description
		end if
	on error goto 0
</SCRIPT>
<%						end if
					end if %>
					<div class="container container1">
					  <div class="left"><%=xxCCName%>:</div>
					  <div class="right"><%
					if encryptmethod="none" then
						if isarray(cnumarr) then
							if UBOUND(cnumarr)>=4 then print htmlspecials(URLDecode(cnumarr(4))&"")
						end if
					elseif encryptmethod="aspencrypt" OR encryptmethod="" then
%><SCRIPT LANGUAGE="VBScript">
	if isarray(cnumarr) then
		if UBOUND(cnumarr)>=4 then Document.Write URLDecode(cnumarr(4))
	end if
</SCRIPT><%			end if %></div>
					</div>
					<div class="container container2">
					  <div class="left"><%=yyCarNum%>:</div>
					  <div class="right"><%
					if ordCNum<>"" then
						if encryptmethod="none" then
							if isarray(cnumarr) then print htmlspecials(cnumarr(0)&"")
						elseif encryptmethod="aspencrypt" OR encryptmethod="" then
%><SCRIPT LANGUAGE="VBScript">
	if isarray(cnumarr) then Document.Write cnumarr(0)
</SCRIPT><%				end if
					else
						print "(no data)"
					end if %></div>
					</div>
					<div class="container container3">
					  <div class="left"><%=yyExpDat%>:</div>
					  <div class="right"><%
					if encryptmethod="none" then
						if isarray(cnumarr) then
							if UBOUND(cnumarr)>=1 then print htmlspecials(cnumarr(1)&"")
						end if
					elseif encryptmethod="aspencrypt" OR encryptmethod="" then
%><SCRIPT LANGUAGE="VBScript">
	if isarray(cnumarr) then
		if UBOUND(cnumarr)>=1 then Document.Write cnumarr(1)
	end if
</SCRIPT><%			end if %></div>
					</div>
					<div class="container container4">
					  <div class="left">CVV Code:</div>
					  <div class="right"><%
					if encryptmethod="none" then
						if isarray(cnumarr) then
							if UBOUND(cnumarr)>=2 then print htmlspecials(cnumarr(2)&"")
						end if
					elseif encryptmethod="aspencrypt" OR encryptmethod="" then
%><SCRIPT LANGUAGE="VBScript">
	if isarray(cnumarr) then
		if UBOUND(cnumarr)>=2 then Document.Write cnumarr(2)
	end if
</SCRIPT><%			end if %></div>
					</div>
					<div class="container container5">
					  <div class="left">Issue Number:</div>
					  <div class="right"><%
					if encryptmethod="none" then
						if isarray(cnumarr) then
							if UBOUND(cnumarr)>=3 then print htmlspecials(cnumarr(3)&"")
						end if
					elseif encryptmethod="aspencrypt" OR encryptmethod="" then
%><SCRIPT LANGUAGE="VBScript">
	if isarray(cnumarr) then
		if UBOUND(cnumarr)>=3 then Document.Write cnumarr(3)
	end if
</SCRIPT><%			end if %></div>
					</div>
<%				end if
				if ordCNum<>"" AND NOT doedit then %>
				  <form method="post" action="adminorders.asp?id=<%=getget("id")%>">
					<input type="hidden" name="delccdets" value="<%=getget("id")%>" />
					<div class="container container6">
					  <div class="left">&nbsp;</div>
					  <div class="right"><input type="submit" value="<%=yyDelCC%>" /></div>
					</div>
				  </form>
<%				end if %>
				</div><%
			end if
		end if ' isprinter

		WSP="" : OWSP="" : percdisc="" : wholesaledisc=FALSE
		if ordClientID<>0 then
			sSQL="SELECT clActions,clPercentDiscount FROM customerlogin WHERE clID="&ordClientID
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if (rs("clActions") AND 8)=8 then
					WSP="pWholesalePrice AS "
					wholesaledisc=TRUE
					if wholesaleoptionpricediff=TRUE then OWSP="optWholesalePriceDiff AS "
				end if
				if (rs("clActions") AND 16)=16 then
					WSP=((100.0-cdbl(rs("clPercentDiscount")))/100.0) & "*"&IIfVr((rs("clActions") AND 8)=8,"pWholesalePrice","pPrice")&" AS "
					percdisc=rs("clPercentDiscount")
					OWSP=((100.0-rs("clPercentDiscount"))/100.0) & "*"&IIfVr((rs("clActions") AND 8)=8 AND wholesaleoptionpricediff,"optWholesalePriceDiff","optPriceDiff")&" AS "
				end if
			end if
			rs.close
		end if
		displayimagecolumn=(isinvoice AND (imgonorderdetails AND 4)=4) OR (NOT isinvoice AND isprinter AND (imgonorderdetails AND 2)=2) OR (NOT isprinter AND (imgonorderdetails AND 1)=1)
		%><div id="producttable" class="producttable">
			<div class="ectorderrow ectorderhead">
			  <div class="ordrowprodidcol"><%=xxPrId%></div>
			  <div class="ordrowprodnamecol"><%=xxPrNm%></div>
			  <div class="ordrowoptionscol"><%=xxPrOpts%></div>
<%		if isinvoice then print "<div class=""ordrowunitpricecol"">" & xxUnitPr & "</div>"
		if NOT (isprinter OR doedit) then print "<div class=""ordrowweightcol""><div class=""largescreenquant ordsmallheader"">Weight</div></div>" %>
			  <div class="ordrowquantcol"><div class="largescreenquant"><%=xxQuant%></div></div>
<%		if NOT isprinter OR isinvoice then print "<div class=""" & IIfVr(doedit,"ordroweditpricecol ","ordrowpricecol ") & "ordsmallheader"">" & IIfVr(doedit, xxUnitPr, xxPrice) & "</div>"
		if doedit then print "<div class=""ordrowdeletecol ordsmallheader"">DEL</div>" %>
			</div>
<%		stockjs=""
		if initialpackweight<>"" AND ordShipCarrier<>5 then totweight=initialpackweight else totweight=0
		if isarray(allorders) then
			totoptpricediff=0
			ordrowclass="orderevenrow"
			for rowcounter=0 to UBOUND(allorders,2)
				if ordrowclass="orderoddrow" then ordrowclass="orderevenrow" else ordrowclass="orderoddrow"
				optpricediff=0 : optweightdiff=0
				cartGiftMessage=trim(allorders(8,rowcounter))
				if allorders(5, rowcounter)=0 AND ordAuthStatus<>"MODWARNOPEN" then stockjs=stockjs & "stock['pid_" & allorders(0, rowcounter) & "']+=" & allorders(3, rowcounter) & ";" & vbCrLf
%>
			<div class="ectorderrow<% if allorders(7,rowcounter)<>0 AND NOT isinvoice then print " ectgiftwraprow"%><%=" "&ordrowclass%>">
			  <div class="ordrowprodidcol"><%
				orddetailsimage=""
				if displayimagecolumn then
					sSQL="SELECT imageSrc FROM productimages WHERE imageProduct='" & escape_string(allorders(0,rowcounter)) & "' AND (imageType=0 OR imageType=1) ORDER BY imageType,imageNumber"
					rs.open sSQL,cnn,0,1
					orddetailsimage=defaultorddetailsimage
					if NOT rs.EOF then orddetailsimage=rs("imageSrc")
					rs.close
					if orddetailsimage<>"" then
						orddetailsimage=IIfVs(lcase(left(orddetailsimage,5))<>"http:" AND lcase(left(orddetailsimage,6))<>"https:" AND left(orddetailsimage,1)<>"/","../") & orddetailsimage
						print "<div class=""orddetailsimg""><img class=""orddetailsimg"" src=""" & orddetailsimage & """ alt="""" /><div>"
					end if
				end if %><div class="smallscreenquant"><%=editspecial(allorders(3,rowcounter),"smallquant"&rowcounter,3,"onkeyup=""document.getElementById('quant"&rowcounter&"').value=this.value"" onchange=""dorecalc(true)""")%> x </div><%
				if doedit then print "<input type=""button"" value=""..."" onclick=""updateoptions("&rowcounter&")"">&nbsp;<input type=""hidden"" name=""cartid"&rowcounter&""" value=""" & htmlspecials(allorders(4,rowcounter)) & """ /><input type=""hidden"" id=""stateexempt"&rowcounter&""" value="""&IIfVr((allorders(6,rowcounter) AND 1)=1,"true","false")&""" /><input type=""hidden"" id=""countryexempt"&rowcounter&""" value="""&IIfVr((allorders(6,rowcounter) AND 2)=2,"true","false")&""" />"
				thelink=""
				if trim(allorders(10,rowcounter)&"")<>"" then cartprodid=allorders(10,rowcounter) else cartprodid=allorders(0,rowcounter)
				sSQL="SELECT pID,"&getlangid("pName",1)&",pStaticPage,pStaticURL,pDisplay FROM products WHERE pDisplay<>0 AND pID='"&escape_string(cartprodid)&"'"
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					thelink=storeurl & getdetailsurl(rs("pID"),rs("pStaticPage"),rs(getlangid("pName",1)),trim(rs("pStaticURL")&""),"","")
				end if
				rs.close
				if doedit OR isprinter then print editspecial(allorders(0,rowcounter),"prodid"&rowcounter,18,"AUTOCOMPLETE=""off"" onkeydown=""return combokey(this,event)"" onkeyup=""combochange(this,event)""") else print IIfVs(allorders(0,rowcounter)<>giftwrappingid,"<a href=""" & thelink & """ target=""_blank"">") & htmlspecials(allorders(0,rowcounter)) & IIfVs(allorders(0,rowcounter)<>giftwrappingid,"</a>")
				if allorders(0,rowcounter)=giftcertificateid then
					sSQL="SELECT gcID FROM giftcertificate WHERE gcCartID=" & allorders(4,rowcounter)
					rs.open sSQL,cnn,0,1
					if NOT rs.EOF then
						print "<input type=""button"" value="""&yyView&""" onclick=""document.location='admingiftcert.asp?id="&rs("gcID")&"'"" style=""margin-left:4px"" />"
					end if
					rs.close
				end if
				if doedit then print showgetoptionsselect("selectprodid"&rowcounter)
				if orddetailsimage<>"" then print "</div></div>"
			%></div>
			  <div class="ordrowprodnamecol"><%
				print replace(editspecial(decodehtmlentities(allorders(1,rowcounter)),"prodname"&rowcounter,24,"AUTOCOMPLETE=""off"" onkeydown=""return combokey(this,event)"" onkeyup=""combochange(this,event)"""),"&amp;","&")
				if doedit then
					print showgetoptionsselect("selectprodname"&rowcounter)
				elseif NOT isinvoice then
					sSQL="SELECT productpackages.pID,quantity,pName,quantity FROM productpackages INNER JOIN products on productpackages.pID=products.pID WHERE packageID='"&escape_string(allorders(0,rowcounter))&"'"
					rs.open sSQL,cnn,0,1
					if NOT rs.EOF then
						print "<div class=""ordpackage"">"
						do while NOT rs.EOF
							print "<div class=""ordpackagerow""><div>&nbsp;&gt;&nbsp;" & rs("pID") & ":</div><div>" & rs("pName") & "</div><div>" & rs("quantity") & "</div></div>"
							rs.movenext
						loop
						print "</div>"
					end if
					rs.close
				end if
			%></div>
			  <div class="ordrowoptionscol"><%
				if doedit then print "<div id=""optionsspan"&rowcounter&""">"
				sSQL="SELECT coOptGroup,coCartOption,coPriceDiff,coWeightDiff,coOptID,optGroup FROM cartoptions LEFT JOIN options ON cartoptions.coOptID=options.optID WHERE coCartID="&allorders(4,rowcounter) & " ORDER BY coID"
				rs2.Open sSQL,cnn,0,1
				if NOT rs2.EOF OR (allorders(7,rowcounter)<>0 AND cartGiftMessage<>"" AND NOT isinvoice) then
					if allorders(5, rowcounter)<>0 AND ordAuthStatus<>"MODWARNOPEN" then stockjs=stockjs & "stock['oid_" & rs2("coOptID") & "']+=" & allorders(3, rowcounter) & ";" & vbCrLf
					if doedit then print "<div class=""optionstable"">"
					do while NOT rs2.EOF
						if doedit then
							print "<div class=""optionstablerow""><div class=""optionstableleft"">" & rs2("coOptGroup") & ":</div><div>"
							if IsNull(rs2("optGroup")) then
								print "xxxxxx"
							else
								sSQL="SELECT optID,"&getlangid("optName",32)&","&OWSP&"optPriceDiff,optType,optFlags,optStock,optTxtCharge,optWeightDiff FROM options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optGroup=" & rs2("optGroup")
								rs3.Open sSQL,cnn,0,1
								if NOT rs3.EOF then
									if abs(rs3("optType"))=1 OR abs(rs3("optType"))=2 OR abs(rs3("optType"))=4 then
										print "<select onchange=""dorecalc(true)"" name=""optn"&rowcounter&"_"&rs2("coOptID")&""" id=""optn"&rowcounter&"_"&rs2("coOptID")&""" size=""1"">"
										do while NOT rs3.EOF
											print "<option value="""&rs3("optID")&"|"&IIfVr((rs3("optFlags") AND 1)=1,(allorders(2,rowcounter)*rs3("optPriceDiff"))/100.0,rs3("optPriceDiff"))&""""
											if rs3("optID")=rs2("coOptID") then print " selected=""selected"""
											print ">"&rs3(getlangid("optName",32))
											if cdbl(rs3("optPriceDiff"))<>0 then
												print " "
												if cdbl(rs3("optPriceDiff")) > 0 then print "+"
												if (rs3("optFlags") AND 1)=1 then
													print FormatNumber((allorders(2,rowcounter)*rs3("optPriceDiff"))/100.0,2)
												else
													print FormatNumber(rs3("optPriceDiff"),2)
												end if
											end if
											print "</option>"
											rs3.MoveNext
										loop
										print "</select>"
									else
										if rs3("optTxtCharge")<>0 then print "<script>opttxtcharge["&rs3("optID")&"]="&rs3("optTxtCharge")&";</script>"
										print "<input type='hidden' name='optn"&rowcounter&"_"&rs2("coOptID")&"' value='"&rs3("optID")&"' /><textarea name='voptn"&rowcounter&"_"&rs2("coOptID")&"' id='voptn"&rowcounter&"_"&rs2("coOptID")&"' cols='30' rows='3'>"
										print htmlspecials(rs2("coCartOption")) & "</textarea>"
									end if
								end if
								rs3.Close
							end if
							print "</div></div>"
						else
							print "<div class=""orderoptline"">" & IIfVs(trim(rs2("coOptGroup")&"")<>"","<div class=""orderoptgrp"">" & rs2("coOptGroup") & "</div>") & "<div class=""orderoption"">" & replace(replace(htmlspecials(rs2("coCartOption")&""),"  ","&nbsp;&nbsp;"),vbLf,"<br />") & "</div></div>"
						end if
						if doedit then
							optpricediff=optpricediff + rs2("coPriceDiff")
						else
							allorders(2,rowcounter)=allorders(2,rowcounter) + rs2("coPriceDiff")
							optweightdiff=optweightdiff+rs2("coWeightDiff")
						end if
						rs2.MoveNext
					loop
					if allorders(7,rowcounter)<>0 AND cartGiftMessage<>"" AND NOT isinvoice then
						print IIfVs(doedit,"<div class=""optionstablerow""><div class=""optionstableleft ordgiftwrapmessaget"">") & "Gift Wrap Message: " & IIfVs(doedit,"</div><div>") & cartGiftMessage & IIfVr(doedit,"</div></div>","<br />")
					end if
					if doedit then print "</div>"
				else
					print "<div>-</div>"
				end if
				rs2.Close
				if doedit then print "</div>" %></div>
<%				if isinvoice then print "<div class=""ordrowunitpricecol"">" & FormatEuroCurrency(allorders(2,rowcounter)) & "</div>"
				prodweight=(allorders(9,rowcounter)+optweightdiff)*allorders(3,rowcounter)
				if isnull(prodweight) then prodweight=0
				totweight=totweight+prodweight
				if NOT (isprinter OR doedit) then print "<div class=""ordrowweightcol""><div class=""largescreenquant"">" & vsround(prodweight,3) & "</div></div>" %>
			  <div class="ordrowquantcol<%=IIfVs(allorders(3,rowcounter)>1," ordmultiquantcol")%>"><div class="largescreenquant"><%=editspecial(allorders(3,rowcounter),"quant"&rowcounter,5,"onkeyup=""document.getElementById('smallquant"&rowcounter&"').value=this.value"" onchange=""dorecalc(true)""")%></div></div>
<%				if NOT isprinter OR isinvoice then %>
			  <div class="<%=IIfVr(doedit,"ordroweditpricecol","ordrowpricecol")%>"><%if doedit then print editnumeric(allorders(2,rowcounter),"price"&rowcounter,7,"class=""orderprice"" onchange=""dorecalc(true)"" ") else print FormatEuroCurrency(allorders(2,rowcounter)*allorders(3,rowcounter))%>
<%						if doedit then
							print "<input type=""hidden"" id=""optdiffspan"&rowcounter&""" value="""&optpricediff&""">"
							totoptpricediff=totoptpricediff + (optpricediff*allorders(3,rowcounter))
						end if
			%></div>
<%				end if
				if doedit then print "<div class=""ordrowdeletecol""><input type=""checkbox"" name=""del_"&rowcounter&""" id=""del_"&rowcounter&""" value=""yes"" /></div>" %>
			</div><%
			next
		end if %>
		</div>
<%		if NOT isprinter OR isinvoice then %>
		<div>
			<div class="ectorderfunctions">
<%			if doedit then %>
				<div style="padding:20px;text-align:center">
<%				print "<input style=""width:30px;"" type=""button"" value=""-"" onclick=""extraproduct('-')""> "&yyMoProd&" <input style=""width:30px;"" type=""button"" value=""+"" onclick=""extraproduct('+')"">"%>
				</div>
				<div style="display:table;width:100%">
<%				if useStockManagement then print "<div class=""ectorderrow""><div style=""width:40%;text-align:right""><input type=""checkbox"" name=""updatestock"" value=""ON"" checked=""checked"" /></div><div>" & yyUpStLv & "</div></div>" %>
					<div class="ectorderrow">
				<div style="text-align:right"><input type="checkbox" id="wholesaledisc" value="ON"<%=IIfVs(wholesaledisc," checked=""checked""")%> /></div><div><%=yyWholPr%></div>
					</div>
					<div class="ectorderrow">
				<div style="text-align:right"><input type="text" id="percdisc" size="3" value="<%=percdisc%>" /></div><div><%=yyPerDis%></div>
					</div>
				</div>
				<div style="padding:20px;text-align:center">
					<input type="button" value="<%=yyRecal%>" onclick="dorecalc(false)">
				</div>
<%			else
				print "&nbsp;"
			end if %>
			</div><div class="ecttotalcolumn">
				<hr class="ordertotals" />
			<div class="ecttotaltable">
<%			if doedit then %>
				<div class="ecttotalrow">
					<div><%=replace(yyOptTot, " ", "&nbsp;")%>:</div>
					<div><span id="optdiffspan"><%=FormatNumber(totoptpricediff, 2)%></span><script>
			var stock=new Array();
<%				optgroups=""
				addcomma=""
				if cstr(theid)<>"0" then
					sSQL="SELECT DISTINCT cartID,pID,pInStock,pStockByOpts FROM cart INNER JOIN products ON cart.cartProdId=products.pID WHERE cartOrderID="&theid
					rs.open sSQL,cnn,0,1
					do while NOT rs.EOF
						print "stock['pid_"&rs("pID")&"']="
						if rs("pStockByOpts")=0 then
							print rs("pInStock")&";"&vbCrLf
						else
							print "'bo';"&vbCrLf
							if mysqlserver=TRUE then
								sSQL="SELECT coID,optStock,coOptID,optGrpID FROM cart INNER JOIN cartoptions ON cart.cartID=cartoptions.coCartID INNER JOIN options ON cartoptions.coOptID=options.optID INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optType IN (-4,-2,-1,1,2,4) AND cartID="&rs("cartID")
							else
								sSQL="SELECT DISTINCT optGrpID FROM cart INNER JOIN (cartoptions INNER JOIN (options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID) ON cartoptions.coOptID=options.optID) ON cart.cartID=cartoptions.coCartID WHERE optType IN (-4,-2,-1,1,2,4) AND cartID="&rs("cartID")
							end if
							rs2.Open sSQL,cnn,0,1
							do while NOT rs2.EOF
								optgroups=optgroups & addcomma & rs2("optGrpID")
								addcomma=","
								rs2.movenext
							loop
							rs2.Close
						end if
						rs.movenext
					loop
					rs.close
				end if
				if optgroups<>"" then
					sSQL="SELECT optID,optStock FROM options WHERE optGroup IN (" & optgroups & ")"
					rs.open sSQL,cnn,0,1
					do while NOT rs.EOF
						print "stock['oid_"&rs("optID")&"']=" & rs("optStock")&";"&vbCrLf
						rs.movenext
					loop
					rs.close
				end if
				print stockjs
				if getget("id")="new" then print "extraproduct('+');" & vbCrLf
%></script></div>
				</div>
<%			end if
			if NOT isprinter AND NOT doedit then %>
				<div class="ecttotalrow">
					<div><%="Total Weight"%>:</div>
					<div><%=vsround(totweight,3)%></div>
				</div>
<%			end if %>
				<div class="ecttotalrow">
					<div><%=replace(xxOrdTot, " ", "&nbsp;")%>:</div>
					<div><%=editnumeric(ordTotal,"ordtotal",7,"")%></div>
				</div>
<%			if isprinter AND combineshippinghandling=TRUE then %>
				<div class="ecttotalrow">
					<div><%=xxShipHa%>:</div>
					<div><%=FormatEuroCurrency(ordShipping+ordHandling)%></div>
				</div>
<%			else
				if ordShipping > 0 OR doedit then %>
				<div class="ecttotalrow">
					<div><%=xxShippg%>:</div>
					<div><%=editnumeric(ordShipping,"ordShipping",7,"")%></div>
<%					if doedit then print "<div><input type=""button"" value="""&"Calculate"&""" onclick=""calcshipping()"" /></div>" %>
				</div>
<%				end if
				if cdbl(ordHandling)<>0.0 OR doedit then %>
				<div class="ecttotalrow">
					<div><%=xxHndlg%>:</div>
					<div><%=editnumeric(ordHandling,"ordHandling",7,"")%></div>
				</div>
<%				end if
			end if
			if cdbl(ordDiscount)<>0.0 OR doedit then %>
				<div class="ecttotalrow">
					<div><%=xxDscnts%>:</div>
					<div><span style="color:#FF0000;font-weight:bold"><%=editnumeric(ordDiscount,"ordDiscount",7,"")%></span></div>
				</div>
<%			end if
			if ordStateTax > 0 OR doedit then %>
				<div class="ecttotalrow">
					<div><%=xxStaTax%>:</div>
					<div><%=editnumeric(ordStateTax,"ordStateTax",7,"")%></div>
<%				if doedit then print "<div><input type=""checkbox"" id=""orderstatetaxexempt""" & IIfVs(NOT orderstatetaxexempt," checked=""checked""") & " title=""Liable / Exempt"" onchange=""document.getElementById('staterate').disabled=!this.checked"" /><input type=""text"" style=""text-align:right;width:24px"" name=""staterate"" id=""staterate"" value="""&statetaxrate&""" " & IIfVs(orderstatetaxexempt," disabled=""disabled""") & "/>%</div>" %>
				</div>
<%			end if
			if ordCountryTax > 0 OR doedit then %>
				<div class="ecttotalrow">
					<div><%=xxCntTax & IIfVs(NOT doedit AND countrytaxrate<>0, " - " & countrytaxrate & "%") %>:</div>
					<div><%=editnumeric(ordCountryTax,"ordCountryTax",7,"")%></div>
<%				if doedit then print "<div><input type=""checkbox"" id=""ordercountrytaxexempt""" & IIfVs(NOT ordercountrytaxexempt," checked=""checked""") & " title=""Liable / Exempt"" onchange=""document.getElementById('countryrate').disabled=!this.checked"" /><input type=""text"" style=""text-align:right;width:24px"" name=""countryrate"" id=""countryrate"" value="""&countrytaxrate&""" " & IIfVs(ordercountrytaxexempt," disabled=""disabled""") & "/>%</div>" %>
				</div>
<%			end if
			if ordHSTTax > 0 OR (doedit AND origCountryID=2) then %>
				<div class="ecttotalrow">
					<div><%=xxHST%>:</div>
					<div><%=editnumeric(ordHSTTax,"ordHSTTax",7,"")%></div>
<%				if doedit then print "<div><input type=""text"" style=""text-align:right;width:10px"" name=""hstrate"" id=""hstrate"" value="""&hsttaxrate&""">%</div>" %>
				</div>
<%			end if %>
				<div class="ecttotalrow">
					<div><%=xxGndTot%>:</div>
					<div class="ordergrandtotal" id="grandtotalspan"><%=FormatEuroCurrency((ordTotal+ordStateTax+ordCountryTax+ordShipping+ordHSTTax+ordHandling)-ordDiscount)%></div>
				</div>
			</div>
			</div>
		</div>
<%		end if ' NOT isprinter OR isinvoice

		if isprinter AND isempty(packingslipfooter) then packingslipfooter=invoicefooter
		if isinvoice AND invoicefooter<>"" then %>
		<div class="orderfooter invoicefooter"><%=invoicefooter%></div>
<%		elseif isprinter AND packingslipfooter<>"" then %>
		<div class="orderfooter packslipfooter"><%=packingslipfooter%></div>
<%		elseif doedit then %>
		<div style="padding:30px;text-align:center"><input type="submit" value="<%=yyUpdate%>" /></div>
<%		end if
		if getget("id")<>"multi" AND getget("id")<>"new" AND NOT doedit then %>
		<div class="bottombuttons no-print">
			<div class="bottombuttonleft"><%
			print "<div class=""bottombuttongroup"">"
			print "<input "&IIfVs(previousid="","disabled=""disabled"" ")&"class=""no-print"" style=""width:110px;margin:2px"" type=""button"" value=""&laquo; "&yyPrev&""" onclick=""researchformgo('"&previousid&"',0)"" />"
			print "<input "&IIfVs(previousbysearch="","disabled=""disabled"" ")&"class=""no-print"" style=""width:110px;margin:2px"" type=""button"" value="""&"Prev. By Search"&""" onclick=""researchformgo('"&previousbysearch&"',0)"" />"
			print "</div><div class=""bottombuttongroup"">"
			print "<input "&IIfVs(NOT(getget("printer")="true" OR getget("invoice")="true"),"disabled=""disabled"" ")&"class=""no-print"" style=""width:110px;margin:2px"" type=""button"" value="""&"Print"&""" onclick=""window.print()"" />"
			print "<input class=""no-print"" style=""width:110px;margin:2px"" type=""button"" value="""&"Orders"&""" onclick=""researchformgo('"&ordID&"',3)"" />"
			print "</div>"
			%></div><div class="bottombuttonright">
<%			print "<div class=""bottombuttongroup"">"
			print "<input "&IIfVs(nextbysearch="","disabled=""disabled"" ")&"class=""no-print"" style=""width:110px;margin:2px"" type=""button"" value="""&"Next By Search"&""" onclick=""researchformgo('"&nextbysearch&"',0)"" />"
			print "<input "&IIfVs(nextid="","disabled=""disabled"" ")&"class=""no-print"" style=""width:110px;margin:2px"" type=""button"" value="""&yyNext&" &raquo;"" onclick=""researchformgo('"&nextid&"',0)"" />"
			print "</div><div class=""bottombuttongroup"">"
			print "<input class=""no-print"" style=""width:110px;margin:2px"" type=""button"" value="""&IIfVr(getget("invoice")="true","Details","Invoice")&""" onclick=""researchformgo('"&ordID&"',"&IIfVr(getget("invoice")="true",4,1)&")"" />"
			print "<input class=""no-print"" style=""width:110px;margin:2px"" type=""button"" value="""&IIfVr(getget("printer")="true","Details","Packing Slip")&""" onclick=""researchformgo('"&ordID&"',"&IIfVr(getget("printer")="true",4,2)&")"" />"
			print "</div>"
%>			</div>
		</div>
<%		end if %>		
	  </div>
<%		if doedit then print "</form>"
	next ' for each objItem in idlist
else
	sSQL="SELECT ordID FROM orders WHERE ordStatus=1"
	if getpost("act")<>"purge" then sSQL=sSQL & " AND ordStatusDate<" & vsusdate(thedate - 3)
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		theid=rs("ordID")
		addcomma=""
		delOptions=""
		sSQL="SELECT cartID FROM cart WHERE cartOrderID="&theid
		rs3.Open sSQL,cnn,0,1
		do while NOT rs3.EOF
			delOptions=delOptions & addcomma & rs3("cartID")
			addcomma=","
			rs3.MoveNext
		loop
		rs3.Close
		if delOptions<>"" then ect_query("DELETE FROM cartoptions WHERE coCartID IN ("&delOptions&")")
		ect_query("DELETE FROM cart WHERE cartOrderID="&theid)
		ect_query("DELETE FROM giftcertificate WHERE gcOrderID="&theid)
		ect_query("DELETE FROM giftcertsapplied WHERE gcaOrdID="&theid)
		ect_query("DELETE FROM orders WHERE ordID="&theid)
		rs.MoveNext
	loop
	rs.close
	if getpost("act")="authorize" then
		ect_query("UPDATE orders set ordAuthNumber='" & escape_string(IIfVr(getpost("authcode")<>"",getpost("authcode"),yyManAut)) & "' WHERE ordID="&getpost("id"))
		ect_query("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&getpost("id"))
		call updateorderstatus(getpost("id"), 3, getpost("emailstat")="1")
	elseif getpost("act")="unpending" then
		ect_query("UPDATE orders set ordAuthStatus='' WHERE ordID="&getpost("id"))
		ect_query("UPDATE orders set ordShipType='"&escape_string(yyMoWarn)&"' WHERE ordShipType='MODWARNOPEN' AND ordID="&getpost("id"))
		ect_query("UPDATE orders set ordAuthNumber='"&escape_string(yyManAut)&"' WHERE ordAuthNumber='' AND ordID="&getpost("id"))
		ect_query("UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&getpost("id"))
		sSQL="SELECT ordStatus FROM orders WHERE ordID="&getpost("id")
		rs.open sSQL,cnn,0,1
		oldordstatus=rs("ordStatus")
		rs.close
		if oldordstatus<3 then call updateorderstatus(getpost("id"), IIfVr(oldordstatus<3,3,oldordstatus), getpost("emailstat")="1")
	elseif getpost("act")="editablefield" then
		response.cookies("editablefield")=getpost("id")
		response.cookies("editablefield").Expires=Date()+365
		if request.servervariables("HTTPS")="on" then response.cookies("editablefield").secure=TRUE
	elseif getpost("act")="searchfield" then
		response.cookies("searchfield")=getpost("id")
		response.cookies("searchfield").Expires=Date()+365
		if request.servervariables("HTTPS")="on" then response.cookies("searchfield").secure=TRUE
	elseif getpost("act")="status" AND getpost("theeditablefield")<>"" AND getpost("theeditablefield")<>"status" then
		maxitems=int(getpost("maxitems"))
		editfield=getpost("theeditablefield")
		for mindex=0 to maxitems-1
			iordid=getpost("ordid" & mindex)
			ect_query("UPDATE orders SET ord"&getpost("theeditablefield")&"='" & escape_string(getpost(editfield & mindex)) & "' WHERE ordID=" & iordid)
		next
	elseif getpost("act")="status" then
		maxitems=int(getpost("maxitems"))
		for mindex=0 to maxitems-1
			call updateorderstatus(getpost("ordid" & mindex), int(request.form("ordStatus" & mindex)), getpost("emailstat")="1")
		next
	end if
	hastodate=FALSE : hasfromdate=FALSE
	todate=trim(request("todate")) : fromdate=trim(request("fromdate"))
	thetodate=thedate : thefromdate=thedate
	call getdates()
	sSQL="SELECT DISTINCT ordID,ordName,ordLastName,payProvName,ordAuthNumber,ordDate,ordStatus,(ordTotal-ordDiscount) AS ordTot,ordTransID,ordAVS,ordCVV,ordPayProvider,ordAuthStatus,ordTrackNum,ordInvoice,ordShipType,ordEmail FROM (orders INNER JOIN payprovider ON payprovider.payProvID=orders.ordPayProvider)"
	if getpost("abandonedstatus")<>"" AND getpost("abandonedstatus")<>"1" then
		sSQL=sSQL&" INNER JOIN abandonedcartemail ON orders.ordID=abandonedcartemail.aceOrderID"
	end if
	whereSQL="" : namesql=""
	origsearchtext=htmlspecials(getrequest("searchtext"))
	call getordsearchwheresql()
	sSQL=sSQL&whereSQL
	editablefield=request.cookies("editablefield")
	sortorder=request.cookies("ordersort")
	if sortorder="oidd" then
		orderSQL=sSQL & " ORDER BY ordID DESC"
	elseif sortorder="orna" then
		orderSQL=sSQL & " ORDER BY ordName"
	elseif sortorder="ornd" then
		orderSQL=sSQL & " ORDER BY ordName DESC"
	elseif sortorder="orda" then
		orderSQL=sSQL & " ORDER BY ordDate"
	elseif sortorder="ordd" then
		orderSQL=sSQL & " ORDER BY ordDate DESC"
	elseif sortorder="oraa" then
		orderSQL=sSQL & " ORDER BY ordAuthNumber"
	elseif sortorder="orad" then
		orderSQL=sSQL & " ORDER BY ordAuthNumber DESC"
	elseif sortorder="orpa" then
		orderSQL=sSQL & " ORDER BY ordPayProvider"
	elseif sortorder="orpd" then
		orderSQL=sSQL & " ORDER BY ordPayProvider DESC"
	elseif sortorder="orsa" OR sortorder="orsd" then
		if editablefield="tracknum" then
			orderSQL=sSQL & " ORDER BY ordTrackNum"
		elseif editablefield="tracknum" then
			orderSQL=sSQL & " ORDER BY ordInvoice"
		elseif editablefield="email" then
			orderSQL=sSQL & " ORDER BY ordEmail"
		else
			orderSQL=sSQL & " ORDER BY ordStatus"
		end if
		if sortorder="orsd" then orderSQL=orderSQL&" DESC"
	else
		orderSQL=sSQL & " ORDER BY ordID"
	end if
	hasdeleted=false
	sSQL="SELECT COUNT(*) AS NumDeleted FROM orders WHERE ordStatus=1"
	rs.open sSQL,cnn,0,1
		if clng(rs("NumDeleted")) > 0 then hasdeleted=TRUE
	rs.close
%>
<script src="popcalendar.js"></script>
<script>
/* <![CDATA[ */
try{languagetext('<%=adminlang%>');}catch(err){}
function delrec(id){
if(confirm("<%=jscheck(yyConDel)%>\n")){
	document.psearchform.id.value=id;
	document.psearchform.act.value="delete";
	document.psearchform.submit();
}
}
function authrec(id, currauth){
var aucode;
if(currauth=='')currauth='<%=yyManAut%>';
if((aucode=prompt("<%=jscheck(yyEntAuth)%>",currauth))!=null){
	document.psearchform.id.value=id;
	document.psearchform.act.value="authorize";
	document.psearchform.authcode.value=aucode;
	document.psearchform.submit();
}
}
function unpendrec(id){
if(confirm("<%=jscheck(yyWarni)%>\n\nThis will not make any changes at your payment processor!\n\nRemove pending status of this order?")){
	document.psearchform.id.value=id;
	document.psearchform.act.value="unpending";
	document.psearchform.submit();
}
}
function unmodwarn(id){
<%		yyModWar="The customer changed cart contents after creating this order.\nBefore authorizing this order check order totals carefully.\n\nPlease click ""OK"" to edit the order and check stock levels as stock has not yet been subtracted for this order."
		if useStockManagement then
			yyModWar=yyModWar&"Please click ""OK"" to edit the order and check stock levels as stock has not yet been subtracted for this order." %>
if(confirm("<%=jscheck(yyWarni)%>\n\n<%=jscheck(yyModWar)%>")){
	document.location='adminorders.asp?doedit=true&id='+id;
}
<%		else
			yyModWar=yyModWar&"Please click ""OK"" to clear the warning status and authorize this order." %>
if(confirm("<%=jscheck(yyWarni)%>\n\n<%=jscheck(yyModWar)%>")){
	document.psearchform.id.value=id;
	document.psearchform.act.value="unpending";
	document.psearchform.submit();
}
<%		end if %>
}
var ctrlset=false;
function setmodstate(evt){
	if(!evt)evt=window.event;
	if(evt.detail&&evt.detail>0) ctrlset=evt.ctrlKey;
}
function checkcontrol(tt,evt){
	//if(!evt)evt=window.event;
	//if(typeof(evt.ctrlKey)!='undefined')ctrlset=evt.ctrlKey;
	if(ctrlset){
		maxitems=document.psearchform.maxitems.value;
		for(index=0;index<maxitems;index++){
			isdisabled=eval('document.psearchform.ordStatus'+index+'.disabled');
			if(! isdisabled){
				if(eval('document.psearchform.ordStatus'+index+'.length') > tt.selectedIndex){
					eval('document.psearchform.ordStatus'+index+'.selectedIndex='+tt.selectedIndex);
					eval('document.psearchform.ordStatus'+index+'.options['+tt.selectedIndex+'].selected=true');
				}
			}
		}
	}
}
function checkprinter(tt,evt){
var thref=tt.href;
<% if netnav then %>
if(evt.ctrlKey || evt.altKey || document.psearchform.ctrlmod[document.psearchform.ctrlmod.selectedIndex].value=="1")thref+="&printer=true";
<% else %>
theevnt=window.event;
if(theevnt.ctrlKey || document.psearchform.ctrlmod[document.psearchform.ctrlmod.selectedIndex].value=="1")thref+="&printer=true";
<% end if %>
if(document.psearchform.ctrlmod[document.psearchform.ctrlmod.selectedIndex].value=="3")thref+="&invoice=true";
if(document.psearchform.ctrlmod[document.psearchform.ctrlmod.selectedIndex].value=="2")thref+="&doedit=true";
document.forms.psearchform.action=thref;
document.forms.psearchform.submit();
return(false);
}
function setdumpformat(fdump){
formatindex=fdump[fdump.selectedIndex].value;
if(formatindex==1)
	document.psearchform.act.value='dumporders';
else if(formatindex==2)
	document.psearchform.act.value='dumpdetails';
else if(formatindex==3)
	document.psearchform.act.value='quickbooks';
else if(formatindex==4)
	document.psearchform.act.value='ouresolutionsxmldump';
else if(formatindex==5){
	dodazzle();
	fdump.selectedIndex=0;
	return;
}
document.psearchform.action='dumporders.asp';
document.psearchform.submit();
fdump.selectedIndex=0;
}
function docheckall(){
	allcbs=document.getElementsByName('ids');
	mainidchecked=document.getElementById('xdocheckall').checked;
	for(i=0;i<allcbs.length;i++){
		allcbs[i].checked=mainidchecked;
	}
	return(true);
}
function checkchecked(printorinvoice){
	allcbs=document.getElementsByName('ids');
	var onechecked=false;
	for(i=0;i<allcbs.length;i++){
		if(allcbs[i].checked)onechecked=true;
	}
	if(onechecked){
		document.forms.psearchform.action='adminorders.asp?'+printorinvoice+'=true&id=multi';
		document.forms.psearchform.submit();
	}else{
		alert("<%=jscheck(yyNoSelO)%>");
	}
}
function changeselectfield(whichfield){
	var editablefield=document.getElementById(whichfield);
	var editfieldval=editablefield[editablefield.selectedIndex].value;
	if(editfieldval=='orsa'||editfieldval=='orsd'){
		changesortorder(editfieldval)
	}else{
		document.psearchform.reset();
		document.psearchform.action='adminorders.asp';
		document.psearchform.id.value=editfieldval;
		document.psearchform.act.value=whichfield;
		document.psearchform.submit();
	}
}
function setCookie(c_name,value,expiredays){
	var exdate=new Date();
	exdate.setDate(exdate.getDate()+expiredays);
	document.cookie=c_name+ "=" +escape(value)+((expiredays==null) ? "" : ";expires="+exdate.toGMTString());
}
function changesortorder(ord){
	setCookie('ordersort',ord,600);
	document.forms.psearchform.submit();
}
var dazzleorightml='<br /><br /><span style="background:#FFFFFF;padding:5px"> Please copy your Dazzle / WorldShip file contents below </span><br /><br /><div style="text-align:center"><textarea id="dazzletextarea" rows="18" cols="120" style="white-space:nowrap;overflow:scroll;" wrap="off"></textarea></div><div style="text-align:center"><input type="button" value="Submit" onclick="processdazzle()" /> <input type="button" value="Cancel" onclick="document.getElementById(\'dazzlediv\').style.display=\'none\'" /></div>';
function dodazzle(){
	document.getElementById('dazzleinner').innerHTML=dazzleorightml;
	document.getElementById('dazzlediv').style.display='';
}
function dazupdajaxcallback(){
	if(ajaxobj.readyState==4){
		var restxt=ajaxobj.responseText;
		if(restxt.search('SUCCESS')!=-1){
			var rowid=restxt.split('|')[1];
			document.getElementById('dazdet'+rowid).style.visibility='hidden';
			document.getElementById('dazdet'+rowid).style.display='none';
			document.getElementById('dazrow'+rowid).style.visibility='hidden';
			document.getElementById('dazrow'+rowid).style.display='none';
			document.getElementById('dazhr'+rowid).style.visibility='hidden';
			document.getElementById('dazhr'+rowid).style.display='none';
		}else
			alert('Error updating');
		if(dazisprocall)dazprocall();
	}
}
function dazzleupd(tordid,ttrnum,rowid,hasduplicate){
	var statussel=document.getElementById('dazordstatus');
	ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
	ajaxobj.onreadystatechange=dazupdajaxcallback;
	ajaxobj.open("GET", "ajaxservice.asp?action=dazzleupd&rowid="+rowid+"&ordid="+tordid+"&trackno="+encodeURIComponent(ttrnum)+"&emstatus="+(document.getElementById('dazemstatus').checked==true?'1':'0')+"&ordstatus="+statussel[statussel.selectedIndex].value+(hasduplicate?"&noemail=true":''), true);
	ajaxobj.send(null);
}
function dazprocall(){
	dazisprocall=true;
	if(dazall.length>0){
		var tind=dazall.pop();
		if(document.getElementById('dazid'+tind)){
			var hasduplicate=false;
			var tordid=document.getElementById('dazid'+tind).value;
			for(var idind=0;idind<dazall.length;idind++){
				if(document.getElementById('dazid'+dazall[idind])){
					var mordid=document.getElementById('dazid'+dazall[idind]).value;
					if(mordid==tordid) hasduplicate=true;
				}
			}
			dazzleupd(tordid,document.getElementById('daztr'+tind).innerHTML,tind,hasduplicate);
		}else dazprocall();
	}
}
var dazisprocall=false;
var statarr=[];
var dazall=[];
function dazajaxcallback(){
	var allstatus='<select id="dazordstatus" size="1" onchange="setCookie(\'dazordstatus\',this[this.selectedIndex].value,600);"><option value="">No Change</option><%
			statarr=""
			for index=0 to UBOUND(allstatus,2)
				if is_numeric(request.cookies("dazordstatus")) then wantstatus=int(request.cookies("dazordstatus")) else wantstatus=0
				if allstatus(0,index)>=3 then print "<option value=""" & allstatus(0,index) & """" & IIfVs(allstatus(0,index)=wantstatus," selected=""selected""") & ">" & jsescape(allstatus(1,index)) & "</option>"
				statarr=statarr&"statarr["&allstatus(0,index)&"]="""&jscheck(allstatus(1,index))&""";"
			next %></select>';
	<%=statarr%>
	if(ajaxobj.readyState==4){
		var restxt=ajaxobj.responseText;
		if(restxt=='ERRORFILEFORMAT'){
			alert('Error in file format. Only Dazzle And WorldShip CSV file formats are supported.');
		}else{
			document.getElementById('dazzleinner').innerHTML='<br />&nbsp;<br /><table style="margin:0 auto;" class="cobtbl" cellspacing="1" cellpadding="3" id="dazzletable"><tr><td class="cobhl" colspan="2">Change Status To:'+allstatus+' | Email Status Change: <input type="checkbox" id="dazemstatus" value="ON" onchange="setCookie(\'dazemstatus\',this.checked?1:0,600);" <% if request.cookies("dazemstatus")="1" then print "checked=""checked"" "%>/></td><td class="cobhl"><input type="button" value="Process All" onclick="dazprocall()" /></td></tr></table><br /><input type="button" value="Close Window" onclick="document.getElementById(\'dazzlediv\').style.display=\'none\'" />';
			var thetable=document.getElementById('dazzletable');
			var tarr=restxt.split('==DAZZLELINE==');
			for(var tind=1;tind<tarr.length;tind++){
				var newrow=thetable.insertRow(-1);
				newrow.id="dazdet"+tind;
				newrow.className='cobhl';
				var tlin=tarr[tind].split('==MATCHLINE==');
				var origdets=tlin[0].split('==ORIGADD==');
				newcell=newrow.insertCell(0);
				newcell.innerHTML=origdets[1];

				newcell=newrow.insertCell(1);
				newcell.innerHTML='<div id="daztr'+tind+'">'+origdets[0]+'</div>';

				newcell=newrow.insertCell(2);
				if(tlin.length<2){
					newcell.innerHTML='No Match';
				}else{
					dazall.push(tind);
					newcell.innerHTML=' - ';
					var newrow=thetable.insertRow(-1);
					newrow.id="dazrow"+tind;
					var ordstatus=0;
					var ordid=0;
					newrow.className='cobll';
					newcell=newrow.insertCell(0);
					newcell.className='cobll';
					var seltxt=tlin.length>2?'<select id="dazsel'+tind+'" size="1" onchange="document.getElementById(\'dazid'+tind+'\').value=this[this.selectedIndex].value.split(\'|\')[0];document.getElementById(\'dazstat'+tind+'\').innerHTML=statarr[this[this.selectedIndex].value.split(\'|\')[1]]">':'';
					for(var tind2=1;tind2<tlin.length;tind2++){
						linspl=tlin[tind2].split('==FULLADD==');
						ordid=linspl[0].split('|')[0];
						seltxt+=(tlin.length>2?'<option value="'+linspl[0]+'">':'')+ordid+' - '+linspl[1]+(tlin.length>2?'</option>':'');
					}
					ordstatus=tlin[1].split('==FULLADD==')[0].split('|')[1];
					newcell.innerHTML=seltxt+(tlin.length>2?'</select>':'')+'<input type="hidden" id="dazid'+tind+'" value="'+tlin[1].split('==FULLADD==')[0].split('|')[0]+'" />';

					newcell=newrow.insertCell(1);
					newcell.className='cobll';
					newcell.id='dazstat'+tind;
					newcell.innerHTML=statarr[ordstatus];

					newcell=newrow.insertCell(2);
					newcell.className='cobll';
					newcell.innerHTML='<input type="button" value="Update" onclick="dazisprocall=false;dazzleupd(document.getElementById(\'dazid'+tind+'\').value,\''+origdets[0]+'\','+tind+',false)" />';
				}
				var newrow=thetable.insertRow(-1);
				newrow.id="dazhr"+tind;
				newrow.className='cobll';
				newcell=newrow.insertCell(0);
				newcell.colSpan=3;
				newcell.innerHTML='<hr width="80%">';
			}
		}
	}
}
function processdazzle(){
	var dazzletext=encodeURIComponent(document.getElementById('dazzletextarea').value);
	if(dazzletext==''){
		alert("No input specified.");
	}else{
		ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
		ajaxobj.onreadystatechange=dazajaxcallback;
		ajaxobj.open("POST", "ajaxservice.asp?action=dazzle", true);
		ajaxobj.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
		ajaxobj.send('dazzletext='+dazzletext);
	}
}
function selectedordersact(tsel){
	var tact=tsel[tsel.selectedIndex].value;
	if(tact=="1")
		checkchecked('printer');
	else if(tact=="2")
		checkchecked('invoice');
	else if(tact=="3")
		dodazzle();
	tsel.selectedIndex=0;
}
function viewproduct(id){
    var form = document.createElement("form");
    var element1 = document.createElement("input");
    var element2 = document.createElement("input");
	var element3 = document.createElement("input");

    form.method = "POST";
    form.action = "adminprods.asp";
	form.target = "_ect";

    element1.value=id;
    element1.name="id";
    form.appendChild(element1);  

    element2.value='1';
    element2.name="posted";
    form.appendChild(element2);
	
	element3.value='modify';
    element3.name="act";
    form.appendChild(element3);

    document.body.appendChild(form);
    form.submit();
	document.body.removeChild(form);
}
function dopurgeorders(){
	if(confirm('Delete orders permanently?')){
		document.psearchform.action='adminorders.asp';
		document.psearchform.act.value='purge';
		return true;
	}else
		return false;
}
/* ]]> */
</script>
<div id="dazzlediv" style="display:none;position:absolute;width:100%;height:2000px;background-image:url(adminimages/opaquepixel.png);top:0px;left:0px;text-align:center;z-index:10000;"><br /><br /><br /><br /><br /><br /><div id="dazzleinner" style="margin:0 auto;background:#FFFFFF;width:800px;height:600px"></div></div>
<%	if NOT success then print "<div style=""text-align:center;color:#FF0000"">"&errmsg&"</div>"
	editablefield=request.cookies("editablefield")
	searchfield=request.cookies("searchfield")
	ordersearchfield=request.cookies("ordersearchfield")
	SESSION("fromdate")=fromdate
	SESSION("todate")=todate
	SESSION("notstatus")=getrequest("notstatus")
	SESSION("notsearchfield")=getrequest("notsearchfield")
	SESSION("searchtext")=origsearchtext
	SESSION("ordStatus")=getrequest("ordStatus")
	SESSION("ordstate")=getrequest("ordstate")
	SESSION("ordcountry")=getrequest("ordcountry")
	SESSION("payprovider")=payprovider %>
<h2><%=yyAdmOrd%></h2>
<form method="post" action="adminorders.asp<% if CurPage<>1 then print "?pg="&CurPage%>" name="psearchform">
			<input type="hidden" name="act" value="" />
			<input type="hidden" name="id" value="" />
			<input type="hidden" name="authcode" value="" />
			<input type="hidden" name="theeditablefield" value="<%=editablefield%>" />
			<input type="hidden" name="thesearchfield" value="<%=searchfield%>" />
            <div class="orderssearch">
			  <div class="orderssearchrow">
                <div style="text-align:right"><input type="button" onclick="popUpCalendar(this, document.forms.psearchform.fromdate, '<%=themask%>', 0)" value="<%=yyOrdFro%>" /></div>
				<div><div style="position:relative;display:inline"><input type="text" class="orddatesel" size="14" name="fromdate" value="<%=fromdate%>" style="vertical-align:middle" /> <input type="button" onclick="document.forms.psearchform.fromdate.value='<%=thedate%>'" value="<%=yyToday%>" /></div></div>
				<div style="text-align:right"><input type="button" onclick="popUpCalendar(this, document.forms.psearchform.todate, '<%=themask%>', -205)" value="<%=yyOrdTil%>" /></div>
				<div><div style="position:relative;display:inline"><input type="text" class="orddatesel" size="14" name="todate" value="<%=todate%>" style="vertical-align:middle" /> <input type="button" onclick="document.forms.psearchform.todate.value='<%=thedate%>'" value="<%=yyToday%>" /></div></div>
			  </div>
			  <div class="orderssearchrow">
				<div style="text-align:center"><input type="checkbox" name="notstatus" title="<%=yyNot%>" value="ON" style="vertical-align:middle" <% if getrequest("notstatus")="ON" then print "checked=""checked"" "%>/> <%=yyOrdSta%></div>
				<div style="text-align:center"><input type="checkbox" name="notsearchfield" title="<%=yyNot%>" value="ON" style="vertical-align:middle" <% if getrequest("notsearchfield")="ON" then print "checked=""checked"" "%>/>
					<select name="searchfield" id="searchfield" size="1" onchange="changeselectfield('searchfield')">
					<option value="state" <% if searchfield="state" then print "selected=""selected"""%>><%=yyState%></option>
					<option value="country" <% if searchfield="country" then print "selected=""selected"""%>><%=yyCountry%></option>
					<option value="payprovider" <% if searchfield="payprovider" OR searchfield="" then print "selected=""selected"""%>><%=yyPayMet%></option>
					</select></div>
				<div style="text-align:right"><select name="ordersearchfield" size="1">
					<option value="ordid" <% if ordersearchfield="ordid" then print "selected=""selected"""%>><%=yySearch&" "&xxOrdId%></option>
					<option value="email" <% if ordersearchfield="email" then print "selected=""selected"""%>><%=yySearch&" "&yyEmail%></option>
					<option value="authcode" <% if ordersearchfield="authcode" then print "selected=""selected"""%>><%=yySearch&" "&yyAutCod%></option>
					<option value="name" <% if ordersearchfield="name" then print "selected=""selected"""%>><%=yySearch&" "&yyName%></option>
					<option value="product" <% if ordersearchfield="product" then print "selected=""selected"""%>><%=yySearch&" "&yyPrName%>/ID</option>
					<option value="address" <% if ordersearchfield="address" then print "selected=""selected"""%>><%=yySearch&" "&yyAddress%></option>
					<option value="zip" <% if ordersearchfield="zip" then print "selected=""selected"""%>><%=yySearch&" "&yyZip%></option>
					<option value="phone" <% if ordersearchfield="phone" then print "selected=""selected"""%>><%=yySearch&" "&yyTelep%></option>
					<option value="invoice" <% if ordersearchfield="invoice" then print "selected=""selected"""%>><%=yySearch&" "&yyInvNum%></option>
					<option value="tracknum" <% if ordersearchfield="tracknum" then print "selected=""selected"""%>><%=yySearch&" "&yyTraNum%></option>
					<option value="affiliate" <% if ordersearchfield="affiliate" then print "selected=""selected"""%>><%=yySearch&" "&yyAffili%></option>
<%					if extraorderfield1<>"" then print "<option value=""extra1"" " & IIfVr(ordersearchfield="extra1", "selected=""selected""", "") & ">" & htmlspecials(left(strip_tags2(extraorderfield1), 16)) & "</option>"
					if extraorderfield2<>"" then print "<option value=""extra2"" " & IIfVr(ordersearchfield="extra2", "selected=""selected""", "") & ">" & htmlspecials(left(strip_tags2(extraorderfield2), 16)) & "</option>"
					if extracheckoutfield1<>"" then print "<option value=""checkout1"" " & IIfVr(ordersearchfield="checkout1", "selected=""selected""", "") & ">" & htmlspecials(left(strip_tags2(extracheckoutfield1), 16)) & "</option>"
					if extracheckoutfield2<>"" then print "<option value=""checkout2"" " & IIfVr(ordersearchfield="checkout2", "selected=""selected""", "") & ">" & htmlspecials(left(strip_tags2(extracheckoutfield2), 16)) & "</option>"
%>					</select></div>
				<div><input class="ordsearchtext" type="text" size="24" name="searchtext" value="<%=origsearchtext%>" /></div>
			  </div>
			  <div class="orderssearchrow">
				<div style="text-align:center;vertical-align:middle">
<%	if editablefield="abandoned" then %>
		<select name="abandonedstatus" size="1">
		<option value="">All Abandoned Carts</option>
		<option value="1"<% if getpost("abandonedstatus")="1" then print " selected=""selected"""%>>Abandoned Carts - No Emails Sent</option>
		<option value="2"<% if getpost("abandonedstatus")="2" then print " selected=""selected"""%>>Abandoned Carts - One Email Sent</option>
		<option value="3"<% if getpost("abandonedstatus")="3" then print " selected=""selected"""%>>Abandoned Carts - Two Emails Sent</option>
		<option value="4"<% if getpost("abandonedstatus")="4" then print " selected=""selected"""%>>Abandoned Carts - Three Emails Sent</option>
		<option value="recovered"<% if getpost("abandonedstatus")="recovered" then print " selected=""selected"""%>>Recovered Carts</option>
		</select>
<%	else %>
		<select name="ordStatus" size="5" multiple="multiple" class="ordstatus"><%
		ordstatus=getrequest("ordStatus")
		if ordstatus<>"" then selstatus=split(ordstatus, ",")
		for index=0 to UBOUND(allstatus,2)
			print "<option value=""" & allstatus(0,index) & """"
			if isarray(selstatus) then
				for ii=0 to UBOUND(selstatus)
					if Int(selstatus(ii))=Int(allstatus(0,index)) then print " selected=""selected"""
				next
			end if
			print ">" & allstatus(1,index) & "</option>"
		next %></select>
<%	end if %>
				</div>
				<div style="text-align:center;vertical-align:middle">
<%	if searchfield="state" then
		ordstate=getrequest("ordstate")
		print "<select name=""ordstate"" size=""5"" multiple=""multiple"">"
		sSQL="SELECT stateID,stateName,stateAbbrev FROM states WHERE stateCountryID=" & origCountryID & " AND stateEnabled=1 ORDER BY stateName"
		rs.open sSQL,cnn,0,1
		if ordstate<>"" then selstatus=split(ordstate, ",") else selstatus=""
		do while NOT rs.EOF
			print "<option value=""" & IIfVr(usestateabbrev=TRUE,rs("stateAbbrev"),rs("stateName")) & """"
			if isarray(selstatus) then
				for ii=0 to UBOUND(selstatus)
					if trim(selstatus(ii))=IIfVr(usestateabbrev=TRUE,rs("stateAbbrev"),rs("stateName")) then print " selected=""selected"""
				next
			end if
			print ">" & rs("stateName") & "</option>"
			rs.MoveNext
		loop
		rs.close
		print "</select>"
	elseif searchfield="country" then
		ordcountry=getrequest("ordcountry")
		print "<select name=""ordcountry"" size=""5"" multiple=""multiple"">"
		sSQL="SELECT countryID,countryName FROM countries WHERE countryEnabled=1 ORDER BY countryOrder DESC, countryName"
		rs.open sSQL,cnn,0,1
		if ordcountry<>"" then selstatus=split(ordcountry, ",") else selstatus=""
		do while NOT rs.EOF
			print "<option value=""" & rs("countryName") & """"
			if isarray(selstatus) then
				for ii=0 to UBOUND(selstatus)
					if trim(selstatus(ii))=rs("countryName") then print " selected=""selected"""
				next
			end if
			print ">" & rs("countryName") & "</option>"
			rs.MoveNext
		loop
		rs.close
		print "</select>"
	else %>
		<select name="payprovider" size="5" multiple="multiple"><%
		sSQL="SELECT payProvID,payProvName FROM payprovider WHERE payProvEnabled=1 ORDER BY payProvOrder"
		rs.open sSQL,cnn,0,1
		if payprovider<>"" then selstatus=split(payprovider, ",") else selstatus=""
		do while NOT rs.EOF
			print "<option value=""" & rs("payProvID") & """"
			if isarray(selstatus) then
				for ii=0 to UBOUND(selstatus)
					if int(selstatus(ii))=rs("payProvID") then print " selected=""selected"""
				next
			end if
			print ">" & rs("payProvName") & "</option>"
			rs.MoveNext
		loop
		rs.close %></select>
<%	end if %>
				</div>
				<div style="text-align:center;vertical-align:middle">
					<select class="orddumpsel" size="1" style="vertical-align:middle;max-width:164px" onchange="setdumpformat(this)">
					<option value=""><%=yySelect%></option>
					<option value="1"><%=yyDmpOrd%></option>
					<option value="2"><%=yyDmpDet%></option>
<%	if false then %>
					<option value="3">Dump orders to Quickbooks format</option>
<%	end if
	if ouresolutionsxml<>"" then print "<option value=""4"">OurESolutions XML format</option>" %>
					<option value="5">WorldShip</option>
<%	if origCountryID=1 then print "<option value=""5"">Dazzle</option>" %>
					</select>
				</div>
				<div style="text-align:center;vertical-align:middle">
					<input style="margin:2px" type="submit" value="<%=yySearch%>" onclick="document.forms.psearchform.action='adminorders.asp';" />
<%	if hasdeleted then %><input style="margin:2px" type="submit" value="<%=yyPurDel%>" onclick="return dopurgeorders()" /><% end if %>
					<input style="margin:2px" type="button" value="<%=yyNewOrd%>" onclick="document.forms.psearchform.action='adminorders.asp?id=new';document.forms.psearchform.submit();" />
				</div>
			  </div>
			</div>

			<div class="orderslist<% if editablefield="abandoned" then print " greenheader"%>">
			  <div class="ordersheadrow">
				<div class="acenter" style="width:3%"><input type="checkbox" id="xdocheckall" value="1" onclick="docheckall()" /></div>
                <div class="acenter"><a href="javascript:changesortorder('<%=IIfVr(sortorder="oida","oidd","oida")%>')"><%=yyOrdId%></a></div>
				<div class="aleft"><a href="javascript:changesortorder('<%=IIfVr(sortorder="orna","ornd","orna")%>')"><%=yyName%></a></div>
				<div class="acenter"><a href="javascript:changesortorder('<%=IIfVr(sortorder="orpa","orpd","orpa")%>')"><%=yyMethod%></a></div>
				<div class="acenter" style="width:3%"><% if editablefield="abandoned" then print "Emails Sent" else print "AVS"%></div>
				<div class="acenter" style="width:3%"><% if editablefield<>"abandoned" then print "CVV"%></div>
				<div class="acenter"><a href="javascript:changesortorder('<%=IIfVr(sortorder="oraa","orad","oraa")%>')"><% if editablefield="abandoned" then print "Amount" else print yyAutCod%></a></div>
				<div class="acenter"><a href="javascript:changesortorder('<%=IIfVr(sortorder="orda","ordd","orda")%>')"><%=yyDate%></a></div>
				<div class="acenter">
					<select class="ordstatus" name="editablefield" id="editablefield" size="1" onchange="changeselectfield('editablefield')">
					<option value="status"><%=yyStatus%></option>
					<option value="tracknum" <% if editablefield="tracknum" then print "selected=""selected"""%>><%=yyTraNum%></option>
					<option value="invoice" <% if editablefield="invoice" then print "selected=""selected"""%>><%=yyInvNum%></option>
					<option value="email" <% if editablefield="email" then print "selected=""selected"""%>><%=yyEmail%></option>
					<option value="abandoned" <% if editablefield="abandoned" then print "selected=""selected"""%>>Abandoned Carts</option>
					<option value="orderlines" <% if editablefield="orderlines" then print "selected=""selected"""%>>Order Lines</option>
					<option value="" disabled="disabled">---------------</option>
					<option value="<%=IIfVr(sortorder="orsa","orsd","orsa")%>">Sort On Column<%=IIfVr(sortorder="orsa"," DESC","")%></option>
					</select>
				</div>
			  </div>
<%	ordTot=0
	rowcounter=0
	rs.CursorLocation=3
	rs.CacheSize=maxordersperpage
	rs.open orderSQL,cnn
	if NOT rs.EOF then
		rs.MoveFirst
		rs.PageSize=maxordersperpage
		iNumOfPages=int((rs.RecordCount + (maxordersperpage-1)) / maxordersperpage)
		rs.AbsolutePage=CurPage
		do while NOT rs.EOF AND rowcounter < rs.PageSize
			if trim(rs("ordAuthNumber")&"")="" then
				startfont="<span style=""color:#FF0000"">"
				endfont="</span>"
			else
				startfont=""
				endfont=""
			end if
			abandonedcartcount=0
			if editablefield="abandoned" then
				sSQL="SELECT COUNT(*) AS tcount FROM abandonedcartemail WHERE aceOrderID=" & rs("ordID")
				rs2.open sSQL,cnn,0,1
				if NOT rs2.EOF then abandonedcartcount=rs2("tcount")
				rs2.close
				if getpost("abandonedstatus")="1" AND abandonedcartcount<>0 then abandonedcartcount=-1
				if getpost("abandonedstatus")="2" AND abandonedcartcount<>1 then abandonedcartcount=-1
				if getpost("abandonedstatus")="3" AND abandonedcartcount<>2 then abandonedcartcount=-1
				if getpost("abandonedstatus")="4" AND abandonedcartcount<>3 then abandonedcartcount=-1
			end if
			if abandonedcartcount<>-1 then
				if cint(rs("ordStatus"))>=3 OR editablefield="abandoned" then ordTot=ordTot+IIfVr(isnull(rs("ordTot")),0,rs("ordTot"))
				if bgcolor="cobll" then bgcolor="cobhl" else bgcolor="cobll"
				if rs("ordAuthStatus")="MODWARNOPEN" OR rs("ordShipType")="MODWARNOPEN" then bgcolor="cobwarn"
%>			  <div class="orderslistrow <%=bgcolor%>">
				<div class="acenter"><input type="checkbox" name="ids" value="<%=rs("ordID")%>"<% if editablefield="abandoned" AND (abandonedcartcount>=3 OR cint(rs("ordStatus"))>2) then print " disabled=""disabled"""%> /></div>
				<div class="acenter"><a onclick="return(checkprinter(this,event));" href="adminorders.asp?id=<%=rs("ordID")%>"><%=""&startfont&rs("ordID")&endfont%></a></div>
				<div><a onclick="return(checkprinter(this,event));" href="adminorders.asp?id=<%=rs("ordID")%>"><%=startfont&htmlspecialsucode(trim(rs("ordName")&" "&rs("ordLastName")))&endfont%></a></div>
				<div class="acenter"><%=startfont&strip_tags2(rs("payProvName")&"")&IIfVr(rs("payProvName")="PayPal" AND trim(rs("ordTransID")&"")<>""," CC","")&endfont%></div>
				<div class="acenter"><%
				if editablefield="abandoned" then
					print abandonedcartcount
				else
					print IIfVr(trim(rs("ordAVS")&"")<>"",strip_tags2(rs("ordAVS")&""),"&nbsp;")
				end if %></div>
				<div class="acenter"><% if trim(rs("ordCVV")&"")<>"" AND editablefield<>"abandoned" then print strip_tags2(rs("ordCVV")) else print "&nbsp;" %></div>
				<div class="acenter orderauthcode"><%
				if editablefield="abandoned" then
					print FormatEuroCurrency(IIfVr(isnull(rs("ordTot")),0,rs("ordTot")))
				elseif rs("ordAuthStatus")="MODWARNOPEN" OR rs("ordShipType")="MODWARNOPEN" then
					isauthorized=false
					print "<input type=""button"" value="""&yyMoWarn&""" onclick=""unmodwarn('"&rs("ordID")&"')"" /><br />"
				else
					if trim(rs("ordAuthStatus")&"")<>"" then print "<input type=""button"" value="""&rs("ordAuthStatus")&""" onclick=""unpendrec('"&rs("ordID")&"')"" /><br />"
					if trim(rs("ordAuthNumber")&"")="" then
						isauthorized=false
						print "<input type='button' name='auth' value='"&yyAuthor&": "&FormatEuroCurrency(IIfVr(isnull(rs("ordTot")),0,rs("ordTot")))&"' onclick=""authrec('"&rs("ordID")&"','')"" />"
					else
						isauthorized=TRUE
						print "<a href=""#"" title="""&FormatEuroCurrency(IIfVr(isnull(rs("ordTot")),0,rs("ordTot")))&""" onclick=""authrec('"&rs("ordID")&"','"&rs("ordAuthNumber")&"');return(false);"">" & startfont & rs("ordAuthNumber") & endfont & "</a>"
					end if
				end if %></div>
				<div class="acenter ordersdate"><%=startfont&Replace(rs("ordDate")&""," ","<br />",1,1)&endfont%></div>
				<div class="acenter"><input type="hidden" name="ordid<%=rowcounter%>" value="<%=rs("ordID")%>" />
<%				if editablefield="orderlines" then
					print "<div style=""margin:5px;padding:5px;border-style:border-box;border:1px solid lightgrey"">"
					sSQL="SELECT cartProdID,cartProdName FROM cart WHERE cartOrderID=" & rs("ordID")
					rs2.open sSQL,cnn,0,1
					do while NOT rs2.EOF
						print "<div><a href=""#"" onclick=""viewproduct('" & jsspecials(rs2("cartProdID")) & "')"">" & htmlspecials(rs2("cartProdName")) & "</a></div>"
						rs2.movenext
					loop
					rs2.close
					print "</div>"
				elseif editablefield="tracknum" then
					print "<input type=""text"" name=""tracknum"&rowcounter&""" size=""24"" value="""&rs("ordTrackNum")&""" tabindex=""" & (rowcounter+1) & """ />"
				elseif editablefield="invoice" then
					print "<input type=""text"" name=""invoice"&rowcounter&""" size=""24"" value="""&rs("ordInvoice")&""" tabindex=""" & (rowcounter+1) & """ />"
				elseif editablefield="email" then
					print "<input type=""text"" name=""email"&rowcounter&""" size=""34"" value="""&rs("ordEmail")&""" tabindex=""" & (rowcounter+1) & """ />"
				elseif editablefield="abandoned" then
					if cint(rs("ordStatus"))>2 then
						print "&nbsp;"
					elseif abandonedcartcount<3 then
						print "<input type=""submit"" value=""Send Abandoned Email"" onclick=""document.psearchform.action='adminorders.asp';document.psearchform.act.value='oneabandoned';document.psearchform.id.value='"&rs("ordID")&"'"" />"
					else
						print "All Sent"
					end if
				else %>
					<select class="ordstatus" name="ordStatus<%=rowcounter%>" size="1" onmousedown="setmodstate(event)" onchange="checkcontrol(this,event)" tabindex="<%=(rowcounter+1)%>"><%
					gotitem=false
					for index=0 to UBOUND(allstatus,2)
						if NOT isauthorized AND allstatus(0,index)>2 then exit for
						if NOT (cint(rs("ordStatus"))<>2 AND allstatus(0,index)=2) then
							print "<option value=""" & allstatus(0,index) & """"
							if cint(rs("ordStatus"))=allstatus(0,index) then
								print " selected=""selected"""
								gotitem=TRUE
							end if
							print ">" & allstatus(1,index) & "</option>"
						end if
					next
					if NOT gotitem then print "<option value="""&rs("ordStatus")&""" selected=""selected"">"&yyUndef&"</option>" %></select>
<%				end if %>
				</div>
			  </div>
<%				rowcounter=rowcounter+1
			end if
			rs.movenext
		loop %>
			  <div class="orderslistrow">
				<div style="vertical-align:top">
					<select class="ordactionsel" size="1" style="position:absolute;margin-top:8px" onchange="selectedordersact(this)">
					<option>Action Select</option>
					<option value="1"><%=yyPakSps%></option>
					<option value="2"><%=yyInvces%></option>
					<option value="3">WorldShip</option>
					<% if origCountryID=1 then print "<option value=""3"">Dazzle</option>" %>
					</select>
				</div>
				<div>&nbsp;</div>
				<div style="vertical-align:top"><select name="ctrlmod" size="1" style="position:absolute;margin-top:8px"><option value="0"><%=yyVieDet%></option><option value="1"><%=yyPPSlip%></option><option value="3"><%=yyPPInv%></option><option value="2" <% if getpost("ctrlmod")="2" then print "selected=""selected"""%>><%=yyEdOrd%></option></select></div>
				<div>&nbsp;</div>
				<div>&nbsp;</div>
				<div>&nbsp;</div>
				<div style="text-align:center"><%=FormatEuroCurrency(ordTot)%></div>
				<div>&nbsp;</div>
				<div style="text-align:center"><input type="hidden" name="maxitems" value="<%=rowcounter%>" /><%
		if editablefield<>"abandoned" then %>
					<input type="checkbox" name="emailstat" value="1"  title="<%=yyEStat%>" style="vertical-align:middle" <% if getpost("emailstat")="1" OR alwaysemailstatus=TRUE then print "checked=""checked"" "%>/>
<%		end if
		if editablefield="abandoned" then %>
				<input type="submit" value="Abandoned Emails to Selected" onclick="document.forms.psearchform.action='adminorders.asp';document.psearchform.act.value='abandoned';" />
<%		else %>
				<input type="submit" value="<%=yyUpdate%>" onclick="document.forms.psearchform.action='adminorders.asp<% if CurPage<>1 then print "?pg="&CurPage%>';document.psearchform.act.value='status';" /> <input type="reset" value="<%=yyReset%>" />
<%		end if %>
				</div>
			  </div>
			</div>
<%	else %>
			</div>
			  <div style="padding:40px 0px;text-align:center"><%=yyNoMat1%></div>
<%		if hasdeleted then %>
			  <div style="padding:20px 0px;text-align:center"><input type="submit" value="<%=yyPurDel%>" onclick="document.psearchform.action='adminorders.asp';document.psearchform.act.value='purge';" /></div>
<%		end if
	end if
	rs.close
	if iNumOfPages>1 then
		pblink="<a class=""ectlink"" href=""adminorders.asp?"
		for each objQS in request.querystring
			if objQS<>"pg" AND getget(objQS)<>"" AND (objQS="searchtext" OR objQS="ordersearchfield" OR objQS="notstatus" OR objQS="notsearchfield" OR objQS="fromdate" OR objQS="todate") then pblink=pblink & urlencode(objQS) & "=" & urlencode(getget(objQS)) & "&amp;"
		next
		for each objQS in request.form
			if objQS<>"pg" AND getpost(objQS)<>"" AND (objQS="searchtext" OR objQS="ordersearchfield" OR objQS="notstatus" OR objQS="notsearchfield" OR objQS="fromdate" OR objQS="todate") then pblink=pblink & urlencode(objQS) & "=" & urlencode(getpost(objQS)) & "&amp;"
		next
		if getrequest("ordStatus")<>"" then pblink=pblink&"ordStatus="&getrequest("ordStatus")&"&amp;"
		if getrequest("payprovider")<>"" then pblink=pblink&"payprovider="&getrequest("payprovider")&"&amp;"
		if getrequest("ordstate")<>"" then pblink=pblink&"ordstate="&getrequest("ordstate")&"&amp;"
		if getrequest("ordcountry")<>"" then pblink=pblink&"ordcountry="&getrequest("ordcountry")&"&amp;"
		pblink=pblink&"pg="
		print "<div class=""orderpagebar"">" & writepagebar(CurPage,iNumOfPages,yyPrev,yyNext,pblink,FALSE) & "</div>"
	end if %>
			<div class="ordernavigation">
				<a href="adminorders.asp?fromdate=<%=DateAdd("m",-1,thefromdate)%>&amp;todate=<%=DateAdd("m",-1,thetodate)%>">- <%=yyMonth%></a> | 
				<a href="adminorders.asp?fromdate=<%=DateValue(thefromdate)-7%>&amp;todate=<%=DateValue(thetodate)-7%>">- <%=yyWeek%></a> | 
				<a href="adminorders.asp?fromdate=<%=DateValue(thefromdate)-1%>&amp;todate=<%=DateValue(thetodate)-1%>">- <%=yyDay%></a> | 
				<a href="adminorders.asp"><%=yyToday%></a> | 
				<a href="adminorders.asp?fromdate=<%=DateValue(thefromdate)+1%>&amp;todate=<%=DateValue(thetodate)+1%>"><%=yyDay%> +</a> | 
				<a href="adminorders.asp?fromdate=<%=DateValue(thefromdate)+7%>&amp;todate=<%=DateValue(thetodate)+7%>"><%=yyWeek%> +</a> | 
				<a href="adminorders.asp?fromdate=<%=DateAdd("m",1,thefromdate)%>&amp;todate=<%=DateAdd("m",1,thetodate)%>"><%=yyMonth%> +</a>
			</div>
		  </form>
<%	if getpost("act")="abandoned" OR getpost("act")="oneabandoned" then print "<script>document.forms.psearchform.submit();</script>"
end if
cnn.Close
set rs=nothing
set rs2=nothing
set rs3=nothing
set cnn=nothing
%>
