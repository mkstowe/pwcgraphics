<%
Response.Buffer=True
Response.Expires=60
Response.Expiresabsolute=Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl="no-cache"
response.Charset="ISO-8859-1"
'=========================================
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
%>
<!--#include file="db_conn_open.asp"-->
<!--#include file="inc/languagefile.asp"-->
<!--#include file="includes.asp"-->
<%
isadmincalc=(getget("action")="admincalc") AND (SESSION("loggedon")<>"")
isaddtocart=(getget("action")="addtocart")
if getget("action")="admincalc" then %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Admin Shipping Calculator</title>
<link rel="stylesheet" type="text/css" href="adminstyle.css"/>
<style type="text/css">
td {font-size:12px;}
div.shiprateline{
width:100%;
float:left;
padding:1px;
}
div.shiptableline{
width:100%;
float:left;
}
div.shiprateradio{
width:10%;
float:left;
}
div.shipratemethod{
width:65%;
float:left;
}
div.shiptablelogo{
height: 10em;
position: relative;
width:80px;
height:60px;
float:left;
}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=<%=adminencoding%>"/>
<script>/* <![CDATA[ */
function updateshiprate(objitem,usrindex){
	window.opener.document.getElementById('ordShipping').value=document.getElementById('shipcost'+usrindex).value;
	window.opener.document.getElementById('shipmethod').value=document.getElementById('shipmethod'+usrindex).value;
	window.opener.dorecalc();
	window.close();
}
function changeshipcarrier(){
	var tobj=document.getElementById('shipcarrselect');
	var tshiptype=tobj[tobj.selectedIndex].value;
	var tquery=window.location.search.substring(1).split('&shiptype=')[0];
<%	querystr=""
	for each objItem in request.querystring
		if objItem<>"shiptype" AND objItem<>"destzip" AND objItem<>"sc" AND objItem<>"sta" AND objItem<>"cl" then querystr=querystr & objItem & "=" & getget(objItem) & "&"
	next %>
	var tcntry=document.getElementById('country');
	var comloc=document.getElementById('commercialloc')?document.getElementById('commercialloc')[document.getElementById('commercialloc').selectedIndex].value:'N';
	var shipstate=document.getElementById('sta') ? document.getElementById('sta')[document.getElementById('sta').selectedIndex].value : '';
	window.location.href='shipservice.asp?<%=querystr%>shiptype='+tshiptype+'&destzip='+document.getElementById('destzip').value+'&sc='+tcntry[tcntry.selectedIndex].value+'&sta='+shipstate+'&cl='+comloc;
}
/* ]]> */</script>
</head>
<body><%
else
	if lcase(adminencoding)<>"utf-8" then response.codepage=65001
	response.charset="utf-8"
end if
%>
<!--#include file="inc/incfunctions.asp"-->
<%	cartisincluded=TRUE
	if isaddtocart then
		checkoutmode="add"
		theid=trim(request.form("id"))
	end if %>
<!--#include file="inc/inccart.asp"-->
<%
if isaddtocart then response.end
adminIntShipping=0 ' So shipping doesn't get changed
handlingeligableitem=FALSE
standalonetestmode=TRUE
debginfo=""
thesessionid=getget("sessionid")
destZip=getget("destzip")
shipCountryID=getget("sc")
if NOT is_numeric(shipCountryID) then shipCountryID=0
shipstateid=getget("sta")
if is_numeric(getget("shiptype")) then shipType=int(getget("shiptype"))
if isadmincalc then
	shippingoptionsasradios=TRUE
	commercialloc_=(getget("cl")="Y")
else
	commercialloc_=SESSION("commercialloc_")
end if
numshipmethods=0
freeshipamnt=0
wantinsurance_=SESSION("wantinsurance_")
saturdaydelivery_=SESSION("saturdaydelivery_")
signaturerelease_=SESSION("signaturerelease_")
willpickup_=SESSION("willpickup_")
rgcpncode=trim(SESSION("cpncode"))
sSQL="SELECT countryID,countryName,countryTax,countryCode,countryFreeShip,countryOrder FROM countries WHERE countryID="&shipCountryID
rs.open sSQL,cnn,0,1
if NOT rs.EOF then
	shipcountry=rs("countryName")
	countryTaxRate=rs("countryTax")
	shipCountryID=rs("countryID")
	shipCountryCode=rs("countryCode")
	freeshipavailtodestination=(rs("countryFreeShip")=1)
	shiphomecountry=(rs("countryID")=origCountryID) OR ((rs("countryID")=1 OR rs("countryID")=2) AND usandcasplitzones)
end if
rs.close
sSQL="SELECT shipInsurance"&IIfVr(shiphomecountry,"Dom","Int")&",insurance"&IIfVr(shiphomecountry,"Dom","Int")&"Min,insurance"&IIfVr(shiphomecountry,"Dom","Int")&"Percent,noCarrier"&IIfVr(shiphomecountry,"Dom","Int")&"Ins FROM admin WHERE adminID=1"
rs.open sSQL,cnn,0,1
if NOT rs.EOF then
	addshippinginsurance=rs("shipInsurance"&IIfVr(shiphomecountry,"Dom","Int"))
	shipinsurancemin=rs("insurance"&IIfVr(shiphomecountry,"Dom","Int")&"Min")
	shipinsurancepercent=rs("insurance"&IIfVr(shiphomecountry,"Dom","Int")&"Percent")
	nocarrierinsurancerates=rs("noCarrier"&IIfVr(shiphomecountry,"Dom","Int")&"Ins")<>0
end if
if addshippinginsurance=3 then forceinsuranceselection=TRUE : addshippinginsurance=2
rs.close
if shipstateid<>"" then
	sSQL="SELECT stateID,stateFreeShip,stateAbbrev,stateName FROM states WHERE "
	if is_numeric(shipstateid) then
		sSQL=sSQL & "stateID=" & shipstateid
	else
		sSQL=sSQL & "stateName='" & escape_string(shipstateid) & "' OR stateAbbrev='" & escape_string(shipstateid) & "'"
	end if
	rs.open sSQL,cnn,0,1
	if rs.EOF then
		shipstateid=""
	else
		if shiphomecountry then freeshipavailtodestination=(freeshipavailtodestination AND (rs("stateFreeShip")=1))
		shipStateAbbrev=rs("stateAbbrev")
		shipstate=rs("stateName")
		shipstateid=rs("stateID")
	end if
	rs.close
else
	shipstateid=""
end if
call getadminshippingparams()
initshippingmethods()
totalgoods=0
alldata=""
success=TRUE
if is_numeric(getget("numshiprate")) then numshiprate=int(getget("numshiprate")) else numshiprate=0
numshiprateingroup=0
if isadmincalc then
	saveLCID=1033
	print "<table class=""cobtbl"" width=""100%"" border=""0"" cellspacing=""2"" cellpadding=""2"">"	
	if splitUSZones OR upsnegdrates then
		alloptions=""
		sSQL="SELECT stateID,stateName FROM states WHERE stateEnabled=1 AND stateCountryID=" & origCountryID & " ORDER BY stateName"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			alloptions=alloptions & "<option value="""" style=""font-weight:bold"">"
			if origCountryID=1 then
				alloptions=alloptions & xxPSelUS
			elseif origCountryID=2 then
				alloptions=alloptions & xxPSelCA
			else
				alloptions=alloptions & xxPlsSel
			end if
			alloptions=alloptions & "</option>"
			do while NOT rs.EOF
				alloptions=alloptions & "<option value=""" & rs("stateID") & """"
				if shipstateid=rs("stateID") then alloptions=alloptions & " selected=""selected"""
				alloptions=alloptions & ">" & rs("stateName") & "</option>" & vbLf
				rs.movenext
			loop
		end if
		rs.close
		if upsnegdrates AND origCountryID<>1 then
			sSQL="SELECT stateID,stateName FROM states WHERE stateEnabled=1 AND stateCountryID=1 ORDER BY stateName"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				alloptions=alloptions & "<option value="""" style=""font-weight:bold"">" & xxPSelUS & "</option>"
				do while NOT rs.EOF
					alloptions=alloptions & "<option value=""" & rs("stateID") & """"
					if shipstateid=rs("stateID") then alloptions=alloptions & " selected=""selected"""
					alloptions=alloptions & ">" & rs("stateName") & "</option>" & vbLf
					rs.movenext
				loop
			end if
			rs.close
		end if
		if upsnegdrates AND origCountryID<>2 then
			sSQL="SELECT stateID,stateName FROM states WHERE stateEnabled=1 AND stateCountryID=2 ORDER BY stateName"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				alloptions=alloptions & "<option value="""" style=""font-weight:bold"">" & xxPSelCA & "</option>"
				do while NOT rs.EOF
					alloptions=alloptions & "<option value=""" & rs("stateID") & """"
					if shipstateid=rs("stateID") then alloptions=alloptions & " selected=""selected"""
					alloptions=alloptions & ">" & rs("stateName") & "</option>" & vbLf
					rs.movenext
				loop
			end if
			rs.close
		end if
		if alloptions<>"" then print "<tr><td class=""cobhl"" align=""right"" width=""50%"">State:</td><td class=""cobll""><select name=""sta"" id=""sta"" size=""1"">" & alloptions & "</select></td></tr>"
	end if
	print "<tr><td class=""cobhl"" align=""right"">Zip:</td><td class=""cobll""><input type=""text"" size=""6"" id=""destzip"" value=""" & destZip & """ /></td></tr>"
	print "<tr><td class=""cobhl"" align=""right"">Country:</td><td class=""cobll""><select name=""country"" id=""country"" size=""1"">"
	sSQL="SELECT countryID,countryName,countryCode,"&getlangid("countryName",8)&" AS cnameshow FROM countries WHERE countryEnabled=1 ORDER BY countryOrder DESC,"&getlangid("countryName",8)
	rs.open sSQL,cnn,0,1
	gotcountry=FALSE
	do while NOT rs.EOF
		print "<option value="""&rs("countryID")&""""
		if shipCountryCode=rs("countryCode") AND NOT gotcountry then print " selected=""selected"""
		if shipcountry=rs("countryName") then gotcountry=TRUE
		cnameshow=rs("cnameshow")
		if cnameshow="United States of America" then cnameshow="USA"
		print ">" & cnameshow & "</option>" & vbLf
		rs.movenext
	loop
	rs.close
	print "</select></td></tr>"
	print "<tr><td class=""cobll"" align=""right""><select size=""1"" id=""shipcarrselect"">"
	sSQL="SELECT altrateid,altratename FROM alternaterates WHERE " & IIfVr(adminAltRates>0,"usealtmethod<>0 OR usealtmethodintl<>0","altrateid IN ("&shipType&","&adminIntShipping&")") & " ORDER BY altrateorder,altrateid"
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		print "<option value=""" & rs("altrateid") & """"
		if rs("altrateid")=shipType then print " selected=""selected"""
		print ">" & rs("altratename") & "</option>"
		rs.movenext
	loop
	rs.close
	print "</select></td>"
	print "<td class=""cobll"">"
	if shipType=3 OR shipType=4 OR shipType>=6 then
		print "<select id=""commercialloc"" size=""1""><option value=""N"">RES</option><option value=""Y"""&IIfVr(getget("cl")="Y"," selected=""selected""","")&">COM</option></select>&nbsp;"
	end if
	print "<input type=""button"" value=""Calculate"" onclick=""changeshipcarrier()"" /></td></tr></table>"
	shipmet="USPS"
	if shipType=1 then shipmet="Flat Rate"
	if shipType=2 then shipmet="Weight Based"
	if shipType=4 then shipmet="UPS"
	if shipType=5 then shipmet="Price Based"
	if shipType=6 then shipmet="Canada Post"
	if shipType=7 then shipmet="FedEx"
	if shipType=8 then shipmet="FedEx SmartPost"
	if shipType=9 then shipmet="DHL"
	print "&nbsp;<br /><table width=""100%"" cellspacing=""2"" cellpadding=""2"" border=""0"" class=""cobtbl""><tr><td align=""center"" class=""cobll"">"
	print "<table cellspacing=""2"" cellpadding=""2"" border=""0""><tr><td align=""right"">" & replace(replace(getshiplogo(shipType),"images","../images"),"&nbsp;","") & "</td><td style=""font-weight:bold"">" & shipmet & " " & xxShippg & "</td></tr></table>"
	productids=""
	redim alldata(12,0)
	rowcounter=0
	for each objItem in request.querystring
		if left(objItem,6)="prodid" then
			prodindex=right(objItem, len(objItem)-6)
			if is_numeric(prodindex) AND is_numeric(getget("quant" & prodindex)) then
				sSQL="SELECT 0 AS cartID,pID,pName,pPrice," & getget("quant" & prodindex) & " AS cartQuantity,pWeight,pShipping,pShipping2,pExemptions,pSection,0 AS topSection,pDims,pTax,'' AS pDescription FROM products WHERE pID='" & escape_string(getget(objItem)) & "'"
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					if UBOUND(alldata,2)<rowcounter then redim preserve alldata(12,rowcounter)
					tmparray=rs.getrows
					for ssind=0 to UBOUND(alldata)
						alldata(ssind,rowcounter)=tmparray(ssind,0)
					next
					optpricediff=0
					optweightdiff=0
					for each optItem in request.querystring
						if left(optItem,len("optn"&prodindex&"_"))="optn"&prodindex&"_" AND is_numeric(getget(optItem)) then
							sSQL="SELECT optID,optPriceDiff,optWeightDiff,optType,optFlags,optRegExp FROM options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optID="&getget(optItem)
							rs2.Open sSQL,cnn,0,1
							if NOT rs2.EOF then
								if abs(rs2("optType"))<> 3 then
									if (rs2("optFlags") AND 1)=0 then optpricediff=optpricediff + IIfVr(trim(rs2("optRegExp")&"")<>"", 0, rs2("optPriceDiff")) else optpricediff=optpricediff + vsround((rs2("optPriceDiff")*alldata(3,rowcounter))/100.0, 2)
									if (rs2("optFlags") AND 2)=0 then optweightdiff=optweightdiff + rs2("optWeightDiff") else optweightdiff=optweightdiff + multShipWeight(alldata(5,rowcounter),rs2("optWeightDiff"))
								end if
							end if
							rs2.Close
						end if
					next
					alldata(3,rowcounter)=alldata(3,rowcounter)+optpricediff
					alldata(5,rowcounter)=alldata(5,rowcounter)+optweightdiff
					rowcounter=rowcounter+1
				end if
				rs.close
			end if
		end if
	next
else
	if mysqlserver=true then
		sSQL="SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity,pWeight,pShipping,pShipping2,pExemptions,pSection,topSection,pDims,pTax,"&getlangid("pDescription",2)&" FROM cart LEFT JOIN products ON cart.cartProdID=products.pID LEFT OUTER JOIN sections ON products.pSection=sections.sectionID WHERE cartCompleted=0 AND " & getsessionsql()
	else
		sSQL="SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity,pWeight,pShipping,pShipping2,pExemptions,pSection,topSection,pDims,pTax,"&getlangid("pDescription",2)&" FROM cart LEFT JOIN (products LEFT OUTER JOIN sections ON products.pSection=sections.sectionID) ON cart.cartProdID=products.pID WHERE cartCompleted=0 AND " & getsessionsql()
	end if
	rs.open sSQL,cnn,0,1
	if NOT (rs.EOF OR rs.BOF) then alldata=rs.getrows
	rs.close
end if
if isarray(alldata) then
	for index=0 to UBOUND(alldata,2)
		if is_numeric(alldata(0,index)) then
			if isnull(alldata(5,index)) then alldata(5,index)=0
			if (alldata(1,index)=giftcertificateid OR alldata(1,index)=donationid) AND isnull(alldata(8,index)) then alldata(8,index)=15
			if alldata(1,index)=giftwrappingid AND isnull(alldata(8,index)) then alldata(8,index)=12
			sSQL="SELECT SUM(coPriceDiff) AS coPrDff FROM cartoptions WHERE coCartID="&alldata(0,index)
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if NOT IsNull(rs("coPrDff")) then alldata(3,index)=cdbl(alldata(3,index))+cdbl(rs("coPrDff"))
			end if
			rs.close
			sSQL="SELECT SUM(coWeightDiff) AS coWghtDff FROM cartoptions WHERE coCartID="&alldata(0,index)
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if NOT IsNull(rs("coWghtDff")) then alldata(5,index)=cdbl(alldata(5,index))+cdbl(rs("coWghtDff"))
			end if
			rs.close
			runTot=(alldata(3,index)*Int(alldata(4,index)))
			totalquantity=totalquantity + alldata(4,index)
			totalgoods=totalgoods+runTot
			thistopcat=0
			if trim(SESSION("clientID"))<>"" then alldata(8,index)=(alldata(8,index) OR (SESSION("clientActions") AND 7))
			if (shipType=2 OR shipType=3 OR shipType=4 OR shipType>=6) AND cdbl(alldata(5,index))<=0.0 then alldata(8,index)=(alldata(8,index) OR 4)
			if (alldata(8,index) AND 1)=1 then statetaxfree=statetaxfree + runTot
			if (alldata(8,index) AND 8)<>8 then handlingeligableitem=TRUE : handlingeligablegoods=handlingeligablegoods + runTot
			if perproducttaxrate=TRUE then
				if isnull(alldata(12,index)) then alldata(12,index)=countryTaxRate
				if (alldata(8,index) AND 2)<>2 then countryTax=countryTax + ((alldata(12,index) * runTot) / 100.0)
			else
				if (alldata(8,index) AND 2)=2 then countrytaxfree=countrytaxfree + runTot
			end if
			if (alldata(8,index) AND 4)=4 then shipfreegoods=shipfreegoods + runTot
			call addproducttoshipping(alldata, index)
		end if
	next
	if is_numeric(getget("orderid")) then call retrieveorderdetails(getget("orderid"), thesessionid)
	if isadmincalc AND shipType=9 AND zipisoptional(shipCountryID) then ordCity=shipstate
else
	errormsg="Error, couldn't find cart."
	success=FALSE
end if
call calculatediscounts(totalgoods, FALSE, rgcpncode)
if totaldiscounts > totalgoods then totaldiscounts=totalgoods
shipsellogo=getshiplogo(shipType)
if success AND calculateshipping() then
	if getget("ratetype")="estimator" then
		if nohandlinginestimator=TRUE then handling=0 : handlingchargepercent=0
		if (IsNumeric(shipinsuranceamt) OR (useuspsinsurancerates=TRUE AND shipType=3)) AND abs(addshippinginsurance)=1 then shipping=shipping + IIfVr(useuspsinsurancerates=TRUE AND shipType=3, getuspsinsurancerate(cdbl(totalgoods)), IIfVr(addshippinginsurance=1,((cdbl(totalgoods)*cdbl(shipinsuranceamt))/100.0),shipinsuranceamt))
		if taxShipping=1 AND showtaxinclusive<>0 then shipping=shipping + (cdbl(shipping)*cdbl(countryTaxRate))/100.0
		calculateshippingdiscounts(FALSE)
		if handlingeligableitem=FALSE then
			handling=0
		else
			if handlingchargepercent<>0 then
				temphandling=(((totalgoods + shipping + handling) - (totaldiscounts + freeshipamnt)) * handlingchargepercent / 100.0)
				if handlingeligablegoods < totalgoods AND totalgoods > 0 then temphandling=temphandling * (handlingeligablegoods / totalgoods)
				handling=handling + temphandling
			end if
			if taxHandling=1 AND showtaxinclusive<>0 then handling=handling + (cdbl(handling)*cdbl(countryTaxRate))/100.0
		end if
		if perproducttaxrate<>TRUE then countryTax=vsround((((totalgoods-countrytaxfree)+IIfVr(taxShipping=2,shipping-freeshipamnt,0)+IIfVr(taxHandling=2,handling,0))-totaldiscounts)*countryTaxRate/100.0, 2)
		countryTax=vsround(countryTax,2)
		handling=vsround(handling,2)
		if is_numeric(getget("best")) then currbest=cdbl(getget("best")) else currbest=100000000
		if ((shipping+handling)-freeshipamnt) < currbest then
			SESSION("xsshipping")=((shipping+handling)-freeshipamnt)
			SESSION("xscountrytax")=countryTax
			SESSION("altrates")=shipType
		end if
		session.LCID=1033
		print "&nbsp;"
		print "SHIPSELPARAM=" & ((shipping+handling)-freeshipamnt)
		print "SHIPSELPARAM=SUCCESS"
		print "SHIPSELPARAM=" & countryTax
		print "SHIPSELPARAM=" & shipType
	else
		if isadmincalc then orderid=0 : print "<table><tr><td>" else orderid=getget("orderid")
		if is_numeric(orderid) then
			freeshippingincludeshandling=FALSE
			insuranceandtaxaddedtoshipping()
			calculateshippingdiscounts(FALSE)
			calculatetaxandhandling()
			cpnmessage=Right(cpnmessage,Len(cpnmessage)-6)
			if shipType>=2 then
				if shippingoptionsasradios<>TRUE then print "<select size=""1"" onchange=""updateshiprate(this,(this.selectedIndex-1)+"&numshiprate&")""><option value="""">"&xxPlsSel&" ("&xxFromSE&": "&FormatEuroCurrency((shipping+IIfVr(combineshippinghandling,handling,0))-freeshipamnt)&")</option>"
				if shipType=>2 then
					for index=0 to UBOUND(intShipping,2)
						if intShipping(3,index)=TRUE then
							if freeshippingincludeshandling=TRUE then handling=0 : handlingchargepercent=0 else handling=orighandling : handlingchargepercent=orighandlingpercent
							if freeshippingapplied AND intShipping(4,index) <> 0 then shipping=0 else shipping=intShipping(2,index)
							calculatetaxandhandling()
							if isadmincalc then
								call writehiddenidvar("shipmethod"&numshiprate,intShipping(0,index))
								call writehiddenidvar("shipcost"&numshiprate,intShipping(2,index))
							end if
							call writeshippingoption(index,vsround(intShipping(2,index), 2), vsround(intShipping(7,index), 2), intShipping(4,index), intShipping(0,index), FALSE, intShipping(1,index))
						end if
					next
				end if
				if shippingoptionsasradios<>TRUE then print "</select>"
			end if
			if NOT isadmincalc then saveshippingoptions()
		end if
		if isadmincalc then
			print "</td></tr></table>"
		else
			print "SHIPSELPARAM="&replace(server.urlencode(shipsellogo),"+","%20")
			print "SHIPSELPARAM=REMOVEME"
			print "SHIPSELPARAM=REMOVEME"
			print "SHIPSELPARAM="&numshiprate
		end if
	end if
else
	success=FALSE
	print errormsg
	if getget("action")<>"admincalc" then
		print "SHIPSELPARAM="&replace(server.urlencode(shipsellogo),"+","%20")
		print "SHIPSELPARAM=ERROR"
	end if
end if
if getget("action")="admincalc" then
	if SESSION("loggedon")="" then
		print "<table width=""100%"" cellspacing=""2"" cellpadding=""2"" border=""0""><tr><td align=""center"">"
		print "&nbsp;<br />&nbsp;<br />&nbsp;<br /><strong>Session Timed Out.</strong><br /><br />Please close this window and click ""Calculate"" again.<br />"
	end if
	print "<br />&nbsp;</td></tr></table>" %>
</body>
</html>
<%
end if
%>