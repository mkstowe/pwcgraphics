<%
Response.Buffer=True
'=========================================
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protect under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
%>
<!--#include file="db_conn_open.asp"-->
<!--#include file="includes.asp"-->
<!--#include file="inc/incfunctions.asp"-->
<%
Server.ScriptTimeout=360
Dim sd, ed, rs, cnn, sSQL, sSQL2, hasdetails, sslok
DIM aFields(4)
function twodp(theval)
	twodp=FormatNumber(theval,2,-1,0,0)
end function
function xmlstrip(name2)
	name2=replace(name2&"","&","chr(11)")
	name2=replace(name2,chr(146),"chr(146)")
	name2=replace(name2,chr(150),"chr(150)")
	name2=replace(name2,"-","chr(45)")
	name2=replace(name2,"'","chr(39)chr(39)")
	name2=replace(name2,"€","chr(128)")
	name2=replace(name2,chr(163),"chr(163)")
	name2=replace(name2,chr(130),"chr(130)")
	name2=replace(name2,chr(138),"chr(138)")
	name2=replace(name2,chr(153),"")
	name2=replace(name2,chr(250),"u")
	name2=replace(name2,chr(225),"a")
	name2=replace(name2,chr(241),"n")
	name2=replace(name2,chr(252),"chr(129)")
	name2=replace(name2,chr(246),"chr(148)")	
	name2=replace(name2,chr(174),"")
	name2=replace(name2,"""","")
	name2=replace(name2,chr(147),"")
	name2=replace(name2,chr(148),"")
	name2=replace(name2,chr(169),"")
	name2=replace(name2,"å","a")
	tmp_str=""
	for i=1 to len(name2)
		ch_code=Asc(Mid(name2,i,1))
		if ch_code>130 then tmp_str=tmp_str & "chr("&ch_code&")" else tmp_str=tmp_str & Mid(name2,i,1)
	next
	xmlstrip=tmp_str
end function
function getsearchparams()
	whereSQL="" : namesql=""
	hasfromdate=FALSE
	hastodate=FALSE
	fromdate=trim(request("fromdate"))
	todate=trim(request("todate"))
	if fromdate<>"" then
		hasfromdate=TRUE
		if IsNumeric(fromdate) then
			thefromdate=(thedate-fromdate)
		else
			if isdate(fromdate) then
				thefromdate=datevalue(fromdate)
			else
				success=false
				errmsg=yyDatInv & " - " & fromdate
			end if
		end if
	else
		thefromdate=thedate
	end if
	if todate<>"" then
		hastodate=TRUE
		if IsNumeric(todate) then
			thetodate=(thedate-todate)
		else
			if isdate(todate) then
				thetodate=datevalue(todate)
			else
				success=false
				errmsg=yyDatInv & " - " & todate
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
	searchtext=escape_string(request.form("searchtext"))
	ordersearchfield=trim(request.form("ordersearchfield"))
	ordstatus=trim(request.form("ordStatus"))
	ordstate=trim(request.form("ordstate"))
	ordcountry=trim(request.form("ordcountry"))
	payprovider=trim(request.form("payprovider"))
	if ordersearchfield="product" AND searchtext<>"" AND NOT hasdetails AND getpost("act")<>"ouresolutionsxmldump" then whereSQL=whereSQL & " INNER JOIN cart ON orders.ordID=cart.cartOrderID "
	if ordersearchfield="ordid" AND searchtext<>"" AND IsNumeric(searchtext) then
		whereSQL=whereSQL & " WHERE ordID=" & searchtext & " "
	else
		if ordstatus<>"" then whereSQL=whereSQL & " WHERE " & IIfVr(request.form("notstatus")="ON","NOT ","") & "(ordStatus IN (" & ordstatus & "))" else whereSQL=whereSQL & " WHERE ordStatus<>1"
		if ordstate<>"" then whereSQL=whereSQL & " AND " & IIfVr(request.form("notsearchfield")="ON","NOT ","") & "(ordState IN ('" & replace(replace(escape_string(ordstate),", ",","),",","','") & "'))"
		if ordcountry<>"" then whereSQL=whereSQL & " AND " & IIfVr(request.form("notsearchfield")="ON","NOT ","") & "(ordCountry IN ('" & replace(replace(escape_string(ordcountry),", ",","),",","','") & "'))"
		if payprovider<>"" then whereSQL=whereSQL & " AND " & IIfVr(request.form("notsearchfield")="ON","NOT ","") & "(ordPayProvider IN ("&payprovider&")) "
		if hasfromdate then
			whereSQL=whereSQL & " AND ordDate BETWEEN " & vsusdate(thefromdate) & " AND " & vsusdate(IIfVr(hastodate, thetodate+1, thefromdate+1))
		elseif searchtext="" AND ordstatus="" AND ordstate="" AND ordcountry="" AND payprovider="" then
			whereSQL=whereSQL & " AND ordDate BETWEEN " & vsusdate(date()) & " AND " & vsusdate(date()+1)
		end if
		if searchtext<>"" then
			if ordersearchfield="ordid" OR ordersearchfield="name" then
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
			if ordersearchfield="ordid" then
				whereSQL=whereSQL & " AND (ordEmail LIKE '%" & searchtext & "%' OR "&namesql&")"
			elseif ordersearchfield="email" then
				whereSQL=whereSQL & " AND ordEmail LIKE '%"&searchtext&"%'"
			elseif ordersearchfield="authcode" then
				whereSQL=whereSQL & " AND (ordAuthNumber LIKE '%"&searchtext&"%' OR ordTransID LIKE '%"&searchtext&"%')"
			elseif ordersearchfield="name" then
				whereSQL=whereSQL & " AND " & namesql
			elseif ordersearchfield="product" AND getpost("act")<>"ouresolutionsxmldump" then
				whereSQL=whereSQL & " AND (cartProdID LIKE '%"&searchtext&"%' OR cartProdName LIKE '%"&searchtext&"%')"
			elseif ordersearchfield="address" then
				whereSQL=whereSQL & " AND (ordAddress LIKE '%"&searchtext&"%' OR ordAddress2 LIKE '%"&searchtext&"%' OR ordCity LIKE '%"&searchtext&"%' OR ordState LIKE '%"&searchtext&"%')"
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
	getsearchparams=whereSQL
end function
if storesessionvalue="" then storesessionvalue="virtualstore"
if NOT disallowlogin then
<!--#include file="inc/incloginfunctions.asp"-->
end if
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE then response.redirect "login.asp"
hasdetails=request.form("act")="dumpdetails"
Response.ContentType="unknown/exe"
if request.form("act")="stockinventory" OR getpost("act")="filteredstock" then
Response.AddHeader "Content-Disposition","attachment;filename=stockinventory.csv"
elseif request.form("act")="productimages" OR request.form("act")="filteredimages" then
Response.AddHeader "Content-Disposition","attachment;filename=productimages.csv"
elseif request.form("act")="dump2COinventory" then
Response.AddHeader "Content-Disposition","attachment;filename=inventory2co.csv"
elseif request.form("act")="fullinventory" OR request.form("act")="filteredinventory" then
Response.AddHeader "Content-Disposition","attachment;filename=inventory.csv"
elseif request.form("act")="catinventory" then
Response.AddHeader "Content-Disposition","attachment;filename=categoryinventory.csv"
elseif request.form("act")="dumpaffiliate" then
Response.AddHeader "Content-Disposition","attachment;filename=affilreport.csv"
elseif request.form("act")="quickbooks" then
elseif request.form("act")="ouresolutionsxmldump" then
Response.AddHeader "Content-Disposition","attachment;filename=oes_ordersdata.xml"
elseif request.form("act")="dumpemails" then
Response.AddHeader "Content-Disposition","attachment;filename=mailinglist.csv"
elseif request.form("act")="dumpevents" then
Response.AddHeader "Content-Disposition","attachment;filename=eventlog.csv"
elseif hasdetails then
Response.AddHeader "Content-Disposition","attachment;filename=orderdetails.csv"
else
Response.AddHeader "Content-Disposition","attachment;filename=dumporders.csv"
end if
sslok=true
if request.servervariables("HTTPS")<>"on" AND (Request.ServerVariables("SERVER_PORT_SECURE") <> "1") AND nochecksslserver<>true then sslok=false
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs2=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
dosaveLCID=session.LCID
session.LCID=1033
if Request.Form("sd")="" then
	sd=Date()
else
	sd=Request.Form("sd")
end if
if Request.Form("ed")="" then
	ed=Date()
else
	ed=Request.Form("ed")
end if
if request.form("act")="dumpaffiliate" then
	tdt=DateValue(sd)
	tdt2=DateValue(ed)+1
	Response.write "Affiliate report for " & sd & " to " & ed & vbCrLf
	Response.write """ID"",""Name"",""Address"",""City"",""State"",""Zip"",""Country"",""Email"",""Total""" & vbCrLf
	if mysqlserver=true then
		sSQL="SELECT affilID,affilName,affilAddress,affilCity,affilState,affilZip,affilCountry,affilEmail,SUM(ordTotal-ordDiscount) AS sumTot FROM affiliates LEFT JOIN orders ON affiliates.affilID=orders.ordAffiliate WHERE ordStatus>=3 AND ordDate BETWEEN " & vsusdate(tdt) & " AND " & vsusdate(tdt2) & " OR orders.ordAffiliate IS NULL GROUP BY affilID ORDER BY affilID"
	else
		sSQL="SELECT affilID,affilName,affilAddress,affilCity,affilState,affilZip,affilCountry,affilEmail,(SELECT Sum(ordTotal-ordDiscount) FROM orders WHERE ordStatus>=3 AND ordAffiliate=affilID AND ordDate BETWEEN " & vsusdate(tdt) & " AND " & vsusdate(tdt2) & ") FROM affiliates ORDER BY affilID"
	end if
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		response.write """"&replace(rs("affilID")&"","""","""""")&""","
		response.write """"&replace(rs("affilName")&"","""","""""")&""","
		response.write """"&replace(rs("affilAddress")&"","""","""""")&""","
		response.write """"&replace(rs("affilCity")&"","""","""""")&""","
		response.write """"&replace(rs("affilState")&"","""","""""")&""","
		response.write """"&replace(rs("affilZip")&"","""","""""")&""","
		response.write """"&replace(rs("affilCountry")&"","""","""""")&""","
		response.write """"&replace(rs("affilEmail")&"","""","""""")&""","
		response.write """"&rs(8)&""""&vbCrLf
		rs.MoveNext
	loop
	rs.close
elseif request.form("act")="stockinventory" OR getpost("act")="filteredstock" then
	stext=getrequest("stext")
	stype=getrequest("stype")
	sprice=getrequest("sprice")
	thecat=getrequest("scat")
	if thecat<>"" then thecat=int(thecat)
	sortorder=request.cookies("psort")
	catorman=request.cookies("pcatorman")
	if getpost("act")="filteredstock" then
		whereand=" WHERE "
		if thecat="" OR sortorder="nsf" then
			sSQL=" FROM products LEFT OUTER JOIN sections ON products.pSection=sections.sectionID"
		elseif mysqlserver=true then
			sSQL=" FROM multisections RIGHT JOIN products ON products.pId=multisections.pId LEFT OUTER JOIN sections ON products.pSection=sections.sectionID"
		else
			sSQL=" FROM " & IIfVs(thecat<>"" AND catorman="man","(") & "multisections RIGHT JOIN (products LEFT OUTER JOIN sections ON products.pSection=sections.sectionID) ON products.pId=multisections.pId" & IIfVs(thecat<>"" AND catorman="man",")")
		end if
		if thecat<>"" then
			if catorman="dis" then
				sSQL=sSQL & " INNER JOIN cpnassign ON products.pID=cpnassign.cpaAssignment" & whereand & "cpnassign.cpaCpnID=" & thecat : whereand=" AND "
			elseif catorman="man" then
				sSQL=sSQL & " INNER JOIN multisearchcriteria ON products.pID=multisearchcriteria.mSCpID" & whereand & "multisearchcriteria.mSCscID=" & thecat : whereand=" AND "
			else
				sectionids=getsectionids(thecat, TRUE)
				if sectionids<>"" then
					if sortorder="nsf" then
						sSQL=sSQL & whereand & "products.pSection IN (" & sectionids & ") "
					else
						sSQL=sSQL & whereand & "(products.pSection IN (" & sectionids & ") OR multisections.pSection IN (" & sectionids & "))"
					end if
					whereand=" AND "
				end if
			end if
		end if
		if sprice<>"" then
			if instr(sprice, "-") > 0 then
				pricearr=split(sprice, "-")
				if NOT is_numeric(pricearr(0)) then pricearr(0)=0
				if NOT is_numeric(pricearr(1)) then pricearr(1)=10000000
				sSQL=sSQL & whereand & "pPrice BETWEEN "&cdbl(replace(pricearr(0),"$",""))&" AND "&cdbl(replace(pricearr(1),"$",""))
				whereand=" AND "
			elseif is_numeric(sprice) then
				sSQL=sSQL & whereand & "pPrice="&cdbl(replace(sprice,"$",""))&" "
				whereand=" AND "
			end if
		end if
		if trim(request("stext"))<>"" then
			sText=escape_string(request("stext"))
			aText=Split(sText)
			if nosearchadmindescription then maxsearchindex=2 else maxsearchindex=3
			aFields(0)="products.pId"
			aFields(1)="pSKU"
			aFields(2)=getlangid("pName",1)
			aFields(3)=getlangid("pDescription",2)
			if request("stype")="exact" then
				sSQL=sSQL & whereand & "(products.pId LIKE '%"&sText&"%' OR pSKU LIKE '%"&sText&"%' OR "&getlangid("pName",1)&" LIKE '%"&sText&"%' OR "&getlangid("pDescription",2)&" LIKE '%"&sText&"%' OR "&getlangid("pLongDescription",2)&" LIKE '%"&sText&"%') "
				whereand=" AND "
			else
				sJoin="AND "
				if request("stype")="any" then sJoin="OR "
				sSQL=sSQL & whereand&"("
				whereand=" AND "
				for index=0 to maxsearchindex
					sSQL=sSQL & "("
					for rowcounter=0 to UBOUND(aText)
						sSQL=sSQL & aFields(index) & " LIKE '%"&aText(rowcounter)&"%' "
						if rowcounter<UBOUND(aText) then sSQL=sSQL & sJoin
					next
					sSQL=sSQL & ") "
					if index < maxsearchindex then sSQL=sSQL & "OR "
				next
				sSQL=sSQL & ") "
			end if
		end if
		if request("disp")="6" then sSQL=sSQL & whereand & "pBackOrder<>0" : whereand=" AND "
		if request("disp")="7" then sSQL=sSQL & whereand & "pBackOrder=0" : whereand=" AND "
		if request("disp")="8" then sSQL=sSQL & whereand & "pGiftWrap<>0" : whereand=" AND "
		if request("disp")="9" then sSQL=sSQL & whereand & "pGiftWrap=0" : whereand=" AND "
		if request("disp")="10" then sSQL=sSQL & whereand & "pRecommend<>0" : whereand=" AND "
		if request("disp")="11" then sSQL=sSQL & whereand & "pRecommend=0" : whereand=" AND "
		if request("disp")="12" then sSQL=sSQL & whereand & "pStaticPage<>0" : whereand=" AND "
		if request("disp")="13" then sSQL=sSQL & whereand & "pStaticPage=0" : whereand=" AND "
		if request("disp")="4" then sSQL=sSQL & whereand & "(rootSection IS NULL OR rootSection=0)" : whereand=" AND "
		if request("disp")="3" then sSQL=sSQL & whereand & "(pInStock<=0 AND pStockByOpts=0)" : whereand=" AND "
		if request("disp")="" OR request("disp")="5" then sSQL=sSQL & whereand & "pDisplay<>0" : whereand=" AND "
		if request("disp")="2" then sSQL=sSQL & whereand & "pDisplay=0" : whereand=" AND "
	else
		sSQL=" FROM products"
	end if
	sSQL2="SELECT " & IIfVs(request.form("act")="filteredstock","DISTINCT ") & "products.pID,pName,pPrice,pInStock,pStockByOpts"
	rs.open sSQL2&sSQL,cnn,0,1
	response.write "pID,pName,pPrice,pInStock,optID,OptionGroup,Option" & vbCrLf
	do while NOT rs.EOF
		if rs("pStockByOpts") <> 0 then
			rs2.Open "SELECT optID,optGrpName,optName,optStock FROM optiongroup INNER JOIN (options INNER JOIN prodoptions ON options.optGroup=prodoptions.poOptionGroup) ON optiongroup.optGrpID=options.optGroup WHERE prodoptions.poProdID='"&escape_string(rs("pID"))&"'",cnn,0,1
			do while NOT rs2.EOF
				response.write """"&replace(rs("pID")&"","""","""""")&""","
				response.write """"&replace(rs("pName")&"","""","""""")&""","
				response.write """"&rs("pPrice")&""","
				response.write rs2("optStock")&","
				response.write trim(rs2("optID"))&","
				response.write """"&replace(rs2("optGrpName")&"","""","""""")&""","
				response.write """"&replace(rs2("optName")&"","""","""""")&""""&vbCrLf
				rs2.MoveNext
			loop
			rs2.Close
		else
			response.write """"&replace(rs("pID")&"","""","""""")&""","
			response.write """"&replace(rs("pName")&"","""","""""")&""","
			response.write """"&rs("pPrice")&""","
			response.write rs("pInStock")&",,,"&vbCrLf
		end if
		rs.MoveNext
	loop
	rs.close
elseif request.form("act")="productimages" OR request.form("act")="filteredimages" then
	thecounter=0
	thecat=getrequest("scat")
	if thecat<>"" then thecat=int(thecat)
	catorman=request.cookies("pcatorman")
	if request.form("act")="filteredimages" then
		sSQL=" FROM productimages"
		if thecat<>"" then
			if catorman="dis" then
				sSQL=sSQL&" INNER JOIN (products INNER JOIN cpnassign ON products.pID=cpnassign.cpaAssignment) ON productimages.imageProduct=products.pID  WHERE cpnassign.cpaCpnID=" & thecat
			elseif catorman="man" then
				sSQL=sSQL&" INNER JOIN (products INNER JOIN multisearchcriteria ON products.pID=multisearchcriteria.mSCpID) ON productimages.imageProduct=products.pID WHERE multisearchcriteria.mSCscID=" & thecat
			else
				sectionids=getsectionids(thecat, TRUE)
				if sectionids<>"" then
					sSQL=sSQL&" INNER JOIN products ON productimages.imageProduct=products.pID WHERE products.pSection IN (" & sectionids & ") "
				end if
			end if
		end if
	else
		sSQL=" FROM productimages"
	end if
	sSQL2="SELECT imageProduct,imageSrc,imageType,imageNumber" & sSQL & " ORDER BY imageProduct,imageType,imageNumber"
	rs.open sSQL2,cnn,0,1
	response.write "imageProduct,imageSrc,imageType,imageNumber" & vbCrLf
	do while NOT rs.EOF
		response.write """"&replace(rs("imageProduct")&"","""","""""")&""","
		response.write """"&replace(rs("imageSrc")&"","""","""""")&""","
		response.write rs("imageType")&","&rs("imageNumber")&vbCrLf
		thecounter=thecounter+1
		if (thecounter MOD 10000)=0 then response.flush
		rs.MoveNext
	loop
	rs.close
elseif request.form("act")="fullinventory" OR request.form("act")="filteredinventory" then
	nomemocolumnseparate=TRUE
	stext=getrequest("stext")
	stype=getrequest("stype")
	sprice=getrequest("sprice")
	thecat=getrequest("scat")
	if thecat<>"" then thecat=int(thecat)
	sortorder=request.cookies("psort")
	catorman=request.cookies("pcatorman")
	columnlist="products.pID,pName"
	for index=2 to adminlanguages+1
		if (adminlangsettings AND 1)=1 then columnlist=columnlist & ",pName"&index
	next
	columnlist=columnlist & ",products.pSection,pPrice,pWholesalePrice,pListPrice,pShipping,pShipping2,pWeight,pDisplay,pSell,pBackOrder,pGiftWrap,pExemptions,pInStock,pDims,pTax,pDropship"
	if digidownloads=TRUE then columnlist=columnlist & ",pDownload"
	if instr(productpagelayout&detailpagelayout,"custom1")>0 then columnlist=columnlist & ",pCustom1"
	if instr(productpagelayout&detailpagelayout,"custom2")>0 then columnlist=columnlist & ",pCustom2"
	if instr(productpagelayout&detailpagelayout,"custom3")>0 then columnlist=columnlist & ",pCustom3"
	columnlist=columnlist & ",pStaticPage,pStockByOpts,pRecommend,pOrder,pSKU,pManufacturer,pSearchParams,pTitle,pMetaDesc,pStaticURL,pMinQuant,pDateAdded"
	memocolumnlist=",pDescription"
	for index=2 to adminlanguages+1
		if (adminlangsettings AND 2)=2 then memocolumnlist=memocolumnlist & ",pDescription"&index
	next
	memocolumnlist=memocolumnlist & ",pLongDescription"
	for index=2 to adminlanguages+1
		if (adminlangsettings AND 4)=4 then memocolumnlist=memocolumnlist & ",pLongDescription"&index
	next
	if nomemocolumnseparate then columnlist=columnlist&memocolumnlist : memocolumnlist=""
	if getpost("act")="filteredinventory" then
		whereand=" WHERE "
		if thecat="" OR sortorder="nsf" then
			sSQL=" FROM products LEFT OUTER JOIN sections ON products.pSection=sections.sectionID"
		elseif mysqlserver=true then
			sSQL=" FROM multisections RIGHT JOIN products ON products.pId=multisections.pId LEFT OUTER JOIN sections ON products.pSection=sections.sectionID"
		else
			sSQL=" FROM " & IIfVs(thecat<>"" AND catorman="man","(") & "multisections RIGHT JOIN (products LEFT OUTER JOIN sections ON products.pSection=sections.sectionID) ON products.pId=multisections.pId" & IIfVs(thecat<>"" AND catorman="man",")")
		end if
		if thecat<>"" then
			if catorman="dis" then
				sSQL=sSQL & " INNER JOIN cpnassign ON products.pID=cpnassign.cpaAssignment" & whereand & "cpnassign.cpaCpnID=" & thecat : whereand=" AND "
			elseif catorman="man" then
				sSQL=sSQL & " INNER JOIN multisearchcriteria ON products.pID=multisearchcriteria.mSCpID" & whereand & "multisearchcriteria.mSCscID=" & thecat : whereand=" AND "
			else
				sectionids=getsectionids(thecat, TRUE)
				if sectionids<>"" then
					if sortorder="nsf" then
						sSQL=sSQL & whereand & "products.pSection IN (" & sectionids & ") "
					else
						sSQL=sSQL & whereand & "(products.pSection IN (" & sectionids & ") OR multisections.pSection IN (" & sectionids & "))"
					end if
					whereand=" AND "
				end if
			end if
		end if
		if sprice<>"" then
			if instr(sprice, "-") > 0 then
				pricearr=split(sprice, "-")
				if NOT is_numeric(pricearr(0)) then pricearr(0)=0
				if NOT is_numeric(pricearr(1)) then pricearr(1)=10000000
				sSQL=sSQL & whereand & "pPrice BETWEEN "&cdbl(replace(pricearr(0),"$",""))&" AND "&cdbl(replace(pricearr(1),"$",""))
				whereand=" AND "
			elseif is_numeric(sprice) then
				sSQL=sSQL & whereand & "pPrice="&cdbl(replace(sprice,"$",""))&" "
				whereand=" AND "
			end if
		end if
		if trim(request("stext"))<>"" then
			sText=escape_string(request("stext"))
			aText=Split(sText)
			if nosearchadmindescription then maxsearchindex=2 else maxsearchindex=3
			aFields(0)="products.pId"
			aFields(1)="pSKU"
			aFields(2)=getlangid("pName",1)
			aFields(3)=getlangid("pDescription",2)
			if request("stype")="exact" then
				sSQL=sSQL & whereand & "(products.pId LIKE '%"&sText&"%' OR pSKU LIKE '%"&sText&"%' OR "&getlangid("pName",1)&" LIKE '%"&sText&"%' OR "&getlangid("pDescription",2)&" LIKE '%"&sText&"%' OR "&getlangid("pLongDescription",2)&" LIKE '%"&sText&"%') "
				whereand=" AND "
			else
				sJoin="AND "
				if request("stype")="any" then sJoin="OR "
				sSQL=sSQL & whereand&"("
				whereand=" AND "
				for index=0 to maxsearchindex
					sSQL=sSQL & "("
					for rowcounter=0 to UBOUND(aText)
						sSQL=sSQL & aFields(index) & " LIKE '%"&aText(rowcounter)&"%' "
						if rowcounter<UBOUND(aText) then sSQL=sSQL & sJoin
					next
					sSQL=sSQL & ") "
					if index < maxsearchindex then sSQL=sSQL & "OR "
				next
				sSQL=sSQL & ") "
			end if
		end if
		if request("disp")="6" then sSQL=sSQL & whereand & "pBackOrder<>0" : whereand=" AND "
		if request("disp")="7" then sSQL=sSQL & whereand & "pBackOrder=0" : whereand=" AND "
		if request("disp")="8" then sSQL=sSQL & whereand & "pGiftWrap<>0" : whereand=" AND "
		if request("disp")="9" then sSQL=sSQL & whereand & "pGiftWrap=0" : whereand=" AND "
		if request("disp")="10" then sSQL=sSQL & whereand & "pRecommend<>0" : whereand=" AND "
		if request("disp")="11" then sSQL=sSQL & whereand & "pRecommend=0" : whereand=" AND "
		if request("disp")="12" then sSQL=sSQL & whereand & "pStaticPage<>0" : whereand=" AND "
		if request("disp")="13" then sSQL=sSQL & whereand & "pStaticPage=0" : whereand=" AND "
		if request("disp")="4" then sSQL=sSQL & whereand & "(rootSection IS NULL OR rootSection=0)" : whereand=" AND "
		if request("disp")="3" then sSQL=sSQL & whereand & "(pInStock<=0 AND pStockByOpts=0)" : whereand=" AND "
		if request("disp")="" OR request("disp")="5" then sSQL=sSQL & whereand & "pDisplay<>0" : whereand=" AND "
		if request("disp")="2" then sSQL=sSQL & whereand & "pDisplay=0" : whereand=" AND "
	else
		sSQL=" FROM products"
	end if
	if sortorder="ida" then
		sSQL=sSQL & " ORDER BY products.pid"
	elseif sortorder="idd" then
		sSQL=sSQL & " ORDER BY products.pid DESC"
	elseif sortorder="" then
		sSQL=sSQL & " ORDER BY pName"
	elseif sortorder="na2" then
		sSQL=sSQL & " ORDER BY pName2"
	elseif sortorder="na3" then
		sSQL=sSQL & " ORDER BY pName3"
	elseif sortorder="nad" then
		sSQL=sSQL & " ORDER BY pName DESC"
	elseif sortorder="pra" then
		sSQL=sSQL & " ORDER BY pPrice"
	elseif sortorder="prd" then
		sSQL=sSQL & " ORDER BY pPrice DESC"
	elseif sortorder="daa" then
		sSQL=sSQL & " ORDER BY pDateAdded"
	elseif sortorder="dad" then
		sSQL=sSQL & " ORDER BY pDateAdded DESC"
	elseif sortorder="poa" then
		sSQL=sSQL & " ORDER BY pOrder"
	elseif sortorder="pod" then
		sSQL=sSQL & " ORDER BY pOrder DESC"
	elseif sortorder="sta" then
		sSQL=sSQL & " ORDER BY products.pInStock"
	elseif sortorder="std" then
		sSQL=sSQL & " ORDER BY products.pInStock DESC"
	end if
	sSQL="SELECT " & IIfVs(request.form("act")="filteredinventory","DISTINCT ") & columnlist & sSQL
	rs.open sSQL,cnn,0,1
	fieldlistarr=split(replace(columnlist & memocolumnlist,"products.",""),",")
	fieldlistcnt=UBOUND(fieldlistarr)
	for index=0 to fieldlistcnt
		response.write """"&fieldlistarr(index)&""""
		if index < fieldlistcnt then response.write ","
	next
	response.write vbCrLf
	thecounter=0
	do while NOT rs.EOF
		if NOT nomemocolumnseparate then rs2.open replace(memocolumnlist,",","SELECT ",1,1) & " FROM products WHERE pID='" & escape_string(rs("pID")) & "'",cnn,0,1
		for index=0 to fieldlistcnt
			if NOT nomemocolumnseparate AND instr(fieldlistarr(index),"Description")>0 then
				response.write """"&replace(rs2(fieldlistarr(index))&"","""","""""")&""""
			else
				fieldtype=rs.fields(fieldlistarr(index)).type
				if fieldtype=11 then
					response.write IIfVr(rs(fieldlistarr(index)),"1","0")
				elseif (fieldtype >= 2 AND fieldtype<=5) OR (fieldtype >= 14 AND fieldtype <= 21) then
					response.write rs(fieldlistarr(index))
				else
					response.write """"&replace(rs(fieldlistarr(index))&"","""","""""")&""""
				end if
			end if
			if index < fieldlistcnt then response.write ","
		next
		if NOT nomemocolumnseparate then rs2.close
		response.write vbCrLf
		thecounter=thecounter+1
		if (thecounter MOD 10)=0 then response.flush
		rs.MoveNext
	loop
	rs.close
elseif request.form("act")="catinventory" then
	fieldlist="sectionID,sectionName"
	for index=2 to adminlanguages+1
		if (adminlangsettings AND 256)=256 then fieldlist=fieldlist & ",sectionName"&index
	next
	fieldlist=fieldlist & ",sectionWorkingName,sectionImage,topSection,sectionOrder,rootSection,sectionDisabled,sectionURL"
	for index=2 to adminlanguages+1
		if (adminlangsettings AND 2048)=2048 then fieldlist=fieldlist & ",sectionURL"&index
	next
	fieldlist=fieldlist & ",sectionDescription"
	for index=2 to adminlanguages+1
		if (adminlangsettings AND 512)=512 then sSQL2=sSQL2 & ",sectionDescription"&index
	next
	sSQL="SELECT " & fieldlist & " FROM sections"
	rs.open sSQL,cnn,0,1
	fieldlistarr=split(fieldlist, ",")
	fieldlistcnt=UBOUND(fieldlistarr)
	for index=0 to fieldlistcnt
		response.write """"&fieldlistarr(index)&""""
		if index < fieldlistcnt then response.write ","
	next
	response.write vbCrLf
	thecounter=0
	do while NOT rs.EOF
		for index=0 to fieldlistcnt
			fieldtype=rs.fields(fieldlistarr(index)).type
			if fieldtype=11 then
				response.write IIfVr(rs(fieldlistarr(index)),"1","0")
			elseif (fieldtype >= 2 AND fieldtype<=5) OR (fieldtype >= 14 AND fieldtype <= 21) then
				response.write rs(fieldlistarr(index))
			else
				response.write """"&replace(rs(fieldlistarr(index))&"","""","""""")&""""
			end if
			if index < fieldlistcnt then response.write ","
		next
		response.write vbCrLf
		thecounter=thecounter+1
		if (thecounter MOD 10)=0 then response.flush
		rs.MoveNext
	loop
	rs.close
elseif request.form("act")="dump2COinventory" then
	sSQL2="SELECT payProvData1 FROM payprovider WHERE payProvID=2"
	rs.open sSQL2,cnn,0,1
	response.write rs("payProvData1") & vbCrLf
	rs.close
	sSQL2="SELECT pID,pName,pPrice,"&IIfVr(digidownloads=TRUE,"pDownload,","")&"pDescription FROM products"
	rs.open sSQL2,cnn,0,1
	do while NOT rs.EOF
		response.write replace(rs("pID"),",","&#44;")&","
		response.write replace(replace(strip_tags2(rs("pName")),",","&#44;"),vbNewline," ")&","
		response.write ","
		response.write rs("pPrice")&","
		response.write ",,"
		if digidownloads=TRUE then
			response.write IIfVr(trim(rs("pDownload")&"")<>"", "N", "Y")&","
		else
			response.write "Y,"
		end if
		response.write replace(replace(strip_tags2(rs("pDescription")&""),",","&#44;"),vbNewline,"\n")&vbCrLf
		rs.MoveNext
	loop
	rs.close
elseif request.form("act")="quickbooks" then
	sSQL2="SELECT ordID,ordName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordExtra1,ordExtra2,ordShipExtra1,ordShipExtra2,ordCheckoutExtra1,ordCheckoutExtra2,ordShipName,ordShipLastName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,ordShipPhone,payProvName,ordAuthNumber,ordTotal,ordDate,ordStateTax,ordCountryTax,ordHSTTax,ordShipping,ordHandling,ordShipType,ordDiscount,ordAddInfo FROM orders INNER JOIN payprovider ON payprovider.payProvID=orders.ordPayProvider"
	sSQL2=sSQL2 & getsearchparams()
	response.write "!TRNS	DATE	ACCNT	NAME	CLASS	AMOUNT	MEMO" & vbCrLf
	response.write "!SPL	DATE	ACCNT	NAME	AMOUNT	MEMO" & vbCrLf
	response.write "!ENDTRNS" & vbCrLf
	rs.open sSQL2,cnn,0,1
	do while NOT rs.EOF
		response.write "TRNS" & vbTab & """" & vsusdate(rs("ordDate")) & """"
		rs.MoveNext
	loop
	rs.close
elseif request.form("act")="ouresolutionsxmldump" then
	response.write "<?xml version=""1.0""?>" & vbCrLf
	response.write "<DATABASE NAME=""DataBaseCopy.mdb"" >" & vbCrLf
	sSQL="SELECT ordID,cartProdId,cartProdName,cartProdPrice,cartQuantity,cartID FROM cart INNER JOIN orders ON cart.cartOrderId=orders.ordID"
	sSQL=sSQL & getsearchparams()
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		theoptionspricediff=0
		sSQL="SELECT coPriceDiff,coOptGroup,coCartOption FROM cartoptions WHERE coCartID=" & rs("cartID")
		rs2.Open sSQL,cnn,0,1
		do while NOT rs2.EOF
			theoptionspricediff=theoptionspricediff + rs2("coPriceDiff")
			rs2.MoveNext
		loop
		rs2.Close
		theunitprice=rs("cartProdPrice")+theoptionspricediff
		sSQL="SELECT pName,pDescription,pDropShip FROM products WHERE pID='"&rs("cartProdID")&"'"
		rs2.Open sSQL,cnn,0,1
		if NOT rs2.EOF then
			prodname=strip_tags2(rs2("pName")&"")
			proddesc=strip_tags2(rs2("pDescription")&"")
			supplier=rs2("pDropShip")
		else
			prodname=""
			proddesc=""
			supplier=0
		end if
		if ouresolutionsxml=1 then
			itemname=strip_tags2(rs("cartProdID")) & "chr(60)brchr(62)" & proddesc
		elseif ouresolutionsxml=3 then
			itemname=strip_tags2(rs("cartProdID"))
		elseif ouresolutionsxml=4 then
			itemname=prodname
		elseif ouresolutionsxml=5 then
			itemname=strip_tags2(rs("cartProdID")) & "chr(60)brchr(62)" & prodname
		else ' default to "2"
			itemname=prodname & "chr(60)brchr(62)" & proddesc
		end if
		rs2.Close
		response.write "<DATA TABLE='oitems' ORDERITEMID='"&rs("cartID")&"' ORDERID='"&rs("ordID")&"' CATALOGID='"&rs("cartID")&"' NUMITEMS='"&rs("cartQuantity")&"' ITEMNAME='"&xmlstrip(itemname)&"' UNITPRICE='"&twodp(theunitprice)&"' DUALPRICE='0' SUPPLIERID='"&supplier&"' ADDRESS='' />" & vbCrLf
		rs.MoveNext
	loop
	rs.close
	sSQL="SELECT ordID,ordName,ordLastName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordExtra1,ordExtra2,ordShipExtra1,ordShipExtra2,ordCheckoutExtra1,ordCheckoutExtra2,ordShipName,ordShipLastName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,ordShipPhone,ordPayProvider,ordAuthNumber,ordTotal,ordDate,ordStateTax,ordCountryTax,ordHSTTax,ordShipping,ordHandling,ordShipType,ordDiscount,ordAffiliate,ordDiscountText,ordStatus,statPrivate,ordAddInfo FROM orders INNER JOIN orderstatus ON orders.ordStatus=orderstatus.statID"
	sSQL=sSQL & getsearchparams()
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		ordGrandTotal=(rs("ordTotal")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordHSTTax")+rs("ordShipping")+rs("ordHandling"))-rs("ordDiscount")
		if usefirstlastname then
			firstname=xmlstrip(trim(rs("ordName")&""))
			lastname=xmlstrip(trim(rs("ordLastName")&""))
		else
			thename=xmlstrip(trim(rs("ordName")&""))
			if thename<>"" then
				if InStr(thename," ") > 0 then
					namearr=Split(thename," ",2)
					firstname=namearr(0)
					lastname=namearr(1)
				else
					firstname=""
					lastname=thename
				end if
			end if
		end if
		response.write "<DATA TABLE='orders' ORDERID='"&rs("ordID")&"' OCUSTOMERID='"&rs("ordID")&"' ODATE='"&DateValue(rs("ordDate"))&"' ORDERAMOUNT='"&twodp(ordGrandTotal)&"' OFIRSTNAME='"&firstname&"' OLASTNAME='"&lastname&"' OEMAIL='"&xmlstrip(rs("ordEmail"))&"' OADDRESS='"&xmlstrip(rs("ordAddress")&IIfVr(trim(rs("ordAddress2")&"")<>"",", " & rs("ordAddress2"), ""))&"' OCITY='"&xmlstrip(rs("ordCity"))&"' OPOSTCODE='"&xmlstrip(rs("ordZip"))&"' OSTATE='"&xmlstrip(rs("ordState"))&"' OCOUNTRY='"&xmlstrip(rs("ordCountry"))&"' OPHONE='"&right(xmlstrip(replace(replace(replace(rs("ordPhone")&""," ", ""),".",""),"-","")), 10)&"' OFAX='' OCOMPANY='"&IIfVr(extra1iscompany=TRUE,xmlstrip(rs("ordExtra1")), "")&"' OCARDTYPE='' "
		if dumpccnumber then
			if sslok=false then
				response.write "OCARDNO='No SSL' OCARDNAME='No SSL' OCARDEXPIRES='No SSL' OCARDADDRESS='No SSL' "
			else
				rs2.Open "SELECT ordCNum,ordPayProvider FROM orders WHERE ordID=" & rs("ordID"),cnn,0,1
				ordCNum=rs2("ordCNum")
				encryptmethod=LCase(encryptmethod&"")
				if encryptmethod="aspencrypt" OR encryptmethod="" then
					response.write "OCARDNO='Encrypted' OCARDNAME='Encrypted' OCARDEXPIRES='Encrypted' OCARDADDRESS='Encrypted' "
				elseif Trim(ordCNum)="" OR IsNull(ordCNum) OR rs2("ordPayProvider")<>10 then
					response.write "OCARDNO='' OCARDNAME='' OCARDEXPIRES='' OCARDADDRESS='' "
				elseif encryptmethod="none" then
					cnumarr=Split(ordCNum, "&")
					if IsArray(cnumarr) then
						response.write "OCARDNO='"&cnumarr(0)&"' OCARDNAME='"&cnumarr(3)&"' OCARDEXPIRES='"&cnumarr(1)&"' OCARDADDRESS='"&rs("ordAddress")&IIfVr(trim(rs("ordAddress2")&"")<>"",", " & rs("ordAddress2"), "")&"' "
					else
						response.write "OCARDNO='' OCARDNAME='' OCARDEXPIRES='' OCARDADDRESS='' "
					end if
				end if
				rs2.Close
			end if
		else
			response.write "OCARDNO='' OCARDNAME='' OCARDEXPIRES='' OCARDADDRESS='' "
		end if
		response.write "OPROCESSED='' OCOMMENT='"&xmlstrip(rs("ordAddInfo"))&"' OTAX='"&twodp(rs("ordStateTax")+rs("ordCountryTax")+rs("ordHSTTax"))&"' OPROMISEDSHIPDATE='' OSHIPPEDDATE='' OSHIPMETHOD='0' OSHIPCOST='"&twodp(rs("ordShipping"))&"' "
		response.write "OSHIPNAME='"&xmlstrip(trim(rs("ordShipName")&" "&rs("ordShipLastName")))&"' OSHIPCOMPANY='' OSHIPEMAIL='' OSHIPMETHODTYPE='"&xmlstrip(rs("ordShipType"))&"' OSHIPADDRESS='"&xmlstrip(rs("ordShipAddress")&IIfVr(trim(rs("ordShipAddress2")&"")<>"",", " & rs("ordShipAddress2"), ""))&"' OSHIPTOWN='"&xmlstrip(rs("ordShipCity"))&"' OSHIPZIP='"&xmlstrip(rs("ordShipZip"))&"' OSHIPCOUNTRY='"&xmlstrip(rs("ordShipCountry"))&"' OSHIPSTATE='"&xmlstrip(rs("ordShipState"))&"' "
		response.write "OPAYMETHOD='"&rs("ordPayProvider")&"' OTHER1='"&IIfVr(extra1iscompany=TRUE,"",xmlstrip(rs("ordExtra1")))&"' OTHER2='"&xmlstrip(rs("ordExtra2"))&"' OTIME='' OAUTHORIZATION='' OERRORS='' ODISCOUNT='"&twodp(rs("ordDiscount"))&"' OSTATUS='"&xmlstrip(rs("statPrivate"))&"' OAFFID='' ODUALTOTAL='0' ODUALTAXES='0' ODUALSHIPPING='0' ODUALDISCOUNT='0' OHANDLING='"&twodp(rs("ordHandling"))&"' COUPON='"&xmlstrip(strip_tags2(rs("ordDiscountText")&""))&"' COUPONDISCOUNT='0' COUPONDISCOUNTDUAL='0' GIFTCERTIFICATE='' GIFTAMOUNTUSED='0' GIFTAMOUNTUSEDDUAL='0' CANCELED='"&IIfVr(rs("ordStatus")<2,"True","False")&"' />" & vbCrLf
		rs.MoveNext
	loop
	rs.close
	response.write "</DATABASE>" & vbCrLf
elseif request.form("act")="dumpevents" then
	if SESSION("loginid")=0 then
		success=TRUE
		sSQL="SELECT userID,eventType,eventDate,eventSuccess,eventOrigin,areaAffected FROM auditlog ORDER BY logID DESC"
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			response.write """"&replace(rs("userID")&"","""","""""")&""","
			response.write """"&replace(rs("eventType")&"","""","""""")&""","
			response.write replace(rs("eventDate")&"","""","""""")&","
			response.write """"&replace(rs("eventSuccess")&"","""","""""")&""","
			response.write """"&replace(rs("eventOrigin")&"","""","""""")&""","
			response.write """"&replace(rs("areaAffected")&"","""","""""")&""""&vbCrLf
			rs.movenext
		loop
		rs.close
	else
		success=FALSE
		response.write "No Access Privileges."
	end if
	call logevent(SESSION("loginuser"),"EVENTLOG",success,"dumporders.asp","DUMP LOG")
elseif request.form("act")="dumpemails" then
	sSQL="SELECT email,mlName,mlIPAddress,mlConfirmDate FROM mailinglist WHERE isconfirmed<>0"
	if request.querystring("entirelist")<>"1" then sSQL=sSQL & " AND selected<>0"
	rs.open sSQL,cnn,0,1
	response.write "Email,Full Name,IP Address,Date Subscribed" & vbCrLf
	do while NOT rs.EOF
		response.write """"&replace(rs("email")&"","""","""""")&""","
		response.write """"&replace(rs("mlName")&"","""","""""")&""","
		response.write """"&replace(rs("mlIPAddress")&"","""","""""")&""","
		thedate=rs("mlConfirmDate")
		if isdate(thedate) then
			thisdate=DatePart("yyyy",thedate) & "-" & IIfVr(DatePart("m",thedate)<10,"0","") & DatePart("m",thedate) & "-" & IIfVr(DatePart("d",thedate)<10,"0","") & DatePart("d",thedate)
			response.write """"& thisdate &""""&vbCrLf
		else
			response.write """"""&vbCrLf
		end if
		rs.movenext
	loop
	rs.close
else
	session.LCID=saveLCID ' If not then dates will not resolve correctly.
	if hasdetails then
		sSQL2="SELECT ordID,ordName,ordLastName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordExtra1,ordExtra2,ordShipExtra1,ordShipExtra2,ordCheckoutExtra1,ordCheckoutExtra2,ordShipName,ordShipLastName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,ordShipPhone,payProvName,ordAuthNumber,ordTotal,ordDate,ordStateTax,ordCountryTax,ordHSTTax,ordShipping,ordHandling,ordShipType,statPrivate,ordAuthStatus,ordTrackNum,ordInvoice,ordDiscount,cartProdId,cartProdName,cartProdPrice,cartQuantity,cartID,ordAddInfo FROM cart INNER JOIN ((orderstatus RIGHT OUTER JOIN orders ON orders.ordStatus=orderstatus.statID) INNER JOIN payprovider ON payprovider.payProvID=orders.ordPayProvider) ON cart.cartOrderId=orders.ordID"
	else
		sSQL2="SELECT ordID,ordName,ordLastName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordExtra1,ordExtra2,ordShipExtra1,ordShipExtra2,ordCheckoutExtra1,ordCheckoutExtra2,ordShipName,ordShipLastName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,ordShipPhone,payProvName,ordAuthNumber,ordTotal,ordDate,ordStateTax,ordCountryTax,ordHSTTax,ordShipping,ordHandling,ordShipType,statPrivate,ordAuthStatus,ordTrackNum,ordInvoice,ordDiscount,ordAddInfo FROM orderstatus RIGHT OUTER JOIN (orders INNER JOIN payprovider ON payprovider.payProvID=orders.ordPayProvider) ON orders.ordStatus=orderstatus.statID"
	end if
	sSQL2=sSQL2 & getsearchparams()
	rs.open sSQL2 & " ORDER BY ordID"&IIfVr(hasdetails,",cartID",""),cnn,0,1
	response.write """OrderID"","
	if extraorderfield1<>"" then response.write """" & replace(extraorderfield1,"""","""""") & ""","
	if usefirstlastname then response.write """FirstName"",""LastName""," else response.write """CustomerName"","
	response.write """Address"","
	if useaddressline2=TRUE then response.write """Address2"","
	response.write """City"",""State"",""Zip"",""Country"",""Email"",""Phone"","
	if extraorderfield2<>"" then response.write """" & replace(extraorderfield2,"""","""""") & ""","
	if extraorderfield1<>"" then response.write """" & replace(extraorderfield1,"""","""""") & ""","
	response.write """ShipName"","
	if usefirstlastname then response.write """ShipLastName"","
	response.write """ShipAddress"","
	if useaddressline2=TRUE then response.write """ShipAddress2"","
	response.write """ShipCity"",""ShipState"",""ShipZip"",""ShipCountry"",""ShipPhone"","
	if extraorderfield2<>"" then response.write """" & replace(extraorderfield2,"""","""""") & ""","
	response.write """PaymentMethod"",""AuthNumber"",""Total"",""Date"",""StateTax"",""CountryTax"","
	if origCountryID=2 then response.write """HST"","
	response.write """Shipping"",""Handling"",""Discounts"",""AddInfo"",""ShippingMethod"",""Status"",""AuthStatus"",""Tracking"",""InvoiceNum"""
	if dumpccnumber then response.write ",""Card Number"",""Expiry Date"",""CVV Code"",""Issue Number"",""Card Name"""
	if hasdetails then response.write ",""ProductID"",""ProductName"",""ProductPrice"",""Quantity"",""Options"""
	response.write vbCrLf
	do while NOT rs.EOF
		response.write rs("ordID")&","
		if extraorderfield1<>"" then response.write """"&replace(rs("ordExtra1")&"","""","""""")&""","
		response.write """"&replace(rs("ordName")&"","""","""""")&""","
		if usefirstlastname then response.write """"&replace(rs("ordLastName")&"","""","""""")&""","
		response.write """"&replace(rs("ordAddress")&"","""","""""")&""","
		if useaddressline2=TRUE then response.write """"&replace(rs("ordAddress2")&"","""","""""")&""","
		response.write """"&replace(rs("ordCity")&"","""","""""")&""","
		response.write """"&replace(rs("ordState")&"","""","""""")&""","
		response.write """"&replace(rs("ordZip")&"","""","""""")&""","
		response.write """"&replace(rs("ordCountry")&"","""","""""")&""","
		response.write """"&replace(rs("ordEmail")&"","""","""""")&""","
		response.write """"&replace(rs("ordPhone")&"","""","""""")&""","
		if extraorderfield2<>"" then response.write """"&replace(rs("ordExtra2")&"","""","""""")&""","
		if extraorderfield1<>"" then response.write """"&replace(rs("ordShipExtra1")&"","""","""""")&""","
		response.write """"&replace(rs("ordShipName")&"","""","""""")&""","
		if usefirstlastname then response.write """"&replace(rs("ordShipLastName")&"","""","""""")&""","
		response.write """"&replace(rs("ordShipAddress")&"","""","""""")&""","
		if useaddressline2=TRUE then response.write """"&replace(rs("ordShipAddress2")&"","""","""""")&""","
		response.write """"&replace(rs("ordShipCity")&"","""","""""")&""","
		response.write """"&replace(rs("ordShipState")&"","""","""""")&""","
		response.write """"&replace(rs("ordShipZip")&"","""","""""")&""","
		response.write """"&replace(rs("ordShipCountry")&"","""","""""")&""","
		response.write """"&replace(rs("ordShipPhone")&"","""","""""")&""","
		if extraorderfield2<>"" then response.write """"&replace(rs("ordShipExtra2")&"","""","""""")&""","
		response.write """"&replace(rs("payProvName")&"","""","""""")&""","
		response.write """"&replace(rs("ordAuthNumber")&"","""","""""")&""","
		response.write """"&rs("ordTotal")&""","
		response.write """"&rs("ordDate")&""","
		response.write """"&rs("ordStateTax")&""","
		response.write """"&rs("ordCountryTax")&""","
		if origCountryID=2 then response.write """"&rs("ordHSTTax")&""","
		response.write """"&rs("ordShipping")&""","
		response.write """"&rs("ordHandling")&""","
		response.write """"&rs("ordDiscount")&""","
		response.write """"&replace(rs("ordAddInfo")&"","""","""""")&""","
		response.write """"&replace(rs("ordShipType")&"","""","""""")&""","
		response.write """"&replace(rs("statPrivate")&"","""","""""")&""","
		response.write """"&replace(rs("ordAuthStatus")&"","""","""""")&""","
		response.write """"&replace(rs("ordTracknum")&"","""","""""")&""","
		response.write """"&replace(rs("ordInvoice")&"","""","""""")&""""
		if dumpccnumber then
			if sslok=false then
				response.write ",No SSL,No SSL,No SSL,No SSL"
			else
				rs2.Open "SELECT ordCNum,ordPayProvider FROM orders WHERE ordID=" & rs("ordID"),cnn,0,1
				ordCNum=rs2("ordCNum")
				encryptmethod=LCase(encryptmethod&"")
				if encryptmethod="aspencrypt" OR encryptmethod="" then
					response.write """Encrypted"",""Encrypted"",""Encrypted"",""Encrypted"",""Encrypted"""
				elseif Trim(ordCNum)="" OR IsNull(ordCNum) OR rs2("ordPayProvider")<>10 then
					response.write ",""(no data)"","""","""","""","""""
				elseif encryptmethod="none" then
					cnumarr=Split(ordCNum, "&")
					if IsArray(cnumarr) then
						response.write ","""""""&cnumarr(0)&""""""""
						if UBOUND(cnumarr)>=1 then response.write ","""""""&cnumarr(1)&"""""""" else response.write ","""""
						if UBOUND(cnumarr)>=2 then response.write ","""&cnumarr(2)&"""" else response.write ","""""
						if UBOUND(cnumarr)>=3 then response.write ","""&cnumarr(3)&"""" else response.write ","""""
						if UBOUND(cnumarr)>=4 then response.write ","""&urldecode(cnumarr(4))&"""" else response.write ","""""
					else
						response.write ",""(no data)"","""","""","""","""""
					end if
				end if
				rs2.Close
			end if
		end if
		if hasdetails then
			theOptions=""
			thePriceDiff=0
			rs2.Open "SELECT coPriceDiff,coOptGroup,coCartOption FROM cartoptions WHERE coCartID=" & rs("cartID") & " ORDER BY coID",cnn,0,1
			do while NOT rs2.EOF
				theOptions=theOptions & "," & """" & replace(rs2("coOptGroup")&"","""","""""") & " - " & replace(rs2("coCartOption"),"""","""""") & """"
				thePriceDiff=thePriceDiff + rs2("coPriceDiff")
				rs2.MoveNext
			loop
			response.write ","""&replace(rs("cartProdId")&"","""","""""")&""""
			response.write ","""&replace(rs("cartProdName")&"","""","""""")&""""
			response.write ","&rs("cartProdPrice")+thePriceDiff
			response.write ","&rs("cartQuantity")
			response.write theOptions
			rs2.Close
		end if
		response.write vbCrLf
		rs.MoveNext
	loop
	rs.close
end if
cnn.Close
set rs=nothing
set rs2=nothing
set cnn=nothing
session.LCID=dosaveLCID
%>
