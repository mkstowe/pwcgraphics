<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
'Build: 7.4.4.003
Dim gasaReferer,gasaThisSite
Dim splitUSZones,countryCurrency,countryNumCurrency,useEuro,storeurl,storeurlssl,handling,handlingchargepercent,adminCanPostUser,adminCanPostLogin,AusPostAPI,adminCanPostPass,packtogether,origZip,shipType,adminShipping,adminIntShipping,saveLCID,delccafter,currRate1,currSymbol1,currRate2,currSymbol2,currRate3,currSymbol3,upsUser,upsPw,adminSecret,cardinalprocessor,cardinalmerchant,cardinalpwd,adminEmailConfirm,reCAPTCHAsitekey,reCAPTCHAsecret,reCAPTCHAuseon,onvacation,mailchimpapikey,mailchimplist
Dim origCountryID,origCountry,origCountryCode,uspsUser,smartPostHub,upsAccess,upsAccount,upsnegdrates,fedexaccount,fedexmeter,fedexuserkey,fedexuserpwd,DHLSiteID,DHLSitePW,DHLAccountNo,adminUnits,adminlanguages,adminlangsettings,ectstorelangarr,storelang,useStockManagement,adminProdsPerPage,countryTax,countryTaxRate,currSymbolText,currDecimalSep,currThousandsSep,currPostAmount,currDecimals,currSymbolHTML,currLastUpdate,currConvUser,currConvPw,emailAddr,allStoreEmails,sendEmail,emailObject,themailhost,smtpport,smtpsecure,theuser,thepass,catalogroot,adminAltRates,prodfilter,sidefilter,prodfiltertext,prodfilterorder,dosortby,sortoptions
htmlemails=0 : blockmultipurchase=0 : blockmaxcartadds=0
Session.LCID=1033
incfunctionsdefined=true : defimagejs=""
function ip2long(ip2lip)
ipret=-1
iparr=split(ip2lip, ".")
if isarray(iparr) then
if UBOUND(iparr)=3 then
if is_numeric(iparr(0)) AND is_numeric(iparr(1)) AND is_numeric(iparr(2)) AND is_numeric(iparr(3)) then
ipret=(iparr(0) * 16777216) + (iparr(1) * 65536) + (iparr(2) * 256) + (iparr(3))
end if
end if
end if
ip2long=ipret
end function
if trim(request.querystring("PARTNER"))<>"" OR trim(request.querystring("REFERER"))<>"" then
	if expireaffiliate="" then expireaffiliate=30
	if trim(request.querystring("PARTNER"))<>"" then thereferer=trim(strip_tags2(request.querystring("PARTNER"))) else thereferer=trim(strip_tags2(request.querystring("REFERER")))
	call setacookie("PARTNER",thereferer,expireaffiliate)
end if
if mysqlserver=TRUE then sqlserver=TRUE
if isempty(extension) then extension=".asp"
storehomeurl="categories" & extension
if xxHomeURL<>"" then storehomeurl=xxHomeURL
if isempty(usecsslayout) then usecsslayout=TRUE
if usecsslayout then nomarkup=TRUE : useproductbodyformat=2 : usesearchbodyformat=2 : usedetailbodyformat=4
if productpagelayout<>"" then useproductbodyformat=2 : usesearchbodyformat=2
if giftcertificateid="" then giftcertificateid="giftcertificate"
if donationid="" then donationid="donation"
if giftwrappingid="" then giftwrappingid="giftwrapping"
if mobilebrowser<>TRUE then mobilebrowser=detectmobilebrowser()
if mobilebrowser then inlinecheckout=TRUE : hideshipaddress=TRUE : usehardaddtocart=IIfVr(usehardcartformobile,TRUE,usehardaddtocart) : disableupdatechecker=TRUE : if mobilebrowsercolumns<>"" then productcolumns=mobilebrowsercolumns
codestr="2952710692840328509902143349209039553396765"
if adminencoding="" then adminencoding="iso-8859-1"
if emailencoding="" then emailencoding=adminencoding
if SESSION("languageid")<>"" then languageid=SESSION("languageid") else if languageid="" then languageid=1
if htmlemails then emlNl="<br />" else emlNl=vbCrLf
if nomarkup=TRUE then sstrong="" : estrong="" else sstrong="<strong>" : estrong="</strong>"
if customeraccounturl="" then customeraccounturl="clientlogin" & extension
if fedextestmode then fedexurl="https://wsbeta.fedex.com:443/web-services" else fedexurl="https://ws.fedex.com:443/web-services"
if loyaltypointvalue="" then loyaltypointvalue=0.0001
if detlinkspacechar="" then detlinkspacechar=" "
if showtaxinclusive=TRUE then showtaxinclusive=1
if NOT is_numeric(showtaxinclusive) then showtaxinclusive=0
if nopriceanywhere=TRUE then noprice=TRUE
if forceclientlogin then enableclientlogin=TRUE
redasterix="<span style=""color:#FF0000"">*</span>"
redstar="<span class=""redstar"" style=""color:#FF1010"">*</span>"
fedexcopyright="FedEx service marks are owned by Federal Express Corporation and are used by permission."
if righttoleft=TRUE then tright="left" : tleft="right" else tright="right" : tleft="left"
if mysqlserver OR sqlserver then txtcollen=8000 else txtcollen=255
pathtohere=request.servervariables("URL")
if instrrev(pathtohere,"/")>0 then pathtohere=left(pathtohere,instrrev(pathtohere, "/"))
if seocaturlpattern="" then seocaturlpattern="/category/%s"
if seoprodurlpattern="" then seoprodurlpattern="/products/%s"
if seomanufacturerpattern="" then seomanufacturerpattern="/manufacturer/%s"
if orhomeurl<>"" then storehomeurl=orhomeurl else if seocategoryurls then storehomeurl=replace(seocaturlpattern,"%s","")
REMOTE_ADDR=ipv6to4(trim(replace(left(IIfVr(request.servervariables("HTTP_CF_CONNECTING_IP")<>"",request.servervariables("HTTP_CF_CONNECTING_IP"),request.servervariables("REMOTE_ADDR")), 48),"'","")))
zipoptional=array(85,91,149,154,200)
function checkcustomerblock()
	if NOT isectadmin then
		set rscb=Server.CreateObject("ADODB.RecordSet")
		set cnncb=Server.CreateObject("ADODB.Connection")
		cnncb.open sDSN
		logip=ip2long(REMOTE_ADDR)
		sSQL="SELECT dcid FROM ipblocking WHERE (dcip1=" & logip & " AND dcip2=0) OR (dcip1<=" & logip & " AND " & logip & "<=dcip2 AND dcip2<>0)"
		rscb.open sSQL,cnncb,0,1
		if NOT rscb.EOF then
			rscb.close
			if customerserviceemail="" then
				sSQL="SELECT adminEmail FROM admin WHERE adminID=1"
				rscb.open sSQL,cnncb,0,1
				if trim(rscb("adminEmail")&"")<>"" then customerserviceemail=split(rscb("adminEmail"),",")(0) else customerserviceemail=""
			end if
			response.clear
			response.status="403 Forbidden"
			print "<html><head><title>System Block</title></head><body><div style=""padding:50px;text-align:center""><p>We apologize but you have been automatically flagged and blocked from this website.</p><p>Please contact us via our customer service email (" & customerserviceemail & ") to have this rectified.</p></div></body></html>"
			response.end
		end if
		rscb.close
		cnncb.close
	end if
end function
call checkcustomerblock()
if SESSION("clientID")<>"" AND sDSN<>"" then
	set clientCnn=Server.CreateObject("ADODB.Connection")
	set clientRS=Server.CreateObject("ADODB.RecordSet")
	clientCnn.open sDSN
	sSQL="SELECT clID FROM customerlogin WHERE clID=" & escape_string(SESSION("clientID")) & " AND clPW='" & escape_string(SESSION("clientPW")) & "'"
	clientRS.open sSQL,clientCnn,0,1
	if clientRS.EOF then SESSION("clientID")=empty
	clientRS.close
	clientCnn.close
	set clientRS=nothing
	set clientCnn=nothing
end if
function getadminsettings()
	if NOT alreadygotadmin then
		sSQL="SELECT adminEmail,htmlemails,adminEmailConfirm,adminProdsPerPage,adminStoreURL,adminStoreURLSSL,adminHandling,adminHandlingPercent,adminDelCC,adminUSZones,adminStockManage,adminShipping,adminIntShipping,adminZipCode,adminUnits,adminlanguages,adminlangsettings,currRate1,currSymbol1,currRate2,currSymbol2,currRate3,currSymbol3,currConvUser,currConvPw,currLastUpdate,adminSecret,countryLCID,countryCurrency,countryNumCurrency,countryName,countryCode,countryID,countryTax,currSymbolText,currDecimalSep,currThousandsSep,currPostAmount,currDecimals,currSymbolHTML,cardinalProcessor,cardinalMerchant,cardinalPwd,catalogRoot,adminAltRates,prodFilter,sideFilter,prodFilterText,prodFilterText2,prodFilterText3,prodFilterOrder,sortOrder,sortOptions,storelang,reCAPTCHAsitekey,reCAPTCHAsecret,reCAPTCHAuseon,onvacation,mailchimpAPIKey,mailchimpList,blockMultiPurchase,blockMaxCartAdds FROM admin INNER JOIN countries ON admin.adminCountry=countries.countryID WHERE adminID=1"
		rs.open sSQL,cnn,0,1
		splitUSZones=(int(rs("adminUSZones"))=1)
		if orlocale<>"" then
			Session.LCID=orlocale
		elseif trim(rs("countryLCID"))<>"0" AND trim(rs("countryLCID"))<>"" then
			on error resume next
			err.number=0
			Session.LCID=cint(rs("countryLCID"))
			if err.number<>0 then response.write "The Locale ID (LCID) " & rs("countryLCID") & " is not available on this server. Please use the orlocale setting to override this.<br />"
			on error goto 0
		end if
		saveLCID=Session.LCID
		countryCurrency=rs("countryCurrency")
		countryNumCurrency=rs("countryNumCurrency")
		useEuro=(countryCurrency="EUR")
		storeurl=trim(rs("adminStoreURL")&"")
		storeurlssl=trim(rs("adminStoreURLSSL")&"")
		useStockManagement=(rs("adminStockManage")<>0)
		adminProdsPerPage=rs("adminProdsPerPage")
		countryTax=cdbl(rs("countryTax"))
		countryTaxRate=cdbl(rs("countryTax"))
		delccafter=int(rs("adminDelCC"))
		handling=cdbl(rs("adminHandling"))
		handlingchargepercent=cdbl(rs("adminHandlingPercent"))
		origZip=rs("adminZipCode")
		shipType=int(rs("adminShipping"))
		adminIntShipping=int(rs("adminIntShipping"))
		origCountry=rs("countryName")
		origCountryCode=rs("countryCode")
		origCountryID=rs("countryID")
		adminUnits=int(rs("adminUnits"))
		htmlemails=rs("htmlemails")<>0
		if htmlemails then emlNl="<br />" else emlNl=vbCrLf
		emailAddr=rs("adminEmail")
		allStoreEmails=rs("adminEmail")
		sendEmail=(rs("adminEmailConfirm") AND 1)=1
		adminEmailConfirm=rs("adminEmailConfirm")
		adminlanguages=int(rs("adminlanguages"))
		adminlangsettings=int(rs("adminlangsettings"))
		storelang=trim(rs("storelang")&"")
		if storelang<>"" then
			ectstorelangarr=split(storelang,"|")
			if UBOUND(ectstorelangarr)>=(languageid-1) then storelang=ectstorelangarr(languageid-1)
		end if
		currRate1=cdbl(rs("currRate1"))
		currSymbol1=trim(rs("currSymbol1")&"")
		currRate2=cdbl(rs("currRate2"))
		currSymbol2=trim(rs("currSymbol2")&"")
		currRate3=cdbl(rs("currRate3"))
		currSymbol3=trim(rs("currSymbol3")&"")
		currConvUser=rs("currConvUser")
		currConvPw=rs("currConvPw")
		currLastUpdate=rs("currLastUpdate")
		currSymbolText=rs("currSymbolText")
		currDecimalSep=rs("currDecimalSep")
		currThousandsSep=rs("currThousandsSep")
		currPostAmount=rs("currPostAmount")
		currDecimals=rs("currDecimals")
		currSymbolHTML=rs("currSymbolHTML")
		adminSecret=rs("adminSecret")
		cardinalprocessor=rs("cardinalProcessor")
		cardinalmerchant=rs("cardinalMerchant")
		cardinalpwd=rs("cardinalPwd")
		catalogroot=rs("catalogRoot")
		adminAltRates=rs("adminAltRates")
		dosortby=rs("sortOrder")
		sortoptions=rs("sortOptions")
		reCAPTCHAsitekey=rs("reCAPTCHAsitekey")
		reCAPTCHAsecret=rs("reCAPTCHAsecret")
		reCAPTCHAuseon=rs("reCAPTCHAuseon")
		onvacation=rs("onvacation")
		blockmultipurchase=rs("blockMultiPurchase")
		blockmaxcartadds=rs("blockMaxCartAdds")
		mailchimpapikey=rs("mailchimpAPIKey")
		mailchimplist=rs("mailchimpList")
		prodfilter=rs("prodFilter")
		sidefilter=rs("sideFilter")
		prodfilterorder=rs("prodFilterOrder")
		prodfiltertext=rs(getlangid("prodFilterText",262144))
		rs.close
	end if
	' Overrides
	if orstoreurl<>"" then storeurl=orstoreurl
	if (left(lcase(storeurl),7)<>"http://") AND (left(lcase(storeurl),8)<>"https://" AND storeurl<>"") then storeurl="http://" & storeurl
	if right(storeurl,1)<>"/" AND storeurl<>"" then storeurl=storeurl & "/"
	if orstoreurlssl<>"" then storeurlssl=orstoreurlssl
	if (left(lcase(storeurlssl),7)<>"http://") AND (left(lcase(storeurlssl),8)<>"https://" AND storeurlssl<>"") then storeurlssl="https://" & storeurlssl
	if right(storeurlssl,1)<>"/" AND storeurlssl<>"" then storeurlssl=storeurlssl & "/"
	if storeurlssl="" AND storeurl<>"" then storeurlssl=storeurl
	if storeurl="" AND storeurlssl<>"" then storeurl=storeurlssl
	if oremailaddr<>"" then allStoreEmails=oremailaddr
	if allStoreEmails<>"" then emailAddr=split(allStoreEmails,",")(0) else emailAddr=""
	if orcatalogroot<>"" then catalogroot=orcatalogroot
	' Language
	if origCountryCode="GB" OR origCountryCode="IE" then
		ssIncTax=replace(ssIncTax,"Tax","VAT")
		xxCntTax="VAT"
	elseif origCountryCode="AU" OR origCountryCode="CA" then
		xxStaTax="PST"
		xxCntTax="GST"
		if storelang="fr" AND origCountryCode="CA" then
			xxStaTax="TVQ"
			xxCntTax="TPS"
		end if
	end if
	if origCountryCode="CA" AND storelang="" then
		xxPostco="Postal Code"
	end if
	getadminsettings=TRUE
end function
function encodeimage(byval imurl)
	imurl=replace(imurl&"","\","/")
	if imurl="prodimages/" then imurl=""
	imurl=replace(replace(replace(replace(replace(imurl,"*","%2A"),"|","%7C"),"<","%3C"),"?","%3F"),">","%3E")
	encodeimage=replace(imurl,"'","\'")
end function
function strip_tags2(mistr)
	set stregexp=new RegExp
	stregexp.pattern="<[^>]+>"
	stregexp.ignorecase=TRUE
	stregexp.global=TRUE
	strip_tags2=stregexp.replace(mistr&"","")
end function
function replaceaccentsansi(ByVal surl)
surl=replace(surl,chr(174),"")
surl=replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(surl,chr(224),"a"),chr(225),"a"),chr(226),"a"),chr(231),"c"),chr(232),"e"),chr(233),"e"),chr(234),"e"),chr(235),"e"),chr(236),"i"),chr(237),"i"),chr(238),"i"),chr(239),"i")
replaceaccentsansi=replace(replace(replace(replace(replace(replace(replace(replace(replace(surl,chr(241),"n"),chr(242),"o"),chr(243),"o"),chr(244),"o"),chr(246),"o"),chr(249),"u"),chr(250),"u"),chr(251),"u"),chr(252),"u")
end function
function replaceaccentsutf(ByVal surl)
surl=replace(replace(surl,"®",""),"™","")
surl=replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(surl,"à","a"),"â","a"),"á","a"),"ç","c"),"è","e"),"ê","e"),"é","e"),"ë","e"),"î","i"),"í","i"),"ï","i")
replaceaccentsutf=replace(replace(replace(replace(replace(replace(replace(replace(replace(surl,"ñ","n"),"ò","o"),"ô","o"),"ó","o"),"ö","o"),"ù","u"),"û","u"),"ú","u"),"ü","u")
end function
function replaceaccents(ByVal surl)
	if lcase(adminencoding)="iso-8859-1" then replaceaccents=replaceaccentsansi(surl) else if lcase(adminencoding)="utf-8" then replaceaccents=replaceaccentsutf(surl) else replaceaccents=surl
end function
function cleanforurl(ByVal surl)
if isempty(urlfillerchar) then urlfillerchar="_"
Set toregexp=new RegExp
toregexp.pattern="<[^>]+>"
toregexp.ignorecase=TRUE
toregexp.global=TRUE
surl=replace(lcase(toregexp.replace(surl, ""))," ",urlfillerchar)
surl=replaceaccents(surl)
toregexp.pattern="[^a-z\"&urlfillerchar&"0-9]"
cleanforurl=toregexp.replace(surl, "")
end function
function xmlencode(ByVal xmlstr)
	xmlencode=vrxmlencode(trim(xmlstr&""))
end function
function vrxmlencode(ByVal xmlstr)
	xmlstr=replace(xmlstr, "&", "&amp;")
	xmlstr=replace(xmlstr, "<", "&lt;")
	xmlstr=replace(xmlstr, ">", "&gt;")
	xmlstr=replace(xmlstr, "'", "&apos;")
	vrxmlencode=replace(xmlstr, """", "&quot;")
end function
function xmlencodecharref(ByVal xmlstr)
	xmlstr=replace(xmlstr&"", "&reg;", "")
	xmlstr=replace(xmlstr, "&", "&#x26;")
	xmlstr=replace(xmlstr, "<", "&#x3c;")
	xmlstr=replace(xmlstr, ">", "&#x3e;")
	tmp_str=""
	for ii=1 to len(xmlstr)
		ch_code=Asc(Mid(xmlstr,ii,1))
		if ch_code<=130 then tmp_str=tmp_str & Mid(xmlstr,ii,1)
	next
	xmlencodecharref=tmp_str
end function
function getlangid(col, bfield)
	if languageid=1 then
		getlangid=col
	else
		if (adminlangsettings AND bfield)<>bfield then getlangid=col else getlangid=col & languageid
	end if
end function
function upsencode(thestr, propcodestr)
	if propcodestr="" then localcodestr=codestr else localcodestr=propcodestr
	newstr=""
	for index=1 to Len(localcodestr)
		thechar=Mid(localcodestr,index,1)
		if NOT is_numeric(thechar) then
			thechar=asc(thechar) MOD 10
		end if
		newstr=newstr & thechar
	next
	localcodestr=newstr
	do while Len(localcodestr) < 40
		localcodestr=localcodestr & localcodestr
	loop
	newstr=""
	for index=1 to Len(thestr)
		thechar=Mid(thestr,index,1)
		newstr=newstr & Chr(asc(thechar)+int(Mid(localcodestr,index,1)))
	next
	upsencode=newstr
end function
function upsdecode(thestr, propcodestr)
	if propcodestr="" then localcodestr=codestr else localcodestr=propcodestr
	newstr=""
	for index=1 to Len(localcodestr)
		thechar=Mid(localcodestr,index,1)
		if NOT is_numeric(thechar) then
			thechar=asc(thechar) MOD 10
		end if
		newstr=newstr & thechar
	next
	localcodestr=newstr
	do while Len(localcodestr) < 40
		localcodestr=localcodestr & localcodestr
	loop
	if IsNull(thestr) then
		upsdecode=""
	else
		newstr=""
		for index=1 to Len(thestr)
			thechar=Mid(thestr,index,1)
			newstr=newstr & Chr(asc(thechar)-int(Mid(localcodestr,index,1)))
		next
		upsdecode=newstr
	end if
end function
function vsusdate(thedate)
	if mysqlserver=true then
		vsusdate="'" & DatePart("yyyy",thedate) & "-" & DatePart("m",thedate) & "-" & DatePart("d",thedate) & "'"
	elseif sqlserver=true then
		vsusdate="CAST('" & DatePart("yyyy",thedate) & IIfVs(DatePart("m",thedate)<10,"0") & DatePart("m",thedate) & IIfVs(DatePart("d",thedate)<10,"0") & DatePart("d",thedate) & "' AS DATETIME)"
	else
		vsusdate="#" & DatePart("m",thedate) & "/" & DatePart("d",thedate) & "/" & DatePart("yyyy",thedate) & "#"
	end if
end function
function vsusdatetime(thedate)
	if mysqlserver=true then
		vsusdatetime="'" & DatePart("yyyy",thedate) & "-" & DatePart("m",thedate) & "-" & DatePart("d",thedate) & " " & DatePart("h",thedate) & ":" & DatePart("n",thedate) & ":" & DatePart("s",thedate) & "'"
	elseif sqlserver=true then
		vsusdatetime="CAST('" & DatePart("yyyy",thedate) & "-" & IIfVs(DatePart("m",thedate)<10,"0") & DatePart("m",thedate) & "-" & IIfVs(DatePart("d",thedate)<10,"0") & DatePart("d",thedate) & "T" & IIfVs(DatePart("h",thedate)<10,"0") & DatePart("h",thedate) & ":" & IIfVs(DatePart("n",thedate)<10,"0") & DatePart("n",thedate) & ":" & IIfVs(DatePart("s",thedate)<10,"0") & DatePart("s",thedate) & "' AS DATETIME)"
	else
		vsusdatetime="#" & DatePart("m",thedate) & "/" & DatePart("d",thedate) & "/" & DatePart("yyyy",thedate) & " " & DatePart("h",thedate) & ":" & DatePart("n",thedate) & ":" & DatePart("s",thedate) & "#"
	end if
end function
function FormatEuroCurrency(amount)
	FormatEuroCurrency=""
	if currPostAmount=0 then FormatEuroCurrency=currSymbolHTML
	session.lcid=1033
	formattednum=FormatNumber(amount,currDecimals,-1,0,IIfVr(currThousandsSep<>"",-1,0))
	formattednum=replace(replace(replace(formattednum,",","x"),".",currDecimalSep),"x",currThousandsSep)
	session.lcid=savelcid
	FormatEuroCurrency=FormatEuroCurrency&formattednum
	if currPostAmount<>0 then FormatEuroCurrency=FormatEuroCurrency&currSymbolHTML
end function
function FormatCurrencyZeroDP(amount)
	FormatCurrencyZeroDP=""
	if currPostAmount=0 then FormatCurrencyZeroDP=currSymbolHTML
	session.lcid=1033
	formattednum=FormatNumber(amount,0,-1,0,IIfVr(currThousandsSep<>"",-1,0))
	formattednum=replace(replace(replace(formattednum,",","x"),".",currDecimalSep),"x",currThousandsSep)
	session.lcid=savelcid
	FormatCurrencyZeroDP=FormatCurrencyZeroDP&formattednum
	if currPostAmount<>0 then FormatCurrencyZeroDP=FormatCurrencyZeroDP&currSymbolHTML
end function
function FormatNumberUS(num,NumDigAfterDec,IncLeadingDig,UseParForNegNum,GroupDig)
	session.lcid=1033
	FormatNumberUS=FormatNumber(num,NumDigAfterDec,IncLeadingDig,UseParForNegNum,GroupDig)
	session.lcid=savelcid
end function
function FormatEmailEuroCurrency(amount)
	FormatEmailEuroCurrency=""
	if currPostAmount=0 then FormatEmailEuroCurrency=currSymbolText
	session.lcid=1033
	formattednum=FormatNumber(amount,currDecimals,-1,0,IIfVr(currThousandsSep<>"",-1,0))
	formattednum=replace(replace(replace(formattednum,",","x"),".",currDecimalSep),"x",currThousandsSep)
	session.lcid=savelcid
	FormatEmailEuroCurrency=FormatEmailEuroCurrency&formattednum
	if currPostAmount<>0 then FormatEmailEuroCurrency=FormatEmailEuroCurrency&currSymbolText
end function
sub do_stock_management(smOrdId)
end sub
sub stock_subtract(smOrdId)
	smOrdId=trim(smOrdId)
	if NOT is_numeric(smOrdId) then smOrdId=0
	sSQL="SELECT cartID,cartProdID,cartQuantity,pStockByOpts FROM cart INNER JOIN products ON cart.cartProdID=products.pID WHERE cartOrderID=" & smOrdId
	rs2.Open sSQL,cnn,0,1
	do while NOT rs2.EOF
		sSQL="UPDATE products SET pNumSales=pNumSales+1 WHERE pID='"&rs2("cartProdID")&"'"
		ect_query(sSQL)
		if useStockManagement then
			if cint(rs2("pStockByOpts"))<>0 then
				sSQL="SELECT coOptID FROM cartoptions INNER JOIN (options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID) ON cartoptions.coOptID=options.optID WHERE optType IN (-4,-2,-1,1,2,4) AND coCartID=" & rs2("cartID")
				rs.open sSQL,cnn,0,1
				do while NOT rs.EOF
					sSQL="UPDATE options SET optStock=optStock-"&rs2("cartQuantity")&" WHERE optID="&rs("coOptID")
					ect_query(sSQL)
					rs.MoveNext
				loop
				rs.close
			else
				sSQL="UPDATE products SET pInStock=pInStock-"&rs2("cartQuantity")&" WHERE pID='"&rs2("cartProdID")&"'"
				ect_query(sSQL)
			end if
			sSQL="SELECT pID,quantity FROM productpackages WHERE packageID='"&escape_string(rs2("cartProdID"))&"'"
			rs.open sSQL,cnn,0,1
			do while NOT rs.EOF
				sSQL="UPDATE products SET pInStock=pInStock-"&(rs2("cartQuantity")*rs("quantity"))&" WHERE pID='"&rs("pID")&"'"
				ect_query(sSQL)
				rs.movenext
			loop
			rs.close
		end if
		rs2.MoveNext
	loop
	rs2.Close
end sub
sub release_stock(smOrdId)
	sSQL="SELECT cartID,cartProdID,cartQuantity,pStockByOpts FROM (cart LEFT JOIN orders ON cart.cartOrderID=orders.ordID) INNER JOIN products ON cart.cartProdID=products.pID WHERE ordAuthStatus<>'MODWARNOPEN' AND cartOrderID=" & smOrdId
	rs2.Open sSQL,cnn,0,1
	do while NOT rs2.EOF
		sSQL="UPDATE products SET pNumSales=pNumSales-1 WHERE pID='"&rs2("cartProdID")&"'"
		ect_query(sSQL)
		if useStockManagement then
			if cint(rs2("pStockByOpts"))<>0 then
				Set clientRS=Server.CreateObject("ADODB.RecordSet")
				sSQL="SELECT coOptID FROM cartoptions INNER JOIN (options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID) ON cartoptions.coOptID=options.optID WHERE optType IN (-4,-2,-1,1,2,4) AND coCartID=" & rs2("cartID")
				clientRS.Open sSQL,cnn,0,1
				do while NOT clientRS.EOF
					sSQL="UPDATE options SET optStock=optStock+"&rs2("cartQuantity")&" WHERE optID="&clientRS("coOptID")
					ect_query(sSQL)
					clientRS.MoveNext
				loop
				clientRS.Close
				Set clientRS=nothing
			else
				sSQL="UPDATE products SET pInStock=pInStock+"&rs2("cartQuantity")&" WHERE pID='"&rs2("cartProdID")&"'"
				ect_query(sSQL)
			end if
			sSQL="SELECT pID,quantity FROM productpackages WHERE packageID='"&escape_string(rs2("cartProdID"))&"'"
			rs3.open sSQL,cnn,0,1
			do while NOT rs3.EOF
				sSQL="UPDATE products SET pInStock=pInStock+"&(rs2("cartQuantity")*rs3("quantity"))&" WHERE pID='"&rs3("pID")&"'"
				ect_query(sSQL)
				rs3.movenext
			loop
			rs3.close
		end if
		rs2.MoveNext
	loop
	rs2.Close
end sub
sub checkaskqextra(aqpindex,aqp,aqpr)
	if aqp<>"" then
		if aqp<>"" AND aqpr then print "if(!efchkextra('askquestionparam"&aqpindex&"',"""&jscheck(strip_tags2(aqp))&"""))return(false);"&vbCrLf
	end if
end sub
sub emailfriendjavascript() %>
<script>
<!--
function openEFWindow(id,askq){
efrdiv=document.createElement('div');
efrdiv.setAttribute('id','efrdiv');
efrdiv.style.zIndex=1000;
efrdiv.style.position='fixed';
efrdiv.style.width='100%';
efrdiv.style.height='100%';
efrdiv.style.top='0px';
efrdiv.style.left='0px';
efrdiv.style.backgroundColor='rgba(140,140,150,0.5)';
document.body.appendChild(efrdiv);
ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
ajaxobj.open("GET", 'emailfriend.asp?lang=<%=storelang%>&'+(askq?'askq=1&':'')+'id='+id, false);
ajaxobj.send(null);
efrdiv.innerHTML=ajaxobj.responseText;
<%		if recaptchaenabled(2) then %>
var captchaWidgetId=grecaptcha.render( 'emfCaptcha', {
	'sitekey' : '<%=reCAPTCHAsitekey%>',
	'expired-callback' : function(){
		emfrecaptchaok=false;
	},
	'callback' : function(response){
		emfrecaptcharesponse=response;
		emfrecaptchaok=true;
	}
});
emfrecaptchaok=false;
<%		end if %>
return false;
}
var emfrecaptchaok=false;
var emfrecaptcharesponse='';
function efchkextra(obid,fldtxt){
	var hasselected=false,fieldtype='';
	var ob=document.getElementById(obid);
	if(ob)fieldtype=(ob.type?ob.type:'radio');
	if(fieldtype=='text'||fieldtype=='textarea'||fieldtype=='password'){
		hasselected=ob.value!='';
	}else if(fieldtype=='select-one'){
		hasselected=ob.selectedIndex!=0;
	}else if(fieldtype=='radio'){
		for(var ii=0;ii<ob.length;ii++)if(ob[ii].checked)hasselected=true;
	}else if(fieldtype=='checkbox')
		hasselected=ob.checked;
	if(!hasselected){
		if(ob.focus)ob.focus();else ob[0].focus();
		alert("<%=jscheck(xxPlsEntr)%> \""+fldtxt+"\".");
		return(false);
	}
	return(true);
}
function efformvalidator(theForm){
	if(document.getElementById('yourname').value==""){
		alert("<%=jscheck(xxPlsEntr)%> \"<%=xxEFNam%>\".");
		document.getElementById('yourname').focus();
		return(false);
	}
	if(document.getElementById('youremail').value==""){
		alert("<%=jscheck(xxPlsEntr)%> \"<%=xxEFEm%>\".");
		document.getElementById('youremail').focus();
		return(false);
	}
	if(document.getElementById('askq').value!='1'){
		if(document.getElementById('friendsemail').value==""){
			alert("<%=jscheck(xxPlsEntr)%> \"<%=xxEFFEm%>\".");
			document.getElementById('friendsemail').focus();
			return(false);
		}
	}else{
		var regex=/[^@]+@[^@]+\.[a-z]{2,}$/i;
		if(!regex.test(document.getElementById('youremail').value)){
			alert("<%=jscheck(xxValEm)%>");
			document.getElementById('youremail').focus();
			return(false);
		}
<%	call checkaskqextra(1,askquestionparam1,askquestionrequired1)
	call checkaskqextra(2,askquestionparam2,askquestionrequired2)
	call checkaskqextra(3,askquestionparam3,askquestionrequired3)
	call checkaskqextra(4,askquestionparam4,askquestionrequired4)
	call checkaskqextra(5,askquestionparam5,askquestionrequired5)
	call checkaskqextra(6,askquestionparam6,askquestionrequired6)
	call checkaskqextra(7,askquestionparam7,askquestionrequired7)
	call checkaskqextra(8,askquestionparam8,askquestionrequired8)
	call checkaskqextra(9,askquestionparam9,askquestionrequired9)
%>	}
<%	if recaptchaenabled(2) then print "if(!emfrecaptchaok){ alert(""" & jscheck(xxRecapt) & """);return(false); }" %>
	return(true);
}
function dosendefdata(){
	if(efformvalidator(document.getElementById('efform'))){
		var ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
		var yourname=document.getElementById("yourname").value;
		var youremail=document.getElementById("youremail").value;
		var friendsemail=(document.getElementById('askq').value=='1'?'':document.getElementById("friendsemail").value);
		var yourcomments=document.getElementById("yourcomments").value;
		postdata="posted=1&efid=" + encodeURIComponent(document.getElementById('efid').value) + (document.getElementById('askq').value=='1'?'&askq=1':'') + "&yourname=" + encodeURIComponent(yourname) + "&youremail=" + encodeURIComponent(youremail) + "&friendsemail=" + encodeURIComponent(friendsemail) + (document.getElementById("origprodid")?"&origprodid="+encodeURIComponent(document.getElementById("origprodid").value):'') + "&yourcomments=" + encodeURIComponent(yourcomments);
		for(var index=0;index<10;index++){
			if(document.getElementById('askquestionparam'+index)){
				var tval,ob=document.getElementById('askquestionparam'+index)
				fieldtype=(ob.type?ob.type:'radio');
				if(fieldtype=='text'||fieldtype=='textarea'||fieldtype=='password'){
					tval=ob.value;
				}else if(fieldtype=='select-one'){
					tval=ob[ob.selectedIndex].value;
				}else if(fieldtype=='radio'){
					for(var ii=0;ii<ob.length;ii++)if(ob[ii].checked)tval=ob[ii].value;
				}else if(fieldtype=='checkbox')
					tval=ob.value;
				postdata+='&askquestionparam'+index+'='+encodeURIComponent(tval);
			}
		}
<%		if recaptchaenabled(2) then print "postdata+='&g-recaptcha-response='+emfrecaptcharesponse;" %>
		ajaxobj.open("POST", "emailfriend.asp?lang=<%=storelang%>",false);
		ajaxobj.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
		ajaxobj.send(postdata);
		document.getElementById('efrcell').innerHTML=ajaxobj.responseText;
	}
}
//-->
</script>
<%
end sub
sub productdisplayscript(doaddprodoptions,isdetail)
if currSymbol1<>"" AND currFormat1="" then currFormat1="%s <span style=""font-weight:bold"">" & currSymbol1 & "</span>"
if currSymbol2<>"" AND currFormat2="" then currFormat2="%s <span style=""font-weight:bold"">" & currSymbol2 & "</span>"
if currSymbol3<>"" AND currFormat3="" then currFormat3="%s <span style=""font-weight:bold"">" & currSymbol3 & "</span>"
	if NOT (pricecheckerisincluded=TRUE) then ' {
		if SESSION("clientID")<>"" AND (wishlistonproducts OR isdetail) then %>
<div id="savelistdiv" class="savecartlist" style="position:absolute;visibility:hidden;top:0px;left:0px;transform:translate(-50%,-50%);width:auto;height:auto;z-index:11000;display:table" onmouseover="oversldiv=true" onmouseout="oversldiv=false;setTimeout('checksldiv()',1000)">
<div style="display:table-row"><div class="cobll" onmouseover="this.className='cobhl'" onmouseout="this.className='cobll'" style="white-space:nowrap;display:table-cell"><a class="ectlink wishlistmenu" href="#" onclick="return subformid(gtid,'0','<%=jsescape(xxMyWisL)%>')"><%=xxMyWisL%></a></div></div>
<%			sSQL="SELECT listID,listName FROM customerlists WHERE listOwner="&SESSION("clientID")
			rs2.Open sSQL,cnn,0,1
			do while NOT rs2.EOF
				response.write "<div style=""display:table-row""><div class=""cobll"" onmouseover=""this.className='cobhl'"" onmouseout=""this.className='cobll'"" style=""white-space:nowrap;display:table-cell""><a class=""ectlink wishlistmenu"" href=""#"" onclick=""return subformid(gtid,'"&rs2("listID")&"','"&jsescape(rs2("listName"))&"')"">"&htmlspecials(rs2("listName"))&"</a></div></div>"
				rs2.movenext
			loop
			rs2.close %>
</div>
<%		end if
		if notifybackinstock then %>
<div id="notifyinstockcover" style="display:none;position:fixed;width:100%;height:100%;background-color:rgba(140,140,150,0.5);top:0px;left:0px;z-index:1000">
	<div class="notifyinstock">
		<div style="padding:6px;float:left;height:31px;font-size:1.4em"><%=xxNotSor%></div>
		<div style="padding:6px;float:right"><img src="images/close.gif" style="cursor:pointer" alt="Close" onclick="closeinstock()" /></div>
		<div style="border-left:6px solid #ffffff;border-right:6px solid #ffffff;padding:6px;background:#eeeeee;clear:both"><%=xxNotCur%></div>
		<div style="padding:5px;font-size:0.8em"><%=xxNotEnt%></div>
		<div style="padding:2px 0px 4px 4px;float:left"><input style="border:1px solid #333;padding:5px;width:220px" id="nsemailadd" type="text" /></div>
		<div style="padding:4px 6px 6px 4px;float:right"><input type="button" class="ectbutton" value="<%=xxEmaiMe%>" onclick="regnotifystock()" /></div>
	</div>
</div>
<%		end if
%><input type="hidden" id="hiddencurr" value="<%=replace(FormatEuroCurrency(0),"""","&quot;")%>" /><div id="opaquediv" style="display:none;position:fixed;width:100%;height:100%;background-color:rgba(140,140,150,0.5);top:0px;left:0px;text-align:center;z-index:10000;"></div><%
if recaptchaenabled(2) then call displayrecaptchajs("efcaptcha",FALSE,FALSE)
%><script>
/* <![CDATA[ */
<%
print "var xxAddWiL=""" & jscheck(xxAddWiL) & """," & _
		"xxBakOpt=""" & jscheck(xxBakOpt) & """," & _
		"xxCarCon=""" & jscheck(xxCarCon) & """," & _
		"xxClkHere=""" & jscheck(xxClkHere) & """," & _
		"xxClsWin=""" & jscheck(xxClsWin) & """," & _
		"xxCntShp=""" & jscheck(xxCntShp) & """," & _
		"xxCntTax=""" & jscheck(xxCntTax) & """," & _
		"xxDigits=""" & jscheck(xxDigits) & """," & _
		"xxDscnts=""" & jscheck(xxDscnts) & """," & _
		"xxEdiOrd=""" & jscheck(xxEdiOrd) & """," & _
		"xxEntMul=""" & jscheck(xxEntMul) & """," & _
		"xxHasAdd=""" & jscheck(xxHasAdd) & """," & _
		"xxInStNo=""" & jscheck(xxInStNo) & """," & _
		"xxInvCha=""" & jscheck(xxInvCha) & """," & _
		"xxListPrice=""" & jscheck(xxListPrice) & """," & _
		"xxNotBaS=""" & jscheck(xxNotBaS) & """," & _
		"xxNotSto=""" & jscheck(xxNotSto) & """," & _
		"xxOpSkTx=""" & jscheck(xxOpSkTx) & """," & _
		"xxOptOOS=""" & jscheck(xxOptOOS) & """," & _
		"xxOutStok=""" & jscheck(xxOutStok) & """," & _
		"xxPrd255=""" & jscheck(xxPrd255) & """," & _
		"xxPrdChs=""" & jscheck(xxPrdChs) & """," & _
		"xxPrdEnt=""" & jscheck(xxPrdEnt) & """," & _
		"xxPrice=""" & jscheck(xxPrice) & """," & _
		"xxSCAdOr=""" & jscheck(xxSCAdOr) & """," & _
		"xxSCBakO=""" & jscheck(xxSCBakO) & """," & _
		"xxSCCarT=""" & jscheck(xxSCCarT) & """," & _
		"xxSCItem=""" & jscheck(xxSCItem) & """," & _
		"xxSCStkW=""" & jscheck(xxSCStkW) & """," & _
		"xxValEm=""" & jscheck(xxValEm) & """;" & vbCrLf
print "var absoptionpricediffs=" & IIfVr(absoptionpricediffs,"true","false") & ";" & vbCrLf
print "var cartpageonhttps=" & IIfVr(cartpageonhttps,"true","false") & ";" & vbCrLf
print "var currDecimalSep='" & jsescape(currDecimalSep) & "';" & vbCrLf
print "var currencyseparator='" & jsescape(currencyseparator) & "';" & vbCrLf
print "var currThousandsSep='" & jsescape(currThousandsSep) & "';" & vbCrLf
print "var ectbody3layouttaxinc=" & IIfVr(ectbody3layouttaxinc,"true","false") & ";" & vbCrLf
print "var extension='" & extension & "';" & vbCrLf
print "var extensionabs='asp';" & vbCrLf
tempStr=FormatEuroCurrency(0)
print "var hasdecimals=" & IIfVr(InStr(tempStr,",")<>0 OR InStr(tempStr,".")<>0,"true","false") & ";" & vbCrLf
print "var hideoptpricediffs=" & IIfVr(hideoptpricediffs,"true","false") & ";" & vbCrLf
print "var imgsoftcartcheckout='" & jsescapel(imageorbutton(imgsoftcartcheckout,xxCOTxt,"sccheckout",IIfVs(cartpageonhttps,storeurlssl)&"cart"&extension, FALSE)) & "';" & vbCrLf
print "var noencodeimages=true;" & vbCrLf
print "var noprice=" & IIfVr(noprice,"true","false") & ";" & vbCrLf
print "var nopriceanywhere=" & IIfVr(nopriceanywhere,"true","false") & ";" & vbCrLf
print "var noshowoptionsinstock=" & IIfVr(noshowoptionsinstock,"true","false") & ";" & vbCrLf
print "var notifybackinstock=" & IIfVr(notifybackinstock,"true","false") & ";" & vbCrLf
print "var noupdateprice=" & IIfVr(noupdateprice,"true","false") & ";" & vbCrLf
print "var pricezeromessage=""" & jscheck(pricezeromessage) & """;" & vbCrLf
print "var showinstock=" & IIfVr(showinstock,"true","false") & ";" & vbCrLf
print "var stockdisplaythreshold=" & IIfVr(stockdisplaythreshold<>"",stockdisplaythreshold,0) & ";" & vbCrLf
print "var showtaxinclusive=" & showtaxinclusive & ";" & vbCrLf
print "var storeurlssl='" & storeurlssl & "';" & vbCrLf
print "var tax=" & replace(countryTaxRate,",",".") & ";" & vbCrLf
print "var txtcollen=" & txtcollen & ";" & vbCrLf
print "var usehardaddtocart=" & IIfVr(usehardaddtocart,"true","false") & ";" & vbCrLf
print "var usestockmanagement=" & IIfVr(useStockManagement,"true","false") & ";" & vbCrLf
print "var yousavetext=""" & jscheck(yousavetext) & """;" & vbCrLf
print "var zero2dps='" & FormatNumber(0,2) & "';" & vbCrLf
print "var currFormat1='" & jsescapel(currFormat1) & "',currFormat2='" & jsescapel(currFormat2) & "',currFormat3='" & jsescapel(currFormat3) & "';" & vbCrLf
Session.LCID=1033
print "var currRate1=" & currRate1 & ",currRate2=" & currRate2 & ",currRate3=" & currRate3 & ";" & vbCrLf
Session.LCID=saveLCID
print "var currSymbol1='" & jsescapel(currSymbol1) & "',currSymbol2='" & jsescapel(currSymbol2) & "',currSymbol3='" & jsescapel(currSymbol3) & "';" & vbCrLf
print "var softcartrelated=" & IIfVr(softcartrelated,"true","false") & ";" & vbCrLf
%>
function updateoptimage(theitem,themenu,opttype){
var imageitemsrc='',mzitem,theopt,theid,imageitem,imlist,imlistl,fn=window['updateprice'+theitem];
fn();
if(opttype==1){
	theopt=document.getElementsByName('optn'+theitem+'x'+themenu)
	for(var i=0; i<theopt.length; i++)
		if(theopt[i].checked)theid=theopt[i].value;
}else{
	theopt=document.getElementById('optn'+theitem+'x'+themenu)
	theid=theopt.options[theopt.selectedIndex].value;
}
<%	if magictool<>"" then %>
if(mzitem=(document.getElementById("zoom1")?document.getElementById("zoom1"):document.getElementById("mz"+(globalquickbuyid!==''?'qb':'prod')+"image"+theitem))){
	if(aIML[theid]){
		<%=magictool%>.update(mzitem,vsdecimg(aIML[theid]),vsdecimg(aIM[theid]));
	}else if(pIM[0]&&pIM[999]){
		imlist=pIM[0];imlistl=pIM[999];
		for(var index=0;index<imlist.length;index++)
			if(imlist[index]==aIM[theid]&&imlistl[index]){<%=magictool%>.update(mzitem.id,vsdecimg(imlistl[index]),vsdecimg(aIM[theid]));return;}
		if(aIM[theid])<%=magictool%>.update(mzitem.id,vsdecimg(aIM[theid]),vsdecimg(aIM[theid]));
	}else if(aIM[theid])
		<%=magictool%>.update(mzitem.id,vsdecimg(aIM[theid]),vsdecimg(aIM[theid]));
}else
<%	end if %>
	if(imageitem=document.getElementById((globalquickbuyid!==''?'qb':'prod')+"image"+theitem)){
		if(aIM[theid]){
			if(typeof(imageitem.src)!='unknown')imageitem.src=vsdecimg(aIM[theid]);
		}
	}
}
function updateprodimage2(isqb,theitem,isnext){
var imlist=pIM[theitem];
if(!pIX[theitem])pIX[theitem]=0;
if(isnext) pIX[theitem]++; else pIX[theitem]--;
if(pIX[theitem]<0) pIX[theitem]=imlist.length-1;
if(pIX[theitem]>=imlist.length) pIX[theitem]=0;
if(document.getElementById((isqb?'qb':'prod')+"image"+theitem)){document.getElementById((isqb?'qb':'prod')+"image"+theitem).src='';document.getElementById((isqb?'qb':'prod')+"image"+theitem).src=vsdecimg(imlist[pIX[theitem]]);}
document.getElementById((isqb?'qb':'extra')+"imcnt"+theitem).innerHTML=pIX[theitem]+1;
<%	if magictool<>"" then %>
if(pIML[theitem]){
	var imlistl=pIML[theitem];
	if(imlistl.length>=pIX[theitem])
		if(mzitem=document.getElementById("mz"+(isqb?'qb':'prod')+"image"+theitem))<%=magictool%>.update(mzitem,vsdecimg(imlistl[pIX[theitem]]),vsdecimg(imlist[pIX[theitem]]));
}
<%	end if %>
return false;
}
<%		if doaddprodoptions then
			if customvalidator<>"" then %>
function customvalidator(theForm){
<%=customvalidator%>
return(true);
}
<%			end if
		end if ' doaddprodoptions
%>
/* ]]> */
</script><%
		pricecheckerisincluded=TRUE
	end if ' } pricecheckerisincluded
end sub
sub updatepricescript()
	prodoptions=""
	sSQL="SELECT poOptionGroup,optType,optFlags,optTxtMaxLen,optAcceptChars,0 AS isDepOpt FROM prodoptions INNER JOIN optiongroup ON optiongroup.optGrpID=prodoptions.poOptionGroup WHERE poProdID='"&escape_string(rs("pID"))&"' AND NOT (poProdID='"&escape_string(giftcertificateid)&"' OR poProdID='"&escape_string(donationid)&"') ORDER BY poID"
	rs2.open sSQL,cnn,0,1
	if NOT rs2.EOF then prodoptions=rs2.getrows
	rs2.close
	if isarray(allimages) then
		if UBOUND(allimages,2)>0 then
			response.write "<script>/* <![CDATA[ */" & vbCrLf & "pIM["&Count&"]=['"&encodeimage(allimages(0,0))&"'"
			extraimages=1
			for index=1 to UBOUND(allimages,2)
				print ",'" & encodeimage(allimages(0,index))&"'":extraimages=extraimages+1
			next
			response.write "];"&vbCrLf
			if isarray(alllgimages) then
				if UBOUND(alllgimages,2)>0 then
					response.write "pIML["&Count&"]=['"&encodeimage(alllgimages(0,0))&"'"
					for index=1 to UBOUND(alllgimages,2)
						print ",'" & encodeimage(alllgimages(0,index))&"'"
					next
					response.write "];"&vbCrLf
				end if
			end if
			response.write "/* ]]> */</script>"
		end if
	end if
end sub
function checkDPs(currcode)
	if currcode="JPY" OR currcode="TWD" then checkDPs=0 else checkDPs=2
end function
sub checkCurrencyRates(currConvUser,currConvPw,currLastUpdate,byRef currRate1,currSymbol1,byRef currRate2,currSymbol2,byRef currRate3,currSymbol3)
	ccsuccess=TRUE
	if currConvUser<>"" AND currLastUpdate < Now()-1 then
		sstr=""
		if currSymbol1<>"" then sstr=sstr & "&curr=" & currSymbol1
		if currSymbol2<>"" then sstr=sstr & "&curr=" & currSymbol2
		if currSymbol3<>"" then sstr=sstr & "&curr=" & currSymbol3
		if sstr="" then
			ect_query("UPDATE admin SET currLastUpdate=" & vsusdate(Now()))
			exit sub
		end if
		sstr="?source=" & countryCurrency & "&user=" & currConvUser & "&pw=" & currConvPw & sstr
		xmlDoc="xml"
		if callxmlfunction("https://www.ecommercetemplates.com/currencyxml.asp" & sstr,"",xmlDoc,"","Msxml2.ServerXMLHTTP",errormsg,FALSE) then
			set t2=xmlDoc.getElementsByTagName("currencyRates").Item(0)
			for j=0 to t2.childNodes.length - 1
				Set n=t2.childNodes.Item(j)
				if n.nodename="currError" then
					call notification_err_msg(n.firstChild.nodeValue)
					ccsuccess=FALSE
				elseif n.nodename="selectedCurrency" then
					currRate=0
					for i=0 To n.childNodes.length - 1
						Set e=n.childNodes.Item(i)
						if e.nodeName="currSymbol" then
							currSymbol=e.firstChild.nodeValue
						elseif e.nodeName="currRate" then
							currRate=e.firstChild.nodeValue
						end if
					next
					saveLCID=Session.LCID
					Session.LCID=1033
					if currSymbol1=currSymbol then
						currRate1=cdbl(currRate)
						ect_query("UPDATE admin SET currRate1="&currRate&" WHERE adminID=1")
					end if
					if currSymbol2=currSymbol then
						currRate2=cdbl(currRate)
						ect_query("UPDATE admin SET currRate2="&currRate&" WHERE adminID=1")
					end if
					if currSymbol3=currSymbol then
						currRate3=cdbl(currRate)
						ect_query("UPDATE admin SET currRate3="&currRate&" WHERE adminID=1")
					end if
					Session.LCID=saveLCID
				end if
			next
			if ccsuccess then ect_query("UPDATE admin SET currLastUpdate=" & vsusdate(Now()))
		end if
		set xmlDoc=nothing
	end if
end Sub
function IIfVr(theExp,theTrue,theFalse)
if theExp then IIfVr=theTrue else IIfVr=theFalse
end function
function IIfVs(theExp,theTrue)
if theExp then IIfVs=theTrue else IIfVs=""
end function
function getsectionids(thesecid, delsections)
	getsectionids=""
	secarr=split(thesecid, ",")
	secid="" : addcomma="" : addcomma2=""
	for each sect in secarr
		if is_numeric(trim(sect)) then secid=secid & addcomma & sect : addcomma=","
	next
	if secid="" then secid="0"
	iterations=0
	iteratemore=TRUE
	if SESSION("clientLoginLevel")<>"" then minloglevel=SESSION("clientLoginLevel") else minloglevel=0
	if delsections then nodel="" else nodel="sectionDisabled<="&minloglevel&" AND "
	do while iteratemore AND iterations<10
		sSQL2="SELECT DISTINCT sectionID,rootSection FROM sections WHERE " & nodel & "(topSection IN ("&secid&")"
		if iterations=0 then sSQL2=sSQL2 & " OR (sectionID IN ("&secid&") AND rootSection=1))" else sSQL2=sSQL2 & ")"
		secid=""
		iteratemore=FALSE
		rs2.Open sSQL2,cnn,0,1
		addcomma=""
		do while NOT rs2.EOF
			if rs2("rootSection")=0 then
				if returnalltopsections then getsectionids=getsectionids & addcomma2 & rs2("sectionID") : addcomma2=","
				iteratemore=TRUE
				secid=secid & addcomma & rs2("sectionID")
				addcomma=","
			else
				getsectionids=getsectionids & addcomma2 & rs2("sectionID")
				addcomma2=","
			end if
			rs2.MoveNext
		loop
		rs2.Close
		iterations=iterations + 1
	loop
	if getsectionids="" then getsectionids="0"
end function
function callxmlfunction(cfurl, cfxml, byref res, cfcert, cxfobj, byref cferr, settimeouts)
	callxmlfunctionstatus=0
	debugres=""
	if proxyserver<>"" AND cxfobj="Msxml2.ServerXMLHTTP" then
		set objHttp=Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
		objHttp.setProxy 2, proxyserver
	else
		set objHttp=Server.CreateObject(cxfobj)
	end if
	if settimeouts=TRUE then
		objHttp.setTimeouts 30000, 30000, 25000, 25000
	elseif settimeouts then
		objHttp.setTimeouts settimeouts*1000, settimeouts*1000, settimeouts*1000, settimeouts*1000
	end if
	objHttp.open IIfVr(cfxml<>"","POST","GET"), cfurl, false
	hascontenttype=FALSE
	if isarray(xmlfnheaders) then
		for each objitem in xmlfnheaders
			objHttp.setRequestHeader objitem(0), objitem(1)
			if objitem(0)="Content-Type" then hascontenttype=TRUE
		next
	end if
	xmlfnheaders=""
	if NOT hascontenttype then objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	' if cfcert<>"" then objHttp.setOption 3, "LOCAL_MACHINE\My\" & cfcert
	if cfcert<>"" then objHttp.SetClientCertificate("LOCAL_MACHINE\My\" & cfcert)
	on error resume next
	err.number=0
	objHttp.Send cfxml
	errnum=err.number
	errdesc=err.description
	on error goto 0
	if errnum<>0 then
		cferr="Error, couldn't connect to "&cfurl&" (" & errnum & ").<br />" & errdesc
		callxmlfunction=FALSE
	elseif objHttp.status<>200 AND objHttp.status<>201 AND NOT iscanadapost then
		callxmlfunctionstatus=objHttp.status
		if isstripedotcom then res=objHttp.responseText
		cferr="Error, invalid response from "&cfurl&" (" & objHttp.status & ")."
		callxmlfunction=FALSE
	else
		callxmlfunctionstatus=objHttp.status
		if res="xml" then set res=objHttp.responseXML else res=objHttp.responseText
		debugres=objHttp.responseText
		callxmlfunction=TRUE
	end if
	on error resume next
	if debugmode then savehtmlemails=htmlemails : htmlemails=FALSE : call dosendemaileo(emailAddr, emailAddr, "", "ASP XML Function Debug", cfxml & vbCrLf & vbCrLf & debugres & vbCrLf & vbCrLf & callxmlfunction,emailObject,themailhost,theuser,thepass) : htmlemails=savehtmlemails
	on error goto 0
	set objHttp=nothing
end function
function getpayprovdetails(ppid,ppdata1,ppdata2,ppdata3,ppdemo,ppmethod)
	getpayprovdetails=getpayprovdetx(ppid,ppdata1,ppdata2,ppdata3,ppdata4,ppdata5,ppdata6,ppflag1,ppflag2,ppflag3,ppbits,ppdemo,ppmethod)
end function
function getpayprovdetx(ppid,ppdata1,ppdata2,ppdata3,ppdata4,ppdata5,ppdata6,ppflag1,ppflag2,ppflag3,ppbits,ppdemo,ppmethod)
	if NOT is_numeric(ppid) then ppid=0
	sSQL="SELECT payProvData1,payProvData2,payProvData3,payProvData4,payProvData5,payProvData6,payProvFlag1,payProvFlag2,payProvFlag3,payProvBits,payProvDemo,payProvMethod FROM payprovider WHERE payProvEnabled=1 AND payProvID=" & replace(ppid,"'","")
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		ppdata1=trim(rs("payProvData1")&"")
		ppdata2=trim(rs("payProvData2")&"")
		ppdata3=trim(rs("payProvData3")&"")
		ppdata4=trim(rs("payProvData4")&"")
		ppdata5=trim(rs("payProvData5")&"")
		ppdata6=trim(rs("payProvData6")&"")
		ppflag1=rs("payProvFlag1")
		ppflag2=rs("payProvFlag2")
		ppflag3=rs("payProvFlag3")
		ppbits=rs("payProvBits")
		ppdemo=(cint(rs("payProvDemo"))=1)
		ppmethod=int(rs("payProvMethod"))
		getpayprovdetx=TRUE
	else
		getpayprovdetx=FALSE
	end if
	rs.close
end function
sub writehiddenvar(hvname,hvval)
response.write "<input type=""hidden"" name=""" & hvname & """ value=""" & htmlspecials(hvval) & """>" & vbCrLf
end sub
function whv(hvname,hvval)
whv="<input type=""hidden"" name=""" & hvname & """ value=""" & htmlspecials(hvval) & """>" & vbCrLf
end function
sub print(ps)
response.write ps
end sub
sub writehiddenidvar(hvname,hvval)
response.write "<input type=""hidden"" name=""" & hvname & """ id=""" & hvname & """ value=""" & htmlspecials(hvval) & """>" & vbCrLf
end sub
function ppsoapheader(username, password, threetokenhash)
ppsoapheader="<" & "?xml version=""1.0"" encoding=""utf-8""?><soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""><soap:Header><RequesterCredentials xmlns=""urn:ebay:api:PayPalAPI""><Credentials xmlns=""urn:ebay:apis:eBLBaseComponents"">" & IIfVr(instr(username,"@AB@")=0, "<Username>" & username & "</Username><Password>" & password & "</Password>" & IIfVr(threetokenhash<>"","<Signature>"&threetokenhash&"</Signature>",""), "<Subject>"&replace(username,"@AB@","")&"</Subject>") & "</Credentials></RequesterCredentials></soap:Header>"
end function
function getoptpricediff(opd, theid, theexp, pprice, ByRef pstock)
	getoptpricediff=cdbl(opd)
	if theexp<>"" AND Left(theexp,1)<>"!" then
		theexp=replace(theexp, "%s", theid)
		if InStr(theexp, " ") > 0 then ' Search and replace
			exparr=split(theexp, " ", 2)
			theid=replace(theid, exparr(0), exparr(1), 1, 1)
		else
			theid=theexp
		end if
		sSQL="SELECT "&WSP&"pPrice,pInStock FROM products WHERE pID='"&escape_string(theid)&"'"
		rs3.Open sSQL,cnn,0,1
		if NOT rs3.EOF then getoptpricediff=rs3("pPrice")-pprice : pstock=rs3("pInStock")
		rs3.Close
	end if
end function
sub addtoaltids(theexp, byref altidarr, byref altids)
	theexp=trim(theexp&"")
	if (theexp<>"") AND Left(theexp,1)<>"!" then
		if NOT isarray(altidarr) then
			altidarr=split(trim(altids))
			altids=""
		end if
		for each theid in altidarr
			if instr(altids,theid & " ")=0 then altids=altids & theid & " "
			theexpa=replace(theexp, "%s", theid)
			if InStr(theexpa, " ") > 0 then ' Search and replace
				exparr=split(theexpa, " ", 2)
				theid=replace(theid, exparr(0), exparr(1), 1, 1)
			else
				theid=theexpa
			end if
			if instr(altids,theid & " ")=0 then altids=altids & theid & " "
		next
	end if
end sub
optjsunique=","
sub addtooptionsjs(byref optionsjs, isdetail, origoptpricediff)
	if instr(optjsunique,","&rs2("optID")&",")=0 then
		if useStockManagement then optionsjs=optionsjs & "oS["&rs2("optID")&"]="&rs2("optStock")&";"
		session.LCID=1033
		if (trim(rs2("optRegExp")&"")="" OR left(trim(rs2("optRegExp")&""),1)="!") AND origoptpricediff<>0 then optionsjs=optionsjs & "op["&rs2("optID")&"]="&origoptpricediff&";"
		session.LCID=saveLCID
		if trim(rs2("optRegExp")&"")<>"" AND left(trim(rs2("optRegExp")&""),1)<>"!" then optionsjs=optionsjs & "or["&rs2("optID")&"]='"&rs2("optRegExp")&"';"
		optionsjs=optionsjs & "ot["&rs2("optID")&"]="""&jscheck(rs2(getlangid("optName",32))&"")&""";"
		if trim(rs2("optAlt"&IIfVr(isdetail,"Large","")&"Image")&"")<>"" then optionsjs=optionsjs & "aIM["&rs2("optID")&"]='"&encodeimage(rs2("optAlt"&IIfVr(isdetail,"Large","")&"Image"))&"';"
		if trim(rs2("optDependants")&"")<>"" then optionsjs=optionsjs & "dOP["&rs2("optID")&"]=["&rs2("optDependants")&"];"
		if magictoolboxproducts<>"" AND NOT isdetail AND trim(rs2("optAltLargeImage")&"")<>"" then optionsjs=optionsjs & "aIML["&rs2("optID")&"]='"&encodeimage(rs2("optAltLargeImage"))&"';"
		optionsjs=optionsjs & vbCrLf
		optjsunique=optjsunique&rs2("optID")&","
	end if
end sub
hasincludedpopcalendar=FALSE
mustincludepopcalendar=FALSE
function getoptionspricediff(thetax)
	optpricediff=0 : pricediff=0 : rowcounter=0
	altids=rs("pID")
	maxindex=UBOUND(prodoptions,2)
	do while rowcounter<=maxindex
		opthasstock=FALSE
		sSQL="SELECT optID,"&getlangid("optName",32)&","&getlangid("optGrpName",16)&","&OWSP&"optPriceDiff,optType,optGrpSelect,optFlags,optTxtMaxLen,optTxtCharge,optStock,optPriceDiff AS optDims,optDefault,optAltImage,optAltLargeImage,optRegExp,"&getlangid("optPlaceholder",16)&","&IIfVs(prodoptions(5,rowcounter)=1,"'' AS ")&"optDependants,optTooltip,optClass FROM options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optGroup="&prodoptions(0,rowcounter)&" ORDER BY optID"
		if mysqlserver then rs2.CursorLocation=3
		rs2.Open sSQL,cnn,1,1
		if NOT rs2.EOF then
			if abs(int(rs2("optType")))=5 then ' Date Picker
			elseif abs(int(rs2("optType")))=3 then ' Text
			elseif abs(int(rs2("optType")))=1 then ' Checkbox / Radio
				do while not rs2.EOF
					origoptpricediff=getoptpricediff(rs2("optPriceDiff"), rs("pID"), trim(rs2("optRegExp")&""), rs("pPrice"), stocknotused)
					call addtoaltids(rs2("optRegExp"), altidarr, altids)
					if cint(rs2("optDefault"))<>0 AND origoptpricediff<>0 AND trim(rs2("optRegExp")&"")="" then
						if (rs2("optFlags") AND 1)=1 then pricediff=(rs("pPrice")*origoptpricediff)/100.0 else pricediff=origoptpricediff
						if showtaxinclusive=2 AND (rs("pExemptions") AND 2)<>2 then pricediff=pricediff+(pricediff*thetax/100.0)
						optpricediff=optpricediff + pricediff
					end if
					rs2.MoveNext
				loop
				set altidarr=nothing
			elseif abs(int(rs2("optType")))=4 then ' Multi
			else ' Select
				gotdefaultdiff=FALSE
				firstpricediff=0
				origoptpricediff=rs2("optPriceDiff")
				if cint(rs2("optGrpSelect"))=0 then
					if (rs2("optFlags") AND 1)=1 then firstpricediff=(rs("pPrice")*origoptpricediff)/100.0 else firstpricediff=origoptpricediff
				end if
				do while not rs2.EOF
					origoptpricediff=getoptpricediff(rs2("optPriceDiff"), rs("pID"), trim(rs2("optRegExp")&""), rs("pPrice"), stocknotused)
					call addtoaltids(rs2("optRegExp"), altidarr, altids)
					if cint(rs2("optDefault"))<>0 AND trim(rs2("optRegExp")&"")="" then
						if origoptpricediff<>0 then
							if (rs2("optFlags") AND 1)=1 then pricediff=(rs("pPrice")*origoptpricediff)/100.0 else pricediff=origoptpricediff
							if showtaxinclusive=2 AND (rs("pExemptions") AND 2)<>2 then pricediff=pricediff+(pricediff*thetax/100.0)
							if cint(rs2("optDefault"))<>0 then optpricediff=optpricediff + pricediff
						end if
						gotdefaultdiff=TRUE
					end if
					rs2.MoveNext
				loop
				set altidarr=nothing
				if NOT gotdefaultdiff then optpricediff=optpricediff + firstpricediff
			end if
		end if
		rs2.Close
		rowcounter=rowcounter+1
	loop
	getoptionspricediff=optpricediff
end function
function displayproductoptions(grpnmstyle,grpnmstyleend,byRef optpricediff,thetax,isdetail, byRef hasmulti, byRef optionsjs)
	optshtml="" : optionsjs="" : defjs="" : dependantoptions=""
	optpricediff=0 : pricediff=0 : rowcounter=0
	altids=rs("pID")
	hasmulti=FALSE
	saveoptionsjs=optionsjs
	saveoptjsunique=optjsunique
	maxindex=UBOUND(prodoptions,2)
	do while rowcounter<=maxindex
		opthasstock=FALSE
		sSQL="SELECT optID,"&getlangid("optName",32)&","&getlangid("optGrpName",16)&","&OWSP&"optPriceDiff,optType,optGrpSelect,optFlags,optTxtMaxLen,optTxtCharge,optStock,optPriceDiff AS optDims,optDefault,optAltImage,optAltLargeImage,optRegExp,"&getlangid("optPlaceholder",16)&","&IIfVs(prodoptions(5,rowcounter)=1,"'' AS ")&"optDependants,optTooltip,optClass FROM options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optGroup="&prodoptions(0,rowcounter)&" ORDER BY optID"
		if mysqlserver then rs2.CursorLocation=3
		rs2.Open sSQL,cnn,1,1
		if NOT rs2.EOF then
			opttooltip=trim(rs2("optTooltip"))
			if opttooltip<>"" then
				if left(opttooltip,2)="##" then
					opttooltip=right(opttooltip,len(opttooltip)-2)
				else
					opttooltip="&nbsp;<span class=""opttooltip"" style=""position:relative"" "&IIfVr(mobilebrowser,"onclick=""thisstyle=this.getElementsByTagName('span')[0].style;thisstyle.display=thisstyle.display=='inline'?'none':'inline'""","onmouseover=""this.getElementsByTagName('span')[0].style.display='inline'"" onmouseout=""this.getElementsByTagName('span')[0].style.display='none'""")&"><img src=""images/ectinfo.png"" alt="""" class=""ectinfo"" /><span style=""display:none;border:1px solid;background:#EEE;position:absolute;z-index:100;min-width:200px"">"&opttooltip&"</span></span>"
				end if
			end if
			if abs(int(rs2("optType")))=5 then ' Date Picker
				opthasstock=TRUE
				optclass=trim(rs2("optClass")&"")
				themask=replace(replace(replace(cstr(dateserial(2003,12,11)),"2003","yyyy"),"12","mm"),"11","dd")
				if ectcaldateformat<>"" then themask=ectcaldateformat
				if NOT hasincludedpopcalendar then mustincludepopcalendar=TRUE
				optshtml=optshtml&"<div class=""optioncontainer " & IIfVr(isdetail,"detail","prod") & "optioncontainer ectdpoption" & IIfVs(optclass<>""," "&optclass) & """"&IIfVs(prodoptions(5,rowcounter)=1," style=""display:none"" id=""divc"&Count&"x"&rowcounter&"""")&"><div class=""optiontext" & IIfVs(isdetail," detailoptiontext") & """>"&grpnmstyle&"<label for=""optn"&Count&"x"&rowcounter&""">"&rs2(getlangid("optGrpName",16))&"</label>"&opttooltip&grpnmstyleend&"</div><div class=""option" & IIfVs(isdetail," detailoption") & """> <input data-optgroup="""&prodoptions(0,rowcounter)&""" "&IIfVs(prodoptions(5,rowcounter)=1,"data-isdep=""1"" ")&"type=""hidden"" name=""optn"&rowcounter&""" value="""&rs2("optID")&""" />"
				optshtml=optshtml&"<div class=""popupcaldiv"" style=""position:relative;display:inline""><input onclick=""popUpCalendar(this,this,'"&themask&"')"" data-optgroup="""&prodoptions(0,rowcounter)&""" "&IIfVs(prodoptions(5,rowcounter)=1,"data-isdep=""1"" ")&"type=""text"" class=""prodoption"&IIfVr(isdetail," detailprodoption","") & """ maxlength=""255"" name=""voptn"&rowcounter&""" id=""optn"&Count&"x"&rowcounter&""" size="""&int(rs2("optDims"))&""" value=""" & htmldisplay(rs2(getlangid("optName",32))) & """" & IIfVs(removedefaultoptiontext," onfocus=""if(this.value=='" & jsescape(rs2(getlangid("optName",32))) & "')this.value='';""") & IIfVs(trim(rs2(getlangid("optPlaceholder",16))&"")<>""," placeholder=""" & rs2(getlangid("optPlaceholder",16)) & """") & IIfVs(prodoptions(5,rowcounter)=1," disabled=""disabled""")&" /></div>"
				optshtml=optshtml&"</div></div>"
			elseif abs(int(rs2("optType")))=3 then ' Text
				opthasstock=TRUE
				optclass=trim(rs2("optClass")&"")
				fieldHeight=cint((cdbl(rs2("optDims"))-int(rs2("optDims")))*100.0)
				optshtml=optshtml&"<div class=""optioncontainer " & IIfVr(isdetail,"detail","prod") & "optioncontainer ecttextoption" & IIfVs(optclass<>""," "&optclass) & """"&IIfVs(prodoptions(5,rowcounter)=1," style=""display:none"" id=""divc"&Count&"x"&rowcounter&"""")&"><div class=""optiontext" & IIfVs(isdetail," detailoptiontext") & """>"&grpnmstyle&"<label for=""optn"&Count&"x"&rowcounter&""">"&rs2(getlangid("optGrpName",16))&"</label>"&opttooltip&grpnmstyleend&"</div><div class=""option" & IIfVs(isdetail," detailoption") & """> <input data-optgroup="""&prodoptions(0,rowcounter)&""" "&IIfVs(prodoptions(5,rowcounter)=1,"data-isdep=""1"" ")&"type=""hidden"" name=""optn"&rowcounter&""" value="""&rs2("optID")&""" />"
				if fieldHeight<>1 then
					optshtml=optshtml&"<textarea data-optgroup="""&prodoptions(0,rowcounter)&""" "&IIfVs(prodoptions(5,rowcounter)=1,"data-isdep=""1"" ")&"class=""prodoption"&IIfVr(isdetail," detailprodoption","")&""" name=""voptn"&rowcounter&""" id=""optn"&Count&"x"&rowcounter&""" cols="""&int(rs2("optDims"))&""" rows="""&fieldHeight&"""" & IIfVs(removedefaultoptiontext," onfocus=""if(this.value=='" & jsescape(rs2(getlangid("optName",32))) & "')this.value='';""") & IIfVs(trim(rs2(getlangid("optPlaceholder",16))&"")<>""," placeholder=""" & rs2(getlangid("optPlaceholder",16)) & """") & IIfVs(prodoptions(5,rowcounter)=1," disabled=""disabled""") & ">"
					optshtml=optshtml&rs2(getlangid("optName",32))&"</textarea>"
				else
					optshtml=optshtml&"<input data-optgroup="""&prodoptions(0,rowcounter)&""" "&IIfVs(prodoptions(5,rowcounter)=1,"data-isdep=""1"" ")&"type=""text"" class=""prodoption"&IIfVr(isdetail," detailprodoption","")&""" maxlength=""255"" name=""voptn"&rowcounter&""" id=""optn"&Count&"x"&rowcounter&""" size="""&int(rs2("optDims"))&""" value=""" & htmldisplay(rs2(getlangid("optName",32))) & """" & IIfVs(removedefaultoptiontext," onfocus=""if(this.value=='" & jsescape(rs2(getlangid("optName",32))) & "')this.value='';""") & IIfVs(trim(rs2(getlangid("optPlaceholder",16))&"")<>""," placeholder=""" & rs2(getlangid("optPlaceholder",16)) & """") & IIfVs(prodoptions(5,rowcounter)=1," disabled=""disabled""")&" />"
				end if
				optshtml=optshtml&"</div></div>"
			elseif abs(int(rs2("optType")))=1 then ' Checkbox / Radio
				isdependent=prodoptions(5,rowcounter)=1
				optshtml=optshtml&"<div class=""optioncontainer " & IIfVr(isdetail,"detail","prod") & "optioncontainer ectradiooption"""&IIfVs(isdependent," style=""display:none"" id=""divc"&Count&"x"&rowcounter&"""")&">"
				if NOT noradiooptionlabel then optshtml=optshtml&"<div class=""optiontext" & IIfVs(isdetail," detailoptiontext") & """>"&grpnmstyle&rs2(getlangid("optGrpName",16))&opttooltip&grpnmstyleend&"</div>"
				optshtml=optshtml&"<div class=""option" & IIfVs(isdetail," detailoption") & """> "
				defjs=defjs & "updateoptimage("&Count&","&rowcounter&",1);"
				index=0
				do while not rs2.EOF
					optclass=trim(rs2("optClass")&"")
					if trim(rs2("optDependants")&"")<>"" then dependantoptions=dependantoptions&","&rs2("optDependants")
					origoptpricediff=getoptpricediff(rs2("optPriceDiff"), rs("pID"), trim(rs2("optRegExp")&""), rs("pPrice"), stocknotused)
					call addtoaltids(rs2("optRegExp"), altidarr, altids)
					optshtml=optshtml&"<div class=""rcoption" & IIfVs((rs2("optFlags") AND 4)=4," rcoptioninline") & IIfVs(optclass<>""," "&optclass) & """><input type="""&IIfVr(rs2.recordcount=1,"checkbox","radio")&""" data-optgroup="""&prodoptions(0,rowcounter)&""" "&IIfVs(isdependent,"disabled=""disabled"" data-isdep=""1"" ")&"class=""prodoption"&IIfVs(isdetail," detailprodoption")&""" style=""vertical-align:middle"" onclick=""updateoptimage("&Count&","&rowcounter&",1)"" name=""optn"&Count&"x"&rowcounter&""" "
					if cint(rs2("optDefault"))<>0 then optshtml=optshtml&"checked=""checked"" "
					optshtml=optshtml&"value='"&rs2("optID")&"' /><span id=""optn"&Count&"x"&rowcounter&"y"&index&""""
					if useStockManagement AND cint(rs("pStockByOpts"))<>0 AND rs2("optStock")<=0 AND trim(rs2("optRegExp")&"")="" then optshtml=optshtml&" class=""oostock""" else opthasstock=true
					optshtml=optshtml&">"&rs2(getlangid("optName",32))
					if NOT isdependent AND hideoptpricediffs<>true AND origoptpricediff<>0 AND trim(rs2("optRegExp")&"")="" then
						optshtml=optshtml&" ("
						if origoptpricediff > 0 then optshtml=optshtml&"+"
						if (rs2("optFlags") AND 1)=1 then pricediff=(rs("pPrice")*origoptpricediff)/100.0 else pricediff=origoptpricediff
						if showtaxinclusive=2 AND (rs("pExemptions") AND 2)<>2 then pricediff=pricediff+(pricediff*thetax/100.0)
						optshtml=optshtml&FormatEuroCurrency(pricediff)&")"
						if cint(rs2("optDefault"))<>0 then optpricediff=optpricediff + pricediff
					end if
					if useStockManagement AND showinstock=TRUE AND noshowoptionsinstock<>TRUE AND cint(rs("pStockByOpts"))<>0 then optshtml=optshtml&replace(xxOpSkTx, "%s", rs2("optStock"))
					optshtml=optshtml&"</span></div>" & vbLf
					index=index + 1
					call addtooptionsjs(optionsjs, isdetail, origoptpricediff)
					rs2.MoveNext
				loop
				set altidarr=nothing
				optshtml=optshtml&"</div></div>"&vbLf
			elseif abs(int(rs2("optType")))=4 then ' Multi
				if multipurchasecolumns="" then multipurchasecolumns=2
				colwid=int(100/multipurchasecolumns)
				if cint(rs2("optGrpSelect"))<>0 AND NOT isdetail then
					hasmulti=2
					optshtml=""
					optionsjs=""
					altids=rs("pID")
					defjs=""
					optionsjs=saveoptionsjs
					optjsunique=saveoptjsunique
					opthasstock=TRUE
				else
					optshtml=optshtml&"<div class=""multioptiontable"""&IIfVs(prodoptions(5,rowcounter)=1," style=""display:none"" id=""divc"&Count&"x"&rowcounter&"""")&">"
					index=0
					do while not rs2.EOF
						optclass=trim(rs2("optClass")&"")
						stocklevel=rs2("optStock")
						origoptpricediff=getoptpricediff(rs2("optPriceDiff"), rs("pID"), trim(rs2("optRegExp")&""), rs("pPrice"), stocklevel)
						call addtoaltids(rs2("optRegExp"), altidarr, altids)
						if useStockManagement AND cint(rs("pStockByOpts"))<>0 AND stocklevel<=0 AND trim(rs2("optRegExp")&"")="" AND cint(rs("pBackOrder"))=0 then oostock=TRUE else oostock=FALSE
						optshtml=optshtml&"<div class=""multioptiontext" & IIfVs(isdetail," detailmultioptiontext") & IIfVs(optclass<>""," "&optclass) & """>"
						if trim(rs2("optAlt"&IIfVr(isdetail,"Large","")&"Image")&"")<>"" then optshtml=optshtml&"&nbsp;&nbsp;<img class=""multiimage"" src="""&trim(rs2("optAlt"&IIfVr(isdetail,"Large","")&"Image"))&""" alt="""" />"
						optshtml=optshtml&"&nbsp;&nbsp;<input data-optgroup="""&prodoptions(0,rowcounter)&""" "&IIfVs(prodoptions(5,rowcounter)=1,"data-isdep=""1"" ")&"type=""text"" maxlength=""5"" name=""optm"&rs2("optID")&""" id=""optm"&Count&"x"&rowcounter&"y"&index&""" size=""1"" "&IIfVr(oostock,"style=""background-color:#EBEBE4"" disabled=""disabled""","")&"/>"
						optshtml=optshtml&"<label for=""optm"&Count&"x"&rowcounter&"y"&index&"""><span id=""optx"&Count&"x"&rowcounter&"y"&index&""" class=""multioption"
						if oostock then optshtml=optshtml&" oostock""" else optshtml=optshtml&"""" : opthasstock=true
						optshtml=optshtml&"> - " & rs2(getlangid("optName",32))
						if hideoptpricediffs<>true AND origoptpricediff<>0 then
							optshtml=optshtml&" ("
							if cdbl(origoptpricediff) > 0 then optshtml=optshtml&"+"
							if (rs2("optFlags") AND 1)=1 AND trim(rs2("optRegExp")&"")="" then pricediff=(rs("pPrice")*origoptpricediff)/100.0 else pricediff=origoptpricediff
							if showtaxinclusive=2 AND (rs("pExemptions") AND 2)<>2 then pricediff=pricediff+(pricediff*thetax/100.0)
							optshtml=optshtml&FormatEuroCurrency(pricediff)&")"
						end if
						optshtml=optshtml&"</span></label></div>"
						index=index + 1
						call addtooptionsjs(optionsjs, isdetail, origoptpricediff)
						rs2.MoveNext
					loop
					hasmulti=1
					optshtml=optshtml&"</div>"
				end if
			else ' Select
				optshtml=optshtml&"<div class=""optioncontainer " & IIfVr(isdetail,"detail","prod") & "optioncontainer ectselectoption"""&IIfVs(prodoptions(5,rowcounter)=1," style=""display:none"" id=""divc"&Count&"x"&rowcounter&"""")&">"
				optshtml=optshtml&IIfVr(NOT (noselectoptionlabel OR optionnameaspleaseselect), "<div class=""optiontext" & IIfVs(isdetail," detailoptiontext") & """>"&grpnmstyle&"<label for=""optn"&Count&"x"&rowcounter&""">"&rs2(getlangid("optGrpName",16))&"</label>"&opttooltip&grpnmstyleend & "</div><div class=""option" & IIfVs(isdetail," detailoption") & """> ", "<div class=""prodoption selectoption"&IIfVr(isdetail," detailprodoption","")&""">") & "<select data-optgroup="""&prodoptions(0,rowcounter)&""" "&IIfVs(prodoptions(5,rowcounter)=1,"data-isdep=""1"" ")&"class=""prodoption"&IIfVr(isdetail," detailprodoption","")&""" onchange=""updateoptimage("&Count&","&rowcounter&",2)"" name=""optn"&rowcounter&""" id=""optn"&Count&"x"&rowcounter&""" "&IIfVs(prodoptions(5,rowcounter)=1,"disabled=""disabled"" ")&"size=""1"">"
				defjs=defjs & "document.getElementById('optn"&Count&"x"&rowcounter&"').onchange();"
				gotdefaultdiff=FALSE
				firstpricediff=0
				origoptpricediff=rs2("optPriceDiff")
				if cint(rs2("optGrpSelect"))<>0 then
					if optionpleaseselecttemplate="" then optionpleaseselecttemplate="%s"
					optshtml=optshtml&"<option value="""">"&IIfVr(optionnameaspleaseselect,replace(optionpleaseselecttemplate,"%s",rs2(getlangid("optGrpName",16))&""),xxPlsSel)&"</option>"
				else
					if (rs2("optFlags") AND 1)=1 then firstpricediff=(rs("pPrice")*origoptpricediff)/100.0 else firstpricediff=origoptpricediff
				end if
				do while not rs2.EOF
					if trim(rs2("optDependants")&"")<>"" then dependantoptions=dependantoptions&","&rs2("optDependants")
					origoptpricediff=getoptpricediff(rs2("optPriceDiff"), rs("pID"), trim(rs2("optRegExp")&""), rs("pPrice"), stocknotused)
					call addtoaltids(rs2("optRegExp"), altidarr, altids)
					optshtml=optshtml&"<option "
					optclass=trim(rs2("optClass")&"")
					if useStockManagement AND cint(rs("pStockByOpts"))<>0 AND rs2("optStock")<=0 AND trim(rs2("optRegExp")&"")="" then optclass=trim(optclass&" oostock") else opthasstock=true
					if optclass<>"" then optshtml=optshtml&"class=""" & optclass & """ "
					optshtml=optshtml&"value="""&rs2("optID")&""""&IIfVr(cint(rs2("optDefault"))<>0," selected=""selected""","")&">"&rs2(getlangid("optName",32))
					if hideoptpricediffs<>true AND trim(rs2("optRegExp")&"")="" then
						if origoptpricediff<>0 then
							optshtml=optshtml&" ("
							if origoptpricediff > 0 then optshtml=optshtml&"+"
							if (rs2("optFlags") AND 1)=1 then pricediff=(rs("pPrice")*origoptpricediff)/100.0 else pricediff=origoptpricediff
							if showtaxinclusive=2 AND (rs("pExemptions") AND 2)<>2 then pricediff=pricediff+(pricediff*thetax/100.0)
							optshtml=optshtml&FormatEuroCurrency(pricediff)&")"
							if cint(rs2("optDefault"))<>0 then optpricediff=optpricediff + pricediff
						end if
						if cint(rs2("optDefault"))<>0 then gotdefaultdiff=TRUE
					end if
					if useStockManagement AND showinstock=TRUE AND noshowoptionsinstock<>TRUE AND cint(rs("pStockByOpts"))<>0 then optshtml=optshtml&replace(xxOpSkTx, "%s", rs2("optStock"))
					optshtml=optshtml&"</option>"&vbCrLf
					call addtooptionsjs(optionsjs, isdetail, origoptpricediff)
					rs2.MoveNext
				loop
				set altidarr=nothing
				if hideoptpricediffs<>true AND NOT gotdefaultdiff then optpricediff=optpricediff + firstpricediff
				optshtml=optshtml&"</select>" & IIfVs((noselectoptionlabel OR optionnameaspleaseselect) AND opttooltip<>"",opttooltip) & "</div>"
				optshtml=optshtml&"</div>"&vbLf
			end if
		end if
		rs2.Close
		optionshavestock=optionshavestock AND (opthasstock OR prodoptions(5,rowcounter)=1)
		dependantoptions=commaseplist(dependantoptions)
		if hasmulti=2 then exit do
		if dependantoptions<>"" then
			sSQL="SELECT optGrpID AS poOptionGroup,optType,optFlags,optTxtMaxLen,optAcceptChars,1 AS isDepOpt FROM optiongroup WHERE optGrpID IN ("&dependantoptions&")"
			rs2.Open sSQL,cnn,0,1
			if NOT rs2.EOF then
				suboptions=rs2.getrows
				depoptsarray=split(dependantoptions,",")
				for soindex=0 to UBOUND(suboptions,2)
					if int(depoptsarray(soindex))<>suboptions(0,soindex) then
						for soindex2=soindex to UBOUND(suboptions,2)
							if int(depoptsarray(soindex))=suboptions(0,soindex2) then
								for soindex3=0 to UBOUND(suboptions)
									tempval=suboptions(soindex3,soindex)
									suboptions(soindex3,soindex)=suboptions(soindex3,soindex2)
									suboptions(soindex3,soindex2)=tempval
								next
								exit for
							end if
						next
					end if
				next
				itemstomove=maxindex-rowcounter
				maxindex=maxindex+UBOUND(suboptions,2)+1
				redim preserve prodoptions(UBOUND(prodoptions),maxindex)
				for soindex=0 to itemstomove-1
					moveto=UBOUND(prodoptions,2)-soindex
					movefrom=rowcounter+itemstomove-soindex
					for soindex2=0 to UBOUND(prodoptions)
						prodoptions(soindex2,moveto)=prodoptions(soindex2,movefrom)
					next
				next
				for soindex=0 to UBOUND(suboptions,2)
					moveto=rowcounter+(UBOUND(suboptions,2)-soindex)+1
					subindex=UBOUND(suboptions,2)-soindex
					for soindex2=0 to UBOUND(prodoptions)
						prodoptions(soindex2,moveto)=suboptions(soindex2,subindex)
					next
				next
			end if
			rs2.Close
			dependantoptions=""
		end if
		rowcounter=rowcounter+1
	loop
	displayproductoptions=optshtml
	sSQL="SELECT pID,"&WSP&"pPrice,pListPrice,pInStock FROM products WHERE pID IN ('"&replace(altids, " ", "','")&"')"
	rs2.Open sSQL,cnn,0,1
	do while NOT rs2.EOF
		sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 1 ")&"imageSrc FROM productimages WHERE imageProduct='"&escape_string(rs2("pID"))&"' AND imageNumber=0 AND imageType="&IIfVr(isdetail,"1","0")&IIfVs(mysqlserver=TRUE," LIMIT 0,1")
		rs3.open sSQL,cnn,0,1
		if NOT rs3.EOF then pi=encodeimage(rs3("imageSrc")) else pi=""
		rs3.close
		if pi<>"" then
			sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 1 ")&"imageSrc FROM productimages WHERE imageProduct='"&escape_string(rs2("pID"))&"' AND imageNumber=0 AND imageType="&IIfVr(isdetail,"2","1")&IIfVs(mysqlserver=TRUE," LIMIT 0,1")
			rs3.open sSQL,cnn,0,1
			if NOT rs3.EOF then pi=pi&"*"&encodeimage(rs3("imageSrc"))
			rs3.close
		end if
		session.LCID=1033
		optionsjs=optionsjs&"sz('"&rs2("pID")&"',"&rs2("pPrice")&","&rs2("pListPrice")&",'"&jsescapel(pi)&"'"
		session.LCID=saveLCID
		if useStockManagement then optionsjs=optionsjs&","&rs2("pInStock")
		optionsjs=optionsjs&");"
		rs2.movenext
	loop
	rs2.Close
	if hasmulti<>2 then defimagejs=defimagejs&"updateprice"&Count&"();"&defjs
	if isarray(prodoptions) then
		optionsjs=optionsjs&"function setvals"&Count&"(){"&vbCrLf
		for rowcounter=0 to UBOUND(prodoptions,2)
			optionsjs=optionsjs&"optacpc["&rowcounter&"]='" & jsescape(prodoptions(4,rowcounter)) & "';optmaxc["&rowcounter&"]=" & prodoptions(3,rowcounter) & ";opttype["&rowcounter&"]=" & prodoptions(1,rowcounter) & ";optperc["&rowcounter&"]=" & IIfVr((prodoptions(2,rowcounter) AND 1)=1, "true", "false") & ";" & vbCrLf
		next
		optionsjs=optionsjs&"}"&vbCrLf
		optionsjs=optionsjs&"function updateprice"&Count&"(){"&vbCrLf
		session.LCID=1033
		optionsjs=optionsjs&"setvals"&Count&"();"
		optionsjs=optionsjs&"dependantopts("&Count&");"
		optionsjs=optionsjs&"updateprice("&(UBOUND(prodoptions,2)+1)&","&Count&","&rs("pPrice")&","&rs("pListPrice")&",'"&rs("pID")&"',"&thetax&","&IIfVr(useStockManagement AND cint(rs("pStockByOpts"))<>0, "true", "false")&","&IIfVr((rs("pExemptions") AND 2)=2, "true", "false")&","&IIfVr(cint(rs("pBackOrder"))<>0, "true", "false")&");"
		session.LCID=saveLCID
		optionsjs=optionsjs&"}"&vbCrLf
	end if
end function
sub displayformvalidator()
	optjs=optjs&"function formvalidator"&Count&"(theForm){"&vbCrLf
	if isarray(prodoptions) then
		optjs=optjs&"setvals"&Count&"();"
		optjs=optjs&"return(ectvalidate(theForm,"&(UBOUND(prodoptions,2)+1)&","&Count&","&IIfVr(useStockManagement AND cint(rs("pStockByOpts"))<>0, "true", "false")&","&IIfVr(cint(rs("pBackOrder"))<>0, "true", "false")&")"&IIfVr(customvalidator<>"","&&customvalidator(theForm)","")&");"
	else
		optjs=optjs&"return("&IIfVr(customvalidator<>"","customvalidator(theForm)","true")&");"
	end if
	optjs=optjs&"}"&vbCrLf
end sub
if enableclientlogin then
	if SESSION("clientID")<>"" then
	elseif is_numeric(getpost("checktmplogin")) then
		Set clientRS=Server.CreateObject("ADODB.RecordSet")
		Set clientCnn=Server.CreateObject("ADODB.Connection")
		clientCnn.open sDSN
		sSQL="SELECT tmploginname FROM tmplogin WHERE tmploginid='" & escape_string(request.form("sessionid")) & "' AND tmploginchk=" & replace(getpost("checktmplogin"),"'","")
		clientRS.Open sSQL,clientCnn,0,1
		if NOT clientRS.EOF then
			SESSION("clientID")=replace(clientRS("tmploginname"),"'","")
			clientRS.Close
			sSQL="SELECT clUserName,clPW,clActions,clLoginLevel,clPercentDiscount FROM customerlogin WHERE clID="&SESSION("clientID")
			clientRS.Open sSQL,clientCnn,0,1
			if NOT clientRS.EOF then
				SESSION("clientPW")=clientRS("clPW")
				SESSION("clientUser")=clientRS("clUserName")
				SESSION("clientActions")=clientRS("clActions")
				SESSION("clientLoginLevel")=clientRS("clLoginLevel")
				SESSION("clientPercentDiscount")=(100.0-cdbl(clientRS("clPercentDiscount")))/100.0
			end if
		end if
		clientRS.Close
		clientCnn.Close
		set clientRS=nothing
		set clientCnn=nothing
	elseif trim(request.cookies("WRITECLL")&"")<>"" then
		Set clientRS=Server.CreateObject("ADODB.RecordSet")
		Set clientCnn=Server.CreateObject("ADODB.Connection")
		clientCnn.open sDSN
		clientEmail=replace(Request.Cookies("WRITECLL"),"'","")
		clientPW=replace(Request.Cookies("WRITECLP"),"'","")
		sSQL="SELECT clID,clUserName,clActions,clLoginLevel,clPercentDiscount FROM customerlogin WHERE (clEmail<>'' AND clEmail='"&clientEmail&"' AND clPW='"&clientPW&"') OR (clEmail='' AND clUserName='"&clientEmail&"' AND clPW='"&clientPW&"')"
		clientRS.Open sSQL,clientCnn,0,1
		if NOT clientRS.EOF then
			SESSION("clientID")=clientRS("clID")
			SESSION("clientUser")=clientRS("clUsername")
			SESSION("clientActions")=clientRS("clActions")
			SESSION("clientLoginLevel")=clientRS("clLoginLevel")
			SESSION("clientPercentDiscount")=(100.0-cdbl(clientRS("clPercentDiscount")))/100.0
		end if
		clientRS.Close
		clientCnn.Close
		set clientRS=nothing
		set clientCnn=nothing
	end if
	if requiredloginlevel<>"" then
		if SESSION("clientLoginLevel")<requiredloginlevel then
			SESSION("clientloginref")=request.servervariables("URL") & IIfVs(request.servervariables("QUERY_STRING")<>"","?"&request.servervariables("QUERY_STRING"))
			response.redirect customeraccounturl
		end if
	end if
	if (SESSION("clientActions") AND 2)=2 then showtaxinclusive=0
end if
function urldecode(encodedstring)
	strIn =encodedstring : strOut="" : intPos=instr(strIn, "+")
	do while intPos
		strLeft="" : strRight=""
		if intPos > 1 then strLeft=Left(strIn, intPos - 1)
		if intPos < len(strIn) then strRight=mid(strIn, intPos + 1)
		strIn=strLeft & " " & strRight
		intPos=instr(strIn, "+")
		intLoop=intLoop + 1
	loop
	intPos=instr(strIn, "%")
	on error resume next
	do while intPos AND Len(strIn)-intPos >= 2
		if intPos > 1 then strOut=strOut & Left(strIn, intPos - 1)
		err.number=0
		if cint("&H" & mid(strIn, intPos + 1, 2))>=32 then strOut=strOut & chr(cint("&H" & mid(strIn, intPos + 1, 2)))
		if err.number<>0 then strOut=strOut & "%" & mid(strIn, intPos + 1, 2)
		if intPos > (len(strIn) - 3) then strIn="" else strIn=mid(strIn, intPos + 3)
		intPos=instr(strIn, "%")
	loop
	on error goto 0
	urldecode=strOut & strIn
end function
function vrmax(a,b)
	if a > b then vrmax=a else vrmax=b
end function
function vrmin(a,b)
	if a < b then vrmin=a else vrmin=b
end function
function getsessionsql()
	getsessionsql=IIfVr(SESSION("clientID")<>"", "cartClientID="&replace(SESSION("clientID"),"'",""), "cartSessionID='"&replace(thesessionid,"'","")&"'")
end function
function getordersessionsql()
	getordersessionsql="ordDate>" & vsusdate(Date()-2) & " AND "&IIfVr(SESSION("clientID")<>"","ordClientID="&replace(SESSION("clientID"),"'",""),"ordSessionID='"&replace(thesessionid,"'","")&"'")
end function
function htmldisplay(thestr)
	htmldisplay=trim(replace(replace(thestr&"",">","&gt;"),"<","&lt;"))
end function
function htmlspecials(thestr)
	htmlspecials=trim(replace(replace(replace(replace(thestr&"","&","&amp;"),">","&gt;"),"<","&lt;"),"""","&quot;"))
end function
function htmlspecialsid(thestr)
	htmlspecialsid=trim(replace(replace(replace(replace(replace(thestr&"","&",""),">",""),"<",""),"""",""),"'",""))
end function
function htmlspecialsucode(thestr)
	' htmlspecialsucode=trim(replace(replace(replace(replace(replace(replace(replace(replace(thestr&"","&","&amp;"),">","&gt;"),"<","&lt;"),"""","&quot;"),"&amp;#","&#"),"&#47;","&amp;#47;"),"&#92;","&amp;#92;"),"&#45;","&amp;#45;"))
	htmlspecialsucode=trim(replace(replace(replace(replace(thestr&"","&","&amp;"),">","&gt;"),"<","&lt;"),"""","&quot;"))
end function
function jsspecials(thestr)
	jsspecials=replace(replace(replace(replace(htmldisplay(thestr),"\","\\"),"'","\'"),vbCR,""),vbLF,"\n")
end function
function jsescape(thestr)
	jsescape=replace(replace(replace(thestr&"","\","\\"),"'","\'"),"<","")
end function
function jsescapel(thestr)
	jsescapel=replace(replace(thestr&"","\","\\"),"'","\'")
end function
sub addtomailinglist(theemail,thename)
	isspam=FALSE
	theemail=trim(lcase(strip_tags2(replace(theemail,"""",""))))
	if instr(theemail,"@")>0 AND instr(theemail, ".")>0 AND len(theemail)>5 then
		if mailchimpapikey<>"" AND mailchimplist<>"" then
			call splitname(thename,firstname,lastname)
			json_data="{""email_address"":" & json_encode(theemail)
			json_data=json_data&", ""ip_signup"":" & json_encode(REMOTE_ADDR)
			json_data=json_data&", ""status"": ""pending"",""merge_fields"":{""FNAME"":" & json_encode(firstname) & ",""LNAME"":" & json_encode(lastname) & "}}"
			xmlfnheaders=array(array("Content-Type","application/json"),array("Authorization", "Basic " & vrbase64_encrypt("anystr:"&mailchimpapikey)))
			mcpwarray=split(mailchimpapikey,"-")
			call callxmlfunction("https://"&mcpwarray(1)&".api.mailchimp.com/3.0/lists/"&mailchimplist&"/members",json_data,res,"","Msxml2.ServerXMLHTTP",errormsg,FALSE)
		else
			confirmdate=date()-365
			sSQL="SELECT email,isconfirmed,mlConfirmDate FROM mailinglist WHERE email='" & escape_string(theemail) & "'"
			rs.open sSQL,cnn,0,1
			emailexists=(NOT rs.EOF)
			if NOT rs.EOF then isconfirmed=(rs("isconfirmed")<>0) : confirmdate=rs("mlConfirmDate") else isconfirmed=FALSE
			rs.close
			emailarr=split(theemail,"@")
			if is_numeric(emailarr(0)) then isspam=TRUE
			if NOT emailexists AND NOT isspam then
				sSQL="SELECT COUNT(*) AS thecnt FROM mailinglist WHERE mlConfirmDate=" & vsusdate(date())&" AND mlIPAddress='"&left(request.servervariables("REMOTE_ADDR"), 48)&"'"
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then thecnt=cint(rs("thecnt")) else thecnt=0
				rs.close
				if thecnt < 3 then ect_query("INSERT INTO mailinglist (email,mlName,isconfirmed,mlConfirmDate,mlIPAddress) VALUES ('" & escape_string(theemail) & "','" & escape_string(thename) & "'," & IIfVr(noconfirmationemail,1,0)&"," & vsusdate(date())&",'"&left(request.servervariables("REMOTE_ADDR"), 48)&"')") else isspam=TRUE
			end if
			if NOT isconfirmed AND NOT noconfirmationemail AND NOT isspam then
				warncheckspamfolder=TRUE
				if confirmdate<>date() then
					ect_query("UPDATE mailinglist SET mlConfirmDate=" & vsusdate(date())&" WHERE email='" & escape_string(theemail) & "'")
					if htmlemails then emlNl="<br />" else emlNl=vbCrLf
					thelink=storeurl & "cart"&extension&"?emailconf="&urlencode(theemail)&"&check="&left(calcmd5(uspsUser&upsUser&origZip&emailObject&checksumtext&":"&theemail), 10)
					if htmlemails then thelink="<a href=""" & thelink & """>" & thelink & "</a>"
					call dosendemaileo(theemail,emailAddr,"",xxMLConf,xxConfEm & emlNl & emlNl & thelink,emailObject,themailhost,theuser,thepass)
				end if
			end if
		end if
	end if
end sub
function unicodehtmltojs(mistr)
	outstr=""
	thelen=len(mistr)
	theind=1
	do while theind <= thelen
		strmrk=instr(theind, mistr, "&#")
		if strmrk > 0 then
			outstr=outstr & mid(mistr, theind, strmrk - theind)
			stremrk=instr(strmrk+2, mistr, ";")
			if stremrk > 0 then
				decnum=mid(mistr, strmrk+2, (stremrk-strmrk)-2)
				if is_numeric(decnum) then
					hexstr=hex(decnum)
					outstr=outstr & "\u" & string(4-len(hexstr), "0") & hexstr
					theind=stremrk+1
				else
					outstr=outstr & mid(mistr, strmrk, stremrk-strmrk)
					theind=stremrk
				end if
			else
				outstr=outstr & mid(mistr, strmrk)
				theind=thelen+1
			end if
		else
			outstr=outstr & mid(mistr, theind)
			theind=thelen+1
		end if
	loop
	unicodehtmltojs=outstr
end function
function jscheck(thetxt)
jscheck=replace(unicodehtmltojs(thetxt),"""","\""")
end function
function imageorlink(theimg,thetext,theclass,thelink,isjs)
	isabsolute=instr(thelink,"http://")>0 OR instr(thelink,"https://")>0
	if theimg="button" then
		imageorlink="<input type=""button"" value="""&thetext&""" class="""&trim("ectbutton " & theclass)&""" onclick="""&IIfVs(NOT isjs,"ectgo" & IIfVs(NOT isabsolute,"no") & "abs('") & thelink & IIfVs(NOT isjs,"')")&""" />"
	elseif theimg<>"" then
		imageorlink="<img style=""cursor:pointer"" src="""&theimg&""" "&IIfVs(theclass<>"","class="""&theclass&""" ")&"onclick="""&IIfVs(NOT isjs,"ectgo" & IIfVs(NOT isabsolute,"no") & "abs('") & thelink & IIfVs(NOT isjs,"')")&""" alt="""&thetext&""" />"
	else
		imageorlink="<a class=""ectlink"&IIfVs(theclass<>""," "&theclass)&""" href="""&IIfVs(isjs,"#"" onclick=""")&thelink&""">"&sstrong&thetext&estrong&"</a>"
	end if
end function
function imageorbuttontag(theimg,thetext,theclass,thelink,isjs)
	isabsolute=instr(thelink,"http://")>0 OR instr(thelink,"https://")>0
	if theimg="link" then
		imageorbuttontag="<a class=""ectlink"&IIfVs(theclass<>""," "&theclass)&""" href="""&IIfVs(isjs,"#"" onclick=""return ")&thelink&""">"&sstrong&thetext&estrong&"</a>"
	elseif theimg<>"" AND theimg<>"button" then
		imageorbuttontag="<img style=""cursor:pointer"" src="""&theimg&""" "&IIfVs(theclass<>"","class="""&theclass&""" ")&"onclick="""&IIfVs(NOT isjs,"ectgo" & IIfVs(NOT isabsolute,"no") & "abs('") & thelink & IIfVs(NOT isjs,"')")&""" alt="""&thetext&""" />"
	else
		imageorbuttontag="<button type=""button"" class="""&trim("ectbutton " & theclass)&""" onclick="""&IIfVs(NOT isjs,"ectgo" & IIfVs(NOT isabsolute,"no") & "abs('") & thelink & IIfVs(NOT isjs,"')")&""">"&thetext&"</button>"
	end if
end function
function imageorbutton(theimg,thetext,theclass,thelink,isjs)
	isabsolute=instr(thelink,"http://")>0 OR instr(thelink,"https://")>0
	if theimg="link" then
		imageorbutton="<a class=""ectlink"&IIfVs(theclass<>""," "&theclass)&""" href="""&IIfVs(isjs,"#"" onclick=""return ")&thelink&""">"&sstrong&thetext&estrong&"</a>"
	elseif theimg<>"" AND theimg<>"button" then
		imageorbutton="<img style=""cursor:pointer"" src="""&theimg&""" "&IIfVs(theclass<>"","class="""&theclass&""" ")&"onclick="""&IIfVs(NOT isjs,"ectgo" & IIfVs(NOT isabsolute,"no") & "abs('") & thelink & IIfVs(NOT isjs,"')")&""" alt="""&thetext&""" />"
	else
		imageorbutton="<input type=""button"" value="""&thetext&""" class="""&trim("ectbutton " & theclass)&""" onclick="""&IIfVs(NOT isjs,"ectgo" & IIfVs(NOT isabsolute,"no") & "abs('") & thelink & IIfVs(NOT isjs,"')")&""" />"
	end if
end function
function imageorsubmit(theimg,thetext,theclass)
	if theimg<>"" AND theimg<>"button" then
		imageorsubmit="<input type=""image"" src="""&theimg&""" alt="""&thetext&""" "&IIfVs(theclass<>"","class="""&theclass&""" ")&"/>"
	else
		imageorsubmit="<input type=""submit"" value="""&thetext&""" class="""&trim("ectbutton " & theclass)&""" />"
	end if
end function
sub dosendemail(seTo,seFrom,seSubject,seBody)
	call dosendemaileo(seTo,seFrom,"",seSubject,seBody,"","","","")
end Sub
sendemailerrnum=0
sendemailerrdesc=""
sub dosendemaileo(seTo,byval seFrom,seReplyTo,seSubject,seBody,emailObject,emailhost,username,password)
	set rsSE=Server.CreateObject("ADODB.RecordSet")
	sSQL="SELECT emailfromname,emailObject,smtpserver,emailUser,emailPass,smtpport,smtpsecure FROM admin WHERE adminID=1"
	rsSE.open sSQL,cnn,0,1
	emailfromname=rsSE("emailfromname")
	emailObject=rsSE("emailObject")
	emailhost=trim(rsSE("smtpserver")&"")
	smtpport=trim(rsSE("smtpport")&"")
	smtpsecure=trim(rsSE("smtpsecure")&"")
	username=trim(rsSE("emailUser")&"")
	password=trim(rsSE("emailPass")&"")
	rsSE.close
	set rsSE=nothing
	sendemailerrnum=0
	sendemailerrdesc=""
	seReplyTo=trim(seReplyTo)
	if instr(seFrom,";")>0 then
		sendersarr=split(seFrom,";")
		seFrom=sendersarr(0)
	end if
	if debugmode<>TRUE then on error resume next
	err.number=0
	if emailObject=0 then
		set EmailObj=Server.CreateObject("CDONTS.NewMail")
		EmailObj.MailFormat=0
		if htmlemails then EmailObj.BodyFormat=0
		EmailObj.To=seTo
		if emailfromname<>"" then EmailObj.From=emailfromname & " <" & seFrom & ">" else EmailObj.From=seFrom
		if seReplyTo<>"" then EmailObj.Value("Reply-To")=seReplyTo
		EmailObj.Subject=seSubject
		EmailObj.Body=seBody
		EmailObj.Send
	elseif emailObject=1 then
		set EmailObj=Server.CreateObject("CDO.Message")
		set iConf=CreateObject("CDO.Configuration")
		scnfg="http://schemas.microsoft.com/cdo/configuration/"
		if NOT (emailhost="your.mailserver.com" OR emailhost="") then
			set Flds=iConf.Fields
			Flds.Item(scnfg&"sendusing")=2
			Flds.Item(scnfg&"smtpserver")=emailhost
			if smtpport<>"" then Flds.Item(scnfg&"smtpserverport")=smtpport
			if smtpsecure="ssl" then Flds.Item(scnfg&"smtpusessl")=TRUE
			if username<>"" AND password<>"" then
				Flds.Item(scnfg&"smtpauthenticate")=1
				Flds.Item(scnfg&"sendusername")=username
				Flds.Item(scnfg&"sendpassword")=password
			end if
			Flds.Update
			EmailObj.Configuration=iConf
			set Flds=nothing
		else
			set Flds=iConf.Fields
			if username<>"" AND password<>"" then
				Flds.Item(scnfg&"smtpauthenticate")=1
				Flds.Item(scnfg&"sendusername")=username
				Flds.Item(scnfg&"sendpassword")=password
			end if
			Flds.Update
			EmailObj.Configuration=iConf
			set Flds=nothing
		end if
		if emailfromname<>"" then EmailObj.From=Chr(34) & emailfromname & Chr(34) & Chr(60) & seFrom & Chr(62) else EmailObj.From=seFrom
		if seReplyTo<>"" then EmailObj.ReplyTo=seReplyTo else EmailObj.ReplyTo=seFrom
		EmailObj.Subject=seSubject
		EmailObj.Fields.Update
		if htmlemails then
			EmailObj.HTMLBody=seBody
			EmailObj.HTMLBodyPart.Charset=emailencoding
			EmailObj.TextBodyPart.Charset=emailencoding
			EmailObj.BodyPart.Charset=emailencoding
		else
			EmailObj.TextBody=seBody
			EmailObj.TextBodyPart.Charset=emailencoding
			EmailObj.BodyPart.Charset=emailencoding
		end if
		EmailObj.To=seTo
		EmailObj.Send
		set iConf=nothing
	elseif emailObject=2 then
		set EmailObj=Server.CreateObject("Persits.MailSender")
		if username<>"" AND password<>"" then
			EmailObj.Username=username
			EmailObj.Password=password
		end if
		EmailObj.Host=emailhost
		if htmlemails then EmailObj.IsHTML=TRUE
		if smtpsecure="tls" then EmailObj.TLS=TRUE
		EmailObj.AddAddress seTo
		EmailObj.From=seFrom
		EmailObj.FromName=IIfVr(emailfromname<>"",emailfromname,seFrom)
		if seReplyTo<>"" then
			EmailObj.AddReplyTo seReplyTo,seReplyTo
		end if
		EmailObj.Subject=seSubject
		if emailencoding<> "iso-8859-1" then
			EmailObj.Charset=emailencoding
		end if
		EmailObj.Body=seBody
		if emailencoding<> "iso-8859-1" then
			EmailObj.ContentTransferEncoding="Quoted-Printable"
		end if
		EmailObj.Send
	elseif emailObject=3 then
		set EmailObj=Server.CreateObject("SMTPsvg.Mailer")
		if htmlemails then EmailObj.ContentType="text/html"
		EmailObj.RemoteHost=emailhost
		EmailObj.AddRecipient seTo, seTo
		EmailObj.FromAddress=seFrom
		if emailfromname<>"" then EmailObj.FromName=emailfromname
		if seReplyTo<>"" then EmailObj.ReplyTo=seReplyTo
		EmailObj.Subject=seSubject
		EmailObj.BodyText=seBody
		EmailObj.SendMail
	elseif emailObject=4 then
		set EmailObj=Server.CreateObject("JMail.SMTPMail")
		if htmlemails then EmailObj.ContentType="text/html"
		EmailObj.silent=true
		EmailObj.Logging=true
		EmailObj.ServerAddress=emailhost
		EmailObj.AddRecipient seTo
		EmailObj.Sender=seFrom
		if emailfromname<>"" then EmailObj.SenderName=emailfromname
		if seReplyTo<>"" then EmailObj.ReplyTo=seReplyTo
		EmailObj.Subject=seSubject
		EmailObj.Body=seBody
		EmailObj.Execute
	elseif emailObject=5 then
		set EmailObj=Server.CreateObject("SoftArtisans.SMTPMail")
		if username<>"" AND password<>"" then
			EmailObj.UserName=username
			EmailObj.Password=password
		end if
		if htmlemails then EmailObj.ContentType="text/html"
		EmailObj.RemoteHost=emailhost
		EmailObj.AddRecipient seTo , seTo
		EmailObj.FromAddress=seFrom
		if emailfromname<>"" then EmailObj.FromName=emailfromname
		if seReplyTo<>"" then EmailObj.ReplyTo=seReplyTo
		EmailObj.Subject=seSubject
		EmailObj.BodyText=seBody
		if NOT EmailObj.SendMail then print "<br /> " & EmailObj.Response
	elseif emailObject=6 then
		set EmailObj=Server.CreateObject("JMail.Message")
		if htmlemails then EmailObj.ContentType="text/html"
		EmailObj.silent=true
		EmailObj.Logging=true
		EmailObj.AddRecipient seTo
		EmailObj.From=seFrom
		if emailfromname<>"" then EmailObj.FromName=emailfromname
		if seReplyTo<>"" then EmailObj.ReplyTo=seReplyTo
		EmailObj.Subject=seSubject
		if htmlemails then EmailObj.HTMLBody=seBody else EmailObj.Body=seBody
		EmailObj.Send(emailhost)
	end if
	sendemailerrnum=err.number
	sendemailerrdesc=err.description
	set EmailObj=nothing
	on error goto 0
end sub
function getgcchar()
	getgcchar=""
	do while getgcchar="" OR getgcchar="O" OR getgcchar="I" OR getgcchar="Q"
		getgcchar=chr(int(26 * Rnd) + 65)
	loop
end function
function getrndchar()
	num=int(36 * Rnd) + 48
	if num>57 then num=num+39
	getrndchar=chr(num)
end function
function substr_count(substr, tsea)
	startpos=1
	substr_count=0
	do while startpos>0
		startpos=instr(startpos,substr,tsea)
		if startpos>0 then substr_count=substr_count+1 : startpos=startpos+1
	loop
end function
function replaceemailtxt(thestr, txtsearch, txtreplace, byref didreplace)
	inbrackets=FALSE
	countinscope=1
	i=instr(thestr, txtsearch)
	didreplace=(i > 0)
	if i > 0 then
		t1=i
		bcount=0
		do while t1>0
			if bcount=0 AND mid(thestr, t1, 1)="{" then exit do
			if mid(thestr, t1, 1)="{" then bcount=bcount+1
			if mid(thestr, t1, 1)="}" then bcount=bcount-1
			t1=t1-1
		loop
		t4=i
		bcount=0
		do while t4<=len(thestr) AND (bcount<>0 OR mid(thestr, t4, 1)<>"}")
			if mid(thestr, t4, 1)="{" then bcount=bcount+1
			if mid(thestr, t4, 1)="}" then bcount=bcount-1
			t4=t4+1
		loop
		if t4>len(thestr) then t4=0
		inbrackets=(t1 > 0 AND t4 > 0)
	end if
	if i=0 then
		replaceemailtxt=thestr
	elseif trim(txtreplace&"")="" then ' want to replace all of txtsearch OR {...txtsearch...}
		if inbrackets then replaceemailtxt=mid(thestr, 1, t1-1) & mid(thestr, t4+1) else replaceemailtxt=replace(thestr, txtsearch, "")
	else ' Want to remove the { and }
		if txtreplace="%ectpreserve%" then txtreplace=""
		if inbrackets then thestr=mid(thestr, 1, t1-1) & mid(thestr, t1+1, (t4-t1)-1) & mid(thestr, t4+1)
		if (txtsearch="%trackingnum%" AND inbrackets) OR left(txtsearch,9)="%statusid" then
			if txtsearch="%trackingnum%" then countinscope=substr_count(mid(thestr, t1, (t4-t1)-1), "%trackingnum%")
			replaceemailtxt=replace(thestr, txtsearch, txtreplace, 1, countinscope, 0)
		else
			replaceemailtxt=replace(thestr, txtsearch, txtreplace)
		end if
	end if
end function
function showproductreviews(disptype, classname)
	showproductreviews="<div class=""smallreviewstars "&classname&"""><a href="""&thedetailslink&"#reviews"">"
	if imgreviewcart="" then call displayreviewicons()
	therating=cint(rs("pTotRating")/rs("pNumRatings"))
	for index=1 to int(therating / 2)
		if imgreviewcart<>"" then
			showproductreviews=showproductreviews & "<img class="""&classname&""" src=""images/s"&imgreviewcart&""" alt="""" />"
		else
			showproductreviews=showproductreviews & "<svg viewBox=""0 0 24 24"" class=""icon"" style=""max-width:30px""><use xlink:href=""#review-icon-full""></use></svg>"
		end if
	next
	ratingover=therating
	if ratingover / 2 > int(ratingover / 2) then
		if imgreviewcart<>"" then
			showproductreviews=showproductreviews & "<img class="""&classname&""" src=""images/s"&replace(imgreviewcart,".","hg.")&""" alt="""" />"
		else
			showproductreviews=showproductreviews & "<svg viewBox=""0 0 24 24"" class=""icon"" style=""max-width:30px""><use xlink:href=""#review-icon-half""></use></svg>"
		end if
		ratingover=ratingover + 1
	end if
	for index=int(ratingover / 2) + 1 to 5
		if imgreviewcart<>"" then
			showproductreviews=showproductreviews & "<img class="""&classname&""" src=""images/s"&replace(imgreviewcart,".","g.")&""" alt="""" />"
		else
			showproductreviews=showproductreviews & "<svg viewBox=""0 0 24 24"" class=""icon"" style=""max-width:30px""><use xlink:href=""#review-icon-empty""></use></svg>"
		end if
	next
	showproductreviews=showproductreviews & "</a><span class=""prodratingtext"">"
	if disptype=2 then showproductreviews=showproductreviews & " <a class=""ectlink prodratinglink"" href="""&thedetailslink&"#reviews"">" & replace(xxBasRat, "%s", rs("pNumRatings")) & "</a>" else if disptype=1 then showproductreviews=showproductreviews & " " & replace(xxBasRat, "%s", rs("pNumRatings")) & " (<a class=""ectlink prodratinglink"" href="""&thedetailslink&"#reviews"">" & xxView & "</a>)"
	showproductreviews=showproductreviews & "</span></div>"
end function
sub splitfirstlastname(thename, byref firstfull, byref lastname)
	if usefirstlastname AND instr(thename, " ")>0 then
		namearr=split(thename," ",2)
		firstfull=namearr(0)
		lastname=namearr(1)
	else
		firstfull=thename
		lastname=""
	end if
end sub
function getcatid(sid,snam,seopattern)
	getcatid=sid
	if seocategoryurls then
		getcatid=IIfVr(instr(snam,"://")>0,snam,replace(seopattern,"%s",rawurlencode(replace(snam," ",detlinkspacechar))))
	elseif usecategoryname AND snam<>"" then
		getcatid=urlencode(snam)
	end if
end function
function cleanupemail(theemail)
	cleanupemail="" : gotat=FALSE
	if len(theemail)<50 then
		theemail=strip_tags2(replace(theemail,"""",""))
		for ixe=1 to len(theemail)
			ch=mid(theemail,ixe,1)
			if ch<>" " AND ch<>"""" AND ch<>"'" AND ch<>"(" AND ch<>")" AND NOT (ch="@" AND gotat) then cleanupemail=cleanupemail & ch
			if ch="@" then gotat=TRUE
		next
		if NOT gotat then cleanupemail=""
	end if
end function
function parse_url(surl,comp)
	parse_url=replace(lcase(surl), "http://", "")
	parse_url=replace(parse_url, "https://", "")
	surlarr=split(parse_url, "?", 2)
	if UBOUND(surlarr) >= 0 then parse_url=surlarr(0)
	if comp=2 then surlarr=split(parse_url, "/", 2) : if UBOUND(surlarr) >= 0 then parse_url=surlarr(0)
end function
function getfullurl(pagepart)
	getfullurl=""
	if left(pagepart,5)<>"http:" AND left(pagepart,6)<>"https:" then
		getfullurl="http"&IIfVs(request.servervariables("HTTPS")="on" OR left(lcase(storeurl),6)="https:","s")&"://"&request.servervariables("HTTP_HOST")
		if left(pagepart,1)<>"/" then getfullurl=getfullurl&IIfVr(instrrev(request.servervariables("URL"),"/")>1,left(request.servervariables("URL"),instrrev(request.servervariables("URL"),"/")),"/")
	end if
	getfullurl=getfullurl&pagepart
end function
sub get_wholesaleprice_sql()
	if SESSION("clientUser")<>"" then
		if (SESSION("clientActions") AND 8)=8 then
			WSP="pWholesalePrice AS "
			TWSP="pWholesalePrice"
			if wholesaleoptionpricediff=TRUE then OWSP="optWholesalePriceDiff AS "
			if nowholesalediscounts=true then nodiscounts=true
		end if
		if (SESSION("clientActions") AND 16)=16 then
			Session.LCID=1033
			WSP=SESSION("clientPercentDiscount") & "*"&IIfVr((SESSION("clientActions") AND 8)=8,"pWholesalePrice","pPrice")&" AS "
			TWSP=SESSION("clientPercentDiscount") & "*"&IIfVr((SESSION("clientActions") AND 8)=8,"pWholesalePrice","pPrice")
			OWSP=SESSION("clientPercentDiscount") & "*"&IIfVr(((SESSION("clientActions") AND 8)=8) AND wholesaleoptionpricediff,"optWholesalePriceDiff","optPriceDiff")&" AS "
			if nowholesalediscounts=true then nodiscounts=true
			Session.LCID=saveLCID
		end if
	end if
end sub
function writepagebar(CurPage,iNumPages,sprev,snext,sLink,nofirstpage)
	Dim i, sStr, startPage, endPage
	startPage=vrmax(1,int(cdbl(CurPage)/10.0)*10)
	endPage=vrmin(iNumPages,(int(cdbl(CurPage)/10.0)*10)+10)
	if CurPage > 1 then
		sStr=sLink & "1""><span class=""pagebarquo pagebarlquo"">&laquo;</span></a> " & sLink & CurPage-1 & """"&IIfVs(CurPage>2," rel=""prev""")&"><span class=""pagebarprevnext pagebarprev"">"&sprev&"</span></a><span class=""pagebarsep""></span>"
	else
		sStr="<span class=""pagebarquo pagebarlquo"">&laquo;</span> <span class=""pagebarprevnext pagebarprev"">"&sprev&"</span><span class=""pagebarsep""></span>"
	end if
	for i=startPage to endPage
		if i=CurPage then
			sStr=sStr & "<span class=""pagebarnum currpage"">" & i & "</span><span class=""pagebarsep""></span>"
		else
			sStr=sStr & sLink & i & """><span class=""pagebarnum"">"
			if i=startPage AND i > 1 then sStr=sStr&"..."
			sStr=sStr & i
			if i=endPage AND i < iNumPages then sStr=sStr&"..."
			sStr=sStr & "</span></a><span class=""pagebarsep""></span>"
		end if
	next
	if CurPage < iNumPages then
		writepagebar=sStr & sLink & CurPage+1 & """ rel=""next""><span class=""pagebarprevnext pagebarnext"">"&snext&"</span></a> " & sLink & iNumPages & """><span class=""pagebarquo pagebarrquo"">&raquo;</span></a>"
	else
		writepagebar=sStr & " <span class=""pagebarprevnext pagebarnext"">"&snext&"</span> <span class=""pagebarquo pagebarrquo"">&raquo;</span>"
	end if
	if nofirstpage then writepagebar=replace(replace(writepagebar,"&amp;pg=1""",""" rel=""start"""),"?pg=1""",""" rel=""start""")
end function
function addtag(tagname, strValue)
	addtag="<" & tagname & ">" & Replace(Replace(strValue&"", "&", "&amp;"), "<", "&lt;") & "</" & tagname & ">"
end function
function escape_string(str)
	escape_string=trim(replace(str&"","'","''"))
	if mysqlserver=TRUE then escape_string=replace(escape_string,"\","\\")
end function
function urlencode(str)
	urlencode=server.urlencode(trim(str&""))
	urlencode=replace(urlencode,"%5F","_")
	urlencode=replace(urlencode,"%2D","-")
	urlencode=replace(urlencode,"%2E",".")
	urlencode=replace(urlencode,"%7E","~")
end function
function rawurlencode(str)
	rawurlencode=replace(urlencode(str),"+","%20")
end function
function getsessionid()
	if is_numeric(persistentcart) then
		if int(persistentcart)<=0 AND nopadsscompliance<>TRUE then
			getsessionid=Session.SessionID
		elseif request.cookies("ectcartcookie")<>"" then
			getsessionid=replace(request.cookies("ectcartcookie"),"'","")
		else
			if persistentcart="" then persistentcart=3
			gotunique=FALSE
			randomize
			do while NOT gotunique
				sequence=""
				for indgsid=0 to 25
					sequence=sequence&getrndchar()
				next
				sSQL="SELECT cartSessionID FROM cart WHERE cartSessionID='" & sequence & "'"
				rs.open sSQL,cnn,0,1
				if rs.EOF then gotunique=TRUE
				rs.close
			loop
			call setacookie("ectcartcookie", sequence, persistentcart)
			getsessionid=sequence
		end if
	else
		getsessionid=Session.SessionID
	end if
end function
function dohashpw(thepw)
	if trim(thepw&"")="" then dohashpw="" else dohashpw=calcmd5("ECT IS BEST"&trim(thepw))
end function
sub logevent(userid,eventtype,eventsuccess,eventorigin,areaaffected)
	if nopadsscompliance<>TRUE then
		sSQL="SELECT logID FROM auditlog WHERE eventType='STARTLOG'"
		rs.open sSQL,cnn,0,1
		if rs.EOF then ect_query("INSERT INTO auditlog (userID,eventType,eventDate,eventSuccess,eventOrigin,areaAffected) VALUES (" & _
			"'" & escape_string(left(userid,48)) & "','STARTLOG'," & vsusdatetime(now) & ",1," & _
			"'" & escape_string(left(eventorigin,48)) & "','" & escape_string(left(areaaffected,48)) & "')")
		rs.close
		sSQL="INSERT INTO auditlog (userID,eventType,eventDate,eventSuccess,eventOrigin,areaAffected) VALUES (" & _
			"'" & escape_string(left(userid,48)) & "','" & escape_string(left(eventtype,48)) & "'," & _
			vsusdatetime(now) & "," & IIfVr(eventsuccess,1,0) & "," & _
			"'" & escape_string(left(eventorigin,48)) & "','" & escape_string(left(areaaffected,48)) & "')"
		ect_query(sSQL)
		ect_query("DELETE FROM auditlog WHERE eventDate<" & vsusdatetime(date()-365))
	end if
end sub
function is_numeric(tstr)
	is_numeric=isnumeric(trim(tstr&"")) AND instr(trim(tstr&""),",")=0
end function
sub splitname(thename, byRef firstname, byRef lastname)
	if InStr(thename," ") > 0 then
		namearr=Split(thename," ",2)
		firstname=namearr(0)
		lastname=namearr(1)
	else
		firstname=""
		lastname=thename
	end if
end sub
function detectmobilebrowser()
	uagent=lcase(Request.ServerVariables("HTTP_USER_AGENT"))
	detectmobilebrowser=FALSE
	if instr(uagent,"android")>0 OR instr(uagent,"blackberry")>0 OR instr(uagent,"iemobile")>0 OR instr(uagent,"iphone")>0 OR instr(uagent,"mobile")>0 OR instr(uagent,"nokia")>0 OR instr(uagent,"opera mini")>0 OR instr(uagent,"pocketpc")>0 OR instr(uagent,"samsung")>0 OR instr(uagent,"symbian")>0 OR instr(uagent,"smartphone")>0 then detectmobilebrowser=TRUE
end function
function getget(tval)
	getget=trim(request.querystring(tval))
end function
function escapeget(tval)
	escapeget=escape_string(trim(request.querystring(tval)))
end function
function getpost(tval)
	getpost=trim(request.form(tval))
end function
function getpostint(tval)
	getpostint=trim(request.form(tval))
	if NOT is_numeric(getpostint) then getpostint=0 else getpostint=int(getpostint)
	if getpostint<-2147483648 then getpostint=-2147483648
	if getpostint>2147483647 then getpostint=2147483647
end function
function escapepost(tval)
	escapepost=escape_string(trim(request.form(tval)))
end function
function getrequest(tval)
	getrequest=trim(request(tval))
end function
function ect_query(ectsql)
	application("postvars")=""
	for each sObj in request.form
		if request.form(sObj)<>"" then
			application("postvars") = application("postvars") & sObj & " : " & request.form(sObj) & vbCrLf
		end if
	next
	application("sqlstatement")=ectsql
	if debugmode then on error resume next
	ect_query=cnn.execute(ectsql)
	if debugmode then
		if err.number<>0 then print "Database Error: " & err.description & "<br />" & ectsql & "<br />" : response.end
	end if
	application("sqlstatement")=""
end function
function getfirstchild(tval)
	if tval.hasChildNodes then getfirstchild=tval.firstChild.nodeValue else getfirstchild=""
end function
function jsurlencode(tstr)
	jsurlencode=replace(urlencode(tstr),"+","%20")
end function
function getdetailsurl(gdid,gdstatic,byval gdname,gdurl,gdqs,gdpathtohere)
	gdname=IIfVr(gdurl<>"",gdurl,gdname&"")
	if seodetailurls then
		getdetailsurl=IIfVr(instr(gdname,"://")>0,gdname,gdpathtohere&rawurlencode(replace(gdname," ",detlinkspacechar))&IIfVs(gdqs<>"","?"&gdqs))
	elseif gdurl<>"" then
		getdetailsurl=gdurl&IIfVs(gdqs<>"","?"&gdqs)
	elseif cint(gdstatic)<>0 then
		getdetailsurl=cleanforurl(gdname)&extension&IIfVs(gdqs<>"","?"&gdqs)
	else
		getdetailsurl="proddetail"&extension&"?prod="&urlencode(IIfVr(usepnamefordetaillinks,replace(gdname," ",detlinkspacechar),gdid))&IIfVs(gdqs<>"","&amp;"&gdqs)
	end if
end function
function commaseplist(inlist)
	set toregexp=new regexp
	toregexp.global=TRUE
	toregexp.pattern="[^0-9,]"
	commaseplist=toregexp.replace(inlist&"","")
	toregexp.pattern=",+"
	commaseplist=toregexp.replace(commaseplist,",")
	toregexp.pattern=",$|^,"
	commaseplist=toregexp.replace(commaseplist,"")
	set toregexp=nothing
end function
sub getperproductdiscounts()
	if noshowdiscounts<>TRUE AND NOT hascheckedperproductdiscounts AND (rs("pExemptions") AND 64)<>64 then
		tdt=Date()
		isglobaldiscount=FALSE : localdiscounts="," : lastcouponname="" : attributelist="" : hascheckedperproductdiscounts=TRUE
		sSQL="SELECT mSCscID FROM multisearchcriteria WHERE mSCpID='"&escape_string(rs("pID"))&"'"
		rs2.open sSQL,cnn,0,1
		do while NOT rs2.EOF
			attributelist=attributelist&rs2("mSCscID")&" "
			rs2.movenext
		loop
		rs2.close
		attributelist=replace(trim(attributelist)," ","','")
		sSQL="SELECT DISTINCT cpnID,"&getlangid("cpnName",1024)&" FROM coupons LEFT OUTER JOIN cpnassign ON coupons.cpnID=cpnassign.cpaCpnID WHERE (cpnSitewide=0 OR cpnSitewide=3) AND cpnNumAvail>0 AND cpnStartDate<=" & vsusdate(tdt)&" AND cpnEndDate>=" & vsusdate(tdt)&" AND cpnIsCoupon=0 AND ((cpaType=2 AND cpaAssignment='"&escape_string(rs("pID"))&"')"
		if NOT isrootsection then sSQL=sSQL & " OR (cpaType=1 AND cpaAssignment IN ('"&replace(topcpnids,",","','")&"') AND NOT cpaAssignment IN ('"&Replace(topsectionids,",","','")&"'))"
		if attributelist<>"" then sSQL=sSQL & " OR (cpaType=3 AND cpaAssignment IN ('"&attributelist&"'))"
		sSQL=sSQL & ") AND ((cpnLoginLevel>=0 AND cpnLoginLevel<="&minloglevel&") OR (cpnLoginLevel<0 AND -1-cpnLoginLevel="&minloglevel&")) ORDER BY "&getlangid("cpnName",1024)
		rs2.Open sSQL,cnn,0,1
		do while NOT rs2.EOF
			for index=0 to maxglobaldiscounts-1
				if globaldiscounts(0,index)=rs2("cpnID") then isglobaldiscount=TRUE : localdiscounts=localdiscounts&rs2("cpnID")&","
			next
			if NOT isglobaldiscount AND rs2(getlangid("cpnName",1024))<>lastcouponname then alldiscounts=alldiscounts & "<div>" & rs2(getlangid("cpnName",1024)) & "</div>" : lastcouponname=rs2(getlangid("cpnName",1024))
			rs2.movenext
		loop
		rs2.close
		if catid<>"0" AND topsectionids<>"" then
			if instr(","&topsectionids&",",","&rs("pSection")&",")=0 then
				for index=0 to maxglobaldiscounts-1
					if instr(localdiscounts,","&globaldiscounts(0,index)&",")=0 then
						if globaldiscounts(2,index)="xxx" then
							rowcounter=0
							otherassignments=""
							sSQL="SELECT cpaAssignment FROM cpnassign WHERE cpaCpnID="&globaldiscounts(0,index)&" AND cpaType=1"
							rs2.Open sSQL,cnn,0,1
							do while NOT rs2.EOF
								otherassignments=otherassignments&rs2("cpaAssignment")&","
								rowcounter=rowcounter+1
								rs2.movenext
							loop
							rs2.close
							if rowcounter>1 then otherassignments=","&getsectionids(left(otherassignments,len(otherassignments)-1),FALSE)&"," else otherassignments=""
							globaldiscounts(2,index)=otherassignments
						end if
						if instr(globaldiscounts(2,index),","&rs("pSection")&",")=0 then noapplydiscounts="<div>" & noapplydiscounts&globaldiscounts(1,index) & "</div>"
					end if
				next
			end if
		end if
	end if
end sub
function twodp(tnum)
    if len(cstr(tnum))<2 then twodp="0"&tnum else twodp=tnum
end function
function iso8601date(tdate)
	iso8601date=DatePart("yyyy",tdate)&"-"&twodp(DatePart("m",tdate))&"-"&twodp(DatePart("d",tdate))
end function
sub displaybmlbanner(pubid,bannerdims) %>
<div class="billmelaterbanner" style="text-align:center"><script data-pp-pubid="<%=pubid%>" data-pp-placementtype="<%=bannerdims%>"> (function (d, t) {
"use strict";
var s=d.getElementsByTagName(t)[0], n=d.createElement(t);
n.src="//paypal.adtag.where.com/merchant.js";
s.parentNode.insertBefore(n, s);
}(document, "script"));
</script></div><%
end sub
function labeltxt(lbltxt,lblid)
	labeltxt="<label class=""ectlabel"" for="""&lblid&""">"&lbltxt&"</label>"
end function
function quantitymarkup(isquickbuy,Count,isdetail,placeholder,addonchange)
if quantityupdown=2 then
quantitymarkup="<div class="""&placeholder&"quantity2div" & IIfVr(isdetail," detail"," prod") & "quantity2div""><div onclick=""quantup('"&IIfVs(isquickbuy,"qb")&Count&"',0)"">-</div><input type=""text"""&IIfVs(NOT isquickbuy," name=""quant""")&" id=""w"&IIfVs(isquickbuy,"qb")&Count&"quant"" maxlength=""5"" value=""1"" title="""&xxQuant&""" " & IIfVs(addonchange,"onchange=""document.getElementById('qnt"&Count&"x').value=this.value"" ") & "class=""quantity2input" & IIfVr(isdetail," detail"," prod") & "quantity2input"" /><div onclick=""quantup('"&IIfVs(isquickbuy,"qb")&Count&"',1)"">+</div></div>"&vbLf
elseif quantityupdown then
quantitymarkup="<div class="""&placeholder&"quantity1div" & IIfVr(isdetail," detail"," prod") & "quantity1div""><input type=""text"""&IIfVs(NOT isquickbuy," name=""quant""")&" id=""w"&IIfVs(isquickbuy,"qb")&Count&"quant"" size=""2"" maxlength=""5"" value=""1"" title="""&xxQuant&""" " & IIfVs(addonchange,"onchange=""document.getElementById('qnt"&Count&"x').value=this.value"" ") & "class=""quantity1input" & IIfVr(isdetail," detail"," prod") & "quantity1input""><div onclick=""quantup('"&IIfVs(isquickbuy,"qb")&Count&"',1)"">+</div><div onclick=""quantup('"&IIfVs(isquickbuy,"qb")&Count&"',0)"">-</div></div>"&vbLf
else
quantitymarkup="<div class="""&placeholder&"quantity0div" & IIfVr(isdetail," detail"," prod") & "quantity0div""><input type=""text"""&IIfVs(NOT isquickbuy," name=""quant""")&" id=""w"&IIfVs(isquickbuy,"qb")&Count&"quant"" size=""2"" maxlength=""5"" value=""1"" title="""&xxQuant&""" " & IIfVs(addonchange,"onchange=""document.getElementById('qnt"&Count&"x').value=this.value"" ") & "class=""quantity0input" & IIfVr(isdetail," detail"," prod") & "quantity0input""></div>"&vbLf
end if
end function
sub displaydiscountexemptions()
	if (rs("pExemptions") AND 80)=80 AND (hasshippingdiscount OR hasproductdiscount) then
		if xxNoDisc<>"" then print "<div class=""ectwarning proddiscountexempt"">"&xxNoDisc&"</div>"
	else
		if (rs("pExemptions") AND 16)=16 AND hasshippingdiscount AND xxNoFrSh<>"" then print "<div class=""ectwarning freeshippingexempt"">"&xxNoFrSh&"</div>"
		if (rs("pExemptions") AND 64)=64 AND hasproductdiscount AND xxNoDsPr<>"" then print "<div class=""ectwarning proddiscountexempt"">"&xxNoDsPr&"</div>"
	end if
end sub
sub displaylayoutarray(layoutarray,isquickbuy)
	dlasavecs=cs
	if isquickbuy then cs=""
	for each layoutoption in layoutarray
		layoutoption=lcase(trim(layoutoption))
		if layoutoption="minquantity" then
			if rs("pMinQuant")>0 then print "<div class=""prodminquant"">" & replace(xxMinQua,"%quant%",rs("pMinQuant")+1) & "</div>"
		elseif layoutoption="productid" then
			if showproductid OR hascustomlayout then print IIfVs(NOT usecsslayout, "<tr><td>") & "<div class="""&cs&"prodid"">" & IIfVs(xxPrId<>"","<span class="""&cs&"prodidlabel"">" & xxPrId & "</span> ") & rs("pID") & "</div>" & IIfVs(NOT usecsslayout, "</td></tr>") & vbLf
		elseif layoutoption="manufacturer" then
			if NOT IsNull(rs(getlangid("scName",131072))) then print IIfVs(NOT usecsslayout, "<tr><td>") & "<div class="""&cs&"prodmanufacturer"">" & IIfVs(xxManLab<>"","<span class="""&cs&"prodmanufacturerlabel"">" & xxManLab & "</span> ") & rs(getlangid("scName",131072)) & "</div>" & IIfVs(NOT usecsslayout, "</td></tr>") & vbLf
		elseif layoutoption="sku" then
			if (showproductsku<>"" OR hascustomlayout) AND trim(rs("pSKU")&"")<>"" then print IIfVs(NOT usecsslayout, "<tr><td>") & "<div class="""&cs&"prodsku"">" & IIfVs(showproductsku<>"","<span class="""&cs&"prodskulabel"">" & showproductsku & "</span> ") & rs("pSKU") & "</div>" & IIfVs(NOT usecsslayout, "</td></tr>") & vbLf
		elseif layoutoption="productimage" then
			if NOT usecsslayout then print "<tr><td width=""100%"" align=""center"" class="""&cs&"prodimage allprodimages"">"
			if NOT isarray(allimages) then
				if usecsslayout then print "<div class="""&cs&"prodimage allprodimages"">&nbsp;</div>" else print "&nbsp;"
			else
				if usecsslayout then
					print "<div class="""&cs&"prodimage allprodimages"">"
				elseif UBOUND(allimages,2)>0 AND NOT thumbnailsonproducts then
					print "<table border=""0"" cellspacing=""1"" cellpadding=""1""><tr><td colspan=""3"">"
				end if
				if (magictoolboxproducts="MagicSlideshow" OR magictoolboxproducts="MagicScroll") AND UBOUND(allimages,2)>0 then
					print "<div class=""" & magictoolboxproducts & """ "&magictooloptionsproducts&">"
					for index=0 to UBOUND(allimages,2)
						largeimage=""
						if isarray(alllgimages) then if UBOUND(alllgimages,2)>=index then largeimage=alllgimages(0,index)
						print "<img" & IIfVs(NOT noschemamarkup," itemprop=""image""") & " src=""" & allimages(0,index) & """ alt="""" "&IIfVs(largeimage<>"" AND magictoolboxproducts="MagicSlideshow","data-fullscreen-image="""&largeimage&""" ")&"/>"
					next
					print "</div>"
				else
					relid=magictooloptionsproducts
					if magictoolboxproducts="MagicThumb" then
						if magictooloptionsproducts="" then relid="rel=""group:g"&Count&""" " else relid=replace(magictooloptionsproducts,"rel=""","rel=""group:g"&Count&";") & " "
					end if
					print IIfVr(magictoolboxproducts<>"" AND magictoolboxproducts<>"MagicSlideshow" AND magictoolboxproducts<>"MagicScroll" AND plargeimage<>"","<a id=""mz" & IIfVr(isquickbuy,"qb","prod") & "image"&Count&""" " & relid & " href="""&plargeimage&""" class=""" & magictoolboxproducts & """>",startlink)&"<img id=""" & IIfVr(isquickbuy,"qb","prod") & "image"&Count&""" class="""&cs&"prodimage allprodimages"" src="""&replace(allimages(0,0),"%s","")&""" alt="""&replace(strip_tags2(rs(getlangid("pName",1))),"""","&quot;")&""" />"&IIfVr(magictoolboxproducts<>"" AND plargeimage<>"","</a>",endlink)
					if UBOUND(allimages,2)>0 AND NOT thumbnailsonproducts then
						print IIfVr(usecsslayout, "<div class="""&cs&"imagenavigator prodimagenavigator"">", "</td></tr><tr><td class="""&cs&"imagenavigator prodimagenavigator"" align=""left"">") & imageorbutton(imgprevimg,xxPrImTx,"previmg","updateprodimage2("&IIfVr(isquickbuy,"true,","false,")&Count&",false)",TRUE) & IIfVs(NOT usecsslayout,"</td><td align=""center"">") & "<span class=""extraimage extraimagenum"" id=""" & IIfVr(isquickbuy,"qb","extra") & "imcnt"&Count&""">1</span> <span class=""extraimage"">"&xxOf&" "&extraimages&"</span>" & IIfVs(NOT usecsslayout,"</td><td align=""right"">") & imageorbutton(imgnextimg,xxNeImTx,"nextimg","updateprodimage2("&IIfVr(isquickbuy,"true,","false,")&Count&",true)",TRUE) & IIfVr(usecsslayout,"</div>","</td></tr></table>")
					end if
				end if
				if usecsslayout then print "</div>"
				if magictoolboxproducts<>"" AND UBOUND(allimages,2)>0 AND thumbnailsonproducts then
					if magictoolboxproducts="MagicThumb" then relid=" rel=""thumb-id:mz" & IIfVr(isquickbuy,"qb","prod") & "image"&Count&"""" else relid=""
					if magictoolboxproducts="MagicZoom" OR magictoolboxproducts="MagicZoomPlus" then relid=" data-zoom-id=""mz" & IIfVr(isquickbuy,"qb","prod") & "image"&Count&""""
					if thumbnailstyleproducts="" then thumbnailstyleproducts="width:50px;padding:2px"
					if usecsslayout then print "<div class=""thumbnailimage productsthumbnail"">" else print "</td></tr><tr><td class=""thumbnailimage productsthumbnail"" align=""center"">"
					if magicscrollthumbnailsproducts then print "<div class=""MagicScroll"">"
					for index=0 to UBOUND(allimages,2)
						if UBOUND(alllgimages,2)>=index then print "<a href=""" & alllgimages(0,index) & """ rev=""" & allimages(0,index) & """" & relid & "><img src=""" & allimages(0,index) & """ style=""" & thumbnailstyleproducts & """ alt="""" /></a>"
					next
					if magicscrollthumbnailsproducts then print "</div>"
					if usecsslayout then print "</div>" else print "</td></tr></table>"
				end if
			end if
			if NOT usecsslayout then print "</td></tr>"
			print vbLf
		elseif layoutoption="productname" then
			if NOT usecsslayout then print "<tr><td width=""100%"">"
			print sstrong & "<div class="""&cs&"prodname"">"&startlink & rs(getlangid("pName",1)) & endlink & xxDot & "</div>" & estrong & vbLf
		elseif layoutoption="discounts" then
			savenoshowdiscounts=noshowdiscounts
			if hascustomlayout then noshowdiscounts=FALSE
			call getperproductdiscounts()
			if alldiscounts<>"" then print IIfVs(xxDsApp<>"","<div class=""eachproddiscountsapply eachproddiscount"">"&xxDsApp&"</div>") & "<div class=""proddiscounts "&cs&"eachproddiscount"">" & alldiscounts & "</div>" & vbLf
			if noapplydiscounts<>"" then print "<div class=""ectwarning discountsnotapply"">"&xxDsNoAp&"</div><div class="""&cs&"prodnoapplydiscounts"">" & noapplydiscounts & "</div>" & vbLf
			call displaydiscountexemptions()
			noshowdiscounts=savenoshowdiscounts
		elseif layoutoption="reviewstars" then
			if ratingsonproductspage=TRUE OR hascustomlayout then
				if rs("pNumRatings")>0 then print showproductreviews(2, cs&"prodrating") else print prodreviewnoratings
			end if
		elseif layoutoption="instock" then
			if useStockManagement AND (showinstock OR hascustomlayout) AND (rs("pInStock")<=clng(stockdisplaythreshold) OR stockdisplaythreshold="") then if cint(rs("pStockByOpts"))=0 then print "<div class="""&cs&"prodinstock"">" & IIfVs(xxInStoc<>"","<span class=""prodinstocklabel"">" & xxInStoc & "</span> ") & vrmax(0,rs("pInStock")) & "</div>" & vbLf
		elseif layoutoption="description" then
			if shortdesc<>"" then print "<div class="""&cs&"proddescription"">" & shortdesc & "</div>" & vbLf else print "<br />"
		elseif layoutoption="options" then
			if hasformvalidator then print "DUPLICATE OPTIONS" else print optionshtml
			hasformvalidator=TRUE
		elseif layoutoption="listprice" then
			if noprice<>TRUE AND rs("pID")<>giftcertificateid AND rs("pID")<>donationid then
				if cdbl(rs("pListPrice"))<>0.0 then
					plistprice=IIfVr(showtaxinclusive=2 AND (rs("pExemptions") AND 2)<>2,rs("pListPrice")+(rs("pListPrice")*thetax/100.0), rs("pListPrice"))
					yousaveprice=plistprice-IIfVr(showtaxinclusive=2 AND (rs("pExemptions") AND 2)<>2,rs("pPrice")+(rs("pPrice")*thetax/100.0),rs("pPrice"))
					print "<div class="""&cs&"listprice"" id="""&IIfVs(isquickbuy,"qb")&"listdivec" & Count & """" & IIfVs(yousaveprice<=0," style=""display:none""") & ">" & Replace(xxListPrice, "%s", FormatEuroCurrency(plistprice)) & IIfVs(yousavetext<>"" AND yousaveprice>0,replace(yousavetext,"%s",FormatEuroCurrency(yousaveprice))) & "</div>"
				end if
			end if
		elseif layoutoption="price" then
			if noprice<>TRUE AND rs("pID")<>giftcertificateid AND rs("pID")<>donationid then
				print "<div class="""&cs&"prodprice""><span class=""prodpricelabel"">" & xxPrice & "</span><span class=""price"" id="""&IIfVs(isquickbuy,"qb")&"pricediv" & Count & """>" & IIfVr(totprice=0 AND pricezeromessage<>"",pricezeromessage,FormatEuroCurrency(IIfVr(showtaxinclusive=2 AND (rs("pExemptions") AND 2)<>2, totprice+(totprice*thetax/100.0), totprice))) & "</span> "
				if showtaxinclusive=1 AND (rs("pExemptions") AND 2)<>2 then print "<span class=""inctax"" id="""&IIfVs(isquickbuy,"qb")&"taxmsg" & Count & """" & IIfVs(totprice=0, " style=""display:none""") & ">" & Replace(ssIncTax,"%s", "<span id="""&IIfVs(isquickbuy,"qb")&"pricedivti" & Count & """>" & IIfVr(totprice=0, "-", FormatEuroCurrency(totprice+(totprice*thetax/100.0))) & "</span> ") & "</span>"
				print "</div>" & vbLf
			end if
		elseif layoutoption="currency" then
			if noprice<>TRUE AND rs("pID")<>giftcertificateid AND rs("pID")<>donationid then
				extracurr=""
				if currRate1<>0 AND currSymbol1<>"" then extracurr=replace(currFormat1, "%s", FormatNumber(totprice*currRate1, checkDPs(currSymbol1))) & currencyseparator
				if currRate2<>0 AND currSymbol2<>"" then extracurr=extracurr & replace(currFormat2, "%s", FormatNumber(totprice*currRate2, checkDPs(currSymbol2))) & currencyseparator
				if currRate3<>0 AND currSymbol3<>"" then extracurr=extracurr & replace(currFormat3, "%s", FormatNumber(totprice*currRate3, checkDPs(currSymbol3)))
				if extracurr<>"" then print "<div class="""&cs&"prodcurrency""><span class=""extracurr"" id="""&IIfVs(isquickbuy,"qb")&"pricedivec" & Count & """>" & IIfVs(totprice<>0, extracurr) & "</span></div>" & vbLf
			end if
		elseif layoutoption="quantity" then
			if NOT (totprice=0 AND nosellzeroprice=TRUE) AND hasmultipurchase=0 AND (isinstock OR isbackorder) then
				print IIfVs(NOT usecsslayout, "<table><tr><td align=""center"">") & quantitymarkup(isquickbuy,Count,FALSE,"",FALSE) & IIfVs(NOT usecsslayout, "</td><td align=""center"">")
			end if
		elseif layoutoption="addtocart" then
			if NOT usecsslayout then print "</td></tr>"
			print replace(IIfVr(isquickbuy,atcmuqb,atcmu),"XXXECTCSPLACEHOLDERXXX",cs)
		elseif layoutoption="addtocartquant" then
			print "<div class=""addtocartquant prodaddtocartquant"">" & quantitymarkup(isquickbuy,Count,FALSE,"",FALSE) & replace(IIfVr(isquickbuy,atcmuqb,atcmu),"XXXECTCSPLACEHOLDERXXXaddtocart",cs&"addtocart") & "</div>"
		elseif layoutoption="custom1" then
			pcustomparam=trim(rs("pCustom1")&"")
			if pcustomparam<>"" then print "<div class="""&cs&"prodcustom1"">" & prodcustomlabel1 & rs("pCustom1") & "</div>" & vbLf
		elseif layoutoption="custom2" then
			pcustomparam=trim(rs("pCustom2")&"")
			if pcustomparam<>"" then print "<div class="""&cs&"prodcustom2"">" & prodcustomlabel2 & rs("pCustom2") & "</div>" & vbLf
		elseif layoutoption="custom3" then
			pcustomparam=trim(rs("pCustom3")&"")
			if pcustomparam<>"" then print "<div class="""&cs&"prodcustom3"">" & prodcustomlabel3 & rs("pCustom3") & "</div>" & vbLf
		elseif layoutoption="detaillink" then
			print "<div class="""&cs&"detaillink"">" & imageorbutton(imgdetaillink,xxPrDets,cs&"detaillink",thedetailslink,FALSE) & "</div>" & vbLf
		elseif layoutoption="dateadded" then
			if NOT isnull(rs("pDateAdded")) then print "<div class="""&cs&"proddateadded"">" & IIfVs(xxDatLab<>"","<div class="""&cs&"proddateaddedlabel"">" & xxDatLab & "</div>") & "<div class="""&cs&"proddateaddeddate"">" & FormatDateTime(rs("pDateAdded"),0) & "</div></div>" & vbLf
		elseif layoutoption="quickbuy" then
			print "<div class=""qbopaque"" id=""qbopaque"&Count&""" style=""display:none"">"
			print "<div class=""qbuywrapper""><div class=""scart scclose""><img src=""images/close.gif"" style=""cursor:pointer"" onclick=""closequickbuy("&Count&")"" alt="""&xxClsWin&""" /></div>"
			call displaylayoutarray(quickbuylayoutarray,TRUE)
			print "<div style=""clear:both""></div></div>"
			print "</div>"
			print "<div class="""&cs&"qbuybutton"">" & imageorbutton(imgquickbuybutton,xxQuiBuy,cs&"qbuybutton","displayquickbuy("&Count&")",TRUE) & "</div>" & vbLf
		elseif layoutoption="quantitypricing" then
			sSQL="SELECT "&WSP&"pPrice,pbQuantity,"&IIfVs(WSP<>"","pbWholesalePercent AS ")&"pbPercent FROM pricebreaks WHERE pbProdID='" & escape_string(rs("pId")) & "' ORDER BY pbQuantity"
			rs2.open sSQL,cnn,0,1
			if NOT rs2.EOF then
				print "<div class=""prodquantpricingwrap""><div class=""prodquantpricing"" style=""display:table"">"
				if xxQuaPri<>"" then print "<div class=""prodqpheading"" style=""display:table-caption"">" & xxQuaPri & "</div>"
				print "<div class=""prodqpheaders"" style=""display:table-row""><div class=""prodqpheadquant"" style=""display:table-cell"">" & xxQuanti & "</div><div class=""prodqpheadprice"" style=""display:table-cell"">" & xxPriQua & "</div></div>"
				quantpricearray=rs2.getrows()
				for index=0 to UBOUND(quantpricearray,2)
					if index<UBOUND(quantpricearray,2) then
						nextquant=quantpricearray(1,index+1)-1
						nextquant=IIfVs(nextquant>quantpricearray(1,index),"-" & nextquant)
					else
						nextquant="+"
					end if
					print "<div class=""prodqprow"" style=""display:table-row""><div class=""prodqpquant"" style=""display:table-cell"">" & quantpricearray(1,index) & nextquant & "</div><div class=""prodqpprice"" style=""display:table-cell"">" & FormatEuroCurrency(IIfVr(quantpricearray(2,index)<>0,rs("pPrice")-((rs("pPrice")*quantpricearray(0,index))/100),quantpricearray(0,index))) & "</div></div>"
				next
				print "</div></div>" & vbLf
			end if
			rs2.close
		elseif left(layoutoption,1)="<" then
			print layoutoption
		else
			print "UNKNOWN LAYOUT OPTION:"&layoutoption&"<br />"
		end if
	next
	cs=dlasavecs
end sub
sub displaysoftlogin()
	if displaysoftlogindone="" then
		location=""
		if instr(request.servervariables("URL"),"/cart"&extension)>0 OR instr(request.servervariables("URL"),"/thanks"&extension)>0 then location="cart"&extension
		if instr(request.servervariables("URL"),"/"&customeraccounturl)>0 then location=customeraccounturl
%>
<script>
var liajaxobj;
function naajaxcallback(){
	if(liajaxobj.readyState==4){
		postdata="email=" + encodeURIComponent(document.getElementById('naemail').value) + "&pw=" + encodeURIComponent(document.getElementById('pass').value);
		document.getElementById('newacctpl').style.display='none';
		document.getElementById('newacctdiv').style.display='';
		if(liajaxobj.responseText.substr(0,7)=='SUCCESS'){
			document.getElementById('newacctdiv').innerHTML="<%=jscheck("<div style=""margin:80px;text-align:center"">"&xxAccSuc&"</div>")%>";
			setTimeout("document.location<%=IIfVr(location<>"","='"&location&"'",".reload()")%>",1200);
		}else{
			document.getElementById('accounterrordiv').innerHTML=liajaxobj.responseText.substr(6);
			var className=document.getElementById('accounterrordiv').className;
			if(className.indexOf('ectwarning')==-1)document.getElementById('accounterrordiv').className+=' ectwarning cartnewaccloginerror';
<%	if recaptchaenabled(8) then print "nacaptchaok=false;grecaptcha.reset(nacaptchawidgetid);" %>
		}
	}
}
var checkedfullname=false;
function checknewaccount(){
var tobj=document.getElementById('naname');
if(tobj.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&xxName)%>\".");
	tobj.focus();
	return(false);
}
var regex=/ /;
if(!checkedfullname && !regex.test(tobj.value)){
	alert("<%=jscheck(xxFulNam&" """&xxName)%>\".");
	tobj.focus();
	checkedfullname=true;
	return(false);
}
var regex=/[^@]+@[^@]+\.[a-z]{2,}$/i;
var tobj=document.getElementById('naemail');
if(!regex.test(tobj.value)){
	alert("<%=jscheck(xxValEm)%>");
	tobj.focus();
	return(false);
}
var tobj=document.getElementById('pass');
<%	if nocustomerloginpwlimit then %>
if(tobj.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&xxPwd)%>\".");
<%	else %>
if(tobj.value.length<6){
	alert("<%=jscheck(replace(xxMinLen,"%s",6)&" """&xxPwd)%>\".");
<%	end if %>
	tobj.focus();
	return(false);
}
<%	if extraclientfield1required then %>
var tobj=document.getElementById('extraclientfield1');
if(tobj.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&extraclientfield1)%>\".");
	tobj.focus();
	return(false);
}
<%	end if
	if extraclientfield2required then %>
var tobj=document.getElementById('extraclientfield2');
if(tobj.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&extraclientfield2)%>\".");
	tobj.focus();
	return(false);
}
<%	end if %>
	postdata="fullname=" + encodeURIComponent(document.getElementById('naname').value) + "&email=" + encodeURIComponent(document.getElementById('naemail').value) + "&pw=" + encodeURIComponent(document.getElementById('pass').value)<%=IIfVs(NOT nomailinglist," + '&allowemail=' + (document.getElementById('allowemail').checked?1:0)")%>;<%
	if recaptchaenabled(8) then
		print vbCrLf & "if(!nacaptchaok){ alert(""" & jscheck(xxRecapt) & """);return(false); }"
		print vbCrLf & "postdata+='&g-recaptcha-response='+encodeURIComponent(nacaptcharesponse);"
	end if
	if trim(extraclientfield1)<>"" then print vbCrLf & "postdata+='&extraclientfield1=' + encodeURIComponent(document.getElementById('extraclientfield1').value);"
	if trim(extraclientfield2)<>"" then print vbCrLf & "postdata+='&extraclientfield2=' + encodeURIComponent(document.getElementById('extraclientfield2').value);"%>
	document.getElementById('newacctdiv').style.display='none';
	document.getElementById('newacctpl').style.display='';
	liajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
	liajaxobj.onreadystatechange=<%=IIfVr(getpost("mode")="checkout","co2newacctcallback","naajaxcallback")%>;
	liajaxobj.open('POST','vsadmin/ajaxservice.asp?action=createaccount',true);
	liajaxobj.setRequestHeader('Content-type','application/x-www-form-urlencoded');
	liajaxobj.send(postdata);
	return false;
}
var lastloginattempt;
function checklogintimeout(){
	var thistime=new Date().getTime();
	var timeremaining=(lastloginattempt+5000)-thistime;
	if(timeremaining<0){
		document.getElementById('cartaccountlogin').disabled=false;
		document.getElementById('fclitspan').innerHTML="<%=jscheck(xxNow)%>";
	}else{
		document.getElementById('fclitspan').innerHTML="<%=jscheck(xxInSecs)%>".replace('%s',Math.ceil(timeremaining/1000));
		setTimeout('checklogintimeout()',1000);
	}
}
function laajaxcallback(){
	var lirefs=[];<%
	if clientloginref0<>"" AND clientloginref0<>"referer" then print "lirefs[0]='" & jsescape(clientloginref0) & "';" & vbCrLf
	if clientloginref1<>"" AND clientloginref1<>"referer" then print "lirefs[1]='" & jsescape(clientloginref1) & "';" & vbCrLf
	if clientloginref2<>"" AND clientloginref2<>"referer" then print "lirefs[2]='" & jsescape(clientloginref2) & "';" & vbCrLf
	if clientloginref3<>"" AND clientloginref3<>"referer" then print "lirefs[3]='" & jsescape(clientloginref3) & "';" & vbCrLf
	if clientloginref4<>"" AND clientloginref4<>"referer" then print "lirefs[4]='" & jsescape(clientloginref4) & "';" & vbCrLf
	if clientloginref5<>"" AND clientloginref5<>"referer" then print "lirefs[5]='" & jsescape(clientloginref5) & "';" & vbCrLf
%>	if(liajaxobj.readyState==4){
		document.getElementById('loginacctpl').style.display='none';
		document.getElementById('loginacctdiv').style.display='';
		if(liajaxobj.responseText.substr(0,10)=='DONELOGOUT'){
			document.getElementById('loginacctdiv').innerHTML="<%=jscheck("<div style=""margin:80px;text-align:center"">"&xxLOSuc&"</div>")%>";
			setTimeout("document.location<%=IIfVr(location<>"","='"&location&"'",".reload()")%>",1000);
		}else if(liajaxobj.responseText.substr(0,7)=='SUCCESS'){
			document.getElementById('loginacctdiv').innerHTML="<%=jscheck("<div style=""margin:80px;text-align:center"">"&xxLISuc&"</div>")%>";
			if(lirefs[liajaxobj.responseText.split(':')[1]])
				setTimeout("document.location='" + lirefs[liajaxobj.responseText.split(':')[1]] + "'",1000);
			else
<%			if SESSION("clientloginref")<>"" then %>
				setTimeout("document.location='<%=jsescape(SESSION("clientloginref"))%>'",1000);
<%			elseif clientloginref<>"" AND clientloginref<>"referer" then %>
				setTimeout("document.location='<%=jsescape(clientloginref)%>'",1000);
<%			else %>
				setTimeout("document.location<%=IIfVr(location<>"","='"&location&"'",".reload()")%>",1000);
<%			end if %>
		}else{
			var responsetext=liajaxobj.responseText.substr(7);
			if(liajaxobj.responseText.substr(6,1)=='1'){
				var thistime=new Date().getTime();
				var timeremaining=(lastloginattempt+5000)-thistime;
				responsetext=responsetext.replace('%s',Math.ceil(timeremaining/1000));
				document.getElementById('cartaccountlogin').disabled=true;
				setTimeout('checklogintimeout()',1000)
			}else
				lastloginattempt=new Date().getTime();
			document.getElementById('liaccterrordiv').innerHTML=responsetext;
			var className=document.getElementById('liaccterrordiv').className;
			if(className.indexOf('ectwarning')==-1)document.getElementById('liaccterrordiv').className+=' ectwarning';
		}
	}
}
function checkloginaccount(){
	var regex=/[^@]+@[^@]+\.[a-z]{2,}$/i;
	var testitem=document.getElementById('liemail');
	if(!regex.test(testitem.value)){
		alert("<%=jscheck(xxValEm)%>");
		testitem.focus();
		return(false);
	}
	testitem=document.getElementById('lipass');
	if(testitem.value==""){
		alert("<%=jscheck(xxPlsEntr&" """&xxPwd)%>\".");
		testitem.focus();
		return(false);
	}
	document.getElementById('loginacctdiv').style.display='none';
	document.getElementById('loginacctpl').style.display='';
	liajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
	liajaxobj.onreadystatechange=laajaxcallback;
	postdata="email=" + encodeURIComponent(document.getElementById('liemail').value) + "&pw=" + encodeURIComponent(document.getElementById('lipass').value) + (document.getElementById('licook').checked?'&licook=ON':'');
	liajaxobj.open('POST','vsadmin/ajaxservice.asp?action=loginaccount&lc=<%=sha256(adminSecret & "ect admin login" & session.sessionid)%>',true);
	liajaxobj.setRequestHeader('Content-type','application/x-www-form-urlencoded');
	liajaxobj.send(postdata);
	return false;
}
function dologoutaccount(){
	document.getElementById('loginacctdiv').style.display='none';
	document.getElementById('alopaquediv').style.display=''
	document.getElementById('loginacctpl').style.display='';
	liajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
	liajaxobj.onreadystatechange=laajaxcallback;
	liajaxobj.open('GET','vsadmin/ajaxservice.asp?action=logoutaccount',true);
	liajaxobj.setRequestHeader('Content-type','application/x-www-form-urlencoded');
	liajaxobj.send(null);
<% if storeurlssl<>storeurl AND request.servervariables("HTTPS")<>"on" then %>
	liajaxobj2=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
	liajaxobj2.open('GET','<%=storeurlssl%>vsadmin/ajaxservice.asp?action=logoutaccount',true);
	liajaxobj2.setRequestHeader('Content-type','application/x-www-form-urlencoded');
	liajaxobj2.send(null);
<% end if %>
	return false;
}
function displaynewaccount(){
	document.getElementById('alopaquediv').style.display='none';
	document.getElementById('acopaquediv').style.display=''
<%	if recaptchaenabled(8) then %>
	if(nacaptchawidgetid===''){
		nacaptchawidgetid=grecaptcha.render('nacaptcha',{
		'sitekey' : '<%=reCAPTCHAsitekey%>',
		'expired-callback' : function(){
			nacaptchaok=false;
		},'callback' : function(response){
			nacaptcharesponse=response;
			nacaptchaok=true;
		}});
	}
	nacaptchaok=false;
<%	end if %>
	return false;
}
var nacaptchawidgetid='';
var nacaptchaok=false;
var nacaptcharesponse=false;
function displayloginaccount(){
	if(document.getElementById('liemail')) document.getElementById('liemail').disabled=false;
	document.getElementById('alopaquediv').style.display='';
	return false;
}
function hideaccounts(){
	if(document.getElementById('liemail')) document.getElementById('liemail').disabled=true;
	document.getElementById('alopaquediv').style.display='none';
	document.getElementById('acopaquediv').style.display='none';
	return false;
}
</script>
<%		print "<div id=""alopaquediv"" style=""display:none;position:fixed;width:100%;height:100%;background-color:rgba(140,140,150,0.5);top:0px;left:0px;z-index:10000"">"
		print "<div class=""accloginwrapper"" style=""margin:120px auto 0 auto;background:#FFF;width:600px;padding:6px;border-radius:5px;box-shadow:1px 1px 5px #333""><div class=""scart scclose""><img src=""images/close.gif"" onclick=""hideaccounts()"" style=""cursor:pointer"" alt=""" & htmlspecials(xxClsWin) & """ /></div>"
		print "<div style=""display:none;text-align:center"" id=""loginacctpl""><img style=""margin:30px"" src=""images/preloader.gif"" alt=""Loading"" /></div>"

		print "<div class=""cartlogin_cntnr"" id=""loginacctdiv""><div class=""cartloginheader"" id=""liaccterrordiv"">" & xxLiDets & "</div>"
		print "<div class=""cartloginemail_cntnr flexvertalign""><div class=""cartloginemailtext"">" & labeltxt(xxEmail,"liemail") & "</div><div class=""cartloginemail""><input type=""text"" name=""liemail"" id=""liemail"" class=""ecttextinput"" size=""31"" placeholder="""&stripnspecials(xxEmail)&""" disabled /></div></div>"
		print "<div class=""cartloginpwd_cntnr flexvertalign""><div class=""cartloginpwdtext"">" & labeltxt(xxPwd,"lipass") & "</div><div class=""cartloginpwd""><input type=""password"" name=""lipass"" id=""lipass"" class=""ecttextinput"" size=""31"" placeholder="""&stripnspecials(xxPwd)&""" autocomplete=""off"" /></div></div>"
		print "<div class=""cartlogincookie_cntnr flexvertalign""><div class=""cartlogincookie""><input type=""checkbox"" id=""licook"" /></div><div class=""cartlogincookietext"">" & labeltxt(xxRemLog,"licook") & "</div></div>"
		print "<div class=""cartloginbuttons_cntnr"">"
		print "<div class=""cartloginlogin"">" & imageorbutton(imgcartaccountlogin,xxSubmt,"cartaccountlogin"" id=""cartaccountlogin","checkloginaccount()",TRUE) & "</div>"
		if allowclientregistration=TRUE then print "<div class=""cartloginnewacct"">" & imageorbutton(imgnewaccount,xxNewAcc,"newaccount","displaynewaccount()",TRUE) & "</div>"
		print "<div class=""cartloginforgotpwd"">" & imageorbutton(imgforgotpassword,xxForPas,"forgotpassword",customeraccounturl&"?mode=lostpassword",FALSE) & "</div></div></div>"
		print "</div>"
		print "</div>"
		' *****
		print "<div id=""acopaquediv"" style=""display:none;position:fixed;width:100%;height:100%;background-color:rgba(140,140,150,0.5);top:0px;left:0px;z-index:10000"">"
		print "<div class=""acccreatewrapper"" style=""margin:120px auto 0 auto;background:#FFF;width:600px;padding:6px;border-radius:5px;box-shadow:1px 1px 5px #333""><div class=""scart scclose""><img src=""images/close.gif"" onclick=""hideaccounts()"" style=""cursor:pointer"" alt=""" & htmlspecials(xxClsWin) & """ /></div>"
		print "<div style=""display:none;text-align:center"" id=""newacctpl""><img style=""margin:30px"" src=""images/preloader.gif"" alt=""Loading"" /></div>"
		print "<div class=""cartnewaccount_cntnr"" id=""newacctdiv"">"
			print "<div class=""cartnewaccountheader"" id=""accounterrordiv"">" & xxNewAcc & "</div>"
			print "<div class=""cartacctloginname_cntnr flexvertalign""><div class=""cartacctloginnametext"">" & redstar & labeltxt(xxName,"naname") & "</div><div class=""cartacctloginname""><input type=""text"" name=""naname"" id=""naname"" class=""ecttextinput"" size=""30"" placeholder="""&stripnspecials(xxName)&""" /></div></div>"
			print "<div class=""cartaccloginemail_cntnr flexvertalign""><div class=""cartaccloginemailtext"">" & redstar & labeltxt(xxEmail,"naemail") & "</div><div class=""cartaccloginemail""><input type=""email"" name=""naemail"" id=""naemail"" class=""ecttextinput"" size=""30"" placeholder="""&stripnspecials(xxEmail)&""" /></div></div>"
			print "<div class=""cartaccloginpassword_cntnr flexvertalign""><div class=""cartaccloginpasswordtext"">" & redstar & labeltxt(xxPwd,"pass") & "</div><div class=""cartaccloginpassword""><input type=""password"" name=""pass"" id=""pass"" class=""ecttextinput"" size=""30"" placeholder="""&stripnspecials(xxPwd)&""" autocomplete=""off"" /></div></div>"
			if extraclientfield1<>"" OR extraclientfield2<>"" then
				if trim(extraclientfield1)<>"" then
					print "<div class=""cartaccloginextra1_cntnr flexvertalign""><div class=""cartaccloginextra1text"">" & IIfVs(extraclientfield1required,redstar) & labeltxt(extraclientfield1,"extraclientfield1") & "</div><div class=""cartaccloginextra1""><input type=""text"" name=""extraclientfield1"" id=""extraclientfield1"" class=""ecttextinput"" size=""30"" placeholder="""&stripnspecials(extraclientfield1)&""" /></div></div>"
				end if
				if trim(extraclientfield2)<>"" then
					print "<div class=""cartaccloginextra2_cntnr flexvertalign""><div class=""cartaccloginextra2text"">" & IIfVs(extraclientfield2required,redstar) & labeltxt(extraclientfield2,"extraclientfield2") & "</div><div class=""cartaccloginextra2""><input type=""text"" name=""extraclientfield2"" id=""extraclientfield2"" class=""ecttextinput"" size=""30"" placeholder="""&stripnspecials(extraclientfield2)&""" /></div></div>"
				end if
			end if
			if nomailinglist<>TRUE then
				print "<div class=""cartaccloginallowpromo_cntnr flexvertalign"">"
				print "<div class=""cartaccloginallowpromo""><input type=""checkbox"" name=""allowemail"" id=""allowemail"" value=""ON""" & IIfVs(allowemaildefaulton=TRUE OR getpost("allowemail")="ON"," checked=""checked""") & " /></div>"
				print "<div class=""cartaccloginallowpromotext"">" & xxAlPrEm & "<div class=""cartacclogineverdivulge"">" & xxNevDiv & "</div></div>"
				print "</div>"
			end if
			if recaptchaenabled(8) then
				print "<div class=""cartaccloginrecaptcha_cntnr flexvertalign"">"
				call displayrecaptchajs("nacaptcha",FALSE,FALSE)
				print "<div class=""cartaccloginrecaptchaspace"">&nbsp;</div>"
				print "<div id=""nacaptcha"" class=""cartaccloginrecaptcha""></div>"
				print "</div>"
			end if
			print "<div class=""cartaccloginalsubmit"">" & imageorbutton(imgcreateaccount,xxCrNwAc,"createaccount","checknewaccount()",TRUE) & "</div>"
		print "</div>"
		print "<div style=""clear:both""></div></div>"
		print "</div>"
		print "<script>document.body.appendChild(document.getElementById('alopaquediv'));document.body.appendChild(document.getElementById('acopaquediv'));</script>"
	end if
	displaysoftlogindone=TRUE
end sub
function checkrecaptcha(recaptchaerr)
	checkrecaptcha=FALSE
	if callxmlfunction("https://www.google.com/recaptcha/api/siteverify","secret=" & reCAPTCHAsecret & "&remoteip=" & REMOTE_ADDR & "&response=" & getpost("g-recaptcha-response"),res,"","WinHTTP.WinHTTPRequest.5.1",recaptchaerr,FALSE) then
		trcaptchaarray=split(res,",")
		for each objitem in trcaptchaarray
			itemarr=split(objitem,":")
			if instr(itemarr(0),"""success""")>0 AND instr(itemarr(1),"true")>0 then checkrecaptcha=TRUE
			if instr(itemarr(0),"""error-codes""")>0 then recaptchaerr=trim(replace(itemarr(1),"}",""))
			if recaptchaerr<>"" then
				recaptchaerr=replace(recaptchaerr,vbLf,"")
				recaptchaerr=replace(recaptchaerr," ","")
				recaptchaerr=replace(recaptchaerr,"[""","<div>")
				recaptchaerr=replace(recaptchaerr,"""]","</div>")
				recaptchaerr=replace(recaptchaerr,""",""","</div><div>")
				recaptchaerr=replace(recaptchaerr,"invalid-input-secret","Your reCaptcha Secret key is incorrect")
			end if
		next
	end if
	if checkrecaptcha=FALSE AND recaptchaerr="" then recaptchaerr=xxRecapt
end function
hasdisplayedrecaptcha=FALSE
recaptchaunique=1
function displayrecaptchajs(recapid,loadimmediately,addunique)
	if addunique then recapid=recapid&recaptchaunique : recaptchaunique=recaptchaunique+1
	if NOT hasdisplayedrecaptcha then
		response.write "<script>var recaptchaids=[];function recaptchaonload(){for(var recapi in recaptchaids){"
		response.write "var restr=recaptchaids[recapi]+""widgetid=grecaptcha.render('""+recaptchaids[recapi]+""',{'sitekey' : '"&reCAPTCHAsitekey&"','expired-callback' : function(){""+recaptchaids[recapi]+""ok=false;},'callback' : function(response){""+recaptchaids[recapi]+""response=response;""+recaptchaids[recapi]+""ok=true;}});"";"
		response.write "eval(restr);"
		response.write "}}</script>"
		response.write "<script src=""https://www.google.com/recaptcha/api.js?render=explicit&amp;onload=recaptchaonload""></script>"
	end if
	recaptchajs="var "&recapid&"ok=false;function "&recapid&"done(){"&recapid&"ok=true;}function "&recapid&"expired(){"&recapid&"ok=false;}"
	if loadimmediately then recaptchajs=recaptchajs&"recaptchaids.push('"&recapid&"');"
	response.write "<script>"&recaptchajs&"</script>"
	hasdisplayedrecaptcha=TRUE
end function
function recaptchaenabled(recapid)
	recaptchaenabled=((reCAPTCHAuseon AND recapid)=recapid AND reCAPTCHAsitekey<>"" AND reCAPTCHAsecret<>"")
end function
function getcategoryurl(gcusectionid,gcusectionname,gcusectionurl,gcurootsection)
	if trim(gcusectionurl&"")<>"" then
		getcategoryurl=getcatid(gcusectionurl,IIfVs(seocategoryurls,gcusectionurl),IIfVr(gcurootsection=1,IIfVr(manufacturers,seomanufacturerpattern,seoprodurlpattern),seocaturlpattern))
	elseif gcurootsection=1 then
		getcategoryurl=IIfVs(NOT seocategoryurls,"products"&extension&"?" & IIfVr(manufacturers,"man","cat") & "=") & getcatid(gcusectionid,gcusectionname,IIfVr(manufacturers,seomanufacturerpattern,seoprodurlpattern))
	else
		getcategoryurl=IIfVs(NOT seocategoryurls,"categories"&extension&"?cat=") & getcatid(gcusectionid,gcusectionname,seocaturlpattern)
	end if
end function
function getcarrierfromtrack(trackno,thelink)
	trackno=replace(trim(trackno)," ","")
	getcarrierfromtrack=""
	set toregexp=new regexp
	toregexp.ignorecase=TRUE
	toregexp.global=TRUE
	toregexp.pattern="^((\d{30})|(9\d{21})|(82\d{8})|(\w\w\d{9}US))$"
	if toregexp.test(trackno) then getcarrierfromtrack="usps"
	if getcarrierfromtrack="" then
		toregexp.pattern="^1Z\w{16}$"
		if toregexp.test(trackno) then getcarrierfromtrack="ups"
	end if
	if getcarrierfromtrack="" then
		toregexp.pattern="^((\d{12})|(\d{15})|(\d{20})|(\d{22}))$"
		if toregexp.test(trackno) then getcarrierfromtrack="fedex"
	end if
	if getcarrierfromtrack="" then
		toregexp.pattern="^(\d{10})$"
		if toregexp.test(trackno) then getcarrierfromtrack="dhl"
	end if
	if getcarrierfromtrack="" then
		toregexp.pattern="^\w\w\d{9}GB$"
		if toregexp.test(trackno) then getcarrierfromtrack="royalmail"
	end if
	if getcarrierfromtrack="" then
		toregexp.pattern="^\w\w\d{9}FR$"
		if toregexp.test(trackno) then getcarrierfromtrack="laposte"
	end if
	if getcarrierfromtrack="" then
		toregexp.pattern="^((\d{16})|(\w\w\d{9}\w\w)|(\w{13}CA))$"
		if toregexp.test(trackno) then getcarrierfromtrack="canadapost"
	end if
	if getcarrierfromtrack="" then
		toregexp.pattern="^(\d{14})$"
		if toregexp.test(trackno) then getcarrierfromtrack="auspost"
	end if
	Set toregexp=Nothing
	thelink=""
	if useinternaltrackinglink then
		thelink="tracking"&extension&"?trackno="
	else
		if getcarrierfromtrack="usps" then
			thelink="https://tools.usps.com/go/TrackConfirmAction.action?tRef=fullpage&tLc=1&text28777=&tLabels="
		elseif getcarrierfromtrack="ups" then
			thelink="https://wwwapps.ups.com/WebTracking/track?trackNums="
		elseif getcarrierfromtrack="fedex" then
			thelink="https://www.fedex.com/apps/fedextrack/?action=track&trackingnumber="
		elseif getcarrierfromtrack="dhl" then
			thelink=IIfVr(tracklink10digit<>"",tracklink10digit,"http://www.dhl.com/en/express/tracking.shtml?AWB=")
		elseif getcarrierfromtrack="canadapost" then
			thelink="http://www.canadapost.ca/cpotools/apps/track/personal/findByTrackNumber?trackingNumber="
		elseif getcarrierfromtrack="auspost" then
			thelink="http://auspost.com.au/track/track.html?id="
		elseif getcarrierfromtrack="royalmail" then
			thelink="http://www.royalmail.com/portal/rm/track?trackNumber="
		elseif getcarrierfromtrack="laposte" then
			thelink="https://www.laposte.fr/outils/suivre-vos-envois?code="
		end if
	end if
end function
function checkpricebreaks(cpbpid,origprice)
	newprice=""
	thetotquant=0
	sSQL="SELECT SUM(cartQuantity) AS totquant FROM cart WHERE cartCompleted=0 AND " & getsessionsql() & " AND cartProdID='"&escape_string(cpbpid)&"'"
	rs2.Open sSQL,cnn,0,1
	if NOT rs2.EOF then
		if IsNull(rs2("totquant")) then thetotquant=0 else thetotquant=rs2("totquant")
	end if
	rs2.Close
	sSQL="SELECT "&WSP&"pPrice,"&IIfVs(WSP<>"","pbWholesalePercent AS ")&"pbPercent FROM pricebreaks WHERE "&thetotquant&">=pbQuantity AND pbProdID='"&escape_string(cpbpid)&"' ORDER BY pbQuantity DESC"
	rs2.Open sSQL,cnn,0,1
	if rs2.EOF then checkpricebreaks=origprice else checkpricebreaks=IIfVr(rs2("pbPercent")<>0,origprice-((origprice*rs2("pPrice"))/100),rs2("pPrice"))
	rs2.Close
	Session.LCID=1033
	sSQL="UPDATE cart SET cartProdPrice="&vsround(checkpricebreaks,2)&" WHERE cartCompleted=0 AND " & getsessionsql() & " AND cartProdID='"&escape_string(cpbpid)&"'"
	ect_query(sSQL)
	sSQL="SELECT cartID FROM cart WHERE cartCompleted=0 AND " & getsessionsql() & " AND cartProdID='"&escape_string(cpbpid)&"'"
	rs2.Open sSQL,cnn,0,1
	do while NOT rs2.EOF
		sSQL="SELECT coCartOption FROM cartoptions WHERE coMultiply<>0 AND coCartID=" & rs2("cartID")
		rs3.Open sSQL,cnn,0,1
		if NOT rs3.EOF then
			totaloptmultiplier=1
			do while NOT rs3.EOF
				if is_numeric(rs3("coCartOption")) then totaloptmultiplier=totaloptmultiplier*cdbl(rs3("coCartOption")) else totaloptmultiplier=0
				rs3.movenext
			loop
			sSQL="UPDATE cart SET cartProdPrice="&vsround(checkpricebreaks*totaloptmultiplier,2)&" WHERE cartID=" & rs2("cartID")
			ect_query(sSQL)
		end if
		rs3.close
		rs2.movenext
	loop
	rs2.Close
	Session.LCID=saveLCID
end function
function ipv6to4(theip)
	set re=new RegExp
	re.pattern="[^a-f0-9]"
	re.ignorecase=TRUE
	re.global=TRUE
	if instr(theip,":")=0 then
		ipv6to4=theip
	else
		do while instr(theip,"::")>0 AND (len(theip)-len(replace(theip,":",""))<7)
			theip=replace(theip,"::",":::",1,1)
		loop
		iparr=split(theip,":")
		for ipv4i=0 to 7
			iparr(ipv4i)=re.replace(iparr(ipv4i),"")
			if trim(iparr(ipv4i))="" then iparr(ipv4i)=0
		next
		iparr(0)=(int("&H" & iparr(0)) XOR int("&H" & iparr(2)) XOR int("&H" & iparr(4)) XOR int("&H" & iparr(6)))
		iparr(1)=(int("&H" & iparr(1)) XOR int("&H" & iparr(3)) XOR int("&H" & iparr(5)) XOR int("&H" & iparr(7)))
		ipv6to4=((int(iparr(1)/256) OR 240) AND 255) & "." & (int(iparr(0)/256) AND 255) & "." & (iparr(1) AND 255) & "." & (iparr(0) AND 255)
	end if
end function
function json_encode(tstr)
	json_encode="""" & trim(replace(tstr&"","""","\""")) & """"
end function
function get_json_val(jsstr, jstag, afterjs)
	get_json_val=""
	if afterjs<>"" then startpos=instr(jsstr,""""&afterjs&"""") else startpos=1
	if startpos<1 then startpos=1
	startpos=instr(startpos,jsstr,""""&jstag&"""")
	if startpos>0 then
		startpos=startpos+len(jstag)+2
		do while startpos<len(jsstr)
			chrat=mid(jsstr,startpos,1)
			if asc(chrat)>32 and chrat<>" " and chrat<>":" then
				gotquotes=(chrat="""")
				endpos=instr(startpos+1,jsstr,IIfVr(gotquotes,"""",","))
				if startpos>0 AND endpos>0 then
					get_json_val=mid(jsstr,startpos+IIfVr(gotquotes,1,0),endpos-startpos-IIfVr(gotquotes,1,0))
					if NOT gotquotes then get_json_val=trim(get_json_val)
				end if
				exit do
			end if
			startpos=startpos+1
		loop
	end if
end function
sub dumpxmlout(sentxml,recvdxml)
	print replace(replace(sentxml,"</","&lt;/"),"<","<br />&lt;")&"<hr />"
	print replace(replace(recvdxml,"</","&lt;/"),"<","<br />&lt;")&"<hr />"
end sub
function breadcrumbselect(currlevel,tlevel)
	if SESSION("clientID")<>"" AND SESSION("clientLoginLevel")<>"" then mll=SESSION("clientLoginLevel") else mll=0
	sSQL="SELECT sectionID,"&getlangid("sectionName",256)&",rootSection,"&IIfVs(languageid<>1,getlangid("sectionurl",2048)&" AS ")&"sectionurl FROM sections WHERE sectionDisabled<=" & mll & " AND topSection=" & tlevel & " ORDER BY "&IIfVr(sortcategoriesalphabetically=TRUE, getlangid("sectionName",256), "sectionOrder")
	rs3.open sSQL,cnn,0,1
	breadcrumbselect="<select class=""breadcrumbcats"" size=""1"" onchange=""ectgocheck(this[this.selectedIndex].value)"">"
	do while NOT rs3.EOF
		caturl=getcategoryurl(rs3("sectionID"),rs3(getlangid("sectionName",256)),rs3("sectionurl"),rs3("rootSection"))
		breadcrumbselect=breadcrumbselect&"<option value="""&caturl&""""
		if currlevel=rs3("sectionID") then breadcrumbselect=breadcrumbselect&" selected=""selected""" : if currcaturl="" then currcaturl=caturl
		breadcrumbselect=breadcrumbselect&">" & rs3(getlangid("sectionName",256)) & "</option>"
		rs3.movenext
	loop
	rs3.close
	breadcrumbselect=breadcrumbselect&"</select>" & vbCrLf
end function
function stripnspecials(sptxt)
	stripnspecials=htmlspecials(strip_tags2(sptxt))
end function
function zeropadint(myint,numdigits)
  zeropadint=right("000000000000"&myint,numdigits)
end function
function uniqid()
	randomize
	uniqid=calcmd5(cstr(rnd))
end function
sub notification_err_msg(tmsg)
	sSQL="UPDATE admin SET adminDeviceNotifAlert='" & escape_string(trim(tmsg)) & "' WHERE adminID=1"
	cnn.execute(sSQL)
end sub
sub ectsetcookie(name,value,expires) ' deprecated
	response.cookies(name)=htmlspecialsid(value)
	if expires<>0 then response.cookies(name).expires=date()+int(expires)
	if request.servervariables("HTTPS")="on" then response.cookies(name).secure=TRUE
end sub
sub setacookie(cname,cval,cdays)
	response.cookies(cname)=cval
	if cdays<>0 then response.cookies(cname).expires=date()+cdays
	if request.servervariables("HTTPS")="on" then response.cookies(cname).secure=TRUE
end sub
function vrxmltag(tagname,tagdata)
	vrxmltag="<" & tagname & ">" & htmlspecials(tagdata) & "</" & tagname & ">"
end function
function getstateorabbreviation(statename)
	sSQL="SELECT stateAbbrev FROM states WHERE (stateCountryID=1 OR stateCountryID=2) AND " & IIfVr(is_numeric(statename), "stateID=" & statename, "(stateName='" & escape_string(statename) & "' OR stateAbbrev='" & escape_string(statename) & "')")
	rs3.open sSQL,cnn,0,1
	if NOT rs3.EOF then statename=rs3("stateAbbrev")
	rs3.close
	getstateorabbreviation=statename
end function
sub updateorderstatus(iordid, ordstatus, emailchanges)
	ordauthno=""
	oldordstatus=999
	payprovider=0
	ordClientID=0
	loyaltypointtotal=0
	savelangid=languageid
	savestorelang=storelang
	replaceone=FALSE
	pointsredeemed=0
	if is_numeric(iordid) AND is_numeric(ordstatus) then
		rs.open "SELECT ordStatus,ordAuthNumber,ordEmail,ordDate,"&getlangid("statPublic",64)&",ordName,ordLastName,ordTrackNum,ordPayProvider,ordLang,ordClientID,loyaltyPoints,ordTotal,ordDiscount,ordInvoice,pointsRedeemed,ordStatusInfo FROM orders INNER JOIN orderstatus ON orders.ordStatus=orderstatus.statID WHERE ordID="&iordid,cnn,0,1
		if NOT rs.EOF then
			oldordstatus=rs("ordStatus")
			ordauthno=rs("ordAuthNumber")
			ordemail=rs("ordEmail")
			orddate=cdate(rs("ordDate"))
			oldstattext=rs(getlangid("statPublic",64))&""
			ordstatinfo=rs("ordStatusInfo")&""
			if htmlemails=TRUE then ordstatinfo=replace(ordstatinfo, vbCrLf, "<br />")
			ordername=trim(rs("ordName")&" "&rs("ordLastName"))
			trackingnum=trim(rs("ordTrackNum")&"")
			payprovider=rs("ordPayProvider")
			languageid=rs("ordLang")+1
			if isarray(ectstorelangarr) then
				if UBOUND(ectstorelangarr)>=(languageid-1) then storelang=ectstorelangarr(languageid-1)
			end if
			ordClientID=rs("ordClientID")
			loyaltypointtotal=rs("loyaltyPoints")
			ordTotal=rs("ordTotal")
			ordDiscount=rs("ordDiscount")
			ordInvoice=rs("ordInvoice")
			if NOT isnull(rs("pointsRedeemed")) then pointsredeemed=rs("pointsRedeemed")
		end if
		rs.close
		rs.open "SELECT "&getlangid("statPublic",64)&" FROM orders INNER JOIN orderstatus ON orders.ordStatus=orderstatus.statID WHERE ordID="&iordid,cnn,0,1
		if NOT rs.EOF then
			oldstattext=rs(getlangid("statPublic",64))&""
		end if
		rs.close
		ect_query("UPDATE cart SET cartCompleted="&IIfVr(ordstatus=2,0,1)&" WHERE cartOrderID=" & iordid)
		if (loyaltypointsnowholesale OR loyaltypointsnopercentdiscount) AND ordClientID<>0 then
			sSQL="SELECT clActions FROM customerlogin WHERE clID=" & ordClientID
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if loyaltypointsnowholesale AND (rs("clActions") AND 8)=8 then ordClientID=0
				if loyaltypointsnopercentdiscount AND (rs("clActions") AND 16)=16 then ordClientID=0
			end if
			rs.close
		end if
		if oldordstatus<>999 AND (oldordstatus<3 AND ordstatus>=3) then
			if ordauthno="" then ect_query("UPDATE orders SET ordAuthNumber='"&escape_string(yyManAut)&"' WHERE ordID=" & iordid)
			sSQL="SELECT cartProdId,cartProdName,cartProdPrice,cartQuantity,cartID FROM cart LEFT JOIN products ON cart.cartProdId=products.pID WHERE cartOrderID=" & iordid
			rs.open sSQL,cnn,0,1
			do while not rs.EOF
				if rs("cartProdId")=giftcertificateid then
					sSQL="UPDATE giftcertificate SET gcAuthorized=1,gcOrigAmount="&rs("cartProdPrice")&",gcRemaining="&rs("cartProdPrice")&" WHERE gcAuthorized=0 AND gcCartID=" & rs("cartID")
					ect_query(sSQL)
				end if
				rs.movenext
			loop
			rs.close
			if loyaltypoints<>"" then
				loyaltypointtotal=int((ordTotal-ordDiscount)*loyaltypoints)
				ect_query("UPDATE orders SET loyaltyPoints=" & loyaltypointtotal & " WHERE ordID=" & iordid)
				if ordClientID<>0 then ect_query("UPDATE customerlogin SET loyaltyPoints=loyaltyPoints+" & loyaltypointtotal & " WHERE clID=" & ordClientID)
			end if
		elseif oldordstatus<>999 AND (oldordstatus>=3 AND ordstatus<3) then
			ect_query("UPDATE giftcertificate SET gcAuthorized=0 WHERE gcOrderID=" & iordid)
			if ordClientID<>0 AND loyaltypoints<>"" then ect_query("UPDATE customerlogin SET loyaltyPoints=loyaltyPoints-" & loyaltypointtotal & " WHERE clID=" & ordClientID)
		end if
		if oldordstatus<>999 AND (oldordstatus<2 AND ordstatus>=2) then
			if ordClientID<>0 then ect_query("UPDATE customerlogin SET loyaltyPoints=loyaltyPoints-" & pointsredeemed & " WHERE clID=" & ordClientID)
		elseif oldordstatus<>999 AND (oldordstatus>=2 AND ordstatus<2) then
			if ordClientID<>0 then ect_query("UPDATE customerlogin SET loyaltyPoints=loyaltyPoints+" & pointsredeemed & " WHERE clID=" & ordClientID)
		end if
		if oldordstatus<>999 AND (oldordstatus<=1 AND ordstatus>1) AND (Date()-orddate) < 365 then stock_subtract(iordid)
		if oldordstatus<>999 AND (oldordstatus>1 AND ordstatus<=1) AND (Date()-orddate) < 365 then release_stock(iordid)
		if oldordstatus<>int(ordstatus) then
			if emailchanges AND ordstatus<>1 then
				rs.open "SELECT "&getlangid("statPublic",64)&",emailstatus FROM orderstatus WHERE statID=" & ordstatus,cnn,0,1
				if NOT rs.EOF then
					newstattext=rs(getlangid("statPublic",64))&""
					emailstatus=cint(rs("emailstatus"))<>0
				else
					emailstatus=FALSE
				end if
				rs.close
				if (adminlangsettings AND 4096)=0 then languageid=1
				if ordstatussubject(languageid)<>"" then emailsubject=ordstatussubject(languageid) else emailsubject="Order status updated"
				ose=ordstatusemail(languageid)
				for uoindex=0 to 18
					replaceone=TRUE
					do while replaceone
						ose=replaceemailtxt(ose, "%statusid" & uoindex & "%", IIfVr(uoindex=ordstatus,"%ectpreserve%",""), replaceone)
					loop
				next
				if storelang="de" OR (storelang="" AND adminlang="de") then session.LCID=1031
				if storelang="es" OR (storelang="" AND adminlang="es") then session.LCID=1034
				if storelang="fr" OR (storelang="" AND adminlang="fr") then session.LCID=1036
				if storelang="it" OR (storelang="" AND adminlang="it") then session.LCID=1040
				if storelang="nl" OR (storelang="" AND adminlang="nl") then session.LCID=1043
				ose=replace(ose, "%orderid%", iordid)
				ose=replace(ose, "%orderdate%", FormatDateTime(orddate, 1) & " " & FormatDateTime(orddate, 4))
				ose=replace(ose, "%oldstatus%", oldstattext)
				ose=replace(ose, "%newstatus%", newstattext)
				ose=replace(ose, "%date%", FormatDateTime(DateAdd("h",dateadjust,Now()), 1) & " " & FormatDateTime(DateAdd("h",dateadjust,Now()), 4))
				session.LCID=saveLCID
				ose=replace(ose, "%ordername%", ordername)
				ose=replaceemailtxt(ose, "%statusinfo%", ordstatinfo, replaceone)
				tracknumarr=split(trackingnum,",")
				for uoindex=0 to UBOUND(tracknumarr)
					ose=replaceemailtxt(ose, "%trackingnum%", tracknumarr(uoindex), replaceone)
				next
				do while instr(ose, "%trackingnum%")>0
					ose=replaceemailtxt(ose, "%trackingnum%", "", replaceone)
				loop
				ose=replaceemailtxt(ose, "%invoicenum%", ordInvoice, replaceone)
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
				if emailstatus then Call DoSendEmailEO(ordemail,emailAddr,"",replace(emailsubject,"%orderid%",iordid),ose,emailObject,themailhost,theuser,thepass)
			end if
		end if
		if oldordstatus<>int(ordstatus) then ect_query("UPDATE orders SET ordStatus=" & ordstatus & ",ordStatusDate=" & vsusdatetime(DateAdd("h",dateadjust,Now())) & " WHERE ordID=" & iordid)
	end if
	languageid=savelangid
	storelang=savestorelang
end sub
function getstatetext(cntid)
	getstatetext=xxStaPro
	if cntid=1 then
		getstatetext=xxStateD
	elseif cntid=2 OR cntid=175 then
		getstatetext=xxProvin
	elseif cntid=142 OR cntid=201 then
		getstatetext=xxCounty
	end if
end function
hasdisplayedsvgicons=FALSE
sub displayreviewicons()
	if NOT hasdisplayedsvgicons then %>
	<svg style="display:none">
		<defs>
			<g id="review-icon-full"><svg xmlns="http://www.w3.org/2000/svg" shape-rendering="geometricPrecision" image-rendering="optimizeQuality" fill-rule="evenodd"><path d="M12 1.2l2.67 8.28 8.7-.02-7.04 5.1 2.7 8.27L12 17.68 4.98 22.8l2.7-8.27-7.04-5.1 8.7.02z"/></svg></g>
			<g id="review-icon-half"><svg xmlns="http://www.w3.org/2000/svg" shape-rendering="geometricPrecision" image-rendering="optimizeQuality" fill-rule="evenodd"><path d="M12 1.2l2.67 8.28 8.7-.02-7.04 5.1 2.7 8.27L12 17.68 4.98 22.8l2.7-8.27-7.04-5.1 8.7.02L12 1.2zm0 3.25v12l5.12 3.73-1.97-6.02 5.13-3.7-6.34.01L12 4.44z"/></svg></g>
			<g id="review-icon-empty"><svg xmlns="http://www.w3.org/2000/svg" shape-rendering="geometricPrecision" image-rendering="optimizeQuality" fill-rule="evenodd"><path d="M12 1.2l2.67 8.28 8.7-.02-7.04 5.1 2.7 8.27L12 17.68 4.98 22.8l2.7-8.27-7.04-5.1 8.7.02L12 1.2zm1.72 8.58L12 4.44l-1.94 6.02-6.34-.01 5.13 3.7-1.97 6.02L12 16.45l5.12 3.73-1.97-6.02 5.13-3.7-6.34.01-.22-.7z"/></svg></g>
		</defs>
	</svg>
<%		hasdisplayedsvgicons=TRUE
	end if
end sub
if SESSION("httpreferer")="" AND request.servervariables("HTTP_REFERER")<>"" then
	httpreferer=left(request.servervariables("HTTP_REFERER"), 255)
	if len(httpreferer)>=255 then
		andpos=instrrev(httpreferer, "&")
		if andpos > 0 then httpreferer=left(httpreferer, andpos-1)
	end if
	SESSION("httpreferer")=httpreferer
end if
%>
<script language="jscript" runat="server">
function vrbase64_encrypt(origstr){
	var tcharset="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
	var str="";
	for(var i=0; i < origstr.length; i += 3){
		triplet=(origstr.charCodeAt(i) << 16) | (origstr.charCodeAt(i+1) << 8) | (origstr.charCodeAt(i+2) << 0)
		for(var j=0; j < 4; j++){
			if(i + j > origstr.length) str += "="; else str += tcharset.charAt((triplet >> 6*(3-j)) & 0x3F);
		}
	}
	return str;
}
function vsround(amnt, decpl){
	return(Math.round(amnt * Math.pow(10,decpl),decpl) / Math.pow(10,decpl));
}
function vsceil(amnt){
	return(Math.ceil(amnt));
}
function long2ip(ip2lip){
	retval=((ip2lip >> 24) & 255) + "." + ((ip2lip >> 16) & 255) + "." + ((ip2lip >> 8) & 255) + "." + (ip2lip & 255);
	return(retval);
}
function check2d(i){
	if(i<10)i="0" + i;
	return i;
}
function getutcdate(theoffset){
	var od=new Date();
	od.setDate(od.getDate()+theoffset);
	var td=od.getUTCFullYear()+"-"+check2d(od.getUTCMonth()+1)+"-"+check2d(od.getUTCDate())+"T";
	td += check2d(od.getUTCHours())+":"+check2d(od.getUTCMinutes())+":"+check2d(od.getUTCSeconds())+"Z";
    return(td);
}

// MD5 Content

/*
 * A JavaScript implementation of the RSA Data Security, Inc. MD5 Message
 * Digest Algorithm, as defined in RFC 1321.
 * Version 1.1 Copyright (C) Paul Johnston 1999 - 2002.
 * Code also contributed by Greg Holt
 * See http://pajhome.org.uk/site/legal.html for details.
 */
/*
 * Add integers, wrapping at 2^32. This uses 16-bit operations internally
 * to work around bugs in some JS interpreters.
 */
function safe_add(x, y){
  var lsw=(x & 0xFFFF) + (y & 0xFFFF)
  var msw=(x >> 16) + (y >> 16) + (lsw >> 16)
  return (msw << 16) | (lsw & 0xFFFF)
}
/*
 * Bitwise rotate a 32-bit number to the left.
 */
function rol(num, cnt){
  return (num << cnt) | (num >>> (32 - cnt))
}
/*
 * These functions implement the four basic operations the algorithm uses.
 */
function cmn(q, a, b, x, s, t){
  return safe_add(rol(safe_add(safe_add(a, q), safe_add(x, t)), s), b)
}
function ffxx(a, b, c, d, x, s, t){
  return cmn((b & c) | ((~b) & d), a, b, x, s, t)
}
function ggxx(a, b, c, d, x, s, t){
  return cmn((b & d) | (c & (~d)), a, b, x, s, t)
}
function hhxx(a, b, c, d, x, s, t){
  return cmn(b ^ c ^ d, a, b, x, s, t)
}
function iixx(a, b, c, d, x, s, t){
  return cmn(c ^ (b | (~d)), a, b, x, s, t)
}
/*
 * Calculate the MD5 of an array of little-endian words, producing an array
 * of little-endian words.
 */
function coreMD5(x){
  var a= 1732584193
  var b=-271733879
  var c=-1732584194
  var d= 271733878
  for(i=0; i < x.length; i += 16){
	var olda=a
	var oldb=b
	var oldc=c
	var oldd=d
	a=ffxx(a, b, c, d, x[i+ 0], 7 , -680876936)
	d=ffxx(d, a, b, c, x[i+ 1], 12, -389564586)
	c=ffxx(c, d, a, b, x[i+ 2], 17,  606105819)
	b=ffxx(b, c, d, a, x[i+ 3], 22, -1044525330)
	a=ffxx(a, b, c, d, x[i+ 4], 7 , -176418897)
	d=ffxx(d, a, b, c, x[i+ 5], 12,  1200080426)
	c=ffxx(c, d, a, b, x[i+ 6], 17, -1473231341)
	b=ffxx(b, c, d, a, x[i+ 7], 22, -45705983)
	a=ffxx(a, b, c, d, x[i+ 8], 7 ,  1770035416)
	d=ffxx(d, a, b, c, x[i+ 9], 12, -1958414417)
	c=ffxx(c, d, a, b, x[i+10], 17, -42063)
	b=ffxx(b, c, d, a, x[i+11], 22, -1990404162)
	a=ffxx(a, b, c, d, x[i+12], 7 ,  1804603682)
	d=ffxx(d, a, b, c, x[i+13], 12, -40341101)
	c=ffxx(c, d, a, b, x[i+14], 17, -1502002290)
	b=ffxx(b, c, d, a, x[i+15], 22,  1236535329)
	a=ggxx(a, b, c, d, x[i+ 1], 5 , -165796510)
	d=ggxx(d, a, b, c, x[i+ 6], 9 , -1069501632)
	c=ggxx(c, d, a, b, x[i+11], 14,  643717713)
	b=ggxx(b, c, d, a, x[i+ 0], 20, -373897302)
	a=ggxx(a, b, c, d, x[i+ 5], 5 , -701558691)
	d=ggxx(d, a, b, c, x[i+10], 9 ,  38016083)
	c=ggxx(c, d, a, b, x[i+15], 14, -660478335)
	b=ggxx(b, c, d, a, x[i+ 4], 20, -405537848)
	a=ggxx(a, b, c, d, x[i+ 9], 5 ,  568446438)
	d=ggxx(d, a, b, c, x[i+14], 9 , -1019803690)
	c=ggxx(c, d, a, b, x[i+ 3], 14, -187363961)
	b=ggxx(b, c, d, a, x[i+ 8], 20,  1163531501)
	a=ggxx(a, b, c, d, x[i+13], 5 , -1444681467)
	d=ggxx(d, a, b, c, x[i+ 2], 9 , -51403784)
	c=ggxx(c, d, a, b, x[i+ 7], 14,  1735328473)
	b=ggxx(b, c, d, a, x[i+12], 20, -1926607734)
	a=hhxx(a, b, c, d, x[i+ 5], 4 , -378558)
	d=hhxx(d, a, b, c, x[i+ 8], 11, -2022574463)
	c=hhxx(c, d, a, b, x[i+11], 16,  1839030562)
	b=hhxx(b, c, d, a, x[i+14], 23, -35309556)
	a=hhxx(a, b, c, d, x[i+ 1], 4 , -1530992060)
	d=hhxx(d, a, b, c, x[i+ 4], 11,  1272893353)
	c=hhxx(c, d, a, b, x[i+ 7], 16, -155497632)
	b=hhxx(b, c, d, a, x[i+10], 23, -1094730640)
	a=hhxx(a, b, c, d, x[i+13], 4 ,  681279174)
	d=hhxx(d, a, b, c, x[i+ 0], 11, -358537222)
	c=hhxx(c, d, a, b, x[i+ 3], 16, -722521979)
	b=hhxx(b, c, d, a, x[i+ 6], 23,  76029189)
	a=hhxx(a, b, c, d, x[i+ 9], 4 , -640364487)
	d=hhxx(d, a, b, c, x[i+12], 11, -421815835)
	c=hhxx(c, d, a, b, x[i+15], 16,  530742520)
	b=hhxx(b, c, d, a, x[i+ 2], 23, -995338651)
	a=iixx(a, b, c, d, x[i+ 0], 6 , -198630844)
	d=iixx(d, a, b, c, x[i+ 7], 10,  1126891415)
	c=iixx(c, d, a, b, x[i+14], 15, -1416354905)
	b=iixx(b, c, d, a, x[i+ 5], 21, -57434055)
	a=iixx(a, b, c, d, x[i+12], 6 ,  1700485571)
	d=iixx(d, a, b, c, x[i+ 3], 10, -1894986606)
	c=iixx(c, d, a, b, x[i+10], 15, -1051523)
	b=iixx(b, c, d, a, x[i+ 1], 21, -2054922799)
	a=iixx(a, b, c, d, x[i+ 8], 6 ,  1873313359)
	d=iixx(d, a, b, c, x[i+15], 10, -30611744)
	c=iixx(c, d, a, b, x[i+ 6], 15, -1560198380)
	b=iixx(b, c, d, a, x[i+13], 21,  1309151649)
	a=iixx(a, b, c, d, x[i+ 4], 6 , -145523070)
	d=iixx(d, a, b, c, x[i+11], 10, -1120210379)
	c=iixx(c, d, a, b, x[i+ 2], 15,  718787259)
	b=iixx(b, c, d, a, x[i+ 9], 21, -343485551)
	a=safe_add(a, olda)
	b=safe_add(b, oldb)
	c=safe_add(c, oldc)
	d=safe_add(d, oldd)
  }
  return [a, b, c, d]
}
/*
 * Convert an array of little-endian words to a hex string.
 */
function binl2hex(binarray){
  var hex_tab="0123456789abcdef"
  var str=""
  for(var i=0; i < binarray.length * 4; i++){
	str += hex_tab.charAt((binarray[i>>2] >> ((i%4)*8+4)) & 0xF) +
		   hex_tab.charAt((binarray[i>>2] >> ((i%4)*8)) & 0xF)
  }
  return str
}
/*
 * Convert an array of little-endian words to a base64 encoded string.
 */
function binl2b64(binarray){
  var tab="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  var str=""
  for(var i=0; i < binarray.length * 32; i += 6){
	str += tab.charAt(((binarray[i>>5] << (i%32)) & 0x3F) |
					  ((binarray[i>>5+1] >> (32-i%32)) & 0x3F))
  }
  return str
}
/*
 * Convert an 8-bit character string to a sequence of 16-word blocks, stored
 * as an array, and append appropriate padding for MD4/5 calculation.
 * If any of the characters are >255, the high byte is silently ignored.
 */
function str2binl(str){
  var nblk=((str.length + 8) >> 6) + 1 // number of 16-word blocks
  var blks=new Array(nblk * 16)
  for(var i=0; i < nblk * 16; i++) blks[i]=0
  for(var i=0; i < str.length; i++)
	blks[i>>2] |= (str.charCodeAt(i) & 0xFF) << ((i%4) * 8)
  blks[i>>2] |= 0x80 << ((i%4) * 8)
  blks[nblk*16-2]=str.length * 8
  return blks
}
/*
 * Convert a wide-character string to a sequence of 16-word blocks, stored as
 * an array, and append appropriate padding for MD4/5 calculation.
 */
function strw2binl(str){
  var nblk=((str.length + 4) >> 5) + 1 // number of 16-word blocks
  var blks=new Array(nblk * 16)
  for(var i=0; i < nblk * 16; i++) blks[i]=0
  for(var i=0; i < str.length; i++)
	blks[i>>1] |= str.charCodeAt(i) << ((i%2) * 16)
  blks[i>>1] |= 0x80 << ((i%2) * 16)
  blks[nblk*16-2]=str.length * 16
  return blks
}
/*
 * External interface
 */
function hexMD5 (str) { return binl2hex(coreMD5( str2binl(str))) }
function hexMD5w(str) { return binl2hex(coreMD5(strw2binl(str))) }
function b64MD5 (str) { return binl2b64(coreMD5( str2binl(str))) }
function b64MD5w(str) { return binl2b64(coreMD5(strw2binl(str))) }
/* Backward compatibility */
function calcmd5(str) { return binl2hex(coreMD5( str2binl(str))) }
function binl2byt(binarray){
var hex_tab="0123456789abcdef";
var  bytarray=new Array(binarray.length * 4);
var str="";
for(var i=0; i < binarray.length * 4; i++){
bytarray[i]=(binarray[i>>2] >> ((i%4)*8+4) & 0xF) << 4 | binarray[i>>2] >> ((i%4)*8) & 0xF;
}
return bytarray;
}
function bytarray2word (barray){
var blks=new Array(barray.length / 4);
for(var i=0; i < blks.length; i++) blks[i]=0
for(var i=0; i < barray.length; i++)
blks[i>>2] |= (barray[i] & 0xFF) << ((i%4) * 8)
//blks[i>>2] |= 0x80 << ((i%4) * 8)
//blks[nblk*16-2]=barray.length * 8
return blks
}
function bytarray2binl (barray){
var nblk=((barray.length + 8) >> 6) + 1 // number of 16-word blocks
var blks=new Array(nblk * 16)
for(var i=0; i < nblk * 16; i++) blks[i]=0
for(var i=0; i < barray.length; i++)
blks[i>>2] |= (barray[i] & 0xFF) << ((i%4) * 8)
blks[i>>2] |= 0x80 << ((i%4) * 8)
blks[nblk*16-2]=barray.length * 8
return blks
}
function b_calcMD5(barray) { return coreMD5(bytarray2binl(barray)) }
function HMAC(key, text){
var hkey,idata,odata;
var ipad= new Array(64);
var opad= new Array (64);
idata=new Array (64 + text.length);
odata=new Array (64 + 16);
if(key.length > 64){
	hkey=calcMD5(key);
}
else
	hkey=key;
for(i=0;i<64;i++){
	idata[i]=ipad[i]=0x36;
	odata[i]=opad[i]=0x5C;
}
for(i=0;i<hkey.length; i++){
	ipad[i] ^= hkey.charCodeAt(i);
	opad[i] ^= hkey.charCodeAt(i);
	idata[i]= ipad[i] & 0xFF;
	odata[i]=opad[i] & 0xFF;
}
for(i=0;i<text.length;i++){
	idata[64+i]=text.charCodeAt(i) & 0xFF;
}
var innerhashout=binl2byt(b_calcMD5(idata));
for(i=0;i<16;i++){
odata[64+i]=innerhashout[i];
}
return binl2hex(b_calcMD5(odata));
}
function GetSecondsSince1970(){
var d=new Date();
var secs= Math.floor(d.getTime() / 1000);
return (secs);
}
function hmac_sha1(key, text){
var hkey,idata,odata;
var ipad= new Array(64);
var opad= new Array (64);
idata=new Array (64 + text.length);
odata=new Array (64 + 16);
if(key.length > 64){
	hkey=sha1(key);
}
else
	hkey=key;
for(i=0;i<64;i++){
	idata[i]=ipad[i]=0x36;
	odata[i]=opad[i]=0x5C;
}
for(i=0;i<hkey.length; i++){
	ipad[i] ^= hkey.charCodeAt(i);
	opad[i] ^= hkey.charCodeAt(i);
	idata[i]= ipad[i] & 0xFF;
	odata[i]=opad[i] & 0xFF;
}
for(i=0;i<text.length;i++) {
	idata[64+i]=text.charCodeAt(i) & 0xFF;
}
var innerhashout=sha1(bytarray2binl(idata));
for(i=0;i<16;i++) {
odata[64+i]=innerhashout[i];
}
return binl2hex(b_calcMD5(odata));
}
/*
 * A JavaScript implementation of the Secure Hash Algorithm, SHA-1, as defined
 * in FIPS PUB 180-1
 * Version 2.1a Copyright Paul Johnston 2000 - 2002.
 * Other contributors: Greg Holt, Andrew Kepert, Ydnar, Lostinet
 * Distributed under the BSD License
 * See http://pajhome.org.uk/crypt/md5 for details.
 */
var hexcase=0;  /* hex output format. 0 - lowercase; 1 - uppercase		*/
var b64pad ="="; /* base-64 pad character. "=" for strict RFC compliance   */
var chrsz  =8;  /* bits per input character. 8 - ASCII; 16 - Unicode	  */
/*
 * These are the functions you'll usually want to call
 * They take string arguments and return either hex or base-64 encoded strings
 */
function hex_sha1(s){return binb2hex(core_sha1(str2binb(s),s.length * chrsz));}
function b64_sha1(s){return binb2b64(core_sha1(str2binb(s),s.length * chrsz));}
function str_sha1(s){return binb2str(core_sha1(str2binb(s),s.length * chrsz));}
function hex_hmac_sha1(key, data){ return binb2hex(core_hmac_sha1(key, data));}
function b64_hmac_sha1(key, data){ return binb2b64(core_hmac_sha1(key, data));}
function str_hmac_sha1(key, data){ return binb2str(core_hmac_sha1(key, data));}
/*
 * Calculate the SHA-1 of an array of big-endian words, and a bit length
 */
function core_sha1(x, len){
  /* append padding */
  x[len >> 5] |= 0x80 << (24 - len % 32);
  x[((len + 64 >> 9) << 4) + 15]=len;
  var w=Array(80);
  var a= 1732584193;
  var b=-271733879;
  var c=-1732584194;
  var d= 271733878;
  var e=-1009589776;
  for(var i=0; i < x.length; i += 16){
	var olda=a;
	var oldb=b;
	var oldc=c;
	var oldd=d;
	var olde=e;
	for(var j=0; j < 80; j++){
	  if(j < 16) w[j]=x[i + j];
	  else w[j]=rol(w[j-3] ^ w[j-8] ^ w[j-14] ^ w[j-16], 1);
	  var t=safe_add(safe_add(rol(a, 5), sha1_ft(j, b, c, d)),
					   safe_add(safe_add(e, w[j]), sha1_kt(j)));
	  e=d;
	  d=c;
	  c=rol(b, 30);
	  b=a;
	  a=t;
	}
	a=safe_add(a, olda);
	b=safe_add(b, oldb);
	c=safe_add(c, oldc);
	d=safe_add(d, oldd);
	e=safe_add(e, olde);
  }
  return Array(a, b, c, d, e);
}
/*
 * Perform the appropriate triplet combination function for the current
 * iteration
 */
function sha1_ft(t, b, c, d){
  if(t < 20) return (b & c) | ((~b) & d);
  if(t < 40) return b ^ c ^ d;
  if(t < 60) return (b & c) | (b & d) | (c & d);
  return b ^ c ^ d;
}
/*
 * Determine the appropriate additive constant for the current iteration
 */
function sha1_kt(t){
  return (t < 20) ?  1518500249 : (t < 40) ?  1859775393 :
		 (t < 60) ? -1894007588 : -899497514;
}
/*
 * Calculate the HMAC-SHA1 of a key and some data
 */
function core_hmac_sha1(key, data){
  var bkey=str2binb(key);
  if(bkey.length > 16) bkey=core_sha1(bkey, key.length * chrsz);
  var ipad=Array(16), opad=Array(16);
  for(var i=0; i < 16; i++){
	ipad[i]=bkey[i] ^ 0x36363636;
	opad[i]=bkey[i] ^ 0x5C5C5C5C;
  }
  var hash=core_sha1(ipad.concat(str2binb(data)), 512 + data.length * chrsz);
  return core_sha1(opad.concat(hash), 512 + 160);
}
/*
 * Convert an 8-bit or 16-bit string to an array of big-endian words
 * In 8-bit function, characters >255 have their hi-byte silently ignored.
 */
function str2binb(str){
  var bin=Array();
  var mask=(1 << chrsz) - 1;
  for(var i=0; i < str.length * chrsz; i += chrsz)
	bin[i>>5] |= (str.charCodeAt(i / chrsz) & mask) << (32 - chrsz - i%32);
  return bin;
}
/*
 * Convert an array of big-endian words to a string
 */
function binb2str(bin){
  var str="";
  var mask=(1 << chrsz) - 1;
  for(var i=0; i < bin.length * 32; i += chrsz)
	str += String.fromCharCode((bin[i>>5] >>> (32 - chrsz - i%32)) & mask);
  return str;
}
/*
 * Convert an array of big-endian words to a hex string.
 */
function binb2hex(binarray){
  var hex_tab=hexcase ? "0123456789ABCDEF" : "0123456789abcdef";
  var str="";
  for(var i=0; i < binarray.length * 4; i++){
	str += hex_tab.charAt((binarray[i>>2] >> ((3 - i%4)*8+4)) & 0xF) +
		   hex_tab.charAt((binarray[i>>2] >> ((3 - i%4)*8  )) & 0xF);
  }
  return str;
}
/*
 * Convert an array of big-endian words to a base-64 string
 */
function binb2b64(binarray){
  var tab="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  var str="";
  for(var i=0; i < binarray.length * 4; i += 3){
	var triplet=(((binarray[i   >> 2] >> 8 * (3 -  i   %4)) & 0xFF) << 16)
				| (((binarray[i+1 >> 2] >> 8 * (3 - (i+1)%4)) & 0xFF) << 8 )
				|  ((binarray[i+2 >> 2] >> 8 * (3 - (i+2)%4)) & 0xFF);
	for(var j=0; j < 4; j++){
	  if(i * 8 + j * 6 > binarray.length * 32) str += b64pad;
	  else str += tab.charAt((triplet >> 6*(3-j)) & 0x3F);
	}
  }
  return str;
}
function SHA256(s){
var chrsz  =8;
var hexcase=0;
function safe_add (x, y) {
var lsw=(x & 0xFFFF) + (y & 0xFFFF);
var msw=(x >> 16) + (y >> 16) + (lsw >> 16);
return (msw << 16) | (lsw & 0xFFFF);
}
function S (X, n) { return ( X >>> n ) | (X << (32 - n)); }
function R (X, n) { return ( X >>> n ); }
function Ch(x, y, z) { return ((x & y) ^ ((~x) & z)); }
function Maj(x, y, z) { return ((x & y) ^ (x & z) ^ (y & z)); }
function Sigma0256(x) { return (S(x, 2) ^ S(x, 13) ^ S(x, 22)); }
function Sigma1256(x) { return (S(x, 6) ^ S(x, 11) ^ S(x, 25)); }
function Gamma0256(x) { return (S(x, 7) ^ S(x, 18) ^ R(x, 3)); }
function Gamma1256(x) { return (S(x, 17) ^ S(x, 19) ^ R(x, 10)); }
function core_sha256 (m, l) {
var K=new Array(0x428A2F98, 0x71374491, 0xB5C0FBCF,
0xE9B5DBA5, 0x3956C25B, 0x59F111F1, 0x923F82A4, 0xAB1C5ED5, 0xD807AA98,
0x12835B01, 0x243185BE, 0x550C7DC3, 0x72BE5D74, 0x80DEB1FE, 0x9BDC06A7,
0xC19BF174, 0xE49B69C1, 0xEFBE4786, 0xFC19DC6, 0x240CA1CC, 0x2DE92C6F,
0x4A7484AA, 0x5CB0A9DC, 0x76F988DA, 0x983E5152, 0xA831C66D, 0xB00327C8,
0xBF597FC7, 0xC6E00BF3, 0xD5A79147, 0x6CA6351, 0x14292967, 0x27B70A85,
0x2E1B2138, 0x4D2C6DFC, 0x53380D13, 0x650A7354, 0x766A0ABB, 0x81C2C92E,
0x92722C85, 0xA2BFE8A1, 0xA81A664B, 0xC24B8B70, 0xC76C51A3, 0xD192E819,
0xD6990624, 0xF40E3585, 0x106AA070, 0x19A4C116, 0x1E376C08, 0x2748774C,
0x34B0BCB5, 0x391C0CB3, 0x4ED8AA4A, 0x5B9CCA4F, 0x682E6FF3, 0x748F82EE,
0x78A5636F, 0x84C87814, 0x8CC70208, 0x90BEFFFA, 0xA4506CEB, 0xBEF9A3F7,
0xC67178F2);
var HASH=new Array(0x6A09E667, 0xBB67AE85, 0x3C6EF372, 0xA54FF53A, 0x510E527F, 0x9B05688C, 0x1F83D9AB, 0x5BE0CD19);
var W=new Array(64);
var a, b, c, d, e, f, g, h, i, j;
var T1, T2;
m[l >> 5] |= 0x80 << (24 - l % 32);
m[((l + 64 >> 9) << 4) + 15]=l;
for ( var i=0; i<m.length; i+=16 ) {
a=HASH[0];
b=HASH[1];
c=HASH[2];
d=HASH[3];
e=HASH[4];
f=HASH[5];
g=HASH[6];
h=HASH[7];
for ( var j=0; j<64; j++) {
if (j < 16) W[j]=m[j + i]; else W[j]=safe_add(safe_add(safe_add(Gamma1256(W[j - 2]), W[j - 7]), Gamma0256(W[j - 15])), W[j - 16]);
T1=safe_add(safe_add(safe_add(safe_add(h, Sigma1256(e)), Ch(e, f, g)), K[j]), W[j]);
T2=safe_add(Sigma0256(a), Maj(a, b, c));
h=g;
g=f;
f=e;
e=safe_add(d, T1);
d=c;
c=b;
b=a;
a=safe_add(T1, T2);
}
HASH[0]=safe_add(a, HASH[0]);
HASH[1]=safe_add(b, HASH[1]);
HASH[2]=safe_add(c, HASH[2]);
HASH[3]=safe_add(d, HASH[3]);
HASH[4]=safe_add(e, HASH[4]);
HASH[5]=safe_add(f, HASH[5]);
HASH[6]=safe_add(g, HASH[6]);
HASH[7]=safe_add(h, HASH[7]);
}
return HASH;
}
function Utf8Encode(string) {
string=string.replace(/\r\n/g,"\n");
var utftext="";
for (var n=0; n < string.length; n++) {
var c=string.charCodeAt(n);
if (c < 128) {
utftext += String.fromCharCode(c);
}
else if((c > 127) && (c < 2048)) {
utftext += String.fromCharCode((c >> 6) | 192);
utftext += String.fromCharCode((c & 63) | 128);
}
else {
utftext += String.fromCharCode((c >> 12) | 224);
utftext += String.fromCharCode(((c >> 6) & 63) | 128);
utftext += String.fromCharCode((c & 63) | 128);
}
}
return utftext;
}
s=Utf8Encode(s);
return binb2hex(core_sha256(str2binb(s), s.length * chrsz));
}
/*
 * A JavaScript implementation of the Secure Hash Algorithm, SHA-256, as defined
 * in FIPS 180-2
 * Version 2.2 Copyright Angel Marin, Paul Johnston 2000 - 2009.
 * Other contributors: Greg Holt, Andrew Kepert, Ydnar, Lostinet
 * Distributed under the BSD License
 * See http://pajhome.org.uk/crypt/md5 for details.
 * Adapted into a WSC for use in classic ASP by Daniel O'Malley
 * (based on an SHA-1 example by Erik Oosterwaal)
 * for use with the Amazon Product Advertising API
 */
function b64_hmac_sha256(k, d){
d=d.replace ( /\s/g, "\n");
return rstr2b64(rstr_hmac_sha256(str2rstr_utf8(k), str2rstr_utf8(d)));
}
function hex_hmac_sha256(k, d){
d=d.replace ( /\s/g, "\n");
return rstr2hex(rstr_hmac_sha256(str2rstr_utf8(k), str2rstr_utf8(d)));
}
/*
 * Calculate the HMAC-sha256 of a key and some data (raw strings)
 */
function rstr_hmac_sha256(key, data)
{
  var bkey=rstr2binb(key);
  if(bkey.length > 16) bkey=binb_sha256(bkey, key.length * 8);
  var ipad=Array(16), opad=Array(16);
  for(var i=0; i < 16; i++)
  {
	ipad[i]=bkey[i] ^ 0x36363636;
	opad[i]=bkey[i] ^ 0x5C5C5C5C;
  }
  var hash=binb_sha256(ipad.concat(rstr2binb(data)), 512 + data.length * 8);
  return binb2rstr(binb_sha256(opad.concat(hash), 512 + 256));
}
function rstr2hex(input)
{
  try { hexcase } catch(e) { hexcase=0; }
  var hex_tab = hexcase ? "0123456789ABCDEF" : "0123456789abcdef";
  var output = "";
  var x;
  for(var i = 0; i < input.length; i++)
  {
    x = input.charCodeAt(i);
    output += hex_tab.charAt((x >>> 4) & 0x0F)
           +  hex_tab.charAt( x        & 0x0F);
  }
  return output;
}
/*
 * Convert a raw string to a base-64 string
 */
function rstr2b64(input)
{
  try { b64pad } catch(e) { b64pad=''; }
  var tab="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  var output="";
  var len=input.length;
  for(var i=0; i < len; i += 3)
  {
	var triplet=(input.charCodeAt(i) << 16)
				| (i + 1 < len ? input.charCodeAt(i+1) << 8 : 0)
				| (i + 2 < len ? input.charCodeAt(i+2)	  : 0);
	for(var j=0; j < 4; j++)
	{
	  if(i * 8 + j * 6 > input.length * 8) output += b64pad;
	  else output += tab.charAt((triplet >>> 6*(3-j)) & 0x3F);
	}
  }
  return output;
}
/*
 * Encode a string as utf-8.
 * For efficiency, this assumes the input is valid utf-16.
 */
function str2rstr_utf8(input)
{
  var output="";
  var i=-1;
  var x, y;
  while(++i < input.length)
  {
	/* Decode utf-16 surrogate pairs */
	x=input.charCodeAt(i);
	y=i + 1 < input.length ? input.charCodeAt(i + 1) : 0;
	if(0xD800 <= x && x <= 0xDBFF && 0xDC00 <= y && y <= 0xDFFF)
	{
	  x=0x10000 + ((x & 0x03FF) << 10) + (y & 0x03FF);
	  i++;
	}
	/* Encode output as utf-8 */
	if(x <= 0x7F)
	  output += String.fromCharCode(x);
	else if(x <= 0x7FF)
	  output += String.fromCharCode(0xC0 | ((x >>> 6 ) & 0x1F),
									0x80 | ( x		 & 0x3F));
	else if(x <= 0xFFFF)
	  output += String.fromCharCode(0xE0 | ((x >>> 12) & 0x0F),
									0x80 | ((x >>> 6 ) & 0x3F),
									0x80 | ( x		 & 0x3F));
	else if(x <= 0x1FFFFF)
	  output += String.fromCharCode(0xF0 | ((x >>> 18) & 0x07),
									0x80 | ((x >>> 12) & 0x3F),
									0x80 | ((x >>> 6 ) & 0x3F),
									0x80 | ( x		 & 0x3F));
  }
  return output;
}
/*
 * Convert a raw string to an array of big-endian words
 * Characters >255 have their high-byte silently ignored.
 */
function rstr2binb(input)
{
  var output=Array(input.length >> 2);
  for(var i=0; i < output.length; i++)
	output[i]=0;
  for(var i=0; i < input.length * 8; i += 8)
	output[i>>5] |= (input.charCodeAt(i / 8) & 0xFF) << (24 - i % 32);
  return output;
}
/*
 * Convert an array of big-endian words to a string
 */
function binb2rstr(input)
{
  var output="";
  for(var i=0; i < input.length * 32; i += 8)
	output += String.fromCharCode((input[i>>5] >>> (24 - i % 32)) & 0xFF);
  return output;
}
/*
 * Main sha256 function, with its support functions
 */
function sha256_S (X, n) {return ( X >>> n ) | (X << (32 - n));}
function sha256_R (X, n) {return ( X >>> n );}
function sha256_Ch(x, y, z) {return ((x & y) ^ ((~x) & z));}
function sha256_Maj(x, y, z) {return ((x & y) ^ (x & z) ^ (y & z));}
function sha256_Sigma0256(x) {return (sha256_S(x, 2) ^ sha256_S(x, 13) ^ sha256_S(x, 22));}
function sha256_Sigma1256(x) {return (sha256_S(x, 6) ^ sha256_S(x, 11) ^ sha256_S(x, 25));}
function sha256_Gamma0256(x) {return (sha256_S(x, 7) ^ sha256_S(x, 18) ^ sha256_R(x, 3));}
function sha256_Gamma1256(x) {return (sha256_S(x, 17) ^ sha256_S(x, 19) ^ sha256_R(x, 10));}
var sha256_K=new Array
(
  1116352408, 1899447441, -1245643825, -373957723, 961987163, 1508970993,
  -1841331548, -1424204075, -670586216, 310598401, 607225278, 1426881987,
  1925078388, -2132889090, -1680079193, -1046744716, -459576895, -272742522,
  264347078, 604807628, 770255983, 1249150122, 1555081692, 1996064986,
  -1740746414, -1473132947, -1341970488, -1084653625, -958395405, -710438585,
  113926993, 338241895, 666307205, 773529912, 1294757372, 1396182291,
  1695183700, 1986661051, -2117940946, -1838011259, -1564481375, -1474664885,
  -1035236496, -949202525, -778901479, -694614492, -200395387, 275423344,
  430227734, 506948616, 659060556, 883997877, 958139571, 1322822218,
  1537002063, 1747873779, 1955562222, 2024104815, -2067236844, -1933114872,
  -1866530822, -1538233109, -1090935817, -965641998
);
function binb_sha256(m, l)
{
  var HASH=new Array(1779033703, -1150833019, 1013904242, -1521486534,
					   1359893119, -1694144372, 528734635, 1541459225);
  var W=new Array(64);
  var a, b, c, d, e, f, g, h;
  var i, j, T1, T2;
  /* append padding */
  m[l >> 5] |= 0x80 << (24 - l % 32);
  m[((l + 64 >> 9) << 4) + 15]=l;
  for(i=0; i < m.length; i += 16)
  {
	a=HASH[0];
	b=HASH[1];
	c=HASH[2];
	d=HASH[3];
	e=HASH[4];
	f=HASH[5];
	g=HASH[6];
	h=HASH[7];
	for(j=0; j < 64; j++)
	{
		if (j < 16) W[j]=m[j + i];
		else W[j]=safe_add(safe_add(safe_add(sha256_Gamma1256(W[j - 2]), W[j - 7]), sha256_Gamma0256(W[j - 15])), W[j - 16]);
		T1=safe_add(safe_add(safe_add(safe_add(h, sha256_Sigma1256(e)), sha256_Ch(e, f, g)), sha256_K[j]), W[j]);
		T2=safe_add(sha256_Sigma0256(a), sha256_Maj(a, b, c));
		h=g;
		g=f;
		f=e;
		e=safe_add(d, T1);
		d=c;
		c=b;
		b=a;
		a=safe_add(T1, T2);
	}
	HASH[0]=safe_add(a, HASH[0]);
	HASH[1]=safe_add(b, HASH[1]);
	HASH[2]=safe_add(c, HASH[2]);
	HASH[3]=safe_add(d, HASH[3]);
	HASH[4]=safe_add(e, HASH[4]);
	HASH[5]=safe_add(f, HASH[5]);
	HASH[6]=safe_add(g, HASH[6]);
	HASH[7]=safe_add(h, HASH[7]);
  }
  return HASH;
}
/**
 * @preserve A JavaScript implementation of the SHA family of hashes, as
 * defined in FIPS PUB 180-4 and FIPS PUB 202, as well as the corresponding
 * HMAC implementation as defined in FIPS PUB 198a
 *
 * Copyright Brian Turek 2008-2017
 * Distributed under the BSD License
 * See http://caligatio.github.com/jsSHA/ for more information
 *
 * Several functions taken from Paul Johnston
 */
	/* Globals */
	var TWO_PWR_32 = 4294967296;
	function Int_64(msint_32, lsint_32)
	{
		this.highOrder = msint_32;
		this.lowOrder = lsint_32;
	}
	function str2packed(str, utfType, existingPacked, existingPackedLen, bigEndianMod)
	{
		var packed, codePnt, codePntArr, byteCnt = 0, i, j, existingByteLen,
			intOffset, byteOffset, shiftModifier, transposeBytes;
		packed = existingPacked || [0];
		existingPackedLen = existingPackedLen || 0;
		existingByteLen = existingPackedLen >>> 3;
		if ("UTF8" === utfType)
		{
			shiftModifier = (bigEndianMod === -1) ? 3 : 0;
			for (i = 0; i < str.length; i += 1)
			{
				codePnt = str.charCodeAt(i);
				codePntArr = [];
				if (0x80 > codePnt)
				{
					codePntArr.push(codePnt);
				}
				else if (0x800 > codePnt)
				{
					codePntArr.push(0xC0 | (codePnt >>> 6));
					codePntArr.push(0x80 | (codePnt & 0x3F));
				}
				else if ((0xd800 > codePnt) || (0xe000 <= codePnt)) {
					codePntArr.push(
						0xe0 | (codePnt >>> 12),
						0x80 | ((codePnt >>> 6) & 0x3f),
						0x80 | (codePnt & 0x3f)
					);
				}
				else
				{
					i += 1;
					codePnt = 0x10000 + (((codePnt & 0x3ff) << 10) | (str.charCodeAt(i) & 0x3ff));
					codePntArr.push(
						0xf0 | (codePnt >>> 18),
						0x80 | ((codePnt >>> 12) & 0x3f),
						0x80 | ((codePnt >>> 6) & 0x3f),
						0x80 | (codePnt & 0x3f)
					);
				}
				for (j = 0; j < codePntArr.length; j += 1)
				{
					byteOffset = byteCnt + existingByteLen;
					intOffset = byteOffset >>> 2;
					while (packed.length <= intOffset)
					{
						packed.push(0);
					}
					/* Known bug kicks in here */
					packed[intOffset] |= codePntArr[j] << (8 * (shiftModifier + bigEndianMod * (byteOffset % 4)));
					byteCnt += 1;
				}
			}
		}
		return {"value" : packed, "binLen" : byteCnt * 8 + existingPackedLen};
	}
	function hex2packed(str, existingPacked, existingPackedLen, bigEndianMod)
	{
		var packed, length = str.length, i, num, intOffset, byteOffset,
			existingByteLen, shiftModifier;
		if (0 !== (length % 2))
		{
			throw new Error("String of HEX type must be in byte increments");
		}
		packed = existingPacked || [0];
		existingPackedLen = existingPackedLen || 0;
		existingByteLen = existingPackedLen >>> 3;
		shiftModifier = (bigEndianMod === -1) ? 3 : 0;
		for (i = 0; i < length; i += 2)
		{
			num = parseInt(str.substr(i, 2), 16);
			if (!isNaN(num))
			{
				byteOffset = (i >>> 1) + existingByteLen;
				intOffset = byteOffset >>> 2;
				while (packed.length <= intOffset)
				{
					packed.push(0);
				}
				packed[intOffset] |= num  << (8 * (shiftModifier + bigEndianMod * (byteOffset % 4)));
			}
			else
			{
				throw new Error("String of HEX type contains invalid characters");
			}
		}
		return {"value" : packed, "binLen" : length * 4 + existingPackedLen};
	}
	function packed2hex(packed, outputLength, bigEndianMod, formatOpts)
	{
		var hex_tab = "0123456789abcdef", str = "",
			length = outputLength / 8, i, srcByte, shiftModifier;
		shiftModifier = (bigEndianMod === -1) ? 3 : 0;
		for (i = 0; i < length; i += 1)
		{
			/* The below is more than a byte but it gets taken care of later */
			srcByte = packed[i >>> 2] >>> (8 * (shiftModifier + bigEndianMod * (i % 4)));
			str += hex_tab.charAt((srcByte >>> 4) & 0xF) +
				hex_tab.charAt(srcByte & 0xF);
		}
		return (formatOpts["outputUpper"]) ? str.toUpperCase() : str;
	}
	function getOutputOpts(options)
	{
		var retVal = {"outputUpper" : false, "b64Pad" : "=", "shakeLen" : -1},
			outputOptions;
		outputOptions = options || {};
		retVal["outputUpper"] = outputOptions["outputUpper"] || false;
		if (true === outputOptions.hasOwnProperty("b64Pad"))
		{
			retVal["b64Pad"] = outputOptions["b64Pad"];
		}
		if ("boolean" !== typeof(retVal["outputUpper"]))
		{
			throw new Error("Invalid outputUpper formatting option");
		}
		if ("string" !== typeof(retVal["b64Pad"]))
		{
			throw new Error("Invalid b64Pad formatting option");
		}
		return retVal;
	}
	function getStrConverter(format, utfType, bigEndianMod)
	{
		var retVal;
		switch (format)
		{
		case "HEX":
			retVal = function(str, existingBin, existingBinLen)
				{
				   return hex2packed(str, existingBin, existingBinLen, bigEndianMod);
				};
			break;
		case "TEXT":
			retVal = function(str, existingBin, existingBinLen)
				{
					return str2packed(str, utfType, existingBin, existingBinLen, bigEndianMod);
				};
			break;
		default:
			throw new Error("format must be HEX, TEXT, B64, BYTES, or ARRAYBUFFER");
		}
		return retVal;
	}
	function rotr_64(x, n)
	{
		var retVal = null, tmp = new Int_64(x.highOrder, x.lowOrder);
		if (32 >= n)
		{
			retVal = new Int_64(
					(tmp.highOrder >>> n) | ((tmp.lowOrder << (32 - n)) & 0xFFFFFFFF),
					(tmp.lowOrder >>> n) | ((tmp.highOrder << (32 - n)) & 0xFFFFFFFF)
				);
		}
		else
		{
			retVal = new Int_64(
					(tmp.lowOrder >>> (n - 32)) | ((tmp.highOrder << (64 - n)) & 0xFFFFFFFF),
					(tmp.highOrder >>> (n - 32)) | ((tmp.lowOrder << (64 - n)) & 0xFFFFFFFF)
				);
		}
		return retVal;
	}
	function shr_64(x, n)
	{
		var retVal = null;
		if (32 >= n)
		{
			retVal = new Int_64(
					x.highOrder >>> n,
					x.lowOrder >>> n | ((x.highOrder << (32 - n)) & 0xFFFFFFFF)
				);
		}
		else
		{
			retVal = new Int_64(
					0,
					x.highOrder >>> (n - 32)
				);
		}
		return retVal;
	}
	function ch_64(x, y, z)
	{
		return new Int_64(
				(x.highOrder & y.highOrder) ^ (~x.highOrder & z.highOrder),
				(x.lowOrder & y.lowOrder) ^ (~x.lowOrder & z.lowOrder)
			);
	}
	function maj_64(x, y, z)
	{
		return new Int_64(
				(x.highOrder & y.highOrder) ^
				(x.highOrder & z.highOrder) ^
				(y.highOrder & z.highOrder),
				(x.lowOrder & y.lowOrder) ^
				(x.lowOrder & z.lowOrder) ^
				(y.lowOrder & z.lowOrder)
			);
	}
	function sigma0_64(x)
	{
		var rotr28 = rotr_64(x, 28), rotr34 = rotr_64(x, 34),
			rotr39 = rotr_64(x, 39);
		return new Int_64(
				rotr28.highOrder ^ rotr34.highOrder ^ rotr39.highOrder,
				rotr28.lowOrder ^ rotr34.lowOrder ^ rotr39.lowOrder);
	}
	function sigma1_64(x)
	{
		var rotr14 = rotr_64(x, 14), rotr18 = rotr_64(x, 18),
			rotr41 = rotr_64(x, 41);
		return new Int_64(
				rotr14.highOrder ^ rotr18.highOrder ^ rotr41.highOrder,
				rotr14.lowOrder ^ rotr18.lowOrder ^ rotr41.lowOrder);
	}
	function gamma0_64(x)
	{
		var rotr1 = rotr_64(x, 1), rotr8 = rotr_64(x, 8), shr7 = shr_64(x, 7);
		return new Int_64(
				rotr1.highOrder ^ rotr8.highOrder ^ shr7.highOrder,
				rotr1.lowOrder ^ rotr8.lowOrder ^ shr7.lowOrder
			);
	}
	function gamma1_64(x)
	{
		var rotr19 = rotr_64(x, 19), rotr61 = rotr_64(x, 61),
			shr6 = shr_64(x, 6);
		return new Int_64(
				rotr19.highOrder ^ rotr61.highOrder ^ shr6.highOrder,
				rotr19.lowOrder ^ rotr61.lowOrder ^ shr6.lowOrder
			);
	}
	function safeAdd_64_2(x, y)
	{
		var lsw, msw, lowOrder, highOrder;
		lsw = (x.lowOrder & 0xFFFF) + (y.lowOrder & 0xFFFF);
		msw = (x.lowOrder >>> 16) + (y.lowOrder >>> 16) + (lsw >>> 16);
		lowOrder = ((msw & 0xFFFF) << 16) | (lsw & 0xFFFF);
		lsw = (x.highOrder & 0xFFFF) + (y.highOrder & 0xFFFF) + (msw >>> 16);
		msw = (x.highOrder >>> 16) + (y.highOrder >>> 16) + (lsw >>> 16);
		highOrder = ((msw & 0xFFFF) << 16) | (lsw & 0xFFFF);
		return new Int_64(highOrder, lowOrder);
	}
	function safeAdd_64_4(a, b, c, d)
	{
		var lsw, msw, lowOrder, highOrder;
		lsw = (a.lowOrder & 0xFFFF) + (b.lowOrder & 0xFFFF) +
			(c.lowOrder & 0xFFFF) + (d.lowOrder & 0xFFFF);
		msw = (a.lowOrder >>> 16) + (b.lowOrder >>> 16) +
			(c.lowOrder >>> 16) + (d.lowOrder >>> 16) + (lsw >>> 16);
		lowOrder = ((msw & 0xFFFF) << 16) | (lsw & 0xFFFF);
		lsw = (a.highOrder & 0xFFFF) + (b.highOrder & 0xFFFF) +
			(c.highOrder & 0xFFFF) + (d.highOrder & 0xFFFF) + (msw >>> 16);
		msw = (a.highOrder >>> 16) + (b.highOrder >>> 16) +
			(c.highOrder >>> 16) + (d.highOrder >>> 16) + (lsw >>> 16);
		highOrder = ((msw & 0xFFFF) << 16) | (lsw & 0xFFFF);
		return new Int_64(highOrder, lowOrder);
	}
	function safeAdd_64_5(a, b, c, d, e)
	{
		var lsw, msw, lowOrder, highOrder;
		lsw = (a.lowOrder & 0xFFFF) + (b.lowOrder & 0xFFFF) +
			(c.lowOrder & 0xFFFF) + (d.lowOrder & 0xFFFF) +
			(e.lowOrder & 0xFFFF);
		msw = (a.lowOrder >>> 16) + (b.lowOrder >>> 16) +
			(c.lowOrder >>> 16) + (d.lowOrder >>> 16) + (e.lowOrder >>> 16) +
			(lsw >>> 16);
		lowOrder = ((msw & 0xFFFF) << 16) | (lsw & 0xFFFF);
		lsw = (a.highOrder & 0xFFFF) + (b.highOrder & 0xFFFF) +
			(c.highOrder & 0xFFFF) + (d.highOrder & 0xFFFF) +
			(e.highOrder & 0xFFFF) + (msw >>> 16);
		msw = (a.highOrder >>> 16) + (b.highOrder >>> 16) +
			(c.highOrder >>> 16) + (d.highOrder >>> 16) +
			(e.highOrder >>> 16) + (lsw >>> 16);
		highOrder = ((msw & 0xFFFF) << 16) | (lsw & 0xFFFF);
		return new Int_64(highOrder, lowOrder);
	}
	function getNewState(variant)
	{
		var retVal = [], H_trunc, H_full, i;
		H_trunc = [
			0xc1059ed8, 0x367cd507, 0x3070dd17, 0xf70e5939,
			0xffc00b31, 0x68581511, 0x64f98fa7, 0xbefa4fa4
		];
		H_full = [
			0x6A09E667, 0xBB67AE85, 0x3C6EF372, 0xA54FF53A,
			0x510E527F, 0x9B05688C, 0x1F83D9AB, 0x5BE0CD19
		];
		retVal = [
			new Int_64(H_full[0], 0xf3bcc908),
			new Int_64(H_full[1], 0x84caa73b),
			new Int_64(H_full[2], 0xfe94f82b),
			new Int_64(H_full[3], 0x5f1d36f1),
			new Int_64(H_full[4], 0xade682d1),
			new Int_64(H_full[5], 0x2b3e6c1f),
			new Int_64(H_full[6], 0xfb41bd6b),
			new Int_64(H_full[7], 0x137e2179)
		];
		return retVal;
	}
	K_sha2 = [
		0x428A2F98, 0x71374491, 0xB5C0FBCF, 0xE9B5DBA5,
		0x3956C25B, 0x59F111F1, 0x923F82A4, 0xAB1C5ED5,
		0xD807AA98, 0x12835B01, 0x243185BE, 0x550C7DC3,
		0x72BE5D74, 0x80DEB1FE, 0x9BDC06A7, 0xC19BF174,
		0xE49B69C1, 0xEFBE4786, 0x0FC19DC6, 0x240CA1CC,
		0x2DE92C6F, 0x4A7484AA, 0x5CB0A9DC, 0x76F988DA,
		0x983E5152, 0xA831C66D, 0xB00327C8, 0xBF597FC7,
		0xC6E00BF3, 0xD5A79147, 0x06CA6351, 0x14292967,
		0x27B70A85, 0x2E1B2138, 0x4D2C6DFC, 0x53380D13,
		0x650A7354, 0x766A0ABB, 0x81C2C92E, 0x92722C85,
		0xA2BFE8A1, 0xA81A664B, 0xC24B8B70, 0xC76C51A3,
		0xD192E819, 0xD6990624, 0xF40E3585, 0x106AA070,
		0x19A4C116, 0x1E376C08, 0x2748774C, 0x34B0BCB5,
		0x391C0CB3, 0x4ED8AA4A, 0x5B9CCA4F, 0x682E6FF3,
		0x748F82EE, 0x78A5636F, 0x84C87814, 0x8CC70208,
		0x90BEFFFA, 0xA4506CEB, 0xBEF9A3F7, 0xC67178F2
	];
	K_sha512 = [
		new Int_64(K_sha2[ 0], 0xd728ae22), new Int_64(K_sha2[ 1], 0x23ef65cd),
		new Int_64(K_sha2[ 2], 0xec4d3b2f), new Int_64(K_sha2[ 3], 0x8189dbbc),
		new Int_64(K_sha2[ 4], 0xf348b538), new Int_64(K_sha2[ 5], 0xb605d019),
		new Int_64(K_sha2[ 6], 0xaf194f9b), new Int_64(K_sha2[ 7], 0xda6d8118),
		new Int_64(K_sha2[ 8], 0xa3030242), new Int_64(K_sha2[ 9], 0x45706fbe),
		new Int_64(K_sha2[10], 0x4ee4b28c), new Int_64(K_sha2[11], 0xd5ffb4e2),
		new Int_64(K_sha2[12], 0xf27b896f), new Int_64(K_sha2[13], 0x3b1696b1),
		new Int_64(K_sha2[14], 0x25c71235), new Int_64(K_sha2[15], 0xcf692694),
		new Int_64(K_sha2[16], 0x9ef14ad2), new Int_64(K_sha2[17], 0x384f25e3),
		new Int_64(K_sha2[18], 0x8b8cd5b5), new Int_64(K_sha2[19], 0x77ac9c65),
		new Int_64(K_sha2[20], 0x592b0275), new Int_64(K_sha2[21], 0x6ea6e483),
		new Int_64(K_sha2[22], 0xbd41fbd4), new Int_64(K_sha2[23], 0x831153b5),
		new Int_64(K_sha2[24], 0xee66dfab), new Int_64(K_sha2[25], 0x2db43210),
		new Int_64(K_sha2[26], 0x98fb213f), new Int_64(K_sha2[27], 0xbeef0ee4),
		new Int_64(K_sha2[28], 0x3da88fc2), new Int_64(K_sha2[29], 0x930aa725),
		new Int_64(K_sha2[30], 0xe003826f), new Int_64(K_sha2[31], 0x0a0e6e70),
		new Int_64(K_sha2[32], 0x46d22ffc), new Int_64(K_sha2[33], 0x5c26c926),
		new Int_64(K_sha2[34], 0x5ac42aed), new Int_64(K_sha2[35], 0x9d95b3df),
		new Int_64(K_sha2[36], 0x8baf63de), new Int_64(K_sha2[37], 0x3c77b2a8),
		new Int_64(K_sha2[38], 0x47edaee6), new Int_64(K_sha2[39], 0x1482353b),
		new Int_64(K_sha2[40], 0x4cf10364), new Int_64(K_sha2[41], 0xbc423001),
		new Int_64(K_sha2[42], 0xd0f89791), new Int_64(K_sha2[43], 0x0654be30),
		new Int_64(K_sha2[44], 0xd6ef5218), new Int_64(K_sha2[45], 0x5565a910),
		new Int_64(K_sha2[46], 0x5771202a), new Int_64(K_sha2[47], 0x32bbd1b8),
		new Int_64(K_sha2[48], 0xb8d2d0c8), new Int_64(K_sha2[49], 0x5141ab53),
		new Int_64(K_sha2[50], 0xdf8eeb99), new Int_64(K_sha2[51], 0xe19b48a8),
		new Int_64(K_sha2[52], 0xc5c95a63), new Int_64(K_sha2[53], 0xe3418acb),
		new Int_64(K_sha2[54], 0x7763e373), new Int_64(K_sha2[55], 0xd6b2b8a3),
		new Int_64(K_sha2[56], 0x5defb2fc), new Int_64(K_sha2[57], 0x43172f60),
		new Int_64(K_sha2[58], 0xa1f0ab72), new Int_64(K_sha2[59], 0x1a6439ec),
		new Int_64(K_sha2[60], 0x23631e28), new Int_64(K_sha2[61], 0xde82bde9),
		new Int_64(K_sha2[62], 0xb2c67915), new Int_64(K_sha2[63], 0xe372532b),
		new Int_64(0xca273ece, 0xea26619c), new Int_64(0xd186b8c7, 0x21c0c207),
		new Int_64(0xeada7dd6, 0xcde0eb1e), new Int_64(0xf57d4f7f, 0xee6ed178),
		new Int_64(0x06f067aa, 0x72176fba), new Int_64(0x0a637dc5, 0xa2c898a6),
		new Int_64(0x113f9804, 0xbef90dae), new Int_64(0x1b710b35, 0x131c471b),
		new Int_64(0x28db77f5, 0x23047d84), new Int_64(0x32caab7b, 0x40c72493),
		new Int_64(0x3c9ebe0a, 0x15c9bebc), new Int_64(0x431d67c4, 0x9c100d4c),
		new Int_64(0x4cc5d4be, 0xcb3e42b6), new Int_64(0x597f299c, 0xfc657e2a),
		new Int_64(0x5fcb6fab, 0x3ad6faec), new Int_64(0x6c44198c, 0x4a475817)
	];
	function roundSHA2(block, H, variant)
	{
		var a, b, c, d, e, f, g, h, T1, T2, numRounds, t, binaryStringMult,
			safeAdd_2, safeAdd_4, safeAdd_5, gamma0, gamma1, sigma0, sigma1,
			ch, maj, Int, W = [], int1, int2, offset, K;
		/* 64-bit variant */
		numRounds = 80;
		binaryStringMult = 2;
		Int = Int_64;
		safeAdd_2 = safeAdd_64_2;
		safeAdd_4 = safeAdd_64_4;
		safeAdd_5 = safeAdd_64_5;
		gamma0 = gamma0_64;
		gamma1 = gamma1_64;
		sigma0 = sigma0_64;
		sigma1 = sigma1_64;
		maj = maj_64;
		ch = ch_64;
		K = K_sha512;
		a = H[0];
		b = H[1];
		c = H[2];
		d = H[3];
		e = H[4];
		f = H[5];
		g = H[6];
		h = H[7];
		for (t = 0; t < numRounds; t += 1)
		{
			if (t < 16)
			{
				offset = t * binaryStringMult;
				int1 = (block.length <= offset) ? 0 : block[offset];
				int2 = (block.length <= offset + 1) ? 0 : block[offset + 1];
				/* Bit of a hack - for 32-bit, the second term is ignored */
				W[t] = new Int(int1, int2);
			}
			else
			{
				W[t] = safeAdd_4(
						gamma1(W[t - 2]), W[t - 7],
						gamma0(W[t - 15]), W[t - 16]
					);
			}
			T1 = safeAdd_5(h, sigma1(e), ch(e, f, g), K[t], W[t]);
			T2 = safeAdd_2(sigma0(a), maj(a, b, c));
			h = g;
			g = f;
			f = e;
			e = safeAdd_2(d, T1);
			d = c;
			c = b;
			b = a;
			a = safeAdd_2(T1, T2);
		}
		H[0] = safeAdd_2(a, H[0]);
		H[1] = safeAdd_2(b, H[1]);
		H[2] = safeAdd_2(c, H[2]);
		H[3] = safeAdd_2(d, H[3]);
		H[4] = safeAdd_2(e, H[4]);
		H[5] = safeAdd_2(f, H[5]);
		H[6] = safeAdd_2(g, H[6]);
		H[7] = safeAdd_2(h, H[7]);
		return H;
	}
	function finalizeSHA2(remainder, remainderBinLen, processedBinLen, H, variant, outputLen)
	{
		var i, appendedMessageLength, offset, retVal, binaryStringInc, totalLen;
		/* 64-bit variant */
		/* The 129 addition is a hack but it works.  The correct number is
		   actually 136 (128 + 8) but the below math fails if
		   remainderBinLen + 136 % 1024 = 0. Since remainderBinLen % 8 = 0,
		   "shorting" the addition is OK. */
		offset = (((remainderBinLen + 129) >>> 10) << 5) + 31;
		binaryStringInc = 32;
		while (remainder.length <= offset)
		{
			remainder.push(0);
		}
		/* Append '1' at the end of the binary string */
		remainder[remainderBinLen >>> 5] |= 0x80 << (24 - remainderBinLen % 32);
		/* Append length of binary string in the position such that the new
		 * length is correct. JavaScript numbers are limited to 2^53 so it's
		 * "safe" to treat the totalLen as a 64-bit integer. */
		totalLen = remainderBinLen + processedBinLen;
		remainder[offset] = totalLen & 0xFFFFFFFF;
		/* Bitwise operators treat the operand as a 32-bit number so need to
		 * use hacky division and round to get access to upper 32-ish bits */
		remainder[offset - 1] = (totalLen / TWO_PWR_32) | 0;
		appendedMessageLength = remainder.length;
		/* This will always be at least 1 full chunk */
		for (i = 0; i < appendedMessageLength; i += binaryStringInc)
		{
			H = roundSHA2(remainder.slice(i, i + binaryStringInc), H, variant);
		}
		retVal = [
			H[0].highOrder, H[0].lowOrder,
			H[1].highOrder, H[1].lowOrder,
			H[2].highOrder, H[2].lowOrder,
			H[3].highOrder, H[3].lowOrder,
			H[4].highOrder, H[4].lowOrder,
			H[5].highOrder, H[5].lowOrder,
			H[6].highOrder, H[6].lowOrder,
			H[7].highOrder, H[7].lowOrder
		];
		return retVal;
	}
	var jsSHA = function(variant,inputFormat,keyInputFormat,options)
	{
		var processedLen = 0, remainder = [], remainderLen = 0, utfType,
			intermediateState, converterFunc, shaVariant = variant, outputBinLen,
			variantBlockSize, roundFunc, finalizeFunc, stateCloneFunc,
			hmacKeySet = false, keyWithIPad = [], keyWithOPad = [], numRounds,
			updatedCalled = false, inputOptions, isSHAKE = false, bigEndianMod = -1;
		inputOptions = options || {};
		utfType = inputOptions["encoding"] || "UTF8";
		numRounds = inputOptions["numRounds"] || 1;
		if ((numRounds !== parseInt(numRounds, 10)) || (1 > numRounds))
		{
			throw new Error("numRounds must a integer >= 1");
		}
		roundFunc = function (block, H) {
			return roundSHA2(block, H, shaVariant);
		};
		finalizeFunc = function (remainder, remainderBinLen, processedBinLen, H, outputLen)
		{
			return finalizeSHA2(remainder, remainderBinLen, processedBinLen, H, shaVariant, outputLen);
		};
		stateCloneFunc = function(state) { return state.slice(); };
		variantBlockSize = 1024;
		outputBinLen = 512;
		converterFunc = getStrConverter(inputFormat, utfType, bigEndianMod);
		intermediateState = getNewState(shaVariant);
		this.setHMACKey = function(key)
		{
			var keyConverterFunc, convertRet, keyBinLen, keyToUse, blockByteSize,
				i, lastArrayIndex, keyOptions;
			if (true === hmacKeySet)
			{
				throw new Error("HMAC key already set");
			}
			if (true === updatedCalled)
			{
				throw new Error("Cannot set HMAC key after calling update");
			}
			keyConverterFunc = getStrConverter(keyInputFormat,"UTF8",bigEndianMod);
			convertRet = keyConverterFunc(key);
			keyBinLen = convertRet["binLen"];
			keyToUse = convertRet["value"];
			blockByteSize = variantBlockSize >>> 3;
			/* These are used multiple times, calculate and store them */
			lastArrayIndex = (blockByteSize / 4) - 1;
			/* Figure out what to do with the key based on its size relative to
			 * the hash's block size */
			if (blockByteSize < (keyBinLen / 8))
			{
				keyToUse = finalizeFunc(keyToUse, keyBinLen, 0,getNewState(shaVariant), outputBinLen);
				/* For all variants, the block size is bigger than the output
				 * size so there will never be a useful byte at the end of the
				 * string */
				while (keyToUse.length <= lastArrayIndex)
				{
					keyToUse.push(0);
				}
				keyToUse[lastArrayIndex] &= 0xFFFFFF00;
			}
			else if (blockByteSize > (keyBinLen / 8))
			{
				/* If the blockByteSize is greater than the key length, there
				 * will always be at LEAST one "useless" byte at the end of the
				 * string */
				while (keyToUse.length <= lastArrayIndex)
				{
					keyToUse.push(0);
				}
				keyToUse[lastArrayIndex] &= 0xFFFFFF00;
			}
			/* Create ipad and opad */
			for (i = 0; i <= lastArrayIndex; i += 1)
			{
				keyWithIPad[i] = keyToUse[i] ^ 0x36363636;
				keyWithOPad[i] = keyToUse[i] ^ 0x5C5C5C5C;
			}
			intermediateState = roundFunc(keyWithIPad, intermediateState);
			processedLen = variantBlockSize;
			hmacKeySet = true;
		};
		this.update = function(srcString)
		{
			var convertRet, chunkBinLen, chunkIntLen, chunk, i, updateProcessedLen = 0,
				variantBlockIntInc = variantBlockSize >>> 5;
			convertRet = converterFunc(srcString, remainder, remainderLen);
			chunkBinLen = convertRet["binLen"];
			chunk = convertRet["value"];
			chunkIntLen = chunkBinLen >>> 5;
			for (i = 0; i < chunkIntLen; i += variantBlockIntInc)
			{
				if (updateProcessedLen + variantBlockSize <= chunkBinLen)
				{
					intermediateState = roundFunc(
						chunk.slice(i, i + variantBlockIntInc),
						intermediateState
					);
					updateProcessedLen += variantBlockSize;
				}
			}
			processedLen += updateProcessedLen;
			remainder = chunk.slice(updateProcessedLen >>> 5);
			remainderLen = chunkBinLen % variantBlockSize;
			updatedCalled = true;
		};
		this.getHash = function(format, options)
		{
			var formatFunc, i, outputOptions, finalizedState;
			if (true === hmacKeySet)
			{
				throw new Error("Cannot call getHash after setting HMAC key");
			}
			outputOptions = getOutputOpts(options);
			formatFunc = function(binarray) {return packed2hex(binarray, outputBinLen, bigEndianMod, outputOptions);};
			finalizedState = finalizeFunc(remainder.slice(), remainderLen, processedLen, stateCloneFunc(intermediateState), outputBinLen);
			for (i = 1; i < numRounds; i += 1)
			{
				finalizedState = finalizeFunc(finalizedState, outputBinLen, 0, getNewState(shaVariant), outputBinLen);
			}
			return formatFunc(finalizedState);
		};
		this.getHMAC = function(format, options)
		{
			var formatFunc,	firstHash, outputOptions, finalizedState;
			if (false === hmacKeySet)
			{
				throw new Error("Cannot call getHMAC without first setting HMAC key");
			}
			outputOptions = getOutputOpts(options);
			formatFunc = function(binarray) {return packed2hex(binarray, outputBinLen, bigEndianMod, outputOptions);};
			firstHash = finalizeFunc(remainder.slice(), remainderLen, processedLen, stateCloneFunc(intermediateState), outputBinLen);
			finalizedState = roundFunc(keyWithOPad, getNewState(shaVariant));
			finalizedState = finalizeFunc(firstHash, outputBinLen, variantBlockSize, finalizedState, outputBinLen);
			return formatFunc(finalizedState);
		};
	};
	function calcHMACSha512(tkey,tmessage,informat,keyformat) {
			var hmacObj = new jsSHA("SHA-512",informat,keyformat);
			hmacObj.setHMACKey(tkey);
			hmacObj.update(tmessage);
			return( hmacObj.getHMAC("HEX"));
	}
</script>
<%
Private m_lOnBits(30)
Private m_l2Power(30)
Private m_bytOnBits(7)
Private m_byt2Power(7)
Private m_InCo(3)
Private m_fbsub(255)
Private m_rbsub(255)
Private m_ptab(255)
Private m_ltab(255)
Private m_ftable(255)
Private m_rtable(255)
Private m_rco(29)
Private m_Nk
Private m_Nb
Private m_Nr
Private m_fi(23)
Private m_ri(23)
Private m_fkey(119)
Private m_rkey(119)
m_InCo(0)=&HB
m_InCo(1)=&HD
m_InCo(2)=&H9
m_InCo(3)=&HE
m_bytOnBits(0)=1
m_bytOnBits(1)=3
m_bytOnBits(2)=7
m_bytOnBits(3)=15
m_bytOnBits(4)=31
m_bytOnBits(5)=63
m_bytOnBits(6)=127
m_bytOnBits(7)=255
m_byt2Power(0)=1
m_byt2Power(1)=2
m_byt2Power(2)=4
m_byt2Power(3)=8
m_byt2Power(4)=16
m_byt2Power(5)=32
m_byt2Power(6)=64
m_byt2Power(7)=128
m_lOnBits(0)=1
m_lOnBits(1)=3
m_lOnBits(2)=7
m_lOnBits(3)=15
m_lOnBits(4)=31
m_lOnBits(5)=63
m_lOnBits(6)=127
m_lOnBits(7)=255
m_lOnBits(8)=511
m_lOnBits(9)=1023
m_lOnBits(10)=2047
m_lOnBits(11)=4095
m_lOnBits(12)=8191
m_lOnBits(13)=16383
m_lOnBits(14)=32767
m_lOnBits(15)=65535
m_lOnBits(16)=131071
m_lOnBits(17)=262143
m_lOnBits(18)=524287
m_lOnBits(19)=1048575
m_lOnBits(20)=2097151
m_lOnBits(21)=4194303
m_lOnBits(22)=8388607
m_lOnBits(23)=16777215
m_lOnBits(24)=33554431
m_lOnBits(25)=67108863
m_lOnBits(26)=134217727
m_lOnBits(27)=268435455
m_lOnBits(28)=536870911
m_lOnBits(29)=1073741823
m_lOnBits(30)=2147483647
m_l2Power(0)=1
m_l2Power(1)=2
m_l2Power(2)=4
m_l2Power(3)=8
m_l2Power(4)=16
m_l2Power(5)=32
m_l2Power(6)=64
m_l2Power(7)=128
m_l2Power(8)=256
m_l2Power(9)=512
m_l2Power(10)=1024
m_l2Power(11)=2048
m_l2Power(12)=4096
m_l2Power(13)=8192
m_l2Power(14)=16384
m_l2Power(15)=32768
m_l2Power(16)=65536
m_l2Power(17)=131072
m_l2Power(18)=262144
m_l2Power(19)=524288
m_l2Power(20)=1048576
m_l2Power(21)=2097152
m_l2Power(22)=4194304
m_l2Power(23)=8388608
m_l2Power(24)=16777216
m_l2Power(25)=33554432
m_l2Power(26)=67108864
m_l2Power(27)=134217728
m_l2Power(28)=268435456
m_l2Power(29)=536870912
m_l2Power(30)=1073741824
Private Function LShift(lValue, iShiftBits)
	If iShiftBits=0 Then
		LShift=lValue
		Exit Function
	ElseIf iShiftBits=31 Then
		If lValue And 1 Then
			LShift=&H80000000
		Else
			LShift=0
		End If
		Exit Function
	ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
		Err.Raise 6
	End If
	If (lValue And m_l2Power(31 - iShiftBits)) Then
		LShift=((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
	Else
		LShift=((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
	End If
End Function
Private Function RShift(lValue, iShiftBits)
	If iShiftBits=0 Then
		RShift=lValue
		Exit Function
	ElseIf iShiftBits=31 Then
		If lValue And &H80000000 Then
			RShift=1
		Else
			RShift=0
		End If
		Exit Function
	ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
		Err.Raise 6
	End If
	RShift=(lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
	If (lValue And &H80000000) Then
		RShift=(RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
	End If
End Function
Private Function LShiftByte(bytValue, bytShiftBits)
	If bytShiftBits=0 Then
		LShiftByte=bytValue
		Exit Function
	ElseIf bytShiftBits=7 Then
		If bytValue And 1 Then
			LShiftByte=&H80
		Else
			LShiftByte=0
		End If
		Exit Function
	ElseIf bytShiftBits < 0 Or bytShiftBits > 7 Then
		Err.Raise 6
	End If
	LShiftByte=((bytValue And m_bytOnBits(7 - bytShiftBits)) * m_byt2Power(bytShiftBits))
End Function
Private Function RShiftByte(bytValue, bytShiftBits)
	If bytShiftBits=0 Then
		RShiftByte=bytValue
		Exit Function
	ElseIf bytShiftBits=7 Then
		If bytValue And &H80 Then
			RShiftByte=1
		Else
			RShiftByte=0
		End If
		Exit Function
	ElseIf bytShiftBits < 0 Or bytShiftBits > 7 Then
		Err.Raise 6
	End If
	RShiftByte=bytValue \ m_byt2Power(bytShiftBits)
End Function
Private Function RotateLeft(lValue, iShiftBits)
	RotateLeft=LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function
Private Function RotateLeftByte(bytValue, bytShiftBits)
	RotateLeftByte=LShiftByte(bytValue, bytShiftBits) Or RShiftByte(bytValue, (8 - bytShiftBits))
End Function
Private Function Pack(b())
	Dim lCount, lTemp
	For lCount=0 To 3
		lTemp=b(lCount)
		Pack=Pack Or LShift(lTemp, (lCount * 8))
	Next
End Function
Private Function PackFrom(b(), k)
	Dim lCount, lTemp
	For lCount=0 To 3
		lTemp=b(lCount + k)
		PackFrom=PackFrom Or LShift(lTemp, (lCount * 8))
	Next
End Function
Private Sub Unpack(a, b())
	b(0)=a And m_lOnBits(7)
	b(1)=RShift(a, 8) And m_lOnBits(7)
	b(2)=RShift(a, 16) And m_lOnBits(7)
	b(3)=RShift(a, 24) And m_lOnBits(7)
End Sub
Private Sub UnpackFrom(a, b(), k)
	b(0 + k)=a And m_lOnBits(7)
	b(1 + k)=RShift(a, 8) And m_lOnBits(7)
	b(2 + k)=RShift(a, 16) And m_lOnBits(7)
	b(3 + k)=RShift(a, 24) And m_lOnBits(7)
End Sub
Private Function xtime(a)
	Dim b
	if (a And &H80) then b=&H1B else b=0
	xtime=LShiftByte(a, 1)
	xtime=xtime Xor b
End Function
Private Function bmul(x, y)
	if x <> 0 And y <> 0 Then bmul=m_ptab((CLng(m_ltab(x)) + CLng(m_ltab(y))) Mod 255) Else bmul=0
End Function
Private Function SubByte(a)
	Dim b(3)
	Unpack a, b
	b(0)=m_fbsub(b(0))
	b(1)=m_fbsub(b(1))
	b(2)=m_fbsub(b(2))
	b(3)=m_fbsub(b(3))
	SubByte=Pack(b)
End Function
Private Function product(x, y)
	Dim xb(3), yb(3)
	Unpack x, xb
	Unpack y, yb
	product=bmul(xb(0), yb(0)) Xor bmul(xb(1), yb(1)) Xor bmul(xb(2), yb(2)) Xor bmul(xb(3), yb(3))
End Function
Private Function InvMixCol(x)
	Dim y, m, b(3)
	m=Pack(m_InCo)
	b(3)=product(m, x)
	m=RotateLeft(m, 24)
	b(2)=product(m, x)
	m=RotateLeft(m, 24)
	b(1)=product(m, x)
	m=RotateLeft(m, 24)
	b(0)=product(m, x)
	y=Pack(b)
	InvMixCol=y
End Function
Private Function ByteSub(x)
	Dim y, z
	z=x
	y=m_ptab(255 - m_ltab(z))
	z=y
	z=RotateLeftByte(z, 1)
	y=y Xor z
	z=RotateLeftByte(z, 1)
	y=y Xor z
	z=RotateLeftByte(z, 1)
	y=y Xor z
	z=RotateLeftByte(z, 1)
	y=y Xor z
	y=y Xor &H63
	ByteSub=y
End Function
Public Sub gentables()
	Dim i, y, b(3), ib
	m_ltab(0)=0
	m_ptab(0)=1
	m_ltab(1)=0
	m_ptab(1)=3
	m_ltab(3)=1
	For i=2 To 255
		m_ptab(i)=m_ptab(i - 1) Xor xtime(m_ptab(i - 1))
		m_ltab(m_ptab(i))=i
	Next
	m_fbsub(0)=&H63
	m_rbsub(&H63)=0
	For i=1 To 255
		ib=i
		y=ByteSub(ib)
		m_fbsub(i)=y
		m_rbsub(y)=i
	Next
	y=1
	For i=0 To 29
		m_rco(i)=y
		y=xtime(y)
	Next
	For i=0 To 255
		y=m_fbsub(i)
		b(3)=y Xor xtime(y)
		b(2)=y
		b(1)=y
		b(0)=xtime(y)
		m_ftable(i)=Pack(b)
		y=m_rbsub(i)
		b(3)=bmul(m_InCo(0), y)
		b(2)=bmul(m_InCo(1), y)
		b(1)=bmul(m_InCo(2), y)
		b(0)=bmul(m_InCo(3), y)
		m_rtable(i)=Pack(b)
	Next
End Sub
Public Sub gkey(nb, nk, key())
	Dim i,j,k,m,N,C1,C2,C3,CipherKey(7)
	m_Nb=nb
	m_Nk=nk
	If m_Nb >= m_Nk Then m_Nr=6 + m_Nb Else m_Nr=6 + m_Nk
	C1=1
	If m_Nb < 8 Then
		C2=2
		C3=3
	Else
		C2=3
		C3=4
	End If
	For j=0 To nb - 1
		m=j * 3
		m_fi(m)=(j + C1) Mod nb
		m_fi(m + 1)=(j + C2) Mod nb
		m_fi(m + 2)=(j + C3) Mod nb
		m_ri(m)=(nb + j - C1) Mod nb
		m_ri(m + 1)=(nb + j - C2) Mod nb
		m_ri(m + 2)=(nb + j - C3) Mod nb
	Next
	N=m_Nb * (m_Nr + 1)
	For i=0 To m_Nk - 1
		j=i * 4
		CipherKey(i)=PackFrom(key, j)
	Next
	For i=0 To m_Nk - 1
		m_fkey(i)=CipherKey(i)
	Next
	j=m_Nk
	k=0
	Do While j < N
		m_fkey(j)=m_fkey(j - m_Nk) XOR SubByte(RotateLeft(m_fkey(j - 1), 24)) XOR m_rco(k)
		If m_Nk <= 6 Then
			i=1
			Do While i < m_Nk And (i + j) < N
				m_fkey(i + j)=m_fkey(i + j - m_Nk) XOR m_fkey(i + j - 1)
				i=i + 1
			Loop
		Else
			i=1
			Do While i < 4 And (i + j) < N
				m_fkey(i + j)=m_fkey(i + j - m_Nk) XOR m_fkey(i + j - 1)
				i=i + 1
			Loop
			If j + 4 < N Then
				m_fkey(j + 4)=m_fkey(j + 4 - m_Nk) XOR SubByte(m_fkey(j + 3))
			End If
			i=5
			Do While i < m_Nk And (i + j) < N
				m_fkey(i + j)=m_fkey(i + j - m_Nk) XOR m_fkey(i + j - 1)
				i=i + 1
			Loop
		End If
		j=j + m_Nk
		k=k + 1
	Loop
	For j=0 To m_Nb - 1
		m_rkey(j + N - nb)=m_fkey(j)
	Next
	i=m_Nb
	Do While i < N - m_Nb
		k=N - m_Nb - i
		For j=0 To m_Nb - 1
			m_rkey(k + j)=InvMixCol(m_fkey(i + j))
		Next
		i=i + m_Nb
	Loop
	j=N - m_Nb
	Do While j < N
		m_rkey(j - N + m_Nb)=m_fkey(j)
		j=j + 1
	Loop
End Sub
Public Sub encrypt(buff())
	Dim i,j,k,m,a(7),b(7),x,y,t
	For i=0 To m_Nb - 1
		j=i * 4
		a(i)=PackFrom(buff, j)
		a(i)=a(i) Xor m_fkey(i)
	Next
	k=m_Nb
	x=a
	y=b
	For i=1 To m_Nr - 1
		For j=0 To m_Nb - 1
			m=j * 3
			y(j)=m_fkey(k) Xor m_ftable(x(j) And m_lOnBits(7)) Xor _
				RotateLeft(m_ftable(RShift(x(m_fi(m)), 8) And m_lOnBits(7)), 8) Xor _
				RotateLeft(m_ftable(RShift(x(m_fi(m + 1)), 16) And m_lOnBits(7)), 16) Xor _
				RotateLeft(m_ftable(RShift(x(m_fi(m + 2)), 24) And m_lOnBits(7)), 24)
			k=k + 1
		Next
		t=x
		x=y
		y=t
	Next
	For j=0 To m_Nb - 1
		m=j * 3
		y(j)=m_fkey(k) Xor m_fbsub(x(j) And m_lOnBits(7)) Xor _
			RotateLeft(m_fbsub(RShift(x(m_fi(m)), 8) And m_lOnBits(7)), 8) Xor _
			RotateLeft(m_fbsub(RShift(x(m_fi(m + 1)), 16) And m_lOnBits(7)), 16) Xor _
			RotateLeft(m_fbsub(RShift(x(m_fi(m + 2)), 24) And m_lOnBits(7)), 24)
		k=k + 1
	Next
	For i=0 To m_Nb - 1
		j=i * 4
		UnpackFrom y(i), buff, j
		x(i)=0
		y(i)=0
	Next
End Sub
Public Sub decrypt(buff())
	Dim i, j, k, m, a(7), b(7), x, y, t
	For i=0 To m_Nb - 1
		j=i * 4
		a(i)=PackFrom(buff, j)
		a(i)=a(i) Xor m_rkey(i)
	Next
	k=m_Nb
	x=a
	y=b
	For i=1 To m_Nr - 1
		For j=0 To m_Nb - 1
			m=j * 3
			y(j)=m_rkey(k) Xor m_rtable(x(j) And m_lOnBits(7)) Xor _
				RotateLeft(m_rtable(RShift(x(m_ri(m)), 8) And m_lOnBits(7)), 8) Xor _
				RotateLeft(m_rtable(RShift(x(m_ri(m + 1)), 16) And m_lOnBits(7)), 16) Xor _
				RotateLeft(m_rtable(RShift(x(m_ri(m + 2)), 24) And m_lOnBits(7)), 24)
			k=k + 1
		Next
		t=x
		x=y
		y=t
	Next
	For j=0 To m_Nb - 1
		m=j * 3
		y(j)=m_rkey(k) Xor m_rbsub(x(j) And m_lOnBits(7)) Xor _
			RotateLeft(m_rbsub(RShift(x(m_ri(m)), 8) And m_lOnBits(7)), 8) Xor _
			RotateLeft(m_rbsub(RShift(x(m_ri(m + 1)), 16) And m_lOnBits(7)), 16) Xor _
			RotateLeft(m_rbsub(RShift(x(m_ri(m + 2)), 24) And m_lOnBits(7)), 24)
		k=k + 1
	Next
	For i=0 To m_Nb - 1
		j=i * 4
		UnpackFrom y(i), buff, j
		x(i)=0
		y(i)=0
	Next
End Sub
Private Function IsInitialized(vArray)
	On Error Resume Next
	IsInitialized=IsNumeric(UBound(vArray))
End Function
Private Sub CopyBytesASP(bytDest, lDestStart, bytSource(), lSourceStart, lLength)
	Dim lCount
	lCount=0
	Do
		bytDest(lDestStart + lCount)=bytSource(lSourceStart + lCount)
		lCount=lCount + 1
	Loop Until lCount=lLength
End Sub
public Sub XORBlock(bytData1,bytData2)
	Dim lCount
	for lCount=0 to 15
		bytData1(lCount)=(bytData1(lCount) XOR bytData2(lCount))
	next
end Sub
Public Function EncryptData(bytMessage, bytPassword)
	Dim bytKey(15)
	Dim bytIn()
	Dim bytOut()
	Dim bytLast(15)
	Dim bytTemp(15)
	Dim lCount
	Dim lLength
	Dim lEncodedLength
	Dim bytLen(3)
	Dim lPosition
	If Not IsInitialized(bytMessage) Then
		Exit Function
	End If
	If Not IsInitialized(bytPassword) Then
		Exit Function
	End If
	For lCount=0 To UBound(bytPassword)
		bytKey(lCount)=bytPassword(lCount)
		If lCount=15 Then
			Exit For
		End If
	Next
	gentables
	gkey 4, 4, bytKey
	lLength=UBound(bytMessage) + 1
	lEncodedLength=lLength
	If lEncodedLength Mod 16 <> 0 Then
		lEncodedLength=lEncodedLength + 16 - (lEncodedLength Mod 16)
	End If
	ReDim bytIn(lEncodedLength - 1)
	ReDim bytOut(lEncodedLength - 1)
	CopyBytesASP bytLast,0,bytPassword,0,16
	Unpack lLength, bytIn
	CopyBytesASP bytIn, 0, bytMessage, 0, lLength
	For lCount=0 To lEncodedLength - 1 Step 16
		CopyBytesASP bytTemp, 0, bytIn, lCount, 16
		XORBlock bytTemp,bytLast 
		Encrypt bytTemp
		CopyBytesASP bytOut, lCount, bytTemp, 0, 16
		CopyBytesASP bytLast,0,bytTemp, 0, 16
	Next
	EncryptData=bytOut
End Function
Function AESEncrypt(sPlain, sPassword)
	Dim bytIn()
	Dim bytOut
	Dim bytPassword()
	Dim lCount
	Dim lLength
	Dim sTemp
	Dim lPadLength
	lLength=Len(sPlain)
	lPadLength=16-(lLength mod 16)
	for lCount=1 to lPadLength
		sPlain=sPlain & Chr(lPadLength)
	next
	lLength=Len(sPlain)
	ReDim bytIn(lLength-1)
	For lCount=1 To lLength
		bytIn(lCount-1)=CByte(AscB(Mid(sPlain,lCount,1)))
	Next
	lLength=Len(sPassword)
	ReDim bytPassword(lLength-1)
	For lCount=1 To lLength
		bytPassword(lCount-1)=CByte(AscB(Mid(sPassword,lCount,1)))
	Next
	bytOut=EncryptData(bytIn, bytPassword)
	sTemp=""
	For lCount=0 To UBound(bytOut)
		sTemp=sTemp & Right("0" & Hex(bytOut(lCount)), 2)
	Next
	AESEncrypt=sTemp
End Function
function Base64Decode(ByVal vCode)
    dim oXML, oNode
    set oXML=CreateObject("Msxml2.DOMDocument.3.0")
    set oNode=oXML.CreateElement("base64")
    oNode.dataType="bin.base64"
    oNode.text=vCode
    Base64Decode=oNode.nodeTypedValue
    set oNode=nothing
    set oXML=nothing
end function
function Base64Encode(sText)
    dim oXML, oNode
    set oXML=CreateObject("Msxml2.DOMDocument.3.0")
    set oNode=oXML.CreateElement("base64")
    oNode.dataType="bin.base64"
    oNode.nodeTypedValue=(sText)
    Base64Encode=oNode.text
    set oNode=nothing
    set oXML=nothing
end function
%>