<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if sDSN="" then response.write "Database connection not set" : response.end
Dim sSQL,rs,rs2,alldata,cnn,rowcounter,success,Count,startlink,endlink,weburl,CurPage,iNumOfPages,subCats,lasttsid,sText,index,aText,aFields(5),relsql(1),currFormat1,currFormat2,currFormat3,aDiscSection()
if request.totalbytes > 10000 OR len(getget("pg"))>8 then
	response.status="400 Bad Request"
	response.end
end if
nosearchrelevance=FALSE : showcategories=FALSE : gotcriteria=FALSE : isrootsection=FALSE
catid="0" : topsectionids="0"
set toregexp=new RegExp
toregexp.pattern="[^,\d]"
toregexp.global=TRUE
scat=left(toregexp.replace(request("scat"),""),16)
sman=left(toregexp.replace(request("sman"),""),16)
toregexp.pattern=",+"
scat=toregexp.replace(scat,",")
sman=toregexp.replace(sman,",")
toregexp.pattern="^,|,$"
scat=toregexp.replace(scat,"")
sman=toregexp.replace(sman,"")
set toregexp=nothing
stype=request("stype")
if stype<>"any" AND stype<>"exact" then stype=""
if left(scat,2)="ms" then thecat=Right(scat,Len(scat)-2) else thecat=scat
if thecat<>"" then catzero=int(split(thecat,",")(0)) else catzero=""
if sman<>"" then manzero=int(split(sman,",")(0)) else manzero=""
cs="" : WSP="" : OWSP=""
TWSP="pPrice"
minprice="" : maxprice=""
if IsEmpty(Count) then Count=0 else Count=(Count+adminProdsPerPage)-(Count MOD adminProdsPerPage)
noautocorrect="autocapitalize=""off"" autocomplete=""off"" spellcheck=""false"" autocorrect=""off"""
if is_numeric(trim(request("sminprice"))) then minprice=cdbl(replace(request("sminprice"),"$",""))
if is_numeric(trim(request("sprice"))) then maxprice=cdbl(replace(request("sprice"),"$",""))
if trim(request("nobox"))="true" then nobox="true" else nobox=""
if lcase(adminencoding)="iso-8859-1" then raquo="»" else raquo="&raquo;"
if magictoolboxproducts<>"" then
	print "<script src=""" & IIfVr(magictoolboxproducts="MagicTouch","http://www.magictoolbox.com/mt/" & magictouchid & "/magictouch.js", lcase(magictoolboxproducts) & "/" & lcase(magictoolboxproducts) & ".js") & """></script>" & magictooloptionsjsproducts
	magictoolboxproducts=replace(magictoolboxproducts,"MagicZoomPlus","MagicZoom")
	magictool=magictoolboxproducts
end if
sub writemenulevel(id,itlevel)
	Dim wmlindex
	if itlevel<10 then
		for wmlindex=0 TO ubound(alldata,2)
			if alldata(2,wmlindex)=id then
				print "<option value='"&alldata(0,wmlindex)&"'"
				if catzero=alldata(0,wmlindex) then print " selected=""selected"">" else print ">"
				for index=0 to itlevel-2
					print raquo & " "
				next
				print alldata(1,wmlindex)&"</option>" & vbCrLf
				if alldata(3,wmlindex)=0 then call writemenulevel(alldata(0,wmlindex),itlevel+1)
			end if
		next
	end if
end sub
sub sortlinesearch(soid, sotext)
	if (sortoptions AND (2 ^ (soid-1)))<>0 then print "<option value="""&soid&""""&IIfVs(dosortby=soid," selected=""selected""")&">"&sotext&"</option>"
end sub
pblink="<a class=""ectlink"" href=""search"&extension&"?nobox="&nobox&"&amp;scat="&urlencode(scat)&"&amp;stext="&urlencode(request("stext"))&IIfVs(stype<>"","&amp;stype="&stype)&"&amp;sprice="&urlencode(maxprice)&IIfVs(minprice<>"","&amp;sminprice="&urlencode(minprice))&IIfVs(sman<>"","&amp;sman="&sman)
nofirstpg=FALSE
if pricecheckerisincluded<>TRUE then pricecheckerisincluded=FALSE
function getlike(fie,t,tjn)
	if left(t, 1)="-" AND usenotsearch then ' pSKU excluded to work around NULL problems
		if fie<>"pSKU" AND fie<>"pSearchParams" then sNOTSQL=sNOTSQL & fie & " LIKE '%"&mid(t, 2)&"%' OR "
	else
		getlike=fie & " LIKE '%"&t&"%' "&tjn
	end if
end function
set rs=Server.CreateObject("ADODB.RecordSet")
set rs2=Server.CreateObject("ADODB.RecordSet")
set rs3=Server.CreateObject("ADODB.RecordSet")
set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
if orprodsperpage<>"" then adminProdsPerPage=orprodsperpage
redim aDiscSection(2,adminProdsPerPage)
call checkCurrencyRates(currConvUser,currConvPw,currLastUpdate,currRate1,currSymbol1,currRate2,currSymbol2,currRate3,currSymbol3)
get_wholesaleprice_sql()
if SESSION("clientLoginLevel")<>"" then minloglevel=SESSION("clientLoginLevel") else minloglevel=0
if nosearchdescription AND nosearchlongdescription then nosearchrelevance=TRUE
scrid="" : numscrid=0
if getpost("posted")="1" OR getget("pg")<>"" then
	if is_numeric(getrequest("sortby")) then SESSION("sortby")=int(getrequest("sortby"))
	if SESSION("sortby")<>"" then dosortby=int(SESSION("sortby"))
	if orsortby<>"" then dosortby=orsortby
	if dosortby=5 then nosearchrelevance=TRUE
	if nosearchbyfilters<>TRUE then
		for each objitem in request.querystring
			if left(objitem,4)="sfil" then
				if is_numeric(right(objitem,len(objitem)-4)) AND is_numeric(getget(objitem)) then
					scrid=scrid&getget(objitem)&","
					pblink=pblink&"&amp;"&objItem&"="&getget(objItem)
				end if
			end if
		next
		if scrid<>"" then
			scrid=left(scrid,len(scrid)-1)
			scridarr=split(scrid,",")
			numscrid=UBOUND(scridarr)+1
			gotcriteria=TRUE
		end if
	end if
	pblink=pblink&"&amp;pg="
	if thecat<>"" then
		sSQL="SELECT DISTINCT "&IIfVs(NOT nosearchrelevance,"0 AS relevanceorder,")&"products.pId,"&getlangid("pName",1)&","&WSP&"pPrice,pWholesalePrice,pOrder,pDateAdded,pManufacturer,pSKU,pPopularity,pNumSales,pNumRatings,pTotRating" & IIfVs(dosortby=17 AND sqlserver AND NOT mysqlserver,",CASE WHEN pNumRatings=0 THEN 0 ELSE pTotRating/pNumRatings END") & " FROM "
		if numscrid>1 then sSQL=sSQL&string(numscrid-1,"(")
		sSQL=sSQL&"(multisections RIGHT JOIN (products INNER JOIN sections ON products.pSection=sections.sectionID) ON products.pId=multisections.pId)" & IIfVs(scrid<>""," INNER JOIN multisearchcriteria ON multisearchcriteria.mSCpID=products.pID")
		for scrindex=1 to numscrid-1
			sSQL=sSQL&") INNER JOIN multisearchcriteria msc"&scrindex&" ON products.pID=msc"&scrindex&".mSCpID"
		next
		sSQL=sSQL&" WHERE sectionDisabled<="&minloglevel&" AND pDisplay<>0 " & IIfVs(ectsiteid<>"", "AND pSiteID=" & ectsiteid & " ") & IIfVs(nosearcharticles,"AND pSchemaType=0 ")
		gotcriteria=true
		sectionids=getsectionids(thecat, false)
		if sectionids<>"" then sSQL=sSQL & "AND (products.pSection IN (" & sectionids & ") OR multisections.pSection IN (" & sectionids & ")) "
	else
		sSQL="SELECT "&IIfVs(NOT nosearchrelevance,"0 AS relevanceorder,")&"products.pId,"&getlangid("pName",1)&","&WSP&"pPrice,pWholesalePrice,pOrder,pDateAdded,pManufacturer,pSKU,pPopularity,pNumSales,pNumRatings,pTotRating" & IIfVs(dosortby=17 AND sqlserver AND NOT mysqlserver,",CASE WHEN pNumRatings=0 THEN 0 ELSE pTotRating/pNumRatings END") & " FROM "
		if numscrid>1 then sSQL=sSQL&string(numscrid-1,"(")
		sSQL=sSQL&"(products INNER JOIN sections ON products.pSection=sections.sectionID)" & IIfVs(scrid<>""," INNER JOIN multisearchcriteria ON multisearchcriteria.mSCpID=products.pID")
		for scrindex=1 to numscrid-1
			sSQL=sSQL&") INNER JOIN multisearchcriteria msc"&scrindex&" ON products.pID=msc"&scrindex&".mSCpID"
		next
		sSQL=sSQL&" WHERE sectionDisabled<="&minloglevel&" AND pDisplay<>0 " & IIfVs(ectsiteid<>"", "AND pSiteID=" & ectsiteid & " ") & IIfVs(nosearcharticles,"AND pSchemaType=0 ")
	end if
	session.LCID=1033
	if is_numeric(trim(request("sprice"))) then
		gotcriteria=true
		sSQL=sSQL & "AND "&TWSP&"<="&cdbl(replace(request("sprice"),"$",""))&" "
	end if
	if minprice<>"" then
		gotcriteria=true
		sSQL=sSQL & "AND "&TWSP&">="&minprice&" "
	end if
	if sman<>"" then
		gotcriteria=true
		sSQL=sSQL & "AND pManufacturer IN ("&sman&") "
	end if
	if scrid<>"" then
		sSQL=sSQL & "AND (multisearchcriteria.mSCscID=" & scridarr(0)
		for scrindex=1 to numscrid-1
			sSQL=sSQL&" AND msc"&scrindex&".mSCscID="&scridarr(scrindex)
		next
		sSQL=sSQL & ") "
	end if
	session.LCID=saveLCID
	if trim(request("stext"))<>"" then
		gotcriteria=true
		sText=escape_string(replace(left(request("stext"), 1024),"[","[[]"))
		aText=Split(sText)
		aFields(0)="products.pId"
		aFields(1)=getlangid("pName",1)
		aFields(2)=getlangid("pDescription",2)
		aFields(3)=getlangid("pLongDescription",4)
		aFields(4)="pSKU"
		aFields(5)="pSearchParams"
		if NOT nosearchrelevance then sSQL=sSQL&"''||SPBLOCKMARKER||''"
		for relindex=0 to IIfVr(nosearchrelevance,0,1)
			if stype="exact" then
				relsql(relindex)="AND "
				if left(sText, 1)="-" AND usenotsearch then relsql(relindex)=relsql(relindex) & "NOT " : sText=mid(sText, 2) : isnot=TRUE else isnot=FALSE
				if relindex=0 OR nosearchrelevance then
					relsql(relindex)=relsql(relindex) & "(" & IIfVs(NOT nosearchprodid,"products.pId LIKE '%"&sText&"%' OR ") & getlangid("pName",1)&" LIKE '%"&sText&"%'"&IIfVr(nosearchparams,""," OR pSearchParams LIKE '%"&sText&"%'")&IIfVr(isnot OR nosearchsku, "", " OR pSKU LIKE '%"&sText&"%'")&IIfVs(NOT nosearchrelevance,") ")
				end if
				if relindex=1 OR nosearchrelevance then
					relsql(relindex)=relsql(relindex) & IIfVs(NOT (nosearchrelevance AND nosearchlongdescription), IIfVr(nosearchrelevance," OR ","(")&IIfVs(nosearchdescription<>TRUE,getlangid("pDescription",2)&" LIKE '%"&sText&"%'")&IIfVs(nosearchlongdescription<>TRUE,IIfVs(nosearchdescription<>TRUE," OR ")&getlangid("pLongDescription",2)&" LIKE '%"&sText&"%'")) & ") "
				end if
			elseif UBOUND(aText) < 24 then
				sNOTSQL="" : sYESSQL=""
				if stype="any" then
					for index=IIfVr(nosearchprodid,1,0) to 5
						tmpSQL=""
						for rowcounter=0 to UBOUND(aText)
							if NOT ((nosearchdescription=TRUE AND index=2) OR (nosearchlongdescription=TRUE AND index=3) OR (nosearchsku=TRUE AND index=4) OR (nosearchparams=TRUE AND index=5)) then
								if ((index=0 OR index=1 OR index=4 OR index=5) AND relindex=0) OR ((index=2 OR index=3) AND relindex=1) OR nosearchrelevance then tmpSQL=tmpSQL & getlike(aFields(index), aText(rowcounter), "OR ")
							end if
						next
						if tmpSQL<>"" then sYESSQL=sYESSQL & "(" & left(tmpSQL, len(tmpSQL)-3) & ") "
						if tmpSQL<>"" then sYESSQL=sYESSQL & "OR "
					next
					if sYESSQL<>"" then sYESSQL=left(sYESSQL,len(sYESSQL)-3)
				else
					for rowcounter=0 to UBOUND(aText)
						tmpSQL=""
						for index=IIfVr(nosearchprodid,1,0) to 5
							if NOT ((nosearchdescription=TRUE AND index=2) OR (nosearchlongdescription=TRUE AND index=3) OR (nosearchsku=TRUE AND index=4) OR (nosearchparams=TRUE AND index=5)) then
								if ((index=0 OR index=1 OR index=4 OR index=5) AND relindex=0) OR ((index=2 OR index=3) AND relindex=1) OR nosearchrelevance then tmpSQL=tmpSQL & getlike(aFields(index), aText(rowcounter), "OR ")
							end if
						next
						if tmpSQL<>"" then sYESSQL=sYESSQL & "(" & left(tmpSQL, len(tmpSQL)-3) & ") "
						if tmpSQL<>"" then sYESSQL=sYESSQL & "AND "
					next
					if sYESSQL<>"" then sYESSQL=left(sYESSQL,len(sYESSQL)-4)
				end if
				relsql(relindex)=""
				if sYESSQL<>"" then relsql(relindex)=relsql(relindex) & "AND (" & sYESSQL & ") "
				if sNOTSQL<>"" then relsql(relindex)=relsql(relindex) & "AND NOT (" & left(sNOTSQL, len(sNOTSQL)-4) & ")"
			end if
		next
		if nosearchrelevance then sSQL=sSQL&relsql(0)
	else
		nosearchrelevance=TRUE
	end if
	if NOT gotcriteria then nosearchrelevance=TRUE
	if dosortby=2 OR dosortby=12 then
		sSortBy=" ORDER BY "&IIfVs(NOT nosearchrelevance,"1,")&IIfVs(NOT mysqlserver,"products.")&"pId"&IIfVs(dosortby=12," DESC")
	elseif dosortby=14 OR dosortby=15 then
		sSortBy=" ORDER BY "&IIfVs(NOT nosearchrelevance,"1,")&"pSKU"&IIfVs(dosortby=15," DESC")
	elseif dosortby=3 OR dosortby=4 then
		sSortBy=" ORDER BY "&IIfVs(NOT nosearchrelevance,"1,")&TWSP&IIfVs(dosortby=4," DESC")&","&IIfVs(NOT mysqlserver,"products.")&"pId"
	elseif dosortby=5 then
		sSortBy=IIfVs(NOT nosearchrelevance,"1,")
	elseif dosortby=6 OR dosortby=7 then
		sSortBy=" ORDER BY "&IIfVs(NOT nosearchrelevance,"1,")&"pOrder"&IIfVs(dosortby=7," DESC")&","&IIfVs(NOT mysqlserver,"products.")&"pId"
	elseif dosortby=8 OR dosortby=9 then
		sSortBy=" ORDER BY "&IIfVs(NOT nosearchrelevance,"1,")&"pDateAdded"&IIfVs(dosortby=9," DESC")&","&IIfVs(NOT mysqlserver,"products.")&"pId"
	elseif dosortby=10 then
		sSortBy=" ORDER BY "&IIfVs(NOT nosearchrelevance,"1,")&"pManufacturer"
	elseif dosortby=16 then
		sSortBy=" ORDER BY "&IIfVs(NOT nosearchrelevance,"1,")&"pNumRatings DESC,"&IIfVs(NOT mysqlserver,"products.")&"pId"
	elseif dosortby=17 then
		sSortBy="CASE WHEN pNumRatings=0 THEN 0 ELSE pTotRating/pNumRatings END"
		if mysqlserver OR NOT sqlserver then sSortBy=IIfVs(NOT sqlserver,"I")&"IF(pNumRatings=0,0,pTotRating/pNumRatings)"
		sSortBy=" ORDER BY "&sSortBy&" DESC,pNumRatings DESC,"&IIfVs(NOT mysqlserver,"products.")&"pId"
	elseif dosortby=18 then
		sSortBy=" ORDER BY "&IIfVs(NOT nosearchrelevance,"1,")&"pNumSales DESC,"&IIfVs(NOT mysqlserver,"products.")&"pId"
	elseif dosortby=19 then
		sSortBy=" ORDER BY "&IIfVs(NOT nosearchrelevance,"1,")&"pPopularity DESC,"&IIfVs(NOT mysqlserver,"products.")&"pId"
	else
		sSortBy=" ORDER BY "&IIfVs(NOT nosearchrelevance,"1,")&getlangid("pName",1)&IIfVs(dosortby=11," DESC")&","&IIfVs(NOT mysqlserver,"products.")&"pId"
	end if
	if NOT gotcriteria then sSQL="SELECT products.pId FROM products INNER JOIN sections ON products.pSection=sections.sectionID WHERE sectionDisabled<="&minloglevel&" AND pDisplay<>0" & IIfVs(nosearcharticles," AND pSchemaType=0")
	if useStockManagement AND noshowoutofstock then sSQL=sSQL & " AND (pInStock>pMinQuant OR pStockByOpts<>0)"
	origSQL=sSQL
	relevantmatches=""
	userelevantmatches=TRUE
	numrelevantmatches=0
	if NOT nosearchrelevance then
		rs.open replace(sSQL,"''||SPBLOCKMARKER||''",relsql(0),1), cnn
		do while NOT rs.EOF
			relevantmatches=relevantmatches&"'"&escape_string(rs("pId"))&"',"
			numrelevantmatches=numrelevantmatches+1
			if numrelevantmatches>100 then userelevantmatches=FALSE : exit do
			rs.movenext
		loop
		rs.close
	end if
	if relevantmatches<>"" then relevantmatches=left(relevantmatches,len(relevantmatches)-1) else userelevantmatches=FALSE
	if gotcriteria then
		if numrelevantmatches>=100 OR nosearchrelevance then
			sSQL=replace(sSQL,"''||SPBLOCKMARKER||''",relsql(0),1)
		else
			sSQL=IIfVs(relevantmatches<>"",replace(sSQL,"''||SPBLOCKMARKER||''",relsql(0),1) & " UNION ALL ") & replace(replace(sSQL,"0 AS relevanceorder,","1 AS relevanceorder,",1),"''||SPBLOCKMARKER||''",relsql(1)&IIfVs(userelevantmatches," AND NOT products.pId IN ("&relevantmatches&")"),1)
		end if
	end if
	rs.CursorLocation=3 ' adUseClient
	rs.CacheSize=adminProdsPerPage
	rs.open sSQL & sSortBy, cnn
	if rs.EOF then
		set stregexp=new RegExp
		stregexp.pattern="LIKE '%([^']{3,}?)s%'"
		stregexp.ignorecase=TRUE
		stregexp.global=TRUE
		relsql0=stregexp.replace(relsql(0)&"","LIKE '%$1%'")
		if relsql0<>relsql(0) then
			rs.close
			relsql(0)=relsql0
			relsql(1)=stregexp.replace(relsql(1),"LIKE '%$1%'")
			sSQL=origSQL
			relevantmatches=""
			userelevantmatches=TRUE
			numrelevantmatches=0
			if NOT nosearchrelevance then
				rs.open replace(sSQL,"''||SPBLOCKMARKER||''",relsql(0),1), cnn
				do while NOT rs.EOF
					relevantmatches=relevantmatches&"'"&escape_string(rs("pId"))&"',"
					numrelevantmatches=numrelevantmatches+1
					if numrelevantmatches>100 then userelevantmatches=FALSE : exit do
					rs.movenext
				loop
				rs.close
			end if
			if relevantmatches<>"" then relevantmatches=left(relevantmatches,len(relevantmatches)-1) else userelevantmatches=FALSE
			if gotcriteria then
				if numrelevantmatches>=100 OR nosearchrelevance then
					sSQL=replace(sSQL,"''||SPBLOCKMARKER||''",relsql(0),1)
				else
					sSQL=IIfVs(relevantmatches<>"",replace(sSQL,"''||SPBLOCKMARKER||''",relsql(0),1) & " UNION ALL ") & replace(replace(sSQL,"0 AS relevanceorder,","1 AS relevanceorder,",1),"''||SPBLOCKMARKER||''",relsql(1)&IIfVs(userelevantmatches," AND NOT products.pId IN ("&relevantmatches&")"),1)
				end if
			end if
			rs.CursorLocation=3 ' adUseClient
			rs.CacheSize=adminProdsPerPage
			rs.open sSQL & sSortBy, cnn
		end if
	end if
	if rs.EOF then
		success=false
		iNumOfPages=0
	else
		success=true
		rs.MoveFirst
		rs.PageSize=adminProdsPerPage
		if NOT is_numeric(getget("pg")) then
			CurPage=1
		else
			CurPage=int(getget("pg"))
			if NOT CurPage>0 then CurPage=1
		end if
		iNumOfPages=int((rs.RecordCount + (adminProdsPerPage-1)) / adminProdsPerPage)
		rs.AbsolutePage=CurPage
	end if
	localcount=0
	if NOT rs.EOF then
		prodlist=""
		addcomma=""
		do while NOT rs.EOF AND localcount<rs.PageSize
			' print "RO: " & rs(0) & " : " & rs("pId") & "<br>"
			prodlist=prodlist & addcomma & "'" & rs("pId") & "'"
			rs.MoveNext
			localcount=localcount+1
			addcomma=","
		loop
		rs.close
		wantmanufacturer=(instr(productpagelayout&quickbuylayout,"manufacturer")>0 OR (useproductbodyformat=3 AND instr(cpdcolumns, "manufacturer")>0) OR (NOT usecsslayout AND xxManLab<>""))
		sSQL="SELECT "&IIfVs(NOT nosearchrelevance,"0 AS relevanceorder,")&"pId,pSKU,"&getlangid("pName",1)&","&WSP&"pPrice,pListPrice,pSection,pSell,pStockByOpts,pStaticPage,pStaticURL,pInStock,pExemptions,pTax,pTotRating,pNumRatings,pBackOrder,pDateAdded,pMinQuant,pCustomCSS,pCustom1,pCustom2,pCustom3,"&IIfVr(wantmanufacturer,getlangid("scName",131072)&",","")&IIfVr(shortdescriptionlimit<>"" AND shortdescriptionlimit=0,"'' AS ","")&getlangid("pDescription",2)&","&getlangid("pLongDescription",4)&" FROM products "&IIfVr(wantmanufacturer,"LEFT OUTER JOIN searchcriteria on products.pManufacturer=searchcriteria.scID ","")&"WHERE pId IN (" & prodlist & ")" & sSortBy
		rs.open sSQL, cnn, 0, 1
	end if
end if
if nobox<>"true" then %>
	  <form method="get" id="stextform" action="search<%=extension%>">		  
		  <input type="hidden" name="pg" value="1" />
		  <div class="searchform">
			<div class="searchheader"><%=xxSrchPr%></div>
			<div class="searchcntnr searchfor_cntnr">
				<div class="searchtext searchfortext"><%=xxSrchFr%></div>
				<div class="searchcontrol searchfor">
					<div style="position:relative">
						<input type="search" name="stext" id="stext" size="20" maxlength="1024" value="<%=htmlspecials(request("stext"))%>" onkeydown="return ectAutoSrchKeydown(this,event,'asp')" onblur="ectAutoHideCombo(this)" <%=noautocorrect%> />
					</div>
<%	if noautosearch<>TRUE then %>
					<div class="autosearch" style="position:absolute;display:none;" id="selectstext"></div>
<%	end if %>
				</div>
			</div>
			<div class="searchcntnr searchprice_cntnr">
				<div class="searchtext searchpricetext"><%=xxSrchMx%></div>
				<div class="searchcontrol searchprice"><input type="number" name="sprice" size="10" maxlength="64" value="<%=htmlspecials(request("sprice"))%>" <%=noautocorrect%> /></div>
			</div>
			<div class="searchcntnr searchtype_cntnr">
				<div class="searchtext searchtypetext"><%=xxSrchTp%></div>
				<div class="searchcontrol searchtype"><select class="searchtype" name="stype" size="1">
					<option value=""><%=xxSrchAl%></option>
					<option value="any" <% if stype="any" then print "selected=""selected"""%>><%=xxSrchAn%></option>
					<option value="exact" <% if stype="exact" then print "selected=""selected"""%>><%=xxSrchEx%></option>
					</select></div>
			</div>
<%	if nocategorysearch<>TRUE then %>
			<div class="searchcntnr searchcategory_cntnr">
				<div class="searchtext searchcategorytext"><%=xxSrchCt%></div>
				<div class="searchcontrol searchcategory"><select class="searchcategory" name="scat" id="scat" size="1">
				  <option value=""><%=xxSrchAC%></option>
<%					sSQL="SELECT sectionID,"&getlangid("sectionName",256)&",topSection,rootSection FROM sections WHERE sectionDisabled<="&minloglevel&" "
					if onlysubcats=true then sSQL=sSQL & "AND rootSection=1 ORDER BY " & getlangid("sectionName",256) else sSQL=sSQL & "ORDER BY " & IIfVr(sortcategoriesalphabetically,getlangid("sectionName",256),"sectionOrder")
					rs2.Open sSQL,cnn,0,1
					if NOT rs2.eof then alldata=rs2.getrows
					rs2.Close
					if IsArray(alldata) then call writemenulevel(catalogroot,1) %>
				  </select></div>
			</div>
<%	end if
	if (prodfilter AND 8)=8 AND sortoptions<>0 AND NOT nosortsearch then %>
			<div class="searchcntnr searchsort_cntnr">
				<div class="searchtext searchsorttext"><%=xxSortBy%></div>
				<div class="searchcontrol searchsort"><select name="sortby" class="searchsort" size="1">
				<option value="0"><%=xxPlsSel%></option>
<%				call sortlinesearch(1, IIfVr(sortoption1<>"",sortoption1,"Sort Alphabetically"))
				call sortlinesearch(11, IIfVr(sortoption11<>"",sortoption11,"Alphabetically (Desc.)"))
				call sortlinesearch(2, IIfVr(sortoption2<>"",sortoption2,"Sort by Product ID"))
				call sortlinesearch(12, IIfVr(sortoption12<>"",sortoption12,"Product ID (Desc.)"))
				call sortlinesearch(14, IIfVr(sortoption14<>"",sortoption14,"Sort By SKU"))
				call sortlinesearch(15, IIfVr(sortoption15<>"",sortoption15,"Sort By SKU (Desc.)"))
				call sortlinesearch(3, IIfVr(sortoption3<>"",sortoption3,"Sort Price (Asc.)"))
				call sortlinesearch(4, IIfVr(sortoption4<>"",sortoption4,"Sort Price (Desc.)"))
				call sortlinesearch(5, IIfVr(sortoption5<>"",sortoption5,"Database Order"))
				call sortlinesearch(6, IIfVr(sortoption6<>"",sortoption6,"Product Order"))
				call sortlinesearch(7, IIfVr(sortoption7<>"",sortoption7,"Product Order (Desc.)"))
				call sortlinesearch(8, IIfVr(sortoption8<>"",sortoption8,"Date Added (Asc.)"))
				call sortlinesearch(9, IIfVr(sortoption9<>"",sortoption9,"Date Added (Desc.)"))
				call sortlinesearch(10, IIfVr(sortoption10<>"",sortoption10,"Sort by Manufacturer"))
				call sortlinesearch(16, IIfVr(sortoption16<>"",sortoption16,"Number of Ratings"))
				call sortlinesearch(17, IIfVr(sortoption17<>"",sortoption17,"Average Rating"))
				call sortlinesearch(18, IIfVr(sortoption18<>"",sortoption18,"Sales Rank"))
				call sortlinesearch(19, IIfVr(sortoption19<>"",sortoption19,"Popularity"))
%>			  </select></div>
			</div>
<%	end if
	if nosearchbyfilters<>TRUE then
		if searchfiltergroups<>"" then searchfiltergroups="WHERE scGroup IN ("&searchfiltergroups&")" : onlysearchfiltermanufacturer=TRUE else searchfiltergroups="WHERE scGroup=0 "
		sSQL="SELECT "&IIfVs(NOT nocountsearchfilter,"COUNT(*) as tcount,") & "scID,"&getlangid("scName",131072)&",scGroup,scgTitle FROM "&IIfVs(NOT nocountsearchfilter,"(")&"(searchcriteria INNER JOIN searchcriteriagroup ON searchcriteria.scGroup=searchcriteriagroup.scgID) "&IIfVr(nocountsearchfilter,IIfVs(onlysearchfiltermanufacturer,searchfiltergroups),"INNER JOIN multisearchcriteria ON multisearchcriteria.mSCscID=searchcriteria.scID) " & IIfVs(onlysearchfiltermanufacturer,searchfiltergroups) & "GROUP BY scID,"&getlangid("scName",131072)&",scGroup,scOrder,scgOrder,scgID,scgTitle ") & "ORDER BY scGroup,scOrder,"&getlangid("scName",131072)
		rs2.open sSQL,cnn,0,1
		if NOT rs2.EOF then
%>			<div class="searchcntnr searchfilters_cntnr">
				<div class="searchtext searchfilterstext"><%=xxSeaFil%></div>
				<div class="searchcontrol searchfilters"><%
			currgroup=-1
			do while NOT rs2.EOF
				if currgroup<>rs2("scGroup") then
					if currgroup<>-1 then print "</select></div>" & vbCrLf
					print "<div class=""searchfiltergroup"&rs2("scGroup")&" searchfiltergroup""><select class=""searchfiltergroup"" name=""sfil"&rs2("scGroup")&""" size=""1""><option style=""font-weight:bold"" value="""">== All " & htmlspecials(rs2("scgTitle")) & " ==</option>"
					currgroup=rs2("scGroup")
				end if
				print "<option value=""" & rs2("scID") & """"
				if getget("sfil"&rs2("scGroup"))=cstr(rs2("scID")) then print " selected=""selected"""
				print ">" & rs2(getlangid("scName",131072))
				if NOT nocountsearchfilter then print " (" & rs2("tcount") & ")"
				print "</option>" & vbCrLf
				rs2.movenext
			loop
			if currgroup<>-1 then print "</select></div>" %></div>
			</div>
<%		end if
		rs2.close
	end if %>
			<div class="searchcntnr ectsearchsubmit"><div class="searchtext"></div><div class="searchcontrol"><%=imageorsubmit(imgsearch,xxSearch,"search")%></div></div>
		  </div>
		</form>
<%
end if
if getpost("posted")="1" OR getget("pg")<>"" then
	print "<div class=""searchresults"">"
	if rs.EOF then print "<div class=""nosearchresults"">" & xxSrchNM & "</div>"
	if NOT rs.EOF then
		if usesearchbodyformat=3 then %>
<!--#include file="incproductbody3.asp"-->
<%		elseif usesearchbodyformat=2 then %>
<!--#include file="incproductbody2.asp"-->
<%		else %>
<!--#include file="incproductbody.asp"-->
<%		end if
		if NOT usecsslayout then print "</td></tr>"
	end if
	print "</div>"
	rs.close
end if
cnn.Close
set rs=nothing
set rs2=nothing
set cnn=nothing
%>
