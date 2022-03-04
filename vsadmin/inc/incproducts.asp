<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if sDSN="" then response.write "Database connection not set" : response.end
Dim rs,alldata,cnn,rowcounter,iNumOfPages,CurPage,Count,weburl,catid,currFormat1,currFormat2,currFormat3,aDiscSection(),aFields(5),globaldiscounts(3,30)
if xxDsNoAp="" then xxDsNoAp="The following discount(s) will not apply:"
isproductspage=TRUE : hasshippingdiscount=FALSE : hasproductdiscount=FALSE
manname="" : catname="" : caturl="" : catrootsection="" : globaldiscounttext="" : catrootsection=1 : maxgroupid=0 : maxglobaldiscounts=0
set rs=Server.CreateObject("ADODB.RecordSet")
set rs2=Server.CreateObject("ADODB.RecordSet")
set rs3=Server.CreateObject("ADODB.RecordSet")
set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
dothrow301=FALSE
if NOT isempty(orprodfilter) then prodfilter=orprodfilter
if getget("cat")<>"" then catid=getget("cat") else catid=""
if getget("man")<>"" then manid=getget("man") else manid=""
set toregexp=new regexp
toregexp.pattern="[^0-9\-\.]"
toregexp.global=TRUE
sprice=toregexp.replace(getget("sprice"),"")
set toregexp=nothing
if NOT alreadygotattributevars then
	numscrid=0 : numscrgroup=0
	scrid=commaseplist(getget("scri"))
	if scrid="" AND is_numeric(explicitmanid) then scrid=explicitmanid
	if scrid<>"" then scridarr=split(scrid,",") : numscrid=UBOUND(scridarr)+1
	for scrind=0 to numscrid-1
		if len(scridarr(scrind))>10 then scrid=""
	next
	alreadygotattributevars=TRUE
end if
if is_numeric(explicitmanid) then manid=explicitmanid
if is_numeric(request("sortby")) then SESSION("sortby")=int(request("sortby"))
if SESSION("sortby")<>"" then dosortby=SESSION("sortby") else if orsortby<>"" then dosortby=orsortby
if seocategoryurls then usecategoryname=TRUE : catid=replace(catid,detlinkspacechar," ") : manid=replace(manid,detlinkspacechar," ")
if bmlbannerproducts<>"" AND paypalpublisherid<>"" then call displaybmlbanner(paypalpublisherid,bmlbannerproducts)
if is_numeric(explicitid) then
	catid=explicitid
elseif usecategoryname AND trim(catid)<>"" then
	sSQL="SELECT sectionID FROM sections WHERE "&IIfVs(seocategoryurls,getlangid("sectionurl",2048)&"='"&escape_string(catid)&"' OR (")&getlangid("sectionName",256)&"='"&escape_string(catid)&"'"&IIfVs(seocategoryurls," AND "&getlangid("sectionurl",2048)&"='')")
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		catname=catid
		catid=rs("sectionID")
	elseif NOT is_numeric(catid) then
		catname="Not Found"
		catid=-1
		response.status="404 Not Found"
	end if
	rs.close
end if
if usecategoryname AND trim(manid)<>"" then
	sSQL="SELECT scID FROM searchcriteria WHERE " & IIfVr(is_numeric(explicitmanid),"scID="&explicitmanid,"(("&getlangid("scURL",8192)&"='' OR "&getlangid("scURL",8192)&" IS NULL) AND "&getlangid("scName",131072)&"='"&escape_string(manid)&"') OR "&getlangid("scURL",8192)&"='"&escape_string(manid)&"'") & " ORDER BY scGroup"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then manname=manid : manid=rs("scID") else manid=-1
	rs.close
end if
if usecategoryname AND catname="" AND manname="" AND seourlsthrow301 AND (is_numeric(catid) OR is_numeric(manid)) then
	if is_numeric(catid) then
		sSQL="SELECT sectionID AS secid,"&getlangid("sectionName",256)&" AS secname,"&getlangid("sectionurl",2048)&" AS securl,rootSection FROM sections WHERE sectionID="&catid
	else
		sSQL="SELECT scID AS secid,"&getlangid("scName",131072)&" AS secname,"&getlangid("scURL",8192)&" AS securl,1 AS rootSection FROM searchcriteria WHERE scID="&manid
	end if
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		if is_numeric(catid) then catid=rs("secid") else manid=rs("secid")
		dothrow301=(explicitid="" AND explicitmanid="") : catname=rs("secname") : caturl=trim(rs("securl")&"") : catrootsection=rs("rootSection")
	end if
	rs.close
end if
if NOT is_numeric(catid) then catid=catalogroot
if is_numeric(manid) OR manufacturers=TRUE then manufacturers=TRUE else manufacturers=FALSE : manid=""
cs="" : WSP="" : OWSP=""
TWSP="pPrice"
sectionheader=""
HTTP_X_ORIGINAL_URL=trim(split(request.servervariables("HTTP_X_ORIGINAL_URL")&"?","?")(0))
if HTTP_X_ORIGINAL_URL="" then HTTP_X_ORIGINAL_URL=trim(split(request.servervariables("HTTP_X_REWRITE_URL")&"?","?")(0))
for each objitem in request.querystring
	if objitem="man" AND is_numeric(getget("man")) then manufacturers=TRUE
next
if ((seocategoryurls AND HTTP_X_ORIGINAL_URL="") OR dothrow301) AND seourlsthrow301 AND NOT is_numeric(explicitid) AND NOT is_numeric(explicitmanid) then
	newloc=getfullurl(getcategoryurl(catid,catname,caturl,catrootsection))
	addand="" : newqs=""
	for each objitem in request.querystring
		if objitem<>"cat" AND NOT (manufacturers AND objitem="man") then newqs=newqs&addand&objitem&"="&urlencode(getget(objitem)) : addand="&"
	next
	response.status="301 Moved Permanently"
	response.addheader "Location", newloc & IIfVs(newqs<>"","?"&newqs)
	response.end
end if
sectionurl=strip_tags2(IIfVr(seocategoryurls AND HTTP_X_ORIGINAL_URL<>"",HTTP_X_ORIGINAL_URL,request.servervariables("URL")))
iNumOfPages=0
if manufacturerpageurl="" then manufacturerpageurl="manufacturers.asp"
if filterpricebands="" then filterpricebands=100
if pricecheckerisincluded<>TRUE then pricecheckerisincluded=FALSE
function isinscrid(cscrid)
	isinscrid=FALSE
	for scrind=0 to numscrid-1
		if cscrid=int(scridarr(scrind)) then isinscrid=TRUE
	next
end function
function getlike(fie,t,tjn)
	if left(t, 1)="-" then ' pSKU excluded to work around NULL problems
		if fie<>"pSKU" then sNOTSQL=sNOTSQL & fie & " LIKE '%"&mid(t, 2)&"%' OR "
	else
		getlike=fie & " LIKE '%"&t&"%' "&tjn
	end if
end function
function sortline(soid, sotext)
	if (sortoptions AND (2 ^ (soid-1)))<>0 then print "<option value="""&soid&""""&IIfVr(dosortby=soid," selected=""selected""","")&">"&sotext&"</option>"
end function
nofirstpg=TRUE
pblink="<a class=""ectlink"" href="""&sectionurl&"?"
for each objQS in request.querystring
	if objQS<>"cat" AND objQS<>"id" AND objQS<>"man" AND objQS<>"pg" then pblink=pblink & urlencode(objQS) & "=" & urlencode(getget(objQS)) & "&amp;"
next
if (catid<>"0" OR (manufacturers AND manid<>"")) AND explicitid="" AND explicitmanid="" AND NOT (seocategoryurls AND HTTP_X_ORIGINAL_URL<>"") then pblink=pblink & IIfVr(manufacturers,"man="&getget("man"),"cat="&getcatid(catid,catname,seoprodurlpattern))&"&amp;pg=" else pblink=pblink & "pg="
if magictoolboxproducts<>"" then
	print "<script src=""" & lcase(magictoolboxproducts) & "/" & lcase(magictoolboxproducts) & ".js""></script>" & magictooloptionsjsproducts
	magictoolboxproducts=replace(magictoolboxproducts,"MagicZoomPlus","MagicZoom")
	magictool=magictoolboxproducts
	if magictoolboxproducts="MagicThumb" then magictooloptionsproducts=replace(magictooloptionsproducts,"data-options=","rel=") else magictooloptionsproducts=replace(magictooloptionsproducts,"rel=","data-options=")
end if
filterurl="" : manfilterurl=""
for each objQS in request.querystring
	if objQS<>"recentview" AND objQS<>"filter" AND objQS<>"pg" AND objQS<>"sortby" AND objQS<>"perpage" AND NOT ((objQS="cat" OR objQS="man") AND seocategoryurls) then
		filterurl=filterurl & urlencode(objQS) & "=" & urlencode(getget(objQS)) & "&"
		if objQS<>"sman" AND objQS<>"scri" AND objQS<>"sprice" then manfilterurl=manfilterurl & urlencode(objQS) & "=" & urlencode(getget(objQS)) & "&"
	end if
next
if filterurl="" then filterurl=sectionurl&"?filter=" else filterurl=sectionurl&"?"&filterurl&"filter="
if manfilterurl="" then manfilterurl=sectionurl&"?" else manfilterurl=sectionurl&"?"&manfilterurl
sub dofilterresults(numfcols)
	if prodfilter<>0 AND NOT (prodfilter=8 AND sortoptions=0) then
		if (prodfilter AND 2)=2 then
			searchcriterialist=""
			currgroupid=-1
			if NOT hascheckedectfilters then
				sSQL="SELECT COUNT("&IIfVr(sqlserver,"DISTINCT products.pID","*")&") as tcount,scID,"&getlangid("scName",131072)&",scGroup,scOrder,scgTitle FROM ((searchcriteria INNER JOIN searchcriteriagroup ON searchcriteria.scGroup=searchcriteriagroup.scgID) " & _
					"INNER JOIN multisearchcriteria ON multisearchcriteria.mSCscID=searchcriteria.scID) "&IIfVr(displayzeroattributes,"LEFT","INNER")&" JOIN "&IIfVs(sectionids<>"","(")&"products"&IIfVs(sectionids<>""," LEFT JOIN multisections ON products.pId=multisections.pId)")&" ON multisearchcriteria.mSCpID=products.pID " & _
					IIfVr(displayzeroattributes,"AND","WHERE")&" pDisplay<>0" & IIfVs(ectsiteid<>"", " AND pSiteID=" & ectsiteid)
				if useStockManagement AND noshowoutofstock then sSQL=sSQL & " AND (pInStock>pMinQuant OR pStockByOpts<>0)"
				if sectionids<>"" then sSQL=sSQL&" AND (products.pSection IN ("&sectionids&") OR multisections.pSection IN (" & sectionids & "))"
				sSQL=sSQL&filtersql&" AND (multisearchcriteria.mSCpID IN (SELECT products.pID FROM "
				if numscrid>1 then sSQL=sSQL&string(numscrid-1,"(")
				sSQL=sSQL&"(products LEFT JOIN multisections ON products.pId=multisections.pId)" & IIfVs(numscrid>0," INNER JOIN multisearchcriteria ON multisearchcriteria.mSCpID=products.pID")
				for scrindex=1 to numscrid-1
					sSQL=sSQL&") INNER JOIN multisearchcriteria msc"&scrindex&" ON products.pID=msc"&scrindex&".mSCpID"
				next
				sSQL=sSQL&" WHERE 1=1"
				if manid<>"0" AND is_numeric(manid) then sSQL=sSQL & " AND pManufacturer=" & manid
				if scSQL<>"" then sSQL=sSQL&" AND " & scSQL
				sSQL=sSQL&")"&IIfVs(numscrid>0," OR scGroup=0")&") GROUP BY scID,"&getlangid("scName",131072)&",scGroup,scOrder,scgOrder,scgID,scgTitle ORDER BY scgOrder,scgID,scOrder,"&getlangid("scName",131072)
				rs2.open sSQL,cnn,0,1
				if NOT rs2.EOF then ectfiltercache=rs2.getrows
				rs2.close
			end if
			hascheckedectfilters=TRUE
			if isarray(ectfiltercache) then
				for cacheindex=0 to UBOUND(ectfiltercache,2)
					if currgroupid<>ectfiltercache(3,cacheindex) then
						maxgroupid=maxgroupid+1
						if searchcriterialist<>"" then searchcriterialist=searchcriterialist&"</select>"
						searchcriterialist=searchcriterialist&"<select name=""scri"" class=""ectselectinput prodfilter"" id=""scri"&maxgroupid&""" onchange=""filterbyman(1)"""&IIfVr(sidefilterstyle="multiple"," size=""5"" multiple=""multiple"""," size=""1""")&"><option value="""" style=""font-weight:bold"">== All " & ectfiltercache(5,cacheindex) & " ==</option>" & vbCrLf
						currgroupid=ectfiltercache(3,cacheindex)
					end if
					searchcriterialist=searchcriterialist&"<option value="""&ectfiltercache(1,cacheindex)&""""&IIfVs(isinscrid(ectfiltercache(1,cacheindex))," selected=""selected""")&">" & ectfiltercache(2,cacheindex) & IIfVs(NOT isinscrid(ectfiltercache(1,cacheindex))," (" & ectfiltercache(0,cacheindex) & ")") & "</option>" & vbCrLf
				next
			else
				prodfilter=prodfilter-2
			end if
			if searchcriterialist<>"" then searchcriterialist=searchcriterialist&"</select>"
		end if
		maxprice=0 : minprice=0
		if (prodfilter AND 4)=4 then
			sSQL="SELECT MAX(" & TWSP & ") AS maxprice,MIN(" & TWSP & ") AS minprice FROM products WHERE pDisplay<>0" & IIfVs(ectsiteid<>"", " AND pSiteID=" & ectsiteid)
			if sectionids<>"" then sSQL="SELECT MAX(" & TWSP & ") AS maxprice,MIN(" & TWSP & ") AS minprice FROM (products LEFT JOIN multisections ON products.pId=multisections.pId) WHERE " & IIfVs(ectsiteid<>"", "pSiteID=" & ectsiteid & " AND ") & "pDisplay<>0 AND (products.pSection IN ("&sectionids&") OR multisections.pSection IN (" & sectionids & "))"
			rs2.open sSQL,cnn,0,1
			if NOT rs2.EOF then if NOT isnull(rs2("maxprice")) then maxprice=rs2("maxprice") : minprice=rs2("minprice")
			if showtaxinclusive=2 then maxprice=maxprice+(maxprice*(countryTaxRate/100.0)) : minprice=minprice+(minprice*(countryTaxRate/100.0))
			rs2.Close
		end if
		textarray=split(prodfiltertext,"&")
		Dim filtertext(10)
		if isarray(textarray) then
			for index=0 to 9
				if UBOUND(textarray)>=index then filtertext(index)=replace(textarray(index),"%26","&")
			next
		end if
		print "<div class=""prodfilterbar"">"
%><script>
/* <![CDATA[ */
function filterbyman(caller){
var furl="<%=replace(replace(manfilterurl,"<",""),"""","\""")%>";
if(document.getElementById('sman')){
	var smanobj=document.getElementById('sman');
	if(smanobj.selectedIndex!=0) furl+='sman='+smanobj[smanobj.selectedIndex].value+'&';
}
<%	for index=1 to maxgroupid %>
	var smanobj=document.getElementById('scri<%=index%>');
	for(var i=1;i<smanobj.length;i++){
        if(smanobj.options[i].selected) furl+='scri='+smanobj.options[i].value+'&';
    }
<%	next %>
if(document.getElementById('spriceobj')){
	var spriceobj=document.getElementById('spriceobj');
	if(spriceobj.selectedIndex!=0) furl+='sprice='+spriceobj[spriceobj.selectedIndex].value+'&';
}
if(document.getElementById('ectfilter')){
	if(document.getElementById('ectfilter').value!='')
		furl+='filter='+encodeURIComponent(document.getElementById('ectfilter').value)+'&';
}
document.location=furl.substr(0,furl.length-1);
}
function changelocation(fact,tobj){
document.location='<%=filterurl%>'.replace(/filter=/,fact+'='+tobj[tobj.selectedIndex].value<% if (prodfilter AND 32)=32 then print "+'&filter='+encodeURIComponent(document.getElementById('ectfilter').value)" %>);
}
function changelocfiltertext(tkeycode,tobj){
if(tkeycode==13)document.location='<%=filterurl%>'+tobj.value;
}
/* ]]> */</script>
<%		if prodfilterorder="" then prodfilterorder="1,2,4,8,16,32"
		filterorderarray=split(prodfilterorder,",")
		for indexfilterorder=0 to UBOUND(filterorderarray)
			select case filterorderarray(indexfilterorder)
			case 2
			if (prodfilter AND 2)=2 then ' Search Criteria
				if filtertext(1)<>"" then print "<div class=""prodfiltergrp ectpfattgrp""><div class=""prodfilter filtertext ectpfatttext"">" & filtertext(1) & "</div>"
				print "<div class=""prodfilter ectpfatt"">" & searchcriterialist & "</div>"
				if filtertext(1)<>"" then print "</div>"
			end if
			case 4
			if (prodfilter AND 4)=4 then ' Price bands
				if filtertext(2)<>"" then print "<div class=""prodfiltergrp ectpfpricegrp""><div class=""prodfilter filtertext ectpfpricetext"">" & filtertext(2) & "</div>"
				rowcounter=2
				currpriceband=getget("sprice")
				print "<div class=""prodfilter ectpfprice"">"
				%><select name="sprice" class="ectselectinput prodfilter" id="spriceobj" size="1" onchange="filterbyman(4)">
				<option value="0"><%=xxPriRan%></option>
<%				if minprice=0 OR filterpricebands>=minprice then %>
				<option value="1"<%if currpriceband="1" then print " selected=""selected"""%>><%=xxFilUnd&" "&FormatCurrencyZeroDP(filterpricebands)%></option>
<%				end if
				if instr(sprice,"-")>0 AND paminprice<>"" AND pamaxprice<>"" AND filterpricebands>=paminprice then %>
				<option value="<%=sprice%>" selected="selected"><%=FormatCurrencyZeroDP(paminprice)&" - "&FormatCurrencyZeroDP(pamaxprice)%></option>
<%				end if
				for index=filterpricebands to maxprice step filterpricebands
					if instr(sprice,"-")>0 AND paminprice<>"" AND pamaxprice<>"" AND index<=paminprice AND (index+filterpricebands)>=paminprice then %>
				<option value="<%=sprice%>" selected="selected"><%=FormatCurrencyZeroDP(paminprice)&" - "&FormatCurrencyZeroDP(pamaxprice)%></option>
<%					end if
					if minprice=0 OR (index+filterpricebands)>=minprice then %>
				<option value="<%=rowcounter%>"<%if currpriceband=cstr(rowcounter) then print " selected=""selected"""%>><%=FormatCurrencyZeroDP(index)&" - "&FormatCurrencyZeroDP(index+filterpricebands)%></option>
<%					end if
					rowcounter=rowcounter+1
					if rowcounter>1000 then exit for
				next %>
			  </select><%
				print "</div>"
				if filtertext(2)<>"" then print "</div>"
			end if
			case 8
			if (prodfilter AND 8)=8 AND sortoptions<>0 then
				if filtertext(3)<>"" then print "<div class=""prodfiltergrp ectpfsortgrp""><div class=""prodfilter filtertext ectpfsorttext"">" & filtertext(3) & "</div>"
				print "<div class=""prodfilter ectpfsort"">"
				%><select class="ectselectinput prodfilter" size="1" onchange="changelocation('sortby',this)">
				<option value="0"><%=xxPlsSel%></option>
<%				call sortline(1, IIfVr(sortoption1<>"",sortoption1,"Sort Alphabetically"))
				call sortline(11, IIfVr(sortoption11<>"",sortoption11,"Alphabetically (Desc.)"))
				call sortline(2, IIfVr(sortoption2<>"",sortoption2,"Sort by Product ID"))
				call sortline(12, IIfVr(sortoption12<>"",sortoption12,"Product ID (Desc.)"))
				call sortline(14, IIfVr(sortoption14<>"",sortoption14,"Sort By SKU"))
				call sortline(15, IIfVr(sortoption15<>"",sortoption15,"Sort By SKU (Desc.)"))
				call sortline(3, IIfVr(sortoption3<>"",sortoption3,"Sort Price (Asc.)"))
				call sortline(4, IIfVr(sortoption4<>"",sortoption4,"Sort Price (Desc.)"))
				call sortline(5, IIfVr(sortoption5<>"",sortoption5,"Database Order"))
				call sortline(6, IIfVr(sortoption6<>"",sortoption6,"Product Order"))
				call sortline(7, IIfVr(sortoption7<>"",sortoption7,"Product Order (Desc.)"))
				call sortline(8, IIfVr(sortoption8<>"",sortoption8,"Date Added (Asc.)"))
				call sortline(9, IIfVr(sortoption9<>"",sortoption9,"Date Added (Desc.)"))
				call sortline(10, IIfVr(sortoption10<>"",sortoption10,"Sort by Manufacturer"))
				call sortline(16, IIfVr(sortoption16<>"",sortoption16,"Number of Ratings"))
				call sortline(17, IIfVr(sortoption17<>"",sortoption17,"Average Rating"))
				call sortline(18, IIfVr(sortoption18<>"",sortoption18,"Sales Rank"))
				call sortline(19, IIfVr(sortoption19<>"",sortoption19,"Popularity"))
%>			  </select><%
				print "</div>"
				if filtertext(3)<>"" then print "</div>"
			end if
			case 16
			if (prodfilter AND 16)=16 then
				if filtertext(4)<>"" then print "<div class=""prodfiltergrp ectpfpagegrp""><div class=""prodfilter filtertext ectpfpagetext"">" & filtertext(4) & "</div>"
				print "<div class=""prodfilter ectpfpage"">"
				%><select class="ectselectinput prodfilter" size="1" onchange="changelocation('perpage',this)">
<%				for index=1 to 5
					print "<option value="""&index&""""&IIfVr(SESSION("perpage")=index," selected=""selected""","")&">"&(prodsperpage*index)&" "&xxPerPag&"</option>"
				next
%>			  </select><%
				print "</div>"
				if filtertext(4)<>"" then print "</div>"
			end if
			case 32
			if (prodfilter AND 32)=32 then
				if filtertext(5)<>"" then print "<div class=""prodfiltergrp ectpfkeywordgrp""><div class=""prodfilter filtertext ectpfkeywordtext"">" & filtertext(5) & "</div>"
				print "<div class=""prodfilter ectpfkeyword"">" %><input onkeydown="changelocfiltertext(event.keyCode,this)" type="text" class="ecttextinput prodfilter" size="20" id="ectfilter" name="filter" value="<%=htmlspecials(getget("filter"))%>" /><%
				print imageorbutton(imgfilterproducts,xxGo,"prodfilter","document.location='"&replace(filterurl,"&","&amp;")&"'+encodeURIComponent(document.getElementById('ectfilter').value)",TRUE)
				print "</div>"
				if filtertext(5)<>"" then print "</div>"
			end if
			end select
		next
		print "</div>"
	end if
end sub
if orprodsperpage<>"" then adminProdsPerPage=orprodsperpage
prodsperpage=adminProdsPerPage
Call checkCurrencyRates(currConvUser,currConvPw,currLastUpdate,currRate1,currSymbol1,currRate2,currSymbol2,currRate3,currSymbol3)
get_wholesaleprice_sql()
tslist=""
thetopts=catid
topsectionids=catid
isrootsection=false
sectiondisabled=false
if SESSION("clientID")<>"" AND SESSION("clientLoginLevel")<>"" then minloglevel=SESSION("clientLoginLevel") else minloglevel=0
if manufacturers then
	sSQL="SELECT "&getlangid("scName",131072)&","&getlangid("scHeader",524288)&" FROM searchcriteria WHERE scID="&manid
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then mfname=rs(getlangid("scName",131072)) : sectionheader=trim(rs(getlangid("scHeader",524288))&"") else mfname="Not Found" : response.status="404 Not Found" : xxNoPrds="<div class=""ectwarning"" style=""text-align:center;padding:50px;overflow:auto;"">This manufacturer could not be found.</div>"
	rs.close
	tslist="<a class=""ectlink"" href="""&storehomeurl&""">"&xxHome&"</a> &raquo; <a class=""ectlink"" href="""&manufacturerpageurl&""">"&xxManuf&"</a> &raquo; " & mfname
	if explicitmanid<>"" then sectionurl=request.servervariables("URL")
	isrootsection=TRUE
else
	for index=0 to 10
		if cstr(thetopts)=cstr(catalogroot) then
			caturl=storehomeurl
			sSQL="SELECT sectionID,topSection,"&getlangid("sectionName",256)&",rootSection,sectionDisabled,"&getlangid("sectionurl",2048)&","&getlangid("sectionHeader",524288)&" FROM sections WHERE sectionID=" & catalogroot
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				xxHome=rs(getlangid("sectionName",256))
				if trim(rs(getlangid("sectionurl",2048))&"")<>"" then caturl=rs(getlangid("sectionurl",2048))
				if rs("sectionID")=int(catid) then sectionheader=rs(getlangid("sectionHeader",524288))
			end if
			rs.close
			tslist="<a class=""ectlink"" href="""&caturl&""">"&xxHome&"</a>" & tslist
			exit for
		elseif index=10 then
			tslist="<strong>Loop</strong>" & tslist
		else
			sSQL="SELECT sectionID,topSection,"&getlangid("sectionName",256)&",rootSection,sectionDisabled,"&IIfVs(languageid<>1,getlangid("sectionurl",2048)&" AS ")&"sectionurl,"&getlangid("sectionHeader",524288)&" FROM sections WHERE sectionID=" & thetopts
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if rs("sectionID")=int(catid) then isrootsection=(rs("rootSection")=1) : sectionheader=rs(getlangid("sectionHeader",524288))
				if rs("sectionDisabled")>minloglevel then catid=-1
				if rs("sectionID")=int(catid) AND isrootsection then
					if dynamicbreadcrumbs then
						tslist="<div class=""ectbreadcrumb""><span class=""navcurrentcat"">&raquo; " & breadcrumbselect(rs("sectionID"),rs("topSection")) & "</span></div>" & tslist
					else
						tslist="<div class=""ectbreadcrumb""><span class=""navcurrentcat"">&raquo; " & rs(getlangid("sectionName",256)) & "</span></div>" & tslist
					end if
					if explicitid<>"" AND trim(rs("sectionurl")&"")<>"" then sectionurl=rs("sectionurl")
					if explicitid="" AND trim(rs("sectionurl"))<>"" AND redirecttostatic=TRUE then
						response.status="301 Moved Permanently"
						response.addheader "Location", rs("sectionurl")
						response.end
					end if
				else
					if dynamicbreadcrumbs then
						tslist="<div class=""ectbreadcrumb"">&raquo; " & breadcrumbselect(rs("sectionID"),rs("topSection")) & "</div>" & tslist
					else
						tslist="<div class=""ectbreadcrumb"">&raquo; <a class=""ectlink"" href=""" & getcategoryurl(rs("sectionID"),rs(getlangid("sectionName",256)),rs("sectionurl"),rs("rootSection")) & """>" & rs(getlangid("sectionName",256)) & "</a></div>" & tslist
					end if
					if trim(rs("sectionurl")&"")<>"" AND explicitid<>"" AND rs("sectionID")=int(catid) then sectionurl=rs("sectionurl")
				end if
				thetopts=rs("topSection")
				topsectionids=topsectionids & "," & thetopts
			else
				if tslist="" AND catname<>"Not Found" then
					catname="Not Found"
					catid=-1
					response.status="404 Not Found"
				end if
				tslist="<a class=""ectlink"" href="""&storehomeurl&""">"&xxHome&"</a> &raquo; Top Section Deleted" & tslist & " &raquo; <a class=""ectlink"" href=""" & IIfVr(seocategoryurls,replace(seoprodurlpattern,"%s",""),"products" & extension) & """>" & xxAlProd & "</a>"
				xxAlProd=""
				rs.close
				exit for
			end if
			rs.close
		end if
	next
end if
if NOT isrootsection AND xxAlProd<>"" then tslist=tslist & "<div class=""ectbreadcrumb"">&raquo; "&xxAlProd&"</div>"
filtersql=""
if getget("filter")<>"" then
	sText=escape_string(left(getget("filter"), 1024))
	aText=Split(sText)
	aFields(0)="products.pId"
	aFields(1)=getlangid("pName",1)
	aFields(2)=getlangid("pDescription",2)
	aFields(3)=getlangid("pLongDescription",4)
	aFields(4)="pSKU"
	aFields(5)="pSearchParams"
	sNOTSQL="" : sYESSQL=""
	for rowcounter=0 to UBOUND(aText)
		tmpSQL=""
		for index=0 to 5
			if NOT ((nosearchdescription=TRUE AND index=2) OR (nosearchlongdescription=TRUE AND index=3) OR (nosearchsku=TRUE AND index=4) OR (nosearchparams=TRUE AND index=5)) then
				tmpSQL=tmpSQL & getlike(aFields(index), aText(rowcounter), "OR ")
			end if
		next
		if tmpSQL<>"" then sYESSQL=sYESSQL & "(" & left(tmpSQL, len(tmpSQL)-3) & ") "
		if tmpSQL<>"" then sYESSQL=sYESSQL & "AND "
	next
	if sYESSQL<>"" then sYESSQL=left(sYESSQL,len(sYESSQL)-4)
	if sYESSQL<>"" then filtersql=" AND (" & sYESSQL & ")"
	if sNOTSQL<>"" then filtersql=filtersql & " AND NOT (" & left(sNOTSQL, len(sNOTSQL)-4) & ")"
end if
paminprice="" : pamaxprice=""
if sprice<>"" then
	taxlevel=1
	if showtaxinclusive=2 then taxlevel=taxlevel+(countryTaxRate/100.0)
	if instr(sprice,"-")>0 then
		spricearr=split(sprice,"-")
		paminprice=spricearr(0)
		pamaxprice=spricearr(1)
		if is_numeric(pamaxprice) then
			if NOT is_numeric(paminprice) then paminprice=0
		else
			pamaxprice=""
		end if
	elseif is_numeric(sprice) then
		priceband=int(sprice)
		paminprice=(priceband-1)*filterpricebands
		pamaxprice=priceband*filterpricebands
	end if
	session.LCID=1033
	if paminprice<>"" AND pamaxprice<>"" then filtersql=filtersql & " AND (("&TWSP&"*"&taxlevel&")>=" & paminprice & " AND ("&TWSP&"*"&taxlevel&")<=" & pamaxprice & ")"
	session.LCID=saveLCID
end if
if (prodfilter AND 1)=1 AND NOT manufacturers then
	manid=getget("sman")
	if NOT is_numeric(manid) then manid=""
end if
if ((prodfilter AND 2)=2 OR (sidefilter AND 2)=2) AND NOT alreadycheckedscsql then
	numscrgroup=0
	numscrid=0
	if scrid<>"" then
		currgroup=-1
		scSQL="(multisearchcriteria.mSCscID IN ("
		sSQL="SELECT scID,scGroup FROM searchcriteria WHERE scID IN (" & scrid & ") ORDER BY scGroup"
		rs.open sSQL,cnn,0,1
		scrid=""
		do while NOT rs.EOF
			scrid=scrid&rs("scID")&","
			if rs("scGroup")<>currgroup then
				addcomma=""
				if currgroup<>-1 then scSQL=scSQL&") AND msc"&numscrgroup&".mSCscID IN ("
				numscrgroup=numscrgroup+1
				currgroup=rs("scGroup")
			end if
			scSQL=scSQL&addcomma&rs("scID")
			addcomma=","
			rs.movenext
		loop
		rs.close
		if scrid<>"" then
			scrid=left(scrid,len(scrid)-1)
			scridarr=split(scrid,",")
			numscrid=UBOUND(scridarr)+1
		end if
		if numscrid=0 then scSQL="" else scSQL=scSQL&"))"
	else
		scSQL=""
	end if
	alreadycheckedscsql=TRUE
end if
if getget("recentview")="true" then
	sSQL="SELECT products.pId,"&getlangid("pName",1)&","&WSP&"pPrice,pOrder,pDateAdded,pManufacturer,pSKU,pPopularity,pNumSales,pNumRatings,pTotRating" & IIfVs(dosortby=17 AND sqlserver AND NOT mysqlserver,",CASE WHEN pNumRatings=0 THEN 0 ELSE pTotRating/pNumRatings END") & " FROM products INNER JOIN recentlyviewed ON recentlyviewed.rvProdID=products.pID WHERE " & IIfVr(SESSION("clientID")<>"", "rvCustomerID="&replace(SESSION("clientID"),"'",""), "(rvCustomerID=0 AND rvSessionID='"&getsessionid()&"')")
else
	rs.open "SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 1 ")&"sectionID FROM sections WHERE sectionDisabled>"&minloglevel&IIfVs(mysqlserver=TRUE," LIMIT 0,1"),cnn,0,1
	if NOT rs.EOF then disabledsections=TRUE else disabledsections=FALSE
	rs.close
	if cstr(catid)=cstr(catalogroot) AND ectsiteid="" then
		sSQL="SELECT pId FROM "
		if numscrgroup>1 then sSQL=sSQL&string(numscrgroup-1,"(")
		sSQL=sSQL&IIfVs(numscrgroup>0 AND disabledsections,"(")&"products" & IIfVs(disabledsections," INNER JOIN sections ON products.pSection=sections.sectionID"&IIfVs(numscrgroup>0,")")) & IIfVs(numscrgroup>0," INNER JOIN multisearchcriteria ON multisearchcriteria.mSCpID=products.pID")
		for scrindex=1 to numscrgroup-1
			sSQL=sSQL&") INNER JOIN multisearchcriteria msc"&scrindex&" ON products.pID=msc"&scrindex&".mSCpID"
		next
		sSQL=sSQL&" WHERE" & IIfVs(disabledsections," sectionDisabled<="&minloglevel&" AND") & IIfVs(ectsiteid<>"", " pSiteID=" & ectsiteid & " AND") & " pDisplay<>0"&filtersql
	else
		sectionids=getsectionids(catid, false)
		sSQL="SELECT DISTINCT"&IIfVs(NOT sqlserver,"ROW")&" products.pId,"&getlangid("pName",1)&","&WSP&"pPrice,pOrder,pDateAdded,pManufacturer,pSKU,pPopularity,pNumSales,pNumRatings,pTotRating" & IIfVs(dosortby=17 AND sqlserver AND NOT mysqlserver,",CASE WHEN pNumRatings=0 THEN 0 ELSE pTotRating/pNumRatings END") & " FROM "
		if numscrgroup>1 then sSQL=sSQL&string(numscrgroup-1,"(")
		sSQL=sSQL&"("&IIfVs(disabledsections,"(")&"products"&IIfVs(disabledsections," INNER JOIN sections ON products.pSection=sections.sectionID)")&" LEFT JOIN multisections ON products.pId=multisections.pId)" & IIfVs(numscrgroup>0," INNER JOIN multisearchcriteria ON multisearchcriteria.mSCpID=products.pID")
		for scrindex=1 to numscrgroup-1
			sSQL=sSQL&") INNER JOIN multisearchcriteria msc"&scrindex&" ON products.pID=msc"&scrindex&".mSCpID"
		next
		sSQL=sSQL&" WHERE" & IIfVs(disabledsections," sectionDisabled<="&minloglevel&" AND") & IIfVs(ectsiteid<>"", " pSiteID=" & ectsiteid & " AND") & " pDisplay<>0"&filtersql&" AND (products.pSection IN (" & sectionids & ") OR multisections.pSection IN (" & sectionids & "))"
	end if
	if manid<>"0" AND is_numeric(manid) then sSQL=sSQL & " AND pManufacturer=" & manid
	if scSQL<>"" then sSQL=sSQL & " AND " & scSQL
	if numscrid>0 OR sprice<>"" OR getget("filter")<>"" then xxNoPrds=xxNoMatc&"<div class=""resetfilters"">"&imageorbutton(resetfilters,xxResFil,"resetfilters",manfilterurl,FALSE)&"</div>"
	if useStockManagement AND noshowoutofstock then sSQL=sSQL & " AND (pInStock>pMinQuant OR pStockByOpts<>0)"
end if
if is_numeric(request("perpage")) then SESSION("perpage")=int(request("perpage"))
if is_numeric(SESSION("perpage")) then adminProdsPerPage=int(SESSION("perpage"))*prodsperpage
if adminProdsPerPage>1000 then adminProdsPerPage=prodsperpage
Redim aDiscSection(2,adminProdsPerPage)
if dosortby=2 OR dosortby=12 then
	sSortBy=" ORDER BY products.pID"&IIfVs(dosortby=12," DESC")
elseif dosortby=14 OR dosortby=15 then
	sSortBy=" ORDER BY pSKU"&IIfVs(dosortby=15," DESC")
elseif dosortby=3 OR dosortby=4 then
	sSortBy=" ORDER BY "&TWSP&IIfVr(dosortby=4," DESC,products.pId",",products.pId")
elseif dosortby=5 then
	sSortBy=""
elseif dosortby=6 OR dosortby=7 then
	sSortBy=" ORDER BY pOrder"&IIfVr(dosortby=7," DESC,products.pId",",products.pId")
elseif dosortby=8 OR dosortby=9 then
	sSortBy=" ORDER BY pDateAdded"&IIfVr(dosortby=9," DESC,products.pId",",products.pId")
elseif dosortby=10 then
	sSortBy=" ORDER BY pManufacturer"
elseif dosortby=16 then
	sSortBy=" ORDER BY pNumRatings DESC,products.pId"
elseif dosortby=17 then
	sSortBy="CASE WHEN pNumRatings=0 THEN 0 ELSE pTotRating/pNumRatings END"
	if mysqlserver OR NOT sqlserver then sSortBy=IIfVs(NOT sqlserver,"I")&"IF(pNumRatings=0,0,pTotRating/pNumRatings)"
	sSortBy=" ORDER BY "&sSortBy&" DESC,pNumRatings DESC,products.pId"
elseif dosortby=18 then
	sSortBy=" ORDER BY pNumSales DESC,products.pId"
elseif dosortby=19 then
	sSortBy=" ORDER BY pPopularity DESC,products.pId"
else
	sSortBy=" ORDER BY "&getlangid("pName",1)&IIfVr(dosortby=11," DESC,products.pId",",products.pId")
end if
rs.CursorLocation=3 ' adUseClient
rs.CacheSize=adminProdsPerPage
rs.open sSQL & sSortBy, cnn
if NOT rs.EOF then
	rs.MoveFirst
	rs.PageSize=adminProdsPerPage
	iNumOfPages=int((rs.RecordCount + (adminProdsPerPage-1)) / adminProdsPerPage)
	if NOT is_numeric(getget("pg")) then CurPage=1 else CurPage=vrmin(vrmax(1, int(getget("pg"))),iNumOfPages)
	rs.AbsolutePage=CurPage
end if
if IsEmpty(Count) then Count=0 else Count=(Count+adminProdsPerPage)-(Count MOD adminProdsPerPage)
if NOT rs.EOF then
	prodlist=""
	addcomma=""
	prodcount=0
	do while NOT rs.EOF AND prodcount<rs.PageSize
		prodlist=prodlist & addcomma & "'" & escape_string(rs("pId")) & "'"
		rs.MoveNext
		prodcount=prodcount+1
		addcomma=","
	loop
	rs.close
	wantmanufacturer=instr(productpagelayout&quickbuylayout,"manufacturer")>0 OR (useproductbodyformat=3 AND instr(cpdcolumns, "manufacturer")>0) OR ((NOT usecsslayout OR useproductbodyformat<>2) AND xxManLab<>"") OR googletagid<>""
	sSQL="SELECT pId,pSKU,"&getlangid("pName",1)&","&WSP&"pPrice,pListPrice,pSection,pSell,pStockByOpts,pStaticPage,pStaticURL,pInStock,pExemptions,pTax,pTotRating,pNumRatings,pBackOrder,pCustomCSS,pCustom1,pCustom2,pCustom3,pDateAdded,pMinQuant,pSchemaType,"&IIfVr(wantmanufacturer,getlangid("scName",131072)&",","")&IIfVr(shortdescriptionlimit<>"" AND shortdescriptionlimit=0,"'' AS ","")&getlangid("pDescription",2)&","&getlangid("pLongDescription",4)&" FROM products "&IIfVr(wantmanufacturer,"LEFT OUTER JOIN searchcriteria on products.pManufacturer=searchcriteria.scID ","")&"WHERE pId IN (" & prodlist & ")" & sSortBy
	rs.open sSQL, cnn, 0, 1
end if
if nowholesalediscounts=TRUE AND SESSION("clientUser")<>"" then
	if ((SESSION("clientActions") AND 8)=8) OR ((SESSION("clientActions") AND 16)=16) then noshowdiscounts=TRUE
end if
if noshowdiscounts<>TRUE then
	Session.LCID=1033
	sSQL="SELECT DISTINCT cpnID,"&getlangid("cpnName",1024)&",cpnType,cpnSitewide,cpaType FROM coupons LEFT OUTER JOIN cpnassign ON coupons.cpnID=cpnassign.cpaCpnID WHERE ("
	addor=""
	if catid<>"0" OR manufacturers then
		sSQL=sSQL & addor & "((cpnSitewide=0 OR cpnSitewide=3) AND cpaType="&IIfVr(manufacturers,3,1)&" AND cpaAssignment IN ('"&IIfVr(manufacturers,manid,replace(topsectionids,",","','"))&"'))"
		addor=" OR "
	end if
	tdt=Date()
	sSQL=sSQL & addor & "(cpnSitewide=1 OR cpnSitewide=2)) AND cpnNumAvail>0 AND cpnStartDate<=" & vsusdate(tdt)&" AND cpnEndDate>=" & vsusdate(tdt)&" AND cpnIsCoupon=0 AND ((cpnLoginLevel>=0 AND cpnLoginLevel<="&minloglevel&") OR (cpnLoginLevel<0 AND -1-cpnLoginLevel="&minloglevel&")) ORDER BY "&getlangid("cpnName",1024)
	Session.LCID=saveLCID
	rs2.Open sSQL,cnn,0,1
	if NOT rs2.EOF then
		lastcouponname=""
		do while NOT rs2.EOF
			if rs2(getlangid("cpnName",1024))<>lastcouponname then globaldiscounttext=globaldiscounttext & "<div class=""adiscount"">" & rs2(getlangid("cpnName",1024)) & "</div>" : lastcouponname=rs2(getlangid("cpnName",1024))
			if rs2("cpnType")=0 then hasshippingdiscount=TRUE else hasproductdiscount=TRUE
			if catid<>"0" OR manid<>"" then
				if (rs2("cpnSitewide")=0 OR rs2("cpnSitewide")=3) AND (rs2("cpaType")=1 OR rs2("cpaType")=3) then
					globaldiscounts(0,maxglobaldiscounts)=rs2("cpnID")
					globaldiscounts(1,maxglobaldiscounts)=rs2(getlangid("cpnName",1024))
					globaldiscounts(2,maxglobaldiscounts)="xxx"
					if maxglobaldiscounts<29 then maxglobaldiscounts=maxglobaldiscounts+1
				end if
			end if
			rs2.movenext
		loop
	end if
	rs2.Close
end if
if ectsiteid<>"" then savecatid=catid
alreadycalculatedattributes=TRUE
		if NOT usecsslayout then
%>	<table border="0" cellspacing="0" cellpadding="0" width="98%" align="center">
		<tr> 
			<td colspan="3" width="100%">
<%		else
			print "<div>"
		end if
		if sectionheader<>"" then print "<div class=""catheader"">" & sectionheader & "</div>"
		if useproductbodyformat=3 then %>
<!--#include file="incproductbody3.asp"-->
<%		elseif useproductbodyformat=2 then %>
<!--#include file="incproductbody2.asp"-->
<%		else %>
<!--#include file="incproductbody.asp"-->
<%		end if
		if NOT usecsslayout then %>
			</td>
		</tr>
	</table>
<%		else
			print "</div>"
		end if
rs.close
cnn.Close
set rs=nothing
set rs2=nothing
set cnn=nothing
%>