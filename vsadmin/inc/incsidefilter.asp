<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
set rs=Server.CreateObject("ADODB.RecordSet")
set rs2=Server.CreateObject("ADODB.RecordSet")
set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
sSQL="SELECT sideFilterOrder,"&getlangid("sideFilterText",262144)&",adminProdsPerPage FROM admin WHERE adminID=1"
rs.open sSQL,cnn,0,1
sidefilterorder=rs("sideFilterOrder")
sidetextarray=split(rs(getlangid("sideFilterText",262144)),"&")
adminProdsPerPage=rs("adminProdsPerPage")
displayzeroattributes=(sqlserver OR mysqlserver) AND displayzeroattributes
DIM sidefiltertext(10)
if isarray(sidetextarray) then
	for index=0 to 9
		if UBOUND(sidetextarray)>=index then sidefiltertext(index)=replace(sidetextarray(index),"%26","&")
	next
end if
rs.close
if getget("cat")<>"" then catid=getget("cat") else catid=""
if getget("man")<>"" then manid=getget("man") else manid=""
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
set toregexp=new regexp
toregexp.pattern="[^0-9\-\.]"
toregexp.global=TRUE
sprice=toregexp.replace(getget("sprice"),"")
set toregexp=nothing
if is_numeric(explicitid) then catid=explicitid
if is_numeric(explicitmanid) then manid=explicitmanid
if is_numeric(request("sortby")) then SESSION("sortby")=int(request("sortby"))
if SESSION("sortby")<>"" then dosortby=SESSION("sortby")
if is_numeric(request("perpage")) then SESSION("perpage")=int(request("perpage"))
if orprodsperpage<>"" then adminProdsPerPage=orprodsperpage
prodsperpage=adminProdsPerPage
if seocategoryurls then usecategoryname=TRUE : catid=replace(catid,detlinkspacechar," ") : manid=replace(manid,detlinkspacechar," ")
if sidefiltermoreafter="" then sidefiltermoreafter=8
if sidefilterclosedheight="" then sidefilterclosedheight=150
TWSP="pPrice"
if filterpricebands="" then filterpricebands=100
function pasortline(soid, sotext)
	if (sortoptions AND (2 ^ (soid-1)))<>0 then print "<option value="""&soid&""""&IIfVr(dosortby=soid," selected=""selected""","")&">"&sotext&"</option>"
end function
function pagetlike(fie,t,tjn)
	if left(t, 1)="-" then ' pSKU excluded to work around NULL problems
		if fie<>"pSKU" then sNOTSQL=sNOTSQL & fie & " LIKE '%"&mid(t, 2)&"%' OR "
	else
		pagetlike=fie & " LIKE '%"&t&"%' "&tjn
	end if
end function
if usecategoryname AND trim(catid)<>"" then
	sSQL="SELECT sectionID FROM sections WHERE "&IIfVs(seocategoryurls,getlangid("sectionurl",2048)&"='"&escape_string(catid)&"' OR (")&getlangid("sectionName",256)&"='"&escape_string(catid)&"'"&IIfVs(seocategoryurls," AND "&getlangid("sectionurl",2048)&"='')")
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then catname=catid : catid=rs("sectionID")
	rs.close
end if
if usecategoryname AND trim(manid)<>"" then
	sSQL="SELECT scID FROM searchcriteria WHERE "&getlangid("scName",131072)&"='"&escape_string(manid)&"' ORDER BY scGroup"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then manname=catid : manid=rs("scID")
	rs.close
end if
if NOT is_numeric(catid) then catid=catalogroot
if is_numeric(manid) OR manufacturers=TRUE then manufacturers=TRUE else manufacturers=FALSE : manid=""

HTTP_X_ORIGINAL_URL=trim(split(request.servervariables("HTTP_X_ORIGINAL_URL")&"?","?")(0))
if HTTP_X_ORIGINAL_URL="" then HTTP_X_ORIGINAL_URL=trim(split(request.servervariables("HTTP_X_REWRITE_URL")&"?","?")(0))
sectionurl=strip_tags2(IIfVr(seocategoryurls AND HTTP_X_ORIGINAL_URL<>"",HTTP_X_ORIGINAL_URL,request.servervariables("URL")))
filterurl="" : manfilterurl=""
for each objQS in request.querystring
	if objQS<>"recentview" AND objQS<>"filter" AND objQS<>"pg" AND objQS<>"sortby" AND objQS<>"perpage" AND NOT ((objQS="cat" OR objQS="man") AND seocategoryurls) then
		filterurl=filterurl & urlencode(objQS) & "=" & urlencode(getget(objQS)) & "&"
		if objQS<>"sman" AND objQS<>"scri" AND objQS<>"sprice" then manfilterurl=manfilterurl & urlencode(objQS) & "=" & urlencode(getget(objQS)) & "&"
	end if
	'if objQS="scri" and is_numeric(getget(objQS)) then allscri=allscri&addandscri&getget(objQS) : addandscri=","
next
if filterurl="" then filterurl=sectionurl&"?filter=" else filterurl=sectionurl&"?"&filterurl&"filter="
if manfilterurl="" then manfilterurl=sectionurl&"?" else manfilterurl=sectionurl&"?"&manfilterurl
filtersql=""
DIM aFieldsPA(5)
if getget("filter")<>"" then
	sText=escape_string(left(getget("filter"), 1024))
	aText=Split(sText)
	aFieldsPA(0)="products.pId"
	aFieldsPA(1)=getlangid("pName",1)
	aFieldsPA(2)=getlangid("pDescription",2)
	aFieldsPA(3)=getlangid("pLongDescription",4)
	aFieldsPA(4)="pSKU"
	aFieldsPA(5)="pSearchParams"
	sNOTSQL="" : sYESSQL=""
	for rowcounter=0 to UBOUND(aText)
		tmpSQL=""
		for index=0 to 5
			if NOT ((nosearchdescription=TRUE AND index=2) OR (nosearchlongdescription=TRUE AND index=3) OR (nosearchsku=TRUE AND index=4) OR (nosearchparams=TRUE AND index=5)) then
				tmpSQL=tmpSQL & pagetlike(aFieldsPA(index), aText(rowcounter), "OR ")
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
if (sidefilter AND 1)=1 AND NOT manufacturers then
	manid=getget("sman")
	if NOT is_numeric(manid) then manid=""
end if
%>
<script>
/* <![CDATA[ */
var furl="<%=replace(replace(manfilterurl,"<",""),"""","\""")%>";
var allscri=[<%=scrid%>];
var scriqs='<%=scrid%>';
var spriceqs='<%=sprice%>';
var spriceset=false;
var filterqs='<%=urlencode(getget("filter"))%>';
var allgroups=[];
function hasgotgroup(mID,gotgroups){
	for(var idind=0;idind<gotgroups.length;idind++){
		if(allgroups[mID]==gotgroups[idind]) return true;
	}
	return false;
}
function getallscr(mID,doadd,allowdupgroup){
	var rettxt='';
	var gotgroups=[];
	gotgroups.push(allgroups[mID]);
	for(var idind=0;idind<allscri.length;idind++){
		if(allowdupgroup||!hasgotgroup(allscri[idind],gotgroups)){
			if(doadd||mID!=allscri[idind])rettxt+=','+allscri[idind];
		}
	}
	return rettxt.substring(1);
}
function ectRAm(mID,addcurrid,removegroup){
	var currscr=getallscr(mID,addcurrid,!removegroup);
	var filter=getfilterqs();
	if(addcurrid) currscr=mID+(currscr!=''?','+currscr:'');
	var newurl=furl+(currscr!=''?'scri='+currscr+'&':'')+getpriceqs()+(filter!=''?'filter='+encodeURIComponent(filter)+'&':'');
	document.location=newurl.substring(0,newurl.length-1);
	return false;
}
<%	if sidefiltermorestyle="mouseover" then %>
function ectMOtGrpTr(grpid){
	var currheight=parseInt(document.getElementById('ectpatgrp'+grpid).style.maxHeight.replace(/px/,''));
	if(isNaN(currheight)) return;
	if(currheight><%=sidefilterclosedheight%>){
		document.getElementById('ectpatgrp'+grpid).style.maxHeight=(currheight-2)+'px';
		setTimeout("ectMOtGrpTr('"+grpid+"');",10);
	}else
		document.getElementById('grpMore'+grpid).style.display='';
}
function ectMOtGrp(tobj,grpid){
tobj.style.maxHeight=tobj.offsetHeight+'px';
setTimeout("ectMOtGrpTr('"+grpid+"');",100);
}
function ectMOGrp(tobj,grpid){
document.getElementById('grpMore'+grpid).style.display='none';
tobj.style.maxHeight='none';
}
<%	end if %>
var timerid;
function adjustprice(ismaxprice,isincrease){
	spriceset=true;
	doadjustprice(ismaxprice,isincrease);
	timerid=setInterval('doadjustprice('+ismaxprice+','+isincrease+')',100);
}
function stopprice(){
	clearInterval(timerid);
}
function doadjustprice(ismaxprice,isincrease){
	var minprice=parseInt(document.getElementById('paminprice').value);
	var maxprice=parseInt(document.getElementById('pamaxprice').value);
	var origmaxprice=parseInt(document.getElementById('origmaxprice').value);
	var addamount;
	if(isNaN(minprice))minprice=0;
	if(isNaN(maxprice))maxprice=origmaxprice;
	if(maxprice-minprice>100000)addamount=10000;
	if(maxprice-minprice>10000)addamount=1000;
	else if(maxprice-minprice>1000)addamount=100;
	else if(maxprice-minprice>100)addamount=10;
	else if(maxprice-minprice>50)addamount=5;
	else addamount=1;
	if(ismaxprice){
		maxprice=Math.max(isincrease?maxprice+addamount:maxprice-addamount,0);
		document.getElementById('pamaxprice').value=maxprice;
		if(maxprice<minprice)document.getElementById('paminprice').value=maxprice;
	}else{
		minprice=Math.max(isincrease?minprice+addamount:minprice-addamount,0);
		document.getElementById('paminprice').value=minprice;
		if(maxprice<minprice)document.getElementById('pamaxprice').value=minprice;
	}
}
function getpriceqs(){
	var newurl='';
	if(spriceset&&document.getElementById('paminprice')){
		var minprice=parseFloat(document.getElementById('paminprice').value);
		var maxprice=parseFloat(document.getElementById('pamaxprice').value);
		if(!isNaN(minprice))
			newurl+='sprice='+minprice+'-'+(isNaN(maxprice)?'':maxprice)+'&';
	}else if(spriceqs!='')
		newurl+='sprice='+spriceqs+'&';
	return(newurl);
}
function getfilterqs(){
	return document.getElementById('paectfilter')?document.getElementById('paectfilter').value:''
}
function filtergo(){
	var newurl='';
	var filter=getfilterqs();
	if(scriqs!='')newurl+='scri='+scriqs+'&';
	newurl+=getpriceqs();
	if(filter!='')newurl+='filter='+encodeURIComponent(filter)+'&';
	document.location=furl+newurl.substring(0,newurl.length-1);
	return false;
}
function changesflocation(fact,tobj){
document.location='<%=filterurl%>'.replace(/filter=/,fact+'='+tobj[tobj.selectedIndex].value<% if (sidefilter AND 32)=32 then print "+(getfilterqs()!=''?'&filter='+encodeURIComponent(getfilterqs()):'')" %>);
}
function changesffiltertext(tkeycode,tobj){
if(tkeycode==13)document.location='<%=filterurl%>'+tobj.value;
}
/* ]]> */
</script>
<%
function isidset(tid)
	isidset=FALSE
	if is_numeric(manid) AND manid=tid then
		isidset=TRUE
	elseif numscrid>0 then
		for scrindex=0 to numscrid-1
			if int(scridarr(scrindex))=tid then isidset=TRUE
		next
	end if
end function
if cstr(catid)<>cstr(catalogroot) then sectionids=getsectionids(catid, false) else sectionids=""
if trim(sidefilterorder&"")="" then sidefilterorder="2,4,8,16,32"
filterorderarray=split(sidefilterorder,",")
if (sidefilter AND 2)=2 AND NOT alreadycheckedscsql then
	numscrid=0
	numscrgroup=0
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
for indexfilterorder=0 to UBOUND(filterorderarray)
	select case filterorderarray(indexfilterorder)
	case 1
	if (sidefilter AND 1)=1 then ' Category Filter
		print "category filter"
	end if
	case 2
	if (sidefilter AND 2)=2 then ' Product Attributes
		if NOT hascheckedectfilters then
			sSQL="SELECT COUNT("&IIfVr(sqlserver,"DISTINCT products.pID","*")&") as tcount,scID,"&getlangid("scName",131072)&",scGroup,scOrder,scgTitle FROM ((searchcriteria INNER JOIN searchcriteriagroup ON searchcriteria.scGroup=searchcriteriagroup.scgID) " & _
				"INNER JOIN multisearchcriteria ON multisearchcriteria.mSCscID=searchcriteria.scID) "&IIfVr(displayzeroattributes,"LEFT","INNER")&" JOIN "&IIfVs(sectionids<>"","(")&"products"&IIfVs(sectionids<>""," LEFT JOIN multisections ON products.pId=multisections.pId)")&" ON multisearchcriteria.mSCpID=products.pID " & _
				IIfVr(displayzeroattributes,"AND","WHERE")&" pDisplay<>0" & IIfVs(ectsiteid<>"", " AND pSiteID=" & ectsiteid)
			if manid<>"0" AND is_numeric(manid) then sSQL=sSQL & " AND pManufacturer=" & manid
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
		'print "<div>" & scSQL & "<br><br>" & sSQL & "</div>"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then ectfiltercache=rs.getrows
			rs.close
		end if
		hascheckedectfilters=TRUE
		currgroupid=-1
		idtogroup="" : grouplist="" : addcomma="" : hasids="|"
		outtxt=""
		lastgroupid=0
		showallgrplist=""
		filternumber=0
		morebuttondisplayed=FALSE
		if isarray(ectfiltercache) then
			for cacheindex=0 to UBOUND(ectfiltercache,2)
				idtogroup=idtogroup&addcomma&ectfiltercache(1,cacheindex)&":"&ectfiltercache(3,cacheindex)
				idisset=isidset(ectfiltercache(1,cacheindex))
				if ectfiltercache(3,cacheindex)<>currgroupid then
					if morebuttondisplayed then
						outtxt=outtxt&"</div>"
						morebuttondisplayed=FALSE
					end if
					filternumber=0
					currgroupid=ectfiltercache(3,cacheindex)
					grouplist=grouplist&addcomma&currgroupid
					if outtxt<>"" then print outtxt & "</div>" & vbCrLf
					outtxt=""
					print "<div id=""ectpatgrp"&currgroupid&""" class=""ectpatgrp"""
					if sidefiltermorestyle="mouseover" then print " onmouseover=""ectMOGrp(this,"&currgroupid&")"" onmouseout=""ectMOtGrp(this,"&currgroupid&")"" style=""max-height:"&sidefilterclosedheight&"px;overflow-y:hidden;position:relative"""
					print "><div class=""ectpattitle"" style=""font-weight:bold;margin-top:5px"">" & ectfiltercache(5,cacheindex) & "<span id=""showallgrp"&currgroupid&""" style=""display:none"" title="""&xxSFClea&"""> (<a href=""#"" onclick=""return ectRAm("&ectfiltercache(1,cacheindex)&",false,true)"">"&xxSFAll&"</a>)</span>" & "</div>" & vbCrLf
					if sidefiltermorestyle="mouseover" then print "<div id=""grpMore"&currgroupid&""" class=""moreattributes"" style=""width:100%;position:absolute;bottom:0;left:0;text-align:center;background-color:#A0A0B0;border:1px solid black;box-sizing:border-box;-moz-box-sizing:border-box"">"&xxSFMore&"</div>" & vbCrLf
				end if
				if filternumber=sidefiltermoreafter AND sidefiltermorestyle<>"none" AND sidefiltermorestyle<>"mouseover" then
					if UBOUND(ectfiltercache,2)-cacheindex=>3 then
						if ectfiltercache(3,cacheindex+3)=currgroupid then
							outtxt=outtxt&"<div class=""grpMoreDiv"" id=""grpMoreDiv"&currgroupid&""">" & imageorbutton(imgsidefiltermore,xxSFMobu,"sidefiltermore","hidesfmorebutton("&currgroupid&")",TRUE) & "</div><div id=""grpMoreLnk"&currgroupid&""" style=""display:none"">"
							morebuttondisplayed=TRUE
						end if
					end if
				end if
				if idisset then showallgrplist=showallgrplist&","&currgroupid
				if sidefilterstyle="multiple" then
					outtxt=outtxt&"<div class=""ectpatcb" & IIfVs(idisset,"set") & IIfVs(ectfiltercache(0,cacheindex)<=0," zeroatt") & """><label style=""cursor:pointer""><input type=""checkbox"" class=""ectpatcb ectpatcbsf"" onchange=""ectRAm("&ectfiltercache(1,cacheindex)&","&IIfVr(idisset,"false","true")&",false)"" " & IIfVs(idisset,"checked=""checked"" ") & "/> " & htmlspecials(ectfiltercache(2,cacheindex)) & IIfVs(NOT idisset AND ectfiltercache(0,cacheindex)>0,"<div class=""ectpacount"">("&ectfiltercache(0,cacheindex)&")</div>") & "</label></div>" & vbCrLf
				elseif sidefilterstyle="selectmenu" then
					outtxt=outtxt&"<option value="""&ectfiltercache(1,cacheindex)&""""&IIfVr(cstr(ectfiltercache(1,cacheindex))=scrid," selected=""selected""","")&">" & htmlspecials(ectfiltercache(2,cacheindex)) & "</option>" & vbCrLf
				else
					outtxt=outtxt&"<div class=""ectpat" & IIfVs(idisset,"set") & IIfVs(ectfiltercache(0,cacheindex)<=0," zeroatt") & """" & IIfVs(NOT idisset," onclick=""ectRAm("&ectfiltercache(1,cacheindex)&",true,true)""") & ">" & htmlspecials(ectfiltercache(2,cacheindex)) & IIfVs(NOT idisset AND ectfiltercache(0,cacheindex)>0,"<div class=""ectpacount"">("&ectfiltercache(0,cacheindex)&")</div>") & "</div>" & vbCrLf
				end if
				addcomma=","
				lastgroupid=ectfiltercache(3,cacheindex)
				filternumber=filternumber+1
			next
		else
			sidefilter=sidefilter-2
		end if
		if morebuttondisplayed then outtxt=outtxt&"</div>"
		if outtxt<>"" then print outtxt & "</div>"
	end if
	case 4
	if (sidefilter AND 4)=4 then ' Price bands
		if sprice="" then
			sSQL="SELECT MAX(" & TWSP & ") AS maxprice,MIN(" & TWSP & ") AS minprice FROM products WHERE pDisplay<>0" & IIfVs(ectsiteid<>"", " AND pSiteID=" & ectsiteid)
			if sectionids<>"" then sSQL="SELECT MAX(" & TWSP & ") AS maxprice,MIN(" & TWSP & ") AS minprice FROM (products LEFT JOIN multisections ON products.pId=multisections.pId) WHERE " & IIfVs(ectsiteid<>"", "pSiteID=" & ectsiteid & " AND ") & "pDisplay<>0 AND (products.pSection IN ("&sectionids&") OR multisections.pSection IN (" & sectionids & "))"
			rs2.open sSQL,cnn,0,1
			if NOT rs2.EOF then
				if NOT isnull(rs2("maxprice")) then
					pamaxprice=rs2("maxprice") : paminprice=rs2("minprice")
					if showtaxinclusive=2 then pamaxprice=pamaxprice+(pamaxprice*(countryTaxRate/100.0)) : paminprice=paminprice+(paminprice*(countryTaxRate/100.0))
					if pamaxprice>100 then pamaxprice=-int(-pamaxprice)
				end if
			end if
			rs2.Close
		end if
		print "<div id=""ectpatgrpPRICE"" class=""ectpatgrp"" style=""overflow:auto""><div class=""ectpattitle"" style=""font-weight:bold;margin-top:5px"">" & sidefiltertext(2) & IIfVs(sprice<>"","<span id=""showallgrpPRI"" title="""&xxSFClea&"""> (<a href=""#"" onclick=""document.getElementById('paminprice').value='';document.getElementById('pamaxprice').value='';spriceset=true;return filtergo()"">"&xxSFAll&"</a>)</span>") & IIfVs(NOT sfpricebuttoninline," / <input type=""button"" class=""ectbutton sidefilter sidefiltergo"" value="""&xxGo&""" onclick=""filtergo()"" />") & "</div>" & vbCrLf
		rowcounter=2
		currpriceband=getget("sprice") %>
		<div style="display:table;margin-top:4px;width:100%">
			<div style="width:50%;text-align:center;display:table-cell">
				<div>
					<input class="sfprice" id="paminprice" type="text" onchange="spriceset=true" value="<%=paminprice%>" placeholder="Min Price" />
				</div>
				<div>
					<img src="images/buttondec.png" onmousedown="adjustprice(false,false)" onmouseup="stopprice()" onmouseout="stopprice()" alt="" /><img src="images/buttoninc.png" onmousedown="adjustprice(false,true)" onmouseup="stopprice()" onmouseout="stopprice()" alt="" />
				</div>
			</div>
			<div style="width:50%;text-align:center;display:table-cell">
				<div>
					<input type="hidden" id="origmaxprice" value="<%=pamaxprice%>" />
					<input class="sfprice" id="pamaxprice" type="text" onchange="spriceset=true" value="<%=pamaxprice%>" placeholder="Max Price" />
				</div>
				<div>
					<img src="images/buttondec.png" onmousedown="adjustprice(true,false)" onmouseup="stopprice()" onmouseout="stopprice()" alt="" /><img src="images/buttoninc.png" onmousedown="adjustprice(true,true)" onmouseup="stopprice()" onmouseout="stopprice()" alt="" />
				</div>
			</div>
<%		if sfpricebuttoninline then %>
			<div style="display:table-cell;vertical-align:middle;padding:2px">
				<input type="button" class="ectbutton sidefilter sidefiltergo" value="<%=xxGo%>" onclick="filtergo()" />
			</div>
<%		end if %>
		</div>
<%		print "</div>"
	end if
	case 8
	if (sidefilter AND 8)=8 then ' Sort Order
		print "<div class=""ectpatgrp"">"
		if sidefiltertext(3)<>"" then print "<div class=""ectpattitle"" style=""font-weight:bold;margin-top:5px"">" & sidefiltertext(3) & "</div>" & vbCrLf
		print "<div>"
		%><select class="sidefilter" size="1" onchange="changesflocation('sortby',this)">
		<option value="0"><%=xxPlsSel%></option>
<%		call pasortline(1, IIfVr(sortoption1<>"",sortoption1,"Sort Alphabetically"))
		call pasortline(11, IIfVr(sortoption11<>"",sortoption11,"Alphabetically (Desc.)"))
		call pasortline(2, IIfVr(sortoption2<>"",sortoption2,"Sort by Product ID"))
		call pasortline(12, IIfVr(sortoption12<>"",sortoption12,"Product ID (Desc.)"))
		call pasortline(14, IIfVr(sortoption14<>"",sortoption14,"Sort By SKU"))
		call pasortline(15, IIfVr(sortoption15<>"",sortoption15,"Sort By SKU (Desc.)"))
		call pasortline(3, IIfVr(sortoption3<>"",sortoption3,"Sort Price (Asc.)"))
		call pasortline(4, IIfVr(sortoption4<>"",sortoption4,"Sort Price (Desc.)"))
		call pasortline(5, IIfVr(sortoption5<>"",sortoption5,"Database Order"))
		call pasortline(6, IIfVr(sortoption6<>"",sortoption6,"Product Order"))
		call pasortline(7, IIfVr(sortoption7<>"",sortoption7,"Product Order (Desc.)"))
		call pasortline(8, IIfVr(sortoption8<>"",sortoption8,"Date Added (Asc.)"))
		call pasortline(9, IIfVr(sortoption9<>"",sortoption9,"Date Added (Desc.)"))
		call pasortline(10, IIfVr(sortoption10<>"",sortoption10,"Sort by Manufacturer"))
		call pasortline(16, IIfVr(sortoption16<>"",sortoption16,"Number of Ratings"))
		call pasortline(17, IIfVr(sortoption17<>"",sortoption17,"Average Rating"))
		call pasortline(18, IIfVr(sortoption18<>"",sortoption18,"Sales Rank"))
		call pasortline(19, IIfVr(sortoption19<>"",sortoption19,"Popularity"))
%>		</select><%
		print "</div></div>"
	end if
	case 16
	if (sidefilter AND 16)=16 then
		print "<div class=""ectpatgrp"">"
		if sidefiltertext(4)<>"" then print "<div class=""ectpattitle"" style=""font-weight:bold;margin-top:5px"">" & sidefiltertext(4) & "</div>"
		print "<div>"
		%><select class="sidefilter" size="1" onchange="changesflocation('perpage',this)">
<%		for index=1 to 5
			print "<option value="""&index&""""&IIfVr(SESSION("perpage")=index," selected=""selected""","")&">"&(prodsperpage*index)&" "&xxPerPag&"</option>"
		next
%>			  </select><%
		print "</div></div>"
	end if
	case 32
	if (sidefilter AND 32)=32 then
		print "<div class=""ectpatgrp"">"
		if sidefiltertext(5)<>"" then print "<div class=""ectpattitle"" style=""font-weight:bold;margin-top:5px"">" & sidefiltertext(5) & "</div>"
		print "<div style=""display:table;width:100%""><div style=""display:table-row""><div style=""display:table-cell""><input onkeydown=""changesffiltertext(event.keyCode,this)"" type=""text"" class=""sidefilter"" style=""width:100%;box-sizing:border-box"" id=""paectfilter"" name=""pafilter"" value="""&htmlspecials(getget("filter"))&""" /></div><div style=""display:table-cell;padding-left:5px"">" & imageorbutton(imgfilterproducts,"&raquo;","sidefilter sidefiltergo","filtergo()",TRUE) & "</div></div></div>"
		print "</div>"
	end if
	end select
next
%>
<script>
/* <![CDATA[ */
var idtogroup='<%=idtogroup%>';
var grouplist='<%=grouplist%>';
var showallgrplist='<%=showallgrplist%>';
var idarr=idtogroup.split(',');
var grouparr=grouplist.split(',');
var showallgrparr=showallgrplist.split(',');
var ECTinGrp=[];
function setgrpHght(grpId,numInGrp){
	document.getElementById('grpMore'+grpId).style.width=document.getElementById('ectpatgrp'+grpId).offsetWidth+'px';
	if(numInGrp<<%=sidefiltermoreafter%>){
		document.getElementById('grpMore'+grpId).style.display='none';
		document.getElementById('ectpatgrp'+grpId).style.maxHeight='none';
		document.getElementById('ectpatgrp'+grpId).onmouseover=null;
		document.getElementById('ectpatgrp'+grpId).onmouseout=null;
	}
}
for(var idind=0;idind<idarr.length;idind++){
	allgroups[idarr[idind].split(':')[0]]=idarr[idind].split(':')[1];
	if(ECTinGrp[idarr[idind].split(':')[1]])ECTinGrp[idarr[idind].split(':')[1]]++; else ECTinGrp[idarr[idind].split(':')[1]]=1;
}
<%	if sidefiltermorestyle="mouseover" then  %>
for(var idind=0;idind<grouparr.length;idind++){
	if(grouparr[idind]!='')setgrpHght(grouparr[idind],ECTinGrp[grouparr[idind]]);
}
<%	elseif sidefiltermorestyle<>"none" then %>
function hidesfmorebutton(tid){
	document.getElementById('grpMoreDiv'+tid).style.display='none';
	document.getElementById('grpMoreLnk'+tid).style.display='';
	return false;
}
<%	end if %>
for(var idind=0;idind<showallgrparr.length;idind++){
	if(showallgrparr[idind]!='') document.getElementById('showallgrp'+showallgrparr[idind]).style.display='';
}
/* ]]> */
</script>
<%
set rs=nothing
set rs2=nothing
set cnn=nothing
alreadycalculatedattributes=TRUE
%>