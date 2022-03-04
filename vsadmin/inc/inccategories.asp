<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if sDSN="" then response.write "Database connection not set" : response.end
dim sSQL,rs,cnn,rowcounter,success,startlink,secdesc
catname="" : caturl="" : catrootsection="" : catrootsection=0
set rs=Server.CreateObject("ADODB.RecordSet")
set rs2=Server.CreateObject("ADODB.RecordSet")
set rs3=Server.CreateObject("ADODB.RecordSet")
set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
dothrow301=FALSE
hascustomcatlayout=trim(categorypagelayout)<>""
if hascustomcatlayout then customlayoutarray=split(categorypagelayout,",")
if getget("cat")<>"" then theid=trim(getget("cat")&"") else theid=""
if seocategoryurls then usecategoryname=TRUE : theid=replace(theid,detlinkspacechar," ")
if is_numeric(explicitid) then
	theid=explicitid
elseif usecategoryname AND theid<>"" then
	sSQL="SELECT sectionID FROM sections WHERE "&IIfVs(seocategoryurls,getlangid("sectionurl",2048)&"='"&escape_string(theid)&"' OR (")&getlangid("sectionName",256)&"='"&escape_string(theid)&"'"&IIfVs(seocategoryurls," AND "&getlangid("sectionurl",2048)&"='')")
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		catname=theid
		theid=rs("sectionID")
	elseif NOT is_numeric(theid) then
		theid=-1
	end if
	rs.close
end if
if (seocategoryurls OR usecategoryname) AND catname="" AND is_numeric(theid) then
	sSQL="SELECT sectionID,"&getlangid("sectionName",256)&","&getlangid("sectionurl",2048)&",rootSection FROM sections WHERE sectionID="&theid
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then dothrow301=explicitid="" : catname=rs(getlangid("sectionName",256)) : theid=rs("sectionID") : caturl=trim(rs(getlangid("sectionurl",2048))&"") : catrootsection=rs("rootSection")
	rs.close
end if
if manufacturerpageurl="" then manufacturerpageurl="manufacturers" & extension
if instr(lcase(request.servervariables("URL")), lcase(manufacturerpageurl))>0 OR getget("man")="all" then manufacturers=TRUE
HTTP_X_ORIGINAL_URL=trim(split(request.servervariables("HTTP_X_ORIGINAL_URL")&"?","?")(0))
if HTTP_X_ORIGINAL_URL="" then HTTP_X_ORIGINAL_URL=trim(split(request.servervariables("HTTP_X_REWRITE_URL")&"?","?")(0))
if ((seocategoryurls AND HTTP_X_ORIGINAL_URL="") OR dothrow301) AND seourlsthrow301 AND NOT is_numeric(explicitid) AND NOT recommendedcategories then
	newloc=getfullurl(getcategoryurl(theid,catname,caturl,catrootsection))
	addand="" : newqs=""
	for each objitem in request.querystring
		if objitem<>"cat" then newqs=newqs&addand&objitem&"="&urlencode(getget(objitem)) : addand="&"
	next
	response.status="301 Moved Permanently"
	response.addheader "Location", newloc & IIfVs(newqs<>"","?"&newqs)
	response.end
end if
if bmlbannercategories<>"" AND paypalpublisherid<>"" then call displaybmlbanner(paypalpublisherid,bmlbannercategories)
if NOT is_numeric(theid) then theid=catalogroot
if NOT is_numeric(categorycolumns) then categorycolumns=1
cellwidth=int(100/categorycolumns)
if usecsslayout then
	usecategoryformat=1
	afterimage=""
	beforedesc=""
elseif usecategoryformat=3 then
	afterimage="<br />"
	beforedesc=""
elseif usecategoryformat=2 then
	afterimage=""
	beforedesc=""
else
	usecategoryformat=1
	afterimage=""
	beforedesc="</td></tr><tr><td class=""catdesc"" colspan=""2"">"
end if
border=0
if IsEmpty(catseparator) then catseparator=IIfVr(usecsslayout,"","<br />&nbsp;")
tslist=""
thetopts=theid
topsectionids=theid
if SESSION("clientID")<>"" AND SESSION("clientLoginLevel")<>"" then minloglevel=SESSION("clientLoginLevel") else minloglevel=0
success=TRUE
if recommendedcategories then
	' Nothing
elseif manufacturers then
	tslist="<a class=""ectlink"" href="""&storehomeurl&""">"&xxHome&"</a> &raquo; " & xxManuf
	xxAlProd=""
else
	for index=0 to 10
		if cstr(thetopts)=cstr(catalogroot) then
			caturl=storehomeurl
			sSQL="SELECT sectionID,topSection,"&getlangid("sectionName",256)&",rootSection,sectionDisabled,"&getlangid("sectionurl",2048)&","&getlangid("sectionHeader",524288) & " FROM sections WHERE sectionID=" & catalogroot
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				xxHome=rs(getlangid("sectionName",256))
				if trim(rs(getlangid("sectionurl",2048))&"")<>"" then caturl=rs(getlangid("sectionurl",2048))
				if rs("sectionID")=int(theid) then
					sectionheader=rs(getlangid("sectionHeader",524288))
					if sectionheader<>"" then print "<div class=""catheader"">" & sectionheader & "</div>"
				end if
			end if
			rs.close
			if cstr(theid)=cstr(catalogroot) then tslist=xxHome&tslist else tslist="<a class=""ectlink"" href="""&caturl&""">"&xxHome&"</a>"&tslist
			exit for
		elseif index=10 then
			tslist="<strong>Loop</strong>" & tslist
		else
			sSQL="SELECT sectionID,topSection,"&getlangid("sectionName",256)&",rootSection,sectionDisabled,"&IIfVs(languageid<>1,getlangid("sectionurl",2048)&" AS ")&"sectionurl" & IIfVs(theid=thetopts,","&getlangid("sectionHeader",524288)) & " FROM sections WHERE sectionID=" & thetopts
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if rs("sectionDisabled")>minloglevel then
					success=FALSE
				elseif rs("sectionID")=int(theid) then
					if dynamicbreadcrumbs then
						tslist="<div class=""ectbreadcrumb"">&raquo; " & breadcrumbselect(rs("sectionID"),rs("topSection")) & "</div>" & tslist
					else
						tslist="<div class=""ectbreadcrumb"">&raquo; " & rs(getlangid("sectionName",256)) & "</div>" & tslist
					end if
					sectionheader=rs(getlangid("sectionHeader",524288))
					if sectionheader<>"" then print "<div class=""catheader"">" & sectionheader & "</div>"
					if explicitid="" AND trim(rs("sectionurl"))<>"" AND redirecttostatic=TRUE then
						response.status="301 Moved Permanently"
						response.addheader "Location", rs(getlangid("sectionurl",2048))
						response.end
					end if
				else
					if dynamicbreadcrumbs then
						tslist="<div class=""ectbreadcrumb"">&raquo; " & breadcrumbselect(rs("sectionID"),rs("topSection")) & "</div>" & tslist
					else
						tslist="<div class=""ectbreadcrumb"">&raquo; <a class=""ectlink"" href=""" & getcategoryurl(rs("sectionID"),rs(getlangid("sectionName",256)),rs("sectionurl"),rs("rootSection")) & """>" & rs(getlangid("sectionName",256)) & "</a></div>" & tslist
					end if
				end if
				thetopts=rs("topSection")
				topsectionids=topsectionids & "," & thetopts
			else
				tslist="Top Section Not Available" & tslist
				rs.close
				exit for
			end if
			rs.close
		end if
	next
end if
if xxAlProd<>"" then tslist=tslist & "<div class=""ectbreadcrumb"">&raquo; <a class=""ectlink"" href="""&IIfVr(seocategoryurls,"","products" & extension)&IIfVr(theid="0" OR cstr(theid)=cstr(catalogroot),IIfVs(seocategoryurls,replace(seoprodurlpattern,"%s","")),IIfVs(NOT seocategoryurls,"?cat=")&getcatid(theid,catname,seoprodurlpattern))&""">"&xxAlProd&"</a></div>"
if manufacturers=TRUE then
	showdiscounts=FALSE
	sSQL="SELECT scID AS sectionID,"&getlangid("scName",131072)&" AS sectionName,'' AS sCustomCSS,1 AS rootSection,scLogo AS sectionImage,scOrder AS sectionOrder,"&getlangid("scURL",8192)&" AS sectionurl,"&getlangid("scDescription",16384)&" AS sectionDescription FROM searchcriteria WHERE scGroup=0 ORDER BY scOrder,"&getlangid("scName",131072)
else
	sSQL="SELECT sectionID,"&IIfVs(languageid<>1,getlangid("sectionName",256)&" AS ")&"sectionName,sCustomCSS,rootSection,sectionImage,sectionOrder,"&IIfVs(languageid<>1,getlangid("sectionurl",2048)&" AS ")&"sectionurl,"&IIfVr(nocategorydescription=TRUE,"'' AS ",IIfVs(languageid<>1,getlangid("sectionDescription",512)&" AS "))&"sectionDescription FROM sections WHERE sectionID<>0 AND "&IIfVr(recommendedcategories,"sRecommend<>0","topSection=" & theid) & " AND sectionDisabled<="&minloglevel&" ORDER BY "&IIfVr(sortcategoriesalphabetically=TRUE, getlangid("sectionName",256), "sectionOrder")
end if
rs.open sSQL,cnn,0,1
if NOT success OR rs.EOF then
	success=FALSE
	mess1=xxNoCats
	tslist="<a class=""ectlink"" href="""&storehomeurl&""">"&xxHome&"</a> "
	if xxAlProd<>"" then tslist=tslist & " &raquo; <a class=""ectlink"" href=""" & IIfVr(seocategoryurls,replace(seoprodurlpattern,"%s",""),"products" & extension) & """>" & xxAlProd & "</a>"
	if usecsslayout then print "<div class=""catnavigation"">" & tslist & "</div>" else print "<td class=""catnavigation""><p class=""catnavigation""><strong>" & tslist & "</strong></p>"
	if rs.EOF AND NOT recommendedcategories then response.status="404 Not Found"
else
	success=TRUE
	mess1=""
end if
if (usecategoryformat=1 OR usecategoryformat=2) then numcolumns=2*categorycolumns else numcolumns=categorycolumns
if NOT usecsslayout then headtable="<table width=""100%"" border=""0"" cellspacing=""3"" cellpadding=""3"">" else headtable=""
if mess1<>"" AND recommendedcategories<>TRUE then
	print headtable : headtable=""
	if NOT usecsslayout then print "<tr><td align=""center""" & IIfVs(numcolumns>1, " colspan="""&numcolumns&"""") & ">"
	print "<div class=""" & IIfVr(success,"categorymessage","categorynotavailable") & """>" & IIfVs(NOT usecsslayout,"<strong>") & mess1 & IIfVs(NOT usecsslayout,"</strong>") & "</div>"
	if NOT usecsslayout then print "</td></tr>"
end if
if nowholesalediscounts=TRUE AND SESSION("clientUser")<>"" then
	if ((SESSION("clientActions") AND 8)=8) OR ((SESSION("clientActions") AND 16)=16) then noshowdiscounts=TRUE
end if
if success then
	tdt=Date()
	if noshowdiscounts<>TRUE AND recommendedcategories<>TRUE then
		if theid="0" then
			sSQL="SELECT DISTINCT "&getlangid("cpnName",1024)&" FROM coupons WHERE (cpnSitewide=1 OR cpnSitewide=2)"
		else
			sSQL="SELECT DISTINCT "&getlangid("cpnName",1024)&" FROM coupons LEFT OUTER JOIN cpnassign ON coupons.cpnID=cpnassign.cpaCpnID WHERE (((cpnSitewide=0 OR cpnSitewide=3) AND cpaType=1 AND cpaAssignment IN ('"&Replace(topsectionids,",","','")&"')) OR cpnSitewide=1 OR cpnSitewide=2)"
		end if
		sSQL=sSQL & " AND cpnNumAvail>0 AND cpnStartDate<=" & vsusdate(tdt)&" AND cpnEndDate>=" & vsusdate(tdt)&" AND cpnIsCoupon=0 AND ((cpnLoginLevel>=0 AND cpnLoginLevel<="&minloglevel&") OR (cpnLoginLevel<0 AND -1-cpnLoginLevel="&minloglevel&"))"
		rs2.open sSQL,cnn,0,1
		if NOT rs2.EOF then
			print headtable : headtable=""
			if NOT usecsslayout then print "<tr><td align=""left"" class=""allcatdiscounts""" & IIfVs(numcolumns>1, " colspan="""&numcolumns&"""") & ">"
			print "<div class=""discountsapply allcatdiscounts"">" & xxDsCat & "</div><div class=""catdiscounts allcatdiscounts"">"
			do while NOT rs2.EOF
				print rs2(getlangid("cpnName",1024)) & "<br />"
				rs2.movenext
			loop
			print "&nbsp;</div>"
			if NOT usecsslayout then print "</td></tr>"
		end if
		rs2.close
	end if
	if NOT usecsslayout AND headtable="" then print "</table>"
	if (IsEmpty(showcategories) OR showcategories=TRUE) AND recommendedcategories<>TRUE then
		if NOT (nobuyorcheckout OR nocheckoutbutton) then print "<div class=""catnavandcheckout catnavcategories"">"
		print "<div class=""catnavigation catnavcategories"">" & tslist & "</div>"
		if NOT (nobuyorcheckout OR nocheckoutbutton) then print "<div class=""catnavcheckout"">" & imageorbutton(imgcheckoutbutton,xxCOTxt,"checkoutbutton","cart"&extension, FALSE) & "</div></div>" & vbCrLf
	end if
	if usecsslayout then print "<div class=""categories"">" else print "<table width=""100%"" border=""0"" cellspacing="""&IIfVr(usecategoryformat=1 AND categorycolumns>1,0,3)&""" cellpadding="""&IIfVr(usecategoryformat=1 AND categorycolumns>1,0,3)&""">"
	tdt=Date()
	columncount=0
	do while NOT rs.EOF
		startlink="<a class=""ectlink"" href=""" & getcategoryurl(rs("sectionID"),rs("sectionName"),rs("sectionurl"),rs("rootSection")) & """>"
		sSQL="SELECT DISTINCT "&getlangid("cpnName",1024)&" FROM coupons LEFT OUTER JOIN cpnassign ON coupons.cpnID=cpnassign.cpaCpnID WHERE (cpnSitewide=0 OR cpnSitewide=3) AND cpnNumAvail>0 AND cpnStartDate<=" & vsusdate(tdt)&" AND cpnEndDate>=" & vsusdate(tdt)&" AND cpnIsCoupon=0 AND cpaType="&IIfVr(manufacturers,3,1)&" AND cpaAssignment='"&rs("sectionID")&"'" & _
			" AND ((cpnLoginLevel>=0 AND cpnLoginLevel<="&minloglevel&") OR (cpnLoginLevel<0 AND -1-cpnLoginLevel="&minloglevel&"))"
		alldiscounts=""
		if noshowdiscounts<>TRUE then
			rs2.open sSQL,cnn,0,1
			do while NOT rs2.EOF
				alldiscounts=alldiscounts & rs2(getlangid("cpnName",1024)) & "<br />"
				rs2.MoveNext
			loop
			rs2.close
		end if
		secdesc=trim(rs("sectionDescription")&"")
		noimage=trim(rs("sectionImage")&"")=""
		if usecsslayout then
			print "<div class=""" & trim("category " & rs("sCustomCSS")) & """>"
		else
			if columncount=0 then print "<tr>"
			if usecategoryformat=1 AND categorycolumns>1 then print "<td width=""" & cellwidth & "%"" valign=""top""><table width=""100%"" border=""0"" cellspacing=""3"" cellpadding=""3""><tr>"
		end if
		if hascustomcatlayout then
			for each layoutoption in customlayoutarray
				layoutoption=lcase(trim(layoutoption))
				if layoutoption="catimage" then
					if NOT noimage then print "<div class=""catimage"">" & startlink & "<img alt="""&htmlspecials(rs("sectionName")&"")&""" class=""catimage"" src="""&rs("sectionImage")&""" /></a></div>"
				elseif layoutoption="catname" then
					print "<div class=""catname"">" & startlink & rs("sectionName") & "</a></div>"
				elseif layoutoption="discounts" then
					if alldiscounts<>"" then print " <div class=""eachcatdiscountsapply eachcatdiscount"">"&xxDsApp&"</div><div class=""catdiscounts eachcatdiscount"">" & alldiscounts & "</div>"
				elseif layoutoption="description" then
					if secdesc<>"" then print "<div class=""catdesc"">" & secdesc & "</div>"
				end if
			next
		else
			if (usecategoryformat=1 OR usecategoryformat=2) AND NOT noimage then
				cellwidth=cellwidth - 5
				print IIfVr(usecsslayout,"<div","<td width=""5%"" align=""right""") & " class=""catimage"">" & startlink&"<img alt="""&replace(rs("sectionName")&"","""","")&""" class=""catimage"" src="""&rs("sectionImage")&""" /></a>" & afterimage & IIfVr(usecsslayout,"</div>","</td>")
			end if
			if NOT usecsslayout then print "<td class=""catname"" width=""" & IIfVr(usecategoryformat=1 AND categorycolumns>1,95,cellwidth) & "%""" & IIfVr((usecategoryformat=1 OR usecategoryformat=2) AND noimage," colspan='2'","") & ">"
			if (usecategoryformat=1 OR usecategoryformat=2) AND NOT noimage then cellwidth=cellwidth + 5
			if usecategoryformat<>1 AND usecategoryformat<>2 AND NOT noimage then print startlink&"<img alt="""&replace(rs("sectionName")&"","""","")&""" class=""catimage"" src="""&rs("sectionImage")&""" /></a>" & afterimage
			if nocategoryname<>TRUE then print IIfVr(usecsslayout,"<div class=""catname"">","<p class=""catname""><strong>")&startlink&rs("sectionName")&"</a>"&xxDot&IIfVs(NOT usecsslayout,"</strong>")
			if alldiscounts<>"" then print " <div class=""eachcatdiscountsapply eachcatdiscount"">"&xxDsApp&"</div><div class=""catdiscounts eachcatdiscount"">" & alldiscounts & "</div>"
			if secdesc="" then print catseparator
			if nocategoryname<>TRUE then print IIfVr(usecsslayout,"</div>","</p>")
			if secdesc<>"" then print IIfVr(usecsslayout,"<div",beforedesc & "<p") &" class=""catdesc"">" & secdesc & catseparator & IIfVr(usecsslayout,"</div>","</p>")
		end if
		print IIfVr(usecsslayout,"</div>","</td>")
		if usecategoryformat=1 AND categorycolumns>1 AND NOT usecsslayout then print "</tr></table></td>"
		columncount=columncount + 1
		if columncount=categorycolumns AND NOT usecsslayout then
			print "</tr>"
			columncount=0
		end if
		rs.movenext
	loop
	if columncount<categorycolumns AND columncount<>0 AND NOT usecsslayout then
		do while columncount<categorycolumns
			print "<td " & IIfVr(usecategoryformat=2, " colspan='2'" , "") & ">&nbsp;</td>"
			columncount=columncount + 1
		loop
		print "</tr>"
	end if
	if usecsslayout then print "</div>"
end if
if NOT usecsslayout then print "</table>"
rs.close
cnn.Close
set rs=nothing
set rs2=nothing
set rs3=nothing
set cnn=nothing
%>