<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if sDSN="" then response.write "Database connection not set" : response.end
function strip_tags(mistr)
	set toregexp=new RegExp
	toregexp.pattern="<[^>]+>"
	toregexp.ignorecase=TRUE
	toregexp.global=TRUE
	mistr=toregexp.replace(mistr&""," ")
	toregexp.pattern="\s\s+"
	strip_tags=toregexp.replace(mistr," ")
	set toregexp=nothing
end function
function mi_escape_string(str)
	mi_escape_string=trim(replace(str&"","'","''"))
	if mysqlserver=TRUE then mi_escape_string=replace(mi_escape_string,"\","\\")
end function
if trim(prodid)<>"" AND trim(explicitid)="" then explicitid=prodid
if prodid="" then prodid=replace(trim(request.querystring("prod")),detlinkspacechar," ")
if trim(explicitid)<>"" then prodid=trim(explicitid)
if catid="" then catid=trim(request.querystring("cat"))
if manid="" then manid=trim(request.querystring("man"))
if seocategoryurls then usecategoryname=TRUE : catid=replace(catid,detlinkspacechar," ")
if seodetailurls then usepnamefordetaillinks=TRUE
if trim(explicitid)<>"" then usepnamefordetaillinks=FALSE
sectionname="" : sectiondescription="" : productid="" : productname="" : productdescription="" : pagetitle="" : metadescription=""
set rs=server.createobject("ADODB.RecordSet")
set cnn=server.createobject("ADODB.Connection")
cnn.open sDSN
if incfunctionsdefined=TRUE then
	alreadygotadmin=getadminsettings()
	sntxt=getlangid("sectionName",256)
	sutxt=getlangid("sectionURL",2048)
	sdtxt=getlangid("sectionDescription",512)
	pntxt=getlangid("pName",1)
	pttxt=getlangid("pTitle",2097152)
	pmtxt=getlangid("pMetaDesc",2097152)
	scnametxt=getlangid("scName",131072)
	if usemetalongdescription=TRUE then pdtxt=getlangid("pLongDescription",4) else pdtxt=getlangid("pDescription",2)
else
	sntxt="sectionName"
	sutxt="sectionURL"
	sdtxt="sectionDescription"
	pntxt="pName"
	pttxt="pTitle"
	pmtxt="pMetaDesc"
	scnametxt="scName"
	if usemetalongdescription=TRUE then pdtxt="pLongDescription" else pdtxt="pDescription"
end if
canonicalnopage=""
if request.querystring("pg")<>"" then
	canonicalnopage="<link rel=""canonical"" href=""" & request.servervariables("URL")
	canonqs=""
	addsep="?"
	for each objitem in request.querystring
		if objitem<>"pg" then
			canonqs=canonqs & addsep & server.urlencode(strip_tags(objitem)) & "=" & server.urlencode(strip_tags(request.querystring(objitem)))
			addsep="&"
		end if
	next
	canonicalnopage=canonicalnopage & canonqs & """ />" & vbCrLf
end if
if usecategoryname AND trim(catid)<>"" then
	sSQL="SELECT sectionID FROM sections WHERE "&IIfVs(seocategoryurls,sutxt&"='"&mi_escape_string(catid)&"' OR (")&sntxt&"='"&mi_escape_string(catid)&"'"&IIfVs(seocategoryurls," AND "&sutxt&"='')")
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then catname=catid : catid=rs("sectionID")
	rs.close
end if
if usecategoryname AND trim(manid)<>"" then
	sSQL="SELECT scID FROM searchcriteria WHERE "&scnametxt&"='"&replace(manid,"'","''")&"'"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then manname=catid : manid=rs("scID")
	rs.close
end if
if prodid<>"" then
	if usepnamefordetaillinks AND trim(request.querystring("prod"))<>"" then pidfield=pntxt else pidfield="pID"
	rs.open "SELECT pID,"&pntxt&","&sntxt&","&pttxt&","&pmtxt&","&pdtxt&" FROM products INNER JOIN sections ON products.pSection=sections.sectionID WHERE "&pidfield&"='"&replace(prodid,"'","''")&"'"&IIfVs(seodetailurls," OR pStaticURL='"&mi_escape_string(prodid)&"'"),cnn,0,1
	if NOT rs.EOF then
		productid=replace(strip_tags(rs("pID")),"""","&quot;")
		productname=replace(strip_tags(rs(pntxt)),"""","&quot;")
		sectionname=replace(strip_tags(rs(sntxt)),"""","&quot;")
		pagetitle=rs(pttxt)&""
		if trim(rs(pmtxt)&"")<>"" then productdescription=replace(rs(pmtxt),"""","&quot;") else productdescription=trim(replace(replace(replace(strip_tags(rs(pdtxt)),"""","&quot;"),vbCrLf," "),vbTab," "))
	end if
	rs.close
	if is_numeric(catid) then
		rs.open "SELECT "&sntxt&" FROM sections WHERE sectionID="&catid,cnn,0,1
		if NOT rs.EOF then sectionname=replace(strip_tags(rs(sntxt)),"""","&quot;")
		rs.close
	end if
	' if instr(detailpagelayout,"socialmedia")>0 then
	if NOT nometaogimage then
		rs.open "SELECT "&IIfVs(NOT mysqlserver,"TOP 1 ")&"imageSrc FROM productimages WHERE imageProduct='" & mi_escape_string(productid) & "' AND imageType IN (0,1" & IIfVs(usegiantogimage,",2") & ") ORDER BY imageType DESC,imageNumber" & IIfVs(mysqlserver," LIMIT 0,1"),cnn,0,1
		if NOT rs.EOF then metaogimage="<meta property=""og:image"" content=""" & IIfVs(instr(rs("imageSrc"),"://")=0,storeurl) & IIfVr(left(rs("imageSrc"),1)="/",right(rs("imageSrc"),len(rs("imageSrc"))-1),rs("imageSrc")) & """ />" & vbCrLf else metaogimage=""
	end if
elseif is_numeric(manid) then
	topsection=""
	if incfunctionsdefined=TRUE then sdtxt=getlangid("scDescription",16384) : sntext=getlangid("scName",131072) else sdtxt="scDescription" : sntext="scName"
	if is_numeric(manid) then sSQL="scID="&manid else sSQL=sntext&"='"&replace(manid,"'","''")&"'"
	rs.open "SELECT "&sntext&","&sdtxt&" FROM searchcriteria WHERE "&sSQL,cnn,0,1
	if NOT rs.EOF then
	sectionname=replace(strip_tags(rs(sntext)),"""","&quot;")
	sectiondescription=replace(strip_tags(rs(sdtxt)),"""","&quot;")
	end if
	rs.close
else ' if is_numeric(catid) OR usecategoryname then
	topsection=0
	if catid="" then catid=0
	if is_numeric(catid) then sSQL="sectionID="&catid else sSQL="sectionName='"&replace(catid,"'","''")&"'"
	rs.open "SELECT "&sntxt&","&sdtxt&",topSection,"&getlangid("sTitle",2097152)&","&getlangid("sMetaDesc",2097152)&" FROM sections WHERE "&sSQL,cnn,0,1
	if NOT rs.EOF then
		sectionname=replace(strip_tags(rs(sntxt)),"""","&quot;")
		sectiondescription=replace(strip_tags(rs(sdtxt)),"""","&quot;")
		topsection=rs("topSection")&""
		pagetitle=rs(getlangid("sTitle",2097152))&""
		if trim(rs(getlangid("sMetaDesc",2097152))&"")<>"" then sectiondescription=replace(rs(getlangid("sMetaDesc",2097152)),"""","&quot;")
	end if
	rs.close
	if topsection<>0 then
		rs.open "SELECT "&sntxt&" FROM sections WHERE sectionID="&topsection,cnn,0,1
		if NOT rs.EOF then topsection=replace(strip_tags(rs(sntxt)),"""","&quot;")
		rs.close
	else
		topsection=""
	end if
end if
cnn.Close
set rs=nothing
set cnn=nothing
%>