<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if sDSN="" then response.write "Database connection not set" : response.end
if request.totalbytes > 10000 then response.end
WSP=""
OWSP=""
TWSP="pPrice"
cs=csstyleprefix
if pricecheckerisincluded<>TRUE then pricecheckerisincluded=FALSE
get_wholesaleprice_sql()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs2=Server.CreateObject("ADODB.RecordSet")
Set rs3=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
if crosssellcolumns="" then if productcolumns="" then crosssellcolumns=3 else crosssellcolumns=productcolumns
if crosssellrows="" then crosssellrows=1
numberofproducts=crosssellcolumns*crosssellrows
productcolumns=crosssellcolumns
cssavenobuyorcheckout=nobuyorcheckout
cssavenoshowdiscounts=noshowdiscounts
cssavenoproductoptions=noproductoptions
cssaveproductpagelayout=productpagelayout
cssaveshowcategories=showcategories
cssavemagictoolboxproducts=magictoolboxproducts
cssaveforcedetailslink=forcedetailslink
if csnobuyorcheckout then nobuyorcheckout=TRUE
if csnoshowdiscounts then noshowdiscounts=TRUE
if csnoproductoptions then noproductoptions=TRUE
if trim(csproductpagelayout)<>"" then productpagelayout=csproductpagelayout
if productpagelayout<>"" then usecsslayout=TRUE : nomarkup=TRUE : sstrong="" : estrong=""
hasshippingdiscount=TRUE : hasproductdiscount=TRUE
if IsEmpty(forcedetailslink) then forcedetailslink=TRUE
iNumOfPages=1
showcategories=FALSE
magictoolboxproducts=""
isrootsection=TRUE
catid="0"
if IsEmpty(Count) then Count=0 else Count=(Count+crosssellcolumns)-(Count MOD crosssellcolumns)
noblankdescription=(instr(","&replace(productpagelayout," ","")&",",",description,")>0 OR instr(","&replace(quickbuylayout," ","")&",",",description,")>0)
if is_numeric(request("sortby")) then SESSION("sortby")=int(request("sortby"))
if SESSION("sortby")<>"" then dosortby=SESSION("sortby")
if orsortby<>"" then dosortby=orsortby
if dosortby=13 AND NOT sqlserver then randomize : dosortby=int(12 * rnd)
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
elseif dosortby=13 then
	sSortBy=IIfVr(mysqlserver," ORDER BY RAND()"," ORDER BY NEWID()")
else
	sSortBy=" ORDER BY "&getlangid("pName",1)&IIfVr(dosortby=11," DESC,products.pId",",products.pId")
end if
if NOT (prodlist<>"") then prodlist=""
if getpost("mode") <> "checkout" AND getpost("mode") <> "add" AND getpost("mode") <> "go" AND getpost("mode") <> "paypalexpress1" AND getpost("mode") <> "authorize" then
	cnn.open sDSN
	thesessionid=getsessionid()
	alreadygotadmin=getadminsettings()
	cssaveprodfilter=prodfilter
	prodfilter=0
	addcomma=""
	if crosssellnotsection<>"" then addcomma=","
	if SESSION("clientID")<>"" AND SESSION("clientLoginLevel")<>"" then minloglevel=SESSION("clientLoginLevel") else minloglevel=0
	rs.open "SELECT "&IIfVr(mysqlserver<>true,"TOP 500 ","")&"sectionID FROM sections WHERE sectionDisabled>"&minloglevel,cnn,0,1
	do while NOT rs.EOF
		crosssellnotsection=crosssellnotsection & addcomma & rs("sectionID")
		addcomma=","
		rs.MoveNext
	loop
	rs.close
	crosssellactionarr=split(crosssellaction, ",")
	for csindex=0 to UBOUND(crosssellactionarr)
		crosssellaction=trim(crosssellactionarr(csindex))
		addcomma="" : relatedlist=""
		if crosssellaction="alsobought" then ' Those who bought what's in your cart also bought.
			if csalsoboughttitle="" then crossselltitle="Customers who bought these products also bought." else crossselltitle=csalsoboughttitle
			if prodlist="" then
				addcomma=""
				sSQL="SELECT cartProdID FROM cart WHERE cartCompleted=0 AND " & getsessionsql()
				rs.open sSQL, cnn, 0, 1
					do while NOT rs.EOF
						prodlist=prodlist & addcomma & "'" & escape_string(rs("cartProdID")) & "'"
						addcomma=","
						rs.MoveNext
					loop
				rs.close
			end if
			addcomma="" : sessionlist="" : thecount=0 : alldone=FALSE
			if prodlist<>"" then
				sSQL="SELECT cartOrderID FROM cart WHERE cartOrderID<>0 AND cartProdID IN ("&prodlist&") AND cartSessionID<>'"&replace(thesessionid,"'","")&"' ORDER BY cartOrderID DESC"
				rs.open sSQL, cnn, 0, 1
				do while NOT rs.EOF AND NOT alldone
					sSQL="SELECT cartProdID FROM cart WHERE cartProdID NOT IN ("&prodlist&IIfVs(relatedlist<>"",","&relatedlist)&") AND cartOrderID=" & rs("cartOrderID")
					rs2.open sSQL, cnn, 0, 1
					do while NOT rs2.EOF
						relatedlist=relatedlist & addcomma & "'" & escape_string(rs2("cartProdID")) & "'"
						addcomma=","
						thecount=thecount+1
						if thecount>=numberofproducts then alldone=TRUE : exit do
						rs2.movenext
					loop
					rs2.close
				rs.MoveNext
				loop
				rs.close
			end if
		elseif crosssellaction="recommended" then ' Top x recommended products (Needs v5.1)
			if csrecommendedtitle="" then crossselltitle="These products are our current recommendations for you." else crossselltitle=csrecommendedtitle
			if prodlist="" then
				addcomma=""
				sSQL="SELECT cartProdID FROM cart WHERE cartCompleted=0 AND " & getsessionsql()
				rs.open sSQL, cnn, 0, 1
					do while NOT rs.EOF
						prodlist=prodlist & addcomma & "'" & escape_string(rs("cartProdID")) & "'"
						addcomma=","
						rs.MoveNext
					loop
				rs.close
			end if
			sSQL="SELECT pID FROM products WHERE " & IIfVs(ectsiteid<>"", "pSiteID=" & ectsiteid & " AND ") & "pDisplay<>0 AND pRecommend<>0"
			if prodlist<>"" then sSQL=sSQL & " AND pID NOT IN (" & prodlist & ")"
			if crosssellnotsection<>"" then sSQL=sSQL & " AND NOT (pSection IN (" & crosssellnotsection & "))"
			rs.open sSQL, cnn, 0, 1
				addcomma="" : relatedlist=""
				do while NOT rs.EOF
					relatedlist=relatedlist & addcomma & "'" & escape_string(rs("pID")) & "'"
					addcomma=","
					rs.MoveNext
				loop
			rs.close
		elseif crosssellaction="related" then ' Products recommended with this product (Would need v5.1)
			if csrelatedtitle="" then crossselltitle="These products are recommended with items in your cart." else crossselltitle=csrelatedtitle
			if prodlist="" then
				addcomma=""
				sSQL="SELECT cartProdID FROM cart WHERE cartCompleted=0 AND " & getsessionsql()
				rs.open sSQL, cnn, 0, 1
					do while NOT rs.EOF
						prodlist=prodlist & addcomma & "'" & escape_string(rs("cartProdID")) & "'"
						addcomma=","
						rs.MoveNext
					loop
				rs.close
			end if
			if prodlist<>"" then
				sSQL="SELECT rpRelProdID FROM relatedprods WHERE rpProdID IN ("&prodlist&") AND rpRelProdID NOT IN ("&prodlist&")"
				if relatedproductsbothways=TRUE then sSQL=sSQL & " UNION SELECT rpProdID FROM relatedprods WHERE rpRelProdID IN ("&prodlist&") AND rpProdID NOT IN ("&prodlist&")"
				rs.open sSQL, cnn, 0, 1
					addcomma="" : relatedlist=""
					do while NOT rs.EOF
						relatedlist=relatedlist & addcomma & "'" & escape_string(rs("rpRelProdID")) & "'"
						addcomma=","
						rs.MoveNext
					loop
				rs.close
			end if
		elseif crosssellaction="bestsellers" then ' Top X best sellers
			if csbestsellerstitle="" then crossselltitle="These are our current best sellers." else crossselltitle=csbestsellerstitle
			sSQL="SELECT "&IIfVr(mysqlserver<>TRUE,"TOP "&numberofproducts,"")&" cartProdID,COUNT(cartProdID) AS pidcount FROM cart INNER JOIN products ON cart.cartProdID=products.pID WHERE cartCompleted<>0 AND " & IIfVs(ectsiteid<>"", "pSiteID=" & ectsiteid & " AND ") & "pDisplay<>0 "&IIfVr(crosssellsection<>"", " AND pSection IN ("&crosssellsection&")", "")&IIfVr(crosssellnotsection<>"", " AND pSection NOT IN ("&crosssellnotsection&")", "")
			if bestsellerlimit<>"" then sSQL=sSQL & " AND cartDateAdded>" & vsusdate(Date()-bestsellerlimit)
			sSQL=sSQL & " GROUP BY cartProdID ORDER BY "&IIfVr(mysqlserver=TRUE,"pidcount","COUNT(cartProdID)")&" DESC"&IIfVs(mysqlserver=TRUE," LIMIT 0,"&numberofproducts)
			relatedlist="" : thecount=0
			rs.open sSQL, cnn, 0, 1
				do while NOT rs.EOF AND thecount<numberofproducts
					relatedlist=relatedlist & addcomma & "'" & escape_string(rs("cartProdID")) & "'"
					addcomma=","
					thecount=thecount+1
					rs.MoveNext
				loop
			rs.close
		else
			if crosssellaction<>"" then print "<p>Unrecognized crosssell action " & crosssellaction & "</p>"
		end if
		if relatedlist<>"" then
			csssaveprodlist=prodlist
			prodlist=relatedlist
			sSQL="SELECT "&IIfVs(mysqlserver<>TRUE AND crosssellsectionmax<>"","TOP "&crosssellsectionmax&" ")&"pId,pSKU,"&getlangid("pName",1)&","&WSP&"pPrice,pListPrice,pSection,pSell,pStockByOpts,pStaticPage,pStaticURL,pInStock,pExemptions,pTax,pTotRating,pNumRatings,pBackOrder,pCustomCSS,pCustom1,pCustom2,pCustom3,pDateAdded,pMinQuant,"&IIfVs(instr(productpagelayout&quickbuylayout,"manufacturer")>0,getlangid("scName",131072)&",")&IIfVs(NOT noblankdescription,"'' AS ")&getlangid("pDescription",2)&","&getlangid("pLongDescription",4)&" FROM products "&IIfVr(instr(productpagelayout&quickbuylayout,"manufacturer")>0,"LEFT OUTER JOIN searchcriteria on products.pManufacturer=searchcriteria.scID ","")&"WHERE " & IIfVs(ectsiteid<>"", "pSiteID=" & ectsiteid & " AND ") & "pDisplay<>0 AND pId IN (" & relatedlist & ")"
			if crosssellnotsection<>"" AND crosssellaction="related" then sSQL=sSQL & " AND NOT (pSection IN (" & crosssellnotsection & "))"
			if useStockManagement AND noshowoutofstock then sSQL=sSQL & " AND (pInStock>pMinQuant OR pStockByOpts<>0)"
			sSQL=sSQL & sSortBy & IIfVs(mysqlserver AND crosssellsectionmax<>""," LIMIT 0,"&crosssellsectionmax)
			rs.CursorLocation=3 ' adUseClient
			rs.CacheSize=numberofproducts
			rs.open sSQL, cnn
			if NOT rs.EOF then
				print "<p class=""cstitle"">"&crossselltitle&"</p>"
				rs.MoveFirst
				rs.PageSize=100
				rs.AbsolutePage=1
				saveAdminProdsPerPage=adminProdsPerPage
				adminProdsPerPage=rs.RecordCount
%>
<!--#include file="incproductbody2.asp"-->
<%				adminProdsPerPage=saveAdminProdsPerPage
			end if
			rs.close
			prodlist=csssaveprodlist
		end if
	next
	prodfilter=cssaveprodfilter
	cnn.Close
end if
forcedetailslink=cssaveforcedetailslink
magictoolboxproducts=cssavemagictoolboxproducts
showcategories=cssaveshowcategories
productpagelayout=cssaveproductpagelayout
noproductoptions=cssavenoproductoptions
nobuyorcheckout=cssavenobuyorcheckout
noshowdiscounts=cssavenoshowdiscounts
set rs=nothing
set rs2=nothing
set rs3=nothing
set cnn=nothing
%>