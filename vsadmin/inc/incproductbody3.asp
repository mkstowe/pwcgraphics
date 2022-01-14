<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if sDSN="" then response.write "Database connection not set" : response.end
prodoptions="" : nooptionshtml="" : optionshtml="" : optjs=""
extraimages=0 : hasmultipurchase=0 : totprice=0
' id,name,discounts,listprice,price,priceinctax,options,quantity,currency,instock,rating,buy
if cpdcolumns="" then cpdcolumns="id,name,discounts,listprice,price,priceinctax,instock,quantity,buy"
cpdarray=split(lcase(cpdcolumns),",")
noproductoptions=TRUE
savetaxinclusive=showtaxinclusive
showtaxinclusive=0
ectbody3layouttaxinc=FALSE
hascurrency=FALSE
noupdateprice=TRUE
if currencyseparator="" then currencyseparator=" "
if SESSION("clientID")="" OR NOT enablewishlists then wishlistonproducts=FALSE
sub docallupdatepricescript()
	call updatepricescript()
	totprice=rs("pPrice")
	hasmultipurchase=0
	optionshtml=""
	optjs=""
	if isarray(prodoptions) then
		if noproductoptions then
			hasmultipurchase=2
		else
			optionshtml=displayproductoptions(sstrong & "<span class=""prodoption"">","</span>" & estrong,optdiff,thetax,FALSE,hasmultipurchase,optjs)
			totprice=totprice+optdiff
		end if
	end if
	call displayformvalidator()
	if optjs<>"" then
		print "<script>/* <![CDATA[ */"&optjs&"/* ]]> */</script>"
		if mustincludepopcalendar AND NOT hasincludedpopcalendar then print "<script>var ectpopcalisproducts=1;" & ectpopcalendarjs & IIfVs(storelang<>"en" AND storelang<>"","var ectpopcallang='"&storelang&"'") & vbLf & "</script><script src=""vsadmin/popcalendar.js""></script>" : hasincludedpopcalendar=TRUE
	end if
	if rs("pID")=giftcertificateid OR rs("pID")=donationid then hasmultipurchase=2
	updatepricecalled=TRUE
end sub
for cpdindex=0 to UBOUND(cpdarray)
	select case trim(cpdarray(cpdindex))
	case "options"
		noproductoptions=FALSE
	case "price"
		noupdateprice=FALSE
	case "priceinctax"
		showtaxinclusive=savetaxinclusive
		ectbody3layouttaxinc=TRUE
	case "currency"
		hascurrency=TRUE
	end select
next
if NOT hascurrency then currSymbol1="" : currSymbol2="" : currSymbol3=""
saveLCID=Session.LCID
call productdisplayscript(NOT noproductoptions,FALSE) %>
		<table width="100%" border="0" cellspacing="3" cellpadding="3">
<%	if IsEmpty(showcategories) OR showcategories=TRUE then %>
		  <tr>
			<td class="prodnavigation" colspan="2" align="left"><% print sstrong & "<p class=""prodnavigation"">" & tslist & "</p>" & estrong %></td>
			<td align="right">&nbsp;<% if nobuyorcheckout<>TRUE then print imageorbutton(imgcheckoutbutton,xxCOTxt,"checkoutbutton","cart"&extension, FALSE)%></td>
		  </tr>
<%	end if
if isproductspage then call dofilterresults(3)
if globaldiscounttext<>"" then %>
		  <tr>
			<td align="left" class="allproddiscounts" colspan="3">
				<div class="discountsapply allproddiscounts"><%=xxDsProd%></div><div class="proddiscounts allproddiscounts"><%
					print globaldiscounttext %></div>
			</td>
		  </tr>
<%
end if
	if iNumOfPages>1 AND pagebarattop=1 then %>
		  <tr>
			<td colspan="3" align="center" class="pagenumbers"><p class="pagenumbers"><%=writepagebar(CurPage,iNumOfPages,xxPrev,xxNext,pblink,nofirstpg) %></p></td>
		  </tr>
<%	end if
	if rs.EOF then
		print "<tr><td colspan=""3"" align=""center""><div class=""noproducts"">" & xxNoPrds & "</div></td></tr>"
	else
	print "<tr><td colspan=""3""><table class=""cobtbl cpd"" width=""98%"" border=""0"" cellspacing=""1"" cellpadding=""3"">"
	if cpdheaders<>"" then
		cpdheadarray=split(cpdheaders,",")
		print "<tr>"
		for cpdindex=0 to UBOUND(cpdheadarray)
			if cpdindex<=UBOUND(cpdarray) then classid=cpdarray(cpdindex) else classid=""
			print "<td class=""cobhl cpdhl""><div class=""cpdhl"&classid&""">"&cpdheadarray(cpdindex)&"</div></td>"
		next
		print "</tr>"
	end if
	localcount=0
	do while NOT rs.EOF AND localcount=0<rs.PageSize
		thedetailslink=getdetailsurl(rs("pId"),rs("pStaticPage"),rs(getlangid("pName",1)),trim(rs("pStaticURL")&""),IIfVs(catid<>"" AND catid<>"0" AND int(catid)<>rs("pSection") AND nocatid<>TRUE,"cat="&catid),pathtohere)
		allimages="" : alllgimages="" : plargeimage=""
		needdetaillink=trim(replace(rs(getlangid("pLongDescription",4))&"","<br />",""))<>""
		rs2.open "SELECT imageSrc FROM productimages WHERE imageType=0 AND imageProduct='" & escape_string(rs("pID")) & "' ORDER BY imageNumber",cnn,0,1
		if NOT rs2.EOF then allimages=rs2.getrows()
		rs2.close
		if magictoolboxproducts<>"" AND isarray(allimages) then
			rs2.open "SELECT imageSrc FROM productimages WHERE imageType=1 AND imageProduct='" & escape_string(rs("pID")) & "' ORDER BY imageNumber",cnn,0,1
			if NOT rs2.EOF then alllgimages=rs2.getrows() : needdetaillink=TRUE : plargeimage=alllgimages(0,0) else if thumbnailsonproducts then alllgimages=allimages : plargeimage=alllgimages(0,0)
			rs2.close
		end if
		if (NOT forcedetailslink AND NOT needdetaillink) OR detailslink<>"" then
			rs2.open "SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 1 ")&"imageSrc FROM productimages WHERE imageType=1 AND imageProduct='" & escape_string(rs("pID")) & "' ORDER BY imageNumber"&IIfVs(mysqlserver=TRUE," LIMIT 0,1"),cnn,0,1
			if NOT rs2.EOF then needdetaillink=TRUE : plargeimage=rs2("imageSrc")
			rs2.close
		end if
		startlink="":endlink=""
		if forcedetailslink OR needdetaillink then
			if detailslink<>"" then
				startlink=replace(replace(detailslink,"%largeimage%", plargeimage),"%pid%", rs("pId"))
				endlink=detailsendlink
			else
				startlink="<a class=""ectlink"" href="""&thedetailslink&""">"
				endlink="</a>"
			end if
		end if
		Session.LCID=1033
		if NOT isrootsection then
			if IsNull(rs("pSection")) then thetopts=0 else thetopts=rs("pSection")
			gotdiscsection=FALSE
			for cpnindex=0 to adminProdsPerPage-1
				if aDiscSection(0,cpnindex)=thetopts then
					gotdiscsection=TRUE
					exit for
				elseif aDiscSection(0,cpnindex)="" then
					exit for
				end if
			next
			aDiscSection(0,cpnindex)=thetopts
			if NOT gotdiscsection then
				topcpnids=thetopts
				for index=0 to 10
					if thetopts=0 then
						exit for
					else
						sSQL="SELECT topSection FROM sections WHERE sectionID=" & thetopts
						rs2.Open sSQL,cnn,0,1
						if NOT rs2.EOF then
							thetopts=rs2("topSection")
							topcpnids=topcpnids & "," & thetopts
						else
							rs2.Close
							exit for
						end if
						rs2.Close
					end if
				next
				aDiscSection(1,cpnindex)=topcpnids
			else
				topcpnids=aDiscSection(1,cpnindex)
			end if
		end if
		alldiscounts="" : noapplydiscounts=""
		hascheckedperproductdiscounts=FALSE
		call getperproductdiscounts()
		Session.LCID=saveLCID
		optionshavestock=TRUE
		print "<tr class=""cpdtr"">"
		if perproducttaxrate=TRUE AND NOT IsNull(rs("pTax")) then thetax=rs("pTax") else thetax=countryTaxRate
		updatepricecalled=FALSE
		for cpdindex=0 to UBOUND(cpdarray)
			select case trim(cpdarray(cpdindex))
			case "id" %>
			<td class="cobll cpdll"><% if NOT updatepricecalled then docallupdatepricescript() %><div class="prod3id"><%=startlink & rs("pID") & endlink %></div></td>
<%			case "sku" %>
			<td class="cobll cpdll"><div class="prod3sku"><%=startlink & rs("pSKU") & endlink %></div></td>
<%			case "manufacturer" %>
			<td class="cobll cpdll"><div class="prod3manufacturer"><%=rs(getlangid("scName",131072))%></div></td>
<%			case "name" %>
			<td class="cobll cpdll"><div class="prod3name"><%=rs(getlangid("pName",1)) %></div></td>
<%			case "description" %>
			<td class="cobll cpdll"><div class="prod3description"><%
				shortdesc=rs(getlangid("pDescription",2))
				if shortdescriptionlimit="" then print shortdesc else if nostripshortdescription<>TRUE then shortdesc=strip_tags2(shortdesc) : print left(shortdesc, shortdescriptionlimit) & IIfVr(len(shortdesc)>shortdescriptionlimit AND shortdescriptionlimit<>0, "...", "") %></div></td>
<%			case "image" %>
			<td class="cobll cpdll"><%
				if NOT updatepricecalled then docallupdatepricescript()
				if NOT isarray(allimages) then
					print "&nbsp;"
				else
					if UBOUND(allimages,2)>0 AND NOT thumbnailsonproducts then print "<table border=""0"" cellspacing=""1"" cellpadding=""1""><tr><td colspan=""3"">"
					if (magictoolboxproducts="MagicSlideshow" OR magictoolboxproducts="MagicScroll") AND UBOUND(allimages,2)>0 then
						print "<div class=""" & magictoolboxproducts & """ "&magictooloptionsproducts&">"
						for index=0 to UBOUND(allimages,2)
							largeimage=""
							if isarray(alllgimages) then if UBOUND(alllgimages,2)>=index then largeimage=alllgimages(0,index)
							print "<img itemprop=""image"" src=""" & allimages(0,index) & """ alt="""" "&IIfVs(largeimage<>"" AND magictoolboxproducts="MagicSlideshow","data-fullscreen-image="""&largeimage&""" ")&"/>"
						next
						print "</div>"
					else
						relid=magictooloptionsproducts
						if magictoolboxproducts="MagicThumb" then
							if magictooloptionsproducts="" then relid="rel=""group:g"&Count&""" " else relid=replace(magictooloptionsproducts,"rel=""","rel=""group:g"&Count&";") & " "
						end if
						print IIfVr(magictoolboxproducts<>"" AND magictoolboxproducts<>"MagicSlideshow" AND magictoolboxproducts<>"MagicScroll" AND plargeimage<>"","<a id=""mzprodimage"&Count&""" " & relid & " href="""&plargeimage&""" class=""" & magictoolboxproducts & """>",startlink)&"<img id=""prodimage"&Count&""" class="""&cs&"prod3image allprodimages"" src="""&replace(allimages(0,0),"%s","")&""" style=""border:0"" alt="""&replace(strip_tags2(rs(getlangid("pName",1))),"""","&quot;")&""" />"&IIfVr(magictoolboxproducts<>"" AND plargeimage<>"","</a>",endlink)
						if UBOUND(allimages,2)>0 AND NOT thumbnailsonproducts then print "</td></tr><tr><td align=""left"">" & imageorbutton(imgprevimg,xxPrImTx,"previmg","updateprodimage("&Count&",false)",TRUE) & "</td><td align=""center""><span class=""extraimage extraimagenum"" id=""extraimcnt"&Count&""">1</span> <span class=""extraimage"">"&xxOf&" "&extraimages&"</span></td><td align=""right"">" & imageorbutton(imgnextimg,xxNeImTx,"nextimg","updateprodimage("&Count&",true)",TRUE) & "</td></tr></table>"
					end if
					if magictoolboxproducts<>"" AND UBOUND(allimages,2)>0 AND thumbnailsonproducts then
						if magictoolboxproducts="MagicThumb" then relid=" rel=""thumb-id:mzprodimage"&Count&"""" else relid=""
						if magictoolboxproducts="MagicZoom" OR magictoolboxproducts="MagicZoomPlus" then relid=" data-zoom-id=""mzprodimage"&Count&""""
						if thumbnailstyleproducts="" then thumbnailstyleproducts="width:50px;padding:2px"
						if usecsslayout then print "<div class=""thumbnailimage productsthumbnail"">" else print "</td></tr><tr><td class=""thumbnailimage productsthumbnail"" align=""center"">"
						if magicscrollthumbnailsproducts then print "<div class=""MagicScroll"">"
						for index=0 to UBOUND(allimages,2)
							if UBOUND(alllgimages,2)>=index then print "<a href=""" & alllgimages(0,index) & """ rev=""" & allimages(0,index) & """" & relid & "><img src=""" & allimages(0,index) & """ style=""" & thumbnailstyleproducts & """ alt="""" /></a>"
						next
						if magicscrollthumbnailsproducts then print "</div>"
						if usecsslayout then print "</div>" else print "</td></tr></table>"
					end if
				end if %></td>
<%			case "discounts" %>
			<td class="cobll cpdll"><div class="prod3discounts"><%
				if alldiscounts<>"" then print alldiscounts
				if noapplydiscounts<>"" then print "<div class=""discountsnotapply"">"&xxDsNoAp&"</div>"&noapplydiscounts
				call displaydiscountexemptions() %></div></td>
<%			case "details" %>
			<td class="cobll cpdll"><div class="prod3details"><% if startlink<>"" then print startlink & xxPrDets&"</a>&nbsp;" else print "&nbsp;" %></div></td>
<%			case "options" %>
			<td class="cobll cpdll">
<%				if NOT updatepricecalled then docallupdatepricescript()
				print "<form method=""post"" id=""ectform"&Count&""" action=""cart"&extension&""" onsubmit=""return formvalidator"&Count&"(this)"">"
				call writehiddenvar("id", rs("pID"))
				call writehiddenvar("mode", "add")
				' call writehiddenvar("frompage", "")
				if wishlistonproducts then call writehiddenvar("listid", "")
				print "<input type=""hidden"" name=""quant"" id=""qnt"&Count&"x"" value="""" />"
				if isarray(prodoptions) then
					if hasmultipurchase=2 then
						print "&nbsp;"
					elseif optionshtml<>"" then
						print "<div class=""prod3options"">" & optionshtml & "</div>"
					end if
				else
					print "&nbsp;"
				end if
				print "</form>"
%>			</td>
<%			case "listprice" %>
			<td class="cobll cpdll"><div class="prod3listprice" id="listdivec<%=Count%>"><%
						if cdbl(rs("pListPrice"))<>0.0 then
							plistprice=IIfVr(showtaxinclusive=2 AND (rs("pExemptions") AND 2)<>2,rs("pListPrice")+(rs("pListPrice")*thetax/100.0), rs("pListPrice"))
							if yousavetext<>"" then yousaveprice=plistprice-IIfVr(showtaxinclusive=2 AND (rs("pExemptions") AND 2)<>2,rs("pPrice")+(rs("pPrice")*thetax/100.0),rs("pPrice")) else yousaveprice=0
							print FormatEuroCurrency(plistprice) & IIfVs(yousavetext<>"" AND yousaveprice>0,replace(yousavetext,"%s",FormatEuroCurrency(yousaveprice)))
						else
							print "&nbsp;"
						end if
%></div></td>
							
<%			case "price" %>
			<td class="cobll cpdll"><% if NOT updatepricecalled then docallupdatepricescript() %><div class="prod3price"><%
						if rs("pID")=giftcertificateid OR rs("pID")=donationid then
							print "-"
						else
							print "<span class=""price"" id=""pricediv" & Count & """>" & IIfVr(totprice=0 AND pricezeromessage<>"", pricezeromessage, FormatEuroCurrency(totprice)) & "</span>"
						end if %></div></td>
<%			case "priceinctax" %>
			<td class="cobll cpdll"><div class="prod3pricetaxinc"><%
						if rs("pID")=giftcertificateid OR rs("pID")=donationid then
							print "-"
						elseif cdbl(totprice)=0 AND pricezeromessage<>"" then
							print "<span class=""price"" id=""pricedivti" & Count & """> &nbsp; </span>"
						else
							print "<span class=""price"" id=""pricedivti" & Count & """>"
							if (rs("pExemptions") AND 2)=2 then print FormatEuroCurrency(totprice) else print FormatEuroCurrency(totprice+(totprice*thetax/100.0))
							print "</span>"
						end if %></div></td>
<%			case "currency" %>
			<td class="cobll cpdll"><%
						extracurr=""
						if currRate1<>0 AND currSymbol1<>"" then extracurr=replace(currFormat1, "%s", FormatNumber(totprice*currRate1, checkDPs(currSymbol1))) & currencyseparator
						if currRate2<>0 AND currSymbol2<>"" then extracurr=extracurr & replace(currFormat2, "%s", FormatNumber(totprice*currRate2, checkDPs(currSymbol2))) & currencyseparator
						if currRate3<>0 AND currSymbol3<>"" then extracurr=extracurr & replace(currFormat3, "%s", FormatNumber(totprice*currRate3, checkDPs(currSymbol3)))
						if totprice=0 AND pricezeromessage<>"" then extracurr=""
						if extracurr<>"" then print "<div class=""prod3currency""><span class=""extracurr"" id=""pricedivec" & Count & """>" & extracurr & "</span></div>"
						%></td>
<%			case "quantity" %>
			<td class="cobll cpdll"><div class="prod3quant" style="white-space:nowrap"><% if hasmultipurchase>0 then print "&nbsp;" else print quantitymarkup(FALSE,Count,FALSE,"",TRUE) %></div></td>
<%			case "instock" %>
			<td class="cobll cpdll"><div class="prod3instock"><% if cint(rs("pStockByOpts"))<>0 OR rs("pID")=giftcertificateid OR rs("pID")=donationid OR (rs("pInStock")>clng(stockdisplaythreshold) AND stockdisplaythreshold<>"") then print "-" else print vrmax(0,rs("pInStock")) %></div></td>
<%			case "rating" %>
			<td class="cobll cpdll"><% if rs("pNumRatings")>0 then print showproductreviews(3, "prod3rating") else print "&nbsp;" %></td>
<%			case "buy" %>
			<td class="cobll cpdll"><% if NOT updatepricecalled then docallupdatepricescript() %><div class="prod3buy"><%
	if useStockManagement then
		if cint(rs("pStockByOpts"))<>0 then isinstock=optionshavestock else isinstock=rs("pInStock")>rs("pMinQuant")
	else
		isinstock=cint(rs("pSell"))<>0
	end if
	if totprice=0 AND nosellzeroprice=TRUE then
		print "&nbsp;"
	else
		if NOT isinstock AND NOT (useStockManagement AND hasmultipurchase=2) AND cint(rs("pBackOrder"))=0 AND notifybackinstock<>TRUE then
			print "<div class=""outofstock"">" & sstrong & xxOutStok & estrong & "</div>"
		elseif hasmultipurchase=2 then
			print imageorbutton(imgconfigoptions,xxConfig,"configbutton",thedetailslink, FALSE)
		else
			isbackorder=NOT isinstock AND cint(rs("pBackOrder"))<>0
			if isbackorder then
				print imageorbutton(imgbackorderbutton,xxBakOrd,"buybutton backorder","subformid("&Count&",'','')", TRUE)
			elseif NOT isinstock AND notifybackinstock then
				print imageorbutton(imgnotifyinstock,xxNotBaS,"notifystock prodnotifystock","return notifyinstock(false,'"&replace(rs("pID"),"'","\'")&"','"&replace(rs("pID"),"'","\'")&"',"&IIfVr(cint(rs("pStockByOpts"))<>0 AND NOT optionshavestock,"-1","0")&")", TRUE)
			else
				print imageorbutton(imgbuybutton,xxAddToC,"buybutton","subformid("&Count&",'','')", TRUE)
			end if
			if wishlistonproducts then print "<div class=""wishlistcontainer productwishlist"">" & imageorlink(imgaddtolist,xxAddLis,"","gtid="&Count&";return displaysavelist(this,event,window)",TRUE) & "</div>"
		end if
	end if %></div></td>
<%			end select
		next
		if noproductoptions then
			nooptionshtml=nooptionshtml & "<form method=""post"" id=""ectform"&Count&""" action=""cart"&extension&""" onsubmit=""return formvalidator"&Count&"(this)"">" & vbCrLf
			nooptionshtml=nooptionshtml & "<input type=""hidden"" name=""quant"" id=""qnt"&Count&"x"" />"
			nooptionshtml=nooptionshtml & "<input type=""hidden"" name=""id"" value="""& rs("pID")&""" />"
			nooptionshtml=nooptionshtml & "<input type=""hidden"" name=""mode"" value=""add"" />"
			if wishlistonproducts then nooptionshtml=nooptionshtml & "<input type=""hidden"" name=""listid"" value="""" />"
			nooptionshtml=nooptionshtml & "</form>" & vbCrLf
		end if
		print "</tr>"
		Count=Count+1
		localcount=localcount+1
		rs.MoveNext
	loop
	print "</table>" & nooptionshtml & "</td></tr>"
	end if
	if iNumOfPages>1 AND nobottompagebar<>TRUE then %>
		  <tr><td colspan="3" align="center" class="pagenumbers"><p class="pagenumbers"><%=writepagebar(CurPage,iNumOfPages,xxPrev,xxNext,pblink,nofirstpg) %></p></td></tr>
<%	end if %>
		</table>
<%	if defimagejs<>"" then print "<script>" & defimagejs & "</script>" %>