<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if sDSN="" then response.write "Database connection not set" : response.end
prodoptions=""
extraimages=0
saveLCID=Session.LCID
if currencyseparator="" then currencyseparator=" "
if SESSION("clientID")="" OR NOT enablewishlists then wishlistonproducts=FALSE
call productdisplayscript(NOT noproductoptions,FALSE)
if IsEmpty(showcategories) OR showcategories=TRUE then
	if NOT (nobuyorcheckout OR nocheckoutbutton) then print "<div class=""catnavandcheckout catnavproducts"">"
	print "<div class=""catnavigation catnavproducts"">" & tslist & "</div>" & vbCrLf
	if NOT (nobuyorcheckout OR nocheckoutbutton) then print "<div class=""catnavcheckout"">" & imageorbutton(imgcheckoutbutton,xxCOTxt,"checkoutbutton","cart"&extension, FALSE) & "</div></div>" & vbCrLf
end if
if isproductspage then call dofilterresults(3)
%>
			<table class="products" width="100%" border="0" cellspacing="3" cellpadding="3">
<%
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
	localcount=0
	savejs=""
	do while NOT rs.EOF AND localcount<rs.PageSize
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
			if isnull(rs("pSection")) then thetopts=0 else thetopts=rs("pSection")
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
%>
              <tr> 
                <td width="26%" rowspan="3" align="center" class="prodimage allprodimages"><%
		if perproducttaxrate=TRUE AND NOT isnull(rs("pTax")) then thetax=rs("pTax") else thetax=countryTaxRate
		call updatepricescript()
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
				print IIfVr(magictoolboxproducts<>"" AND magictoolboxproducts<>"MagicSlideshow" AND magictoolboxproducts<>"MagicScroll" AND plargeimage<>"","<a id=""mzprodimage"&Count&""" " & relid & " href="""&plargeimage&""" class=""" & magictoolboxproducts & """>",startlink)&"<img id=""prodimage"&Count&""" class="""&cs&"prodimage allprodimages"" src="""&replace(allimages(0,0),"%s","")&""" style=""border:0"" alt="""&replace(strip_tags2(rs(getlangid("pName",1))),"""","&quot;")&""" />"&IIfVr(magictoolboxproducts<>"" AND plargeimage<>"","</a>",endlink)
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
		end if %></td><td width="59%">
<%					if showproductid=TRUE then print "<div class=""prodid"">" & xxPrId & ": " & rs("pID") & "</div>"
					if xxManLab<>"" then
						if NOT isnull(rs(getlangid("scName",131072))) then print "<div class=""prodmanufacturer"">" & xxManLab & ": " & rs(getlangid("scName",131072)) & "</div>"
					end if
					if showproductsku<>"" AND trim(rs("pSKU")&"")<>"" then print "<div class=""prodsku"">" & showproductsku & ": " & rs("pSKU") & "</div>"
					print "<div class=""prodname"">"&startlink&rs(getlangid("pName",1))&endlink&xxDot
					if alldiscounts<>"" then print " <span class=""discountsapply"">"&xxDsApp&"</span></div>" & "<div class=""proddiscounts"">" & alldiscounts & "</div>" else print "</div>"
					if noapplydiscounts<>"" then print " <span class=""discountsnotapply"">"&xxDsNoAp&"</span></div><div class=""prodnoapplydiscounts"">" & noapplydiscounts & "</div>" else print "</div>"
					call displaydiscountexemptions()
					if useStockManagement AND showinstock=TRUE AND (rs("pInStock")<=clng(stockdisplaythreshold) OR stockdisplaythreshold="") then if cint(rs("pStockByOpts"))=0 then print "<div class=""prodinstock""><strong>" & xxInStoc & ":</strong> " & vrmax(0,rs("pInStock")) & "</div>"
					if ratingsonproductspage=TRUE AND rs("pNumRatings")>0 then print showproductreviews(1, "prodrating") %>
                </td>
				<td width="15%" align="right" valign="top"><%
            		if startlink<>"" then
                		print "<p>" & startlink & xxPrDets & "</a>&nbsp;</p>"
                	else
                		print "&nbsp;"
                	end if
              %></td>
			  </tr>
			  <tr>
				<td colspan="2" class="proddescription"><form method="post" id="ectform<%=Count%>" action="cart<%=extension%>" onsubmit="return formvalidator<%=Count%>(this)"><%
	call writehiddenvar("id", rs("pID"))
	call writehiddenvar("mode", "add")
	' call writehiddenvar("frompage", "")
	if wishlistonproducts then call writehiddenvar("listid", "")
	print "<input type=""hidden"" name=""quant"" id=""qnt"&Count&"x"" value="""" />"
	print "<div class=""proddescription"">"
	shortdesc=rs(getlangid("pDescription",2))
	if shortdescriptionlimit="" then print shortdesc else if nostripshortdescription<>TRUE then shortdesc=strip_tags2(shortdesc) : print left(shortdesc, shortdescriptionlimit) & IIfVr(len(shortdesc)>shortdescriptionlimit AND shortdescriptionlimit<>0, "...", "")
	print "</div>"
	optionshavestock=TRUE
	totprice=rs("pPrice")
	hasmultipurchase=0
	optjs=""
	if isarray(prodoptions) then
		if noproductoptions then
			hasmultipurchase=2
		else
			optionshtml=displayproductoptions(optdiff,thetax,FALSE,hasmultipurchase,optjs)
			if optionshtml<>"" then print "<div class=""prodoptions"">" & optionshtml & "</div>"
			totprice=totprice + optdiff
		end if
	end if
	call displayformvalidator()
	savejs=savejs&optjs
	if mustincludepopcalendar AND NOT hasincludedpopcalendar then print "<script>var ectpopcalisproducts=1;" & ectpopcalendarjs & IIfVs(storelang<>"en" AND storelang<>"","var ectpopcallang='"&storelang&"'") & vbLf & "</script><script src=""vsadmin/popcalendar.js""></script>" : hasincludedpopcalendar=TRUE
%>		</form></td>
			  </tr>
			  <tr>
				<td width="59%" align="center"><%
					if noprice=TRUE OR rs("pID")=giftcertificateid OR rs("pID")=donationid  then
						print "&nbsp;"
					else
						if cdbl(rs("pListPrice"))<>0.0 then
							plistprice=IIfVr(showtaxinclusive=2 AND (rs("pExemptions") AND 2)<>2,rs("pListPrice")+(rs("pListPrice")*thetax/100.0), rs("pListPrice"))
							if yousavetext<>"" then yousaveprice=plistprice-IIfVr(showtaxinclusive=2 AND (rs("pExemptions") AND 2)<>2,rs("pPrice")+(rs("pPrice")*thetax/100.0),rs("pPrice")) else yousaveprice=0
							print "<div class=""listprice"" id=""listdivec" & Count & """>" & Replace(xxListPrice, "%s", FormatEuroCurrency(plistprice)) & IIfVs(yousavetext<>"" AND yousaveprice>0,replace(yousavetext,"%s",FormatEuroCurrency(yousaveprice))) & "</div>"
						end if
						print "<div class=""prodprice""><strong>" & xxPrice&IIfVs(xxPrice<>"",":") & "</strong> <span class=""price"" id=""pricediv" & Count & """>" & IIfVr(totprice=0 AND pricezeromessage<>"",pricezeromessage,FormatEuroCurrency(IIfVr(showtaxinclusive=2 AND (rs("pExemptions") AND 2)<>2, totprice+(totprice*thetax/100.0), totprice))) & "</span> "
						if showtaxinclusive=1 AND (rs("pExemptions") AND 2)<>2 then print "<span class=""inctax"" id=""taxmsg" & Count & """" & IIfVs(totprice=0, " style=""display:none""") & ">" & Replace(ssIncTax,"%s", "<span id=""pricedivti" & Count & """>" & IIfVr(totprice=0, "-", FormatEuroCurrency(totprice+(totprice*thetax/100.0))) & "</span> ") & "</span>"
						print "</div>"
						extracurr=""
						if currRate1<>0 AND currSymbol1<>"" then extracurr=replace(currFormat1, "%s", FormatNumber(totprice*currRate1, checkDPs(currSymbol1))) & currencyseparator
						if currRate2<>0 AND currSymbol2<>"" then extracurr=extracurr & replace(currFormat2, "%s", FormatNumber(totprice*currRate2, checkDPs(currSymbol2))) & currencyseparator
						if currRate3<>0 AND currSymbol3<>"" then extracurr=extracurr & replace(currFormat3, "%s", FormatNumber(totprice*currRate3, checkDPs(currSymbol3)))
						if extracurr<>"" then print "<div class=""prodcurrency""><span class=""extracurr"" id=""pricedivec" & Count & """>" & IIfVr(totprice=0, "", extracurr) & "</span></div>"
					end if %>
                </td>
			    <td align="right" valign="bottom" style="white-space:nowrap;"><%
		if nobuyorcheckout=TRUE then
			print "&nbsp;"
		else
			if rs("pID")=giftcertificateid OR rs("pID")=donationid then hasmultipurchase=2
			if useStockManagement then
				if cint(rs("pStockByOpts"))<>0 then isinstock=optionshavestock else isinstock=rs("pInStock")>rs("pMinQuant")
			else
				isinstock=cint(rs("pSell")) <> 0
			end if
			if totprice=0 AND nosellzeroprice=TRUE then
				print "&nbsp;"
			else
				if NOT isinstock AND NOT (useStockManagement AND hasmultipurchase=2) AND cint(rs("pBackOrder"))=0 AND notifybackinstock<>TRUE then
					print "<div class=""outofstock"">" & xxOutStok & "</div>"
				elseif hasmultipurchase=2 then
					print imageorbutton(imgconfigoptions,xxConfig,"configbutton",thedetailslink,FALSE)
				else
					isbackorder=NOT isinstock AND cint(rs("pBackOrder"))<>0
					if showquantonproduct AND hasmultipurchase=0 AND (isinstock OR isbackorder) then print "<table><tr><td align=""center"">" & quantitymarkup(FALSE,Count,FALSE,"",TRUE) & "</td><td align=""center"">"
					if isbackorder then
						print imageorbutton(imgbackorderbutton,xxBakOrd,"buybutton backorder","subformid("&Count&",'','')", TRUE)
					elseif NOT isinstock AND notifybackinstock then
						print imageorbutton(imgnotifyinstock,xxNotBaS,"notifystock prodnotifystock","return notifyinstock(false,'"&replace(rs("pID"),"'","\'")&"','"&replace(rs("pID"),"'","\'")&"',"&IIfVr(cint(rs("pStockByOpts"))<>0 AND NOT optionshavestock,"-1","0")&")", TRUE)
					else
						print imageorbutton(imgbuybutton,xxAddToC,"buybutton","subformid("&Count&",'','')", TRUE)
					end if
					if wishlistonproducts then print "<div class=""wishlistcontainer productwishlist"">" & imageorlink(imgaddtolist,xxAddLis,"","gtid="&Count&";return displaysavelist(this,event,window)",TRUE) & "</div>"
					if showquantonproduct AND hasmultipurchase=0 AND (isinstock OR isbackorder) then print "</td></tr></table>"
				end if
			end if
		end if %></td>
			  </tr>
<%		Count=Count+1
		localcount=localcount+1
		rs.MoveNext
	loop
	end if
	if iNumOfPages>1 AND nobottompagebar<>TRUE then %>
		  <tr><td colspan="3" align="center" class="pagenumbers"><p class="pagenumbers"><%=writepagebar(CurPage,iNumOfPages,xxPrev,xxNext,pblink,nofirstpg) %></p></td></tr>
<%	end if %>
		</table>
<%	if savejs<>"" then print "<script>/* <![CDATA[ */"&savejs&"/* ]]> */</script>"
	if defimagejs<>"" then print "<script>"&defimagejs&"</script>" %>