<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if sDSN="" then response.write "Database connection not set" : response.end
currFormat1="" : currFormat2="" : currFormat3=""
prodoptions=""
extraimages=0
hasmultipurchase=FALSE
hascustomlayout=FALSE
saveLCID=Session.LCID
savequickbuylayout=quickbuylayout
saveusecsslayout=usecsslayout
if productpagelayout<>"" then usecsslayout=TRUE : nomarkup=TRUE : sstrong="" : estrong=""
if productpagelayout="" OR NOT usecsslayout then productpagelayout="productid,sku,productimage,productname,discounts,reviewstars" & IIfVs(showinstock,",instock") & ",description" & IIfVs(NOT noproductoptions,",options")&",listprice,price,currency" & IIfVs(NOT nobuyorcheckout,",addtocart") else hascustomlayout=TRUE
if quickbuylayout="" then quickbuylayout="productimage,productid,productname,description,options,price,detaillink" & IIfVs(NOT nobuyorcheckout,",addtocart")
if instr(1,productpagelayout,"quickbuy",1)=0 then quickbuylayout=""
if instr(1,productpagelayout&quickbuylayout,"detaillink",1)>0 then forcedetailslink=TRUE
if instr(1,productpagelayout&quickbuylayout,"instock",1)>0 then showinstock=TRUE
customlayoutarray=split(productpagelayout,",")
quickbuylayoutarray=split(quickbuylayout,",")
localcount=0 : savejs="" : analyticsoutput=""
if currencyseparator="" then currencyseparator=" "
if SESSION("clientID")="" OR NOT enablewishlists then wishlistonproducts=FALSE
call productdisplayscript(NOT noproductoptions OR instr(quickbuylayout,"options")>0,FALSE)
if productcolumns="" then productcolumns=1
if IsEmpty(showcategories) OR showcategories=TRUE then
	if NOT (nobuyorcheckout OR nocheckoutbutton) then print "<div class=""catnavandcheckout catnavproducts"">"
	print "<div class=""catnavigation catnavproducts"">" & tslist & "</div>" & vbCrLf
	if NOT (nobuyorcheckout OR nocheckoutbutton) then print "<div class=""catnavcheckout"">" & imageorbutton(imgcheckoutbutton,xxCOTxt,"checkoutbutton","cart"&extension, FALSE) & "</div></div>" & vbCrLf
end if
if isproductspage then call dofilterresults(productcolumns)
if NOT usecsslayout then print "<table class=""" & cs & "products"" width=""100%"" border=""0"" cellspacing=""3"" cellpadding=""3"">"
saveprodlist=""
if globaldiscounttext<>"" then
	if NOT usecsslayout then print "<tr><td align=""left"" class=""allproddiscounts"" colspan=""" & productcolumns & """>"
	print "<div class=""discountsapply allproddiscounts"">" & xxDsProd & "</div><div class=""proddiscounts allproddiscounts"">" & globaldiscounttext & "</div>"
	if NOT usecsslayout then print "</td></tr>"
end if
	if iNumOfPages > 1 AND pagebarattop=1 then
		if usecsslayout then print "<div class=""pagenumbers"">" & vbCrLf else print "<tr><td colspan=""" & productcolumns & """ align=""center"" class=""pagenumbers""><p class=""pagenumbers"">"
		print writepagebar(CurPage,iNumOfPages,xxPrev,xxNext,pblink,nofirstpg)
		if usecsslayout then print "</div>" & vbCrLf else print "</p></td></tr>"
	end if
	if usecsslayout then print "<div class=""" & cs & "products"">"
	if rs.EOF then
		print IIfVs(NOT usecsslayout, "<tr><td colspan=""3"">") & "<div class=""noproducts"" style=""text-align:center"">" & xxNoPrds & "</div>" & IIfVs(NOT usecsslayout, "</td></tr>")
	else
		do while NOT rs.EOF AND localcount<rs.PageSize
			saveprodlist=saveprodlist&rs("pId")&" "
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
				rs2.open "SELECT " & IIfVs(mysqlserver<>TRUE, "TOP 1") & " imageSrc FROM productimages WHERE imageType=1 AND imageProduct='" & escape_string(rs("pID")) & "' ORDER BY imageNumber" & IIfVs(mysqlserver=TRUE, " LIMIT 0,1"),cnn,0,1
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
			if (localcount MOD productcolumns=0) AND NOT usecsslayout then print "<tr>"
			if NOT usecsslayout then print "<td width=""" & Int(100 / productcolumns) & "%"" align=""center"" valign=""top"" class=""" & cs & "product"">"
			print "<div class=""" & cs & "product " & IIfVs(rs("pSchemaType")=1,"prodarticle ") & trim(rs("pCustomCSS") & " " & rs("pId")) & """>" & vbCrLf
			if perproducttaxrate=TRUE AND NOT isnull(rs("pTax")) then thetax=rs("pTax") else thetax=countryTaxRate
			call updatepricescript()
			shortdesc=trim(rs(getlangid("pDescription",2)))
			if shortdescriptionlimit<>"" then if nostripshortdescription<>TRUE then shortdesc=strip_tags2(shortdesc) : shortdesc=left(shortdesc, shortdescriptionlimit) & IIfVs(len(shortdesc)>shortdescriptionlimit AND shortdescriptionlimit<>0, "...")
			print "<form method=""post"" id=""ectform" & Count & """ action=""cart"&extension&""" onsubmit=""return formvalidator" & Count & "(this)"">"
			if NOT usecsslayout then print "<table width=""100%"" border=""0"" cellspacing=""4"" cellpadding=""4"">"
			hasformvalidator=FALSE : isbackorder=FALSE
			optionshavestock=TRUE : isinstock=TRUE
			hasmultipurchase=0
			atcmu="" : atcmuqb="" : optionshtml="" : optjs=""
			' Options Markup
			totprice=rs("pPrice")
			if isarray(prodoptions) then
				if NOT noproductoptions OR instr(quickbuylayout,"options")>0 then
					if abs(prodoptions(1,0))=4 then thestyle="" else thestyle=" width=""100%"""
					optionshtml=displayproductoptions(optdiff,thetax,FALSE,hasmultipurchase,optjs)
					if optionshtml<>"" then optionshtml="<div class="""&IIfVs(instr(quickbuylayout,"options")=0,cs)&"prodoptions""" & IIfVs(NOT usecsslayout," style=""float:left;width:98%;""") & ">" & optionshtml & "</div>"
					if NOT noproductoptions then totprice=totprice+optdiff
				end if
				if noproductoptions then hasmultipurchase=2
			end if
			call displayformvalidator()
			if optjs<>"" then
				savejs=savejs&optjs
				if mustincludepopcalendar AND NOT hasincludedpopcalendar then optionshtml=optionshtml & "<script>var ectpopcalisproducts=1;" & ectpopcalendarjs & IIfVs(storelang<>"en" AND storelang<>"","var ectpopcallang='"&storelang&"'") & vbLf & "</script><script src=""vsadmin/popcalendar.js""></script>" : hasincludedpopcalendar=TRUE
			end if
			' Add to Cart Markup
			if NOT nobuyorcheckout OR instr(1,productpagelayout&quickbuylayout,"addtocart",1)>0 then
				if rs("pID")=giftcertificateid OR rs("pID")=donationid then hasmultipurchase=2
				if NOT usecsslayout then atcmu=atcmu &"<tr><td align=""center"">"
				if useStockManagement then
					if cint(rs("pStockByOpts"))<>0 then isinstock=optionshavestock else isinstock=rs("pInStock")>rs("pMinQuant")
				else
					isinstock=cint(rs("pSell"))<>0
				end if
				if totprice=0 AND nosellzeroprice=TRUE then
					atcmu=atcmu & "&nbsp;"
					atcmuqb=atcmu
				else
					if usecsslayout then atcmu=atcmu & "<div class=""XXXECTCSPLACEHOLDERXXXaddtocart"">"
					if NOT isinstock AND NOT (useStockManagement AND hasmultipurchase=2) AND cint(rs("pBackOrder"))=0 AND notifybackinstock<>TRUE then
						atcmu=atcmu & "<input class=""ectbutton outofstock prodoutofstock"" type=""button"" value="""&xxOutStok&""" disabled=""disabled"" />"
						atcmuqb=atcmu
					else
						atcmuqb=atcmu
						isbackorder=NOT isinstock AND cint(rs("pBackOrder"))<>0
						call writehiddenvar("id", rs("pID"))
						call writehiddenvar("mode", "add")
						if wishlistonproducts then call writehiddenvar("listid", "")
						if NOT hascustomlayout AND showquantonproduct AND hasmultipurchase=0 AND (isinstock OR isbackorder) then
							atcmuqb=atcmuqb & IIfVs(NOT usecsslayout, "<table><tr><td align=""center"">") & quantitymarkup(FALSE,Count,FALSE,"XXXECTCSPLACEHOLDERXXX",FALSE) & IIfVs(NOT usecsslayout, "</td><td align=""center"">")
						end if
						if isbackorder then
							if usehardaddtocart then atcmuqb=atcmuqb & imageorsubmit(imgbackorderbutton,xxBakOrd,"buybutton backorder") else atcmuqb=atcmuqb & imageorbuttontag(imgbackorderbutton,xxBakOrd,"buybutton backorder","subformid("&Count&",'','')",TRUE)
						elseif NOT isinstock AND notifybackinstock then
							atcmuqb=atcmuqb & imageorbuttontag(imgnotifyinstock,xxNotBaS,"notifystock prodnotifystock","return notifyinstock(false,'"&replace(rs("pID"),"'","\'")&"','"&replace(rs("pID"),"'","\'")&"',"&IIfVr(cint(rs("pStockByOpts"))<>0 AND NOT optionshavestock,"-1","0")&")", TRUE)
						else
							if custombuybutton<>"" then
								atcmuqb=atcmuqb & custombuybutton
							else
								if usehardaddtocart then atcmuqb=atcmuqb & imageorsubmit(imgbuybutton,xxAddToC,"buybutton") else atcmuqb=atcmuqb & imageorbuttontag(imgbuybutton,xxAddToC,"buybutton"" id=""ectaddcart"&Count,"subformid("&Count&",'','')",TRUE)
							end if
						end if
						if wishlistonproducts then atcmuqb=atcmuqb & "<div class=""wishlistcontainer productwishlist"">" & imageorbuttontag(imgaddtolist,xxAddLis,"prodwishlist","gtid="&Count&";return displaysavelist(this,event,window)",TRUE) & "</div>"
						if showquantonproduct AND hasmultipurchase=0 AND (isinstock OR isbackorder) then atcmuqb=atcmuqb & IIfVs(NOT usecsslayout, "</td></tr></table>")
						if hasmultipurchase=2 then
							if usecsslayout then atcmu=atcmu & "<div class=""configbutton"">"
							atcmu=atcmu & imageorbuttontag(imgconfigoptions,xxConfig,"configbutton",thedetailslink, FALSE)
							if usecsslayout then atcmu=atcmu & "</div>"
						else
							atcmu=atcmuqb
						end if
					end if
					if usecsslayout then atcmu=atcmu & "</div>" : atcmuqb=atcmuqb & "</div>"
				end if
				if NOT usecsslayout then atcmu=atcmu & "</td></tr>" : atcmuqb=atcmuqb & "</td></tr>"
			end if
			call displaylayoutarray(customlayoutarray,FALSE)
			if NOT hasformvalidator then
				optjs="" : defimagejs="" : prodoptions=""
				optionshavestock=TRUE
				totprice=rs("pPrice")
				hasmultipurchase=0
				call displayformvalidator()
				savejs=savejs&optjs
				if mustincludepopcalendar AND NOT hasincludedpopcalendar then print "<script>var ectpopcalisproducts=1;" & ectpopcalendarjs & IIfVs(storelang<>"en" AND storelang<>"","var ectpopcallang='"&storelang&"'") & vbLf & "</script><script src=""vsadmin/popcalendar.js""></script>" : hasincludedpopcalendar=TRUE
			end if
			if NOT usecsslayout then print "</table>"
			print "</form></div>"
			if NOT usecsslayout then print "</td>"
			call addtogoogleanalyticsplist(rs,analyticsoutput,Count+1)
			Count=Count+1
			localcount=localcount+1
			rs.MoveNext
			if (localcount MOD productcolumns=0) AND NOT usecsslayout then print "</tr>"
		loop
		if savejs<>"" then print "<script>/* <![CDATA[ */"&savejs&"/* ]]> */</script>"
		if ectsiteid<>"" then SESSION("saveprodlist")=saveprodlist
		if (localcount MOD productcolumns<>0) AND NOT usecsslayout then
			do while localcount MOD productcolumns<>0
				print "<td class="""&cs&"noproduct"" width="""&Int(100 / productcolumns)&"%"" align=""center"">&nbsp;</td>"
				localcount=localcount+1
			loop
			print "</tr>"
		end if
	end if
	if usecsslayout then print "</div>"
	if iNumOfPages>1 AND nobottompagebar<>TRUE then
		if usecsslayout then print "<div class=""pagenumbers"">" & vbCrLf else print "<tr><td colspan=""" & productcolumns & """ align=""center"" class=""pagenumbers""><p class=""pagenumbers"">"
		print writepagebar(CurPage,iNumOfPages,xxPrev,xxNext,pblink,nofirstpg)
		if usecsslayout then print "</div>" & vbCrLf else print "</p></td></tr>"
	end if
	if NOT usecsslayout then print "</table>"
	if defimagejs<>"" then print "<script>"&defimagejs&"</script>"
	quickbuylayout=savequickbuylayout
	usecsslayout=saveusecsslayout
	if googletagid<>"" AND analyticsoutput<>"" then
		print "<script>gtag(""event"",""view_item_list"",{items:[" & analyticsoutput & "]});</script>"&vbLf
	end if
%>