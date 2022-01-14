<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if sDSN="" then response.write "Database connection not set" : response.end
Dim sSQL,rs,alldata,cnn,rowcounter,iNumOfPages,CurPage,Count,weburl,longdesc,currFormat1,currFormat2,currFormat3
skuschemaidentifier=trim("sku " & skuschemaidentifier)
if trim(explicitid)<>"" then prodid=trim(explicitid) else prodid=replace(getget("prod"),detlinkspacechar," ")
WSP=""
OWSP=""
TWSP="pPrice"
Count=0
hasmultipurchase=FALSE
isarticle=FALSE
optionshtml=""
if pricecheckerisincluded<>TRUE then pricecheckerisincluded=FALSE
hascustomlayout=FALSE
if detailpagelayout="" OR NOT usecsslayout then detailpagelayout="productimage,productid,manufacturer,sku,productname,discounts,instock,description,listprice,price,currency,options,addtocart,previousnext,emailfriend"&IIfVs(showsearchwords,",searchwords") else hascustomlayout=TRUE
customlayoutarray=split(detailpagelayout,",")
if socialmediabuttons="" then
	if hascustomlayout AND instr(detailpagelayout,"socialmedia")>0 then
		socialmediabuttons="facebook,linkedin,twitter,askaquestion"
	elseif useemailfriend OR useaskaquestion then
		socialmediabuttons=IIfVs(useemailfriend,"emailfriend")&IIfVs(useaskaquestion,IIfVs(useemailfriend,",")&"askaquestion")
	end if
end if
if instr(detailpagelayout,"review")>0 then enablecustomerratings=TRUE
emailfriendoraskquestion=instr(socialmediabuttons,"emailfriend")>0 OR instr(socialmediabuttons,"askaquestion")>0
if numcustomerratings="" then numcustomerratings=6
reviewsshown=FALSE
if wishlistonproducts=TRUE then wishlistondetail=TRUE
if SESSION("clientID")="" OR enablewishlists=FALSE OR wishlistondetail="" then wishlistondetail=FALSE
if SESSION("clientID")<>"" AND SESSION("clientLoginLevel")<>"" then minloglevel=SESSION("clientLoginLevel") else minloglevel=0
if seodetailurls then usepnamefordetaillinks=TRUE
if bmlbannerdetails<>"" AND paypalpublisherid<>"" then call displaybmlbanner(paypalpublisherid,bmlbannerdetails)
function displaytabs(thedesc)
	hasdesctab=(instr(thedesc, "<ecttab")>0)
	hasdesschema=TRUE
	if hasdesctab OR ecttabsspecials<>"" OR ecttabs<>"" OR defaultdescriptiontab<>"" then
		if defaultdescriptiontab="" then
			defaultdescriptiontab="<ecttab title="""&xxDescr&""" special=""ectdescription"">"
		elseif instr(defaultdescriptiontab," special=""ectdescription""")=0 then
			defaultdescriptiontab=replace(defaultdescriptiontab,">"," special=""ectdescription"">",1,1)
		end if
		if NOT hasdesctab AND thedesc<>"" then
			thedesc=defaultdescriptiontab & thedesc
		elseif instr(thedesc," itemprop=""description""")=0 AND NOT noschemamarkup then
			hasdesschema=FALSE
		end if
		if instr(ecttabsspecials, "%tabs%")>0 then thedesc=replace(ecttabsspecials,"%tabs%",thedesc) else thedesc=thedesc&ecttabsspecials
		if ecttabs="slidingpanel" then
			displaytabs="<div class=""slidingTabPanelWrapper""><ul class=""slidingTabPanel"">"
			tabcontent="<div id=""slidingPanel""><div"&IIfVs(NOT hasdesschema," itemprop=""description""")&">"
		else
			displaytabs="<div class=""TabbedPanels"" id=""TabbedPanels1""><ul class=""TabbedPanelsTabGroup"">"
			tabcontent="<div class=""TabbedPanelsContentGroup"""&IIfVs(NOT hasdesschema," itemprop=""description""")&">"
		end if
		dind=instr(1, thedesc, "<ecttab", 1)
		tabindex=1
		do while dind<>0
			dind=dind+8
			dind2=instr(dind, thedesc, ">")
			if dind2<>0 then
				dclass="" : did="" : dtitle="" : dimage="" : dimageov="" : dspecial=""
				tproperties=mid(thedesc,dind,dind2-dind)
				pind=instr(1, tproperties, "title=", 1)
				if pind<>0 then
					pind=instr(pind, tproperties, """")+1
					pind2=instr(pind, tproperties, """")
					dtitle=mid(tproperties,pind,pind2-pind)
				end if
				pind=instr(1, tproperties, "img=", 1)
				if pind<>0 then
					pind=instr(pind, tproperties, """")+1
					pind2=instr(pind, tproperties, """")
					dimage=mid(tproperties,pind,pind2-pind)
				end if
				pind=instr(1, tproperties, "imgov=", 1)
				if pind<>0 then
					pind=instr(pind, tproperties, """")+1
					pind2=instr(pind, tproperties, """")
					dimageov=mid(tproperties,pind,pind2-pind)
				end if
				pind=instr(1, tproperties, "special=", 1)
				if pind<>0 then
					pind=instr(pind, tproperties, """")+1
					pind2=instr(pind, tproperties, """")
					dspecial=mid(tproperties,pind,pind2-pind)
				end if
				pind=instr(1, tproperties, "id=", 1)
				if pind<>0 then
					pind=instr(pind, tproperties, """")+1
					pind2=instr(pind, tproperties, """")
					did=mid(tproperties,pind,pind2-pind)
				end if
				pind=instr(1, tproperties, "class=", 1)
				if pind<>0 then
					pind=instr(pind, tproperties, """")+1
					pind2=instr(pind, tproperties, """")
					dclass=mid(tproperties,pind,pind2-pind)
				end if
				dind2=dind2+1
				dind=instr(dind2,thedesc, "<ecttab", 1)
				if dind=0 then dcontent=mid(thedesc,dind2) else dcontent=mid(thedesc,dind2,dind-dind2)
				hascontent=TRUE
				isdescriptiontab=FALSE
				if dspecial="reviews" then
					if enablecustomerratings then
						sSQL="SELECT rtID,rtRating,rtPosterName,rtHeader,rtDate,rtComments FROM ratings WHERE rtApproved<>0 AND rtProdID='"&escape_string(prodid)&"'"
						if ratingslanguages<>"" then sSQL=sSQL & " AND rtLanguage+1 IN ("&ratingslanguages&")" else if languageid<>"" then sSQL=sSQL & " AND rtLanguage="&(int(languageid)-1) else sSQL=sSQL & " AND rtLanguage=0"
						sSQL=sSQL & " ORDER BY rtDate DESC,rtRating DESC"
						dcontent=IIfVr(usecsslayout, "<div class=""reviewtab"">", "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">") & showreviews(sSQL,FALSE) & IIfVr(usecsslayout, "</div>", "</table>")
						reviewsshown=TRUE
					else
						hascontent=FALSE
					end if
				elseif dspecial="quantitypricing" then
					quantpri=pddquantitypricing()
					if quantpri<>"" then dcontent=dcontent & quantpri else hascontent=FALSE
				elseif dspecial="related" then
					dcontent=IIfVr(usecsslayout, "<div class=""reltab"">", "<table class=""reltab"" width=""100%"">")
					if relatedtabtemplate="" then
						if usecsslayout then
							relatedtabtemplate="<div class=""reltabproduct""><div class=""reltabimage"">%img%</div><div class=""reltabname"">%name%" & IIfVs(NOT noprice," - %price%") & "</div>" & _
								"<div class=""reltabdescription"">%description%</div></div>"
						else
							relatedtabtemplate="<tr><td class=""reltabimage"" rowspan=""2"">%img%</td><td class=""reltabname"">%name%" & IIfVs(NOT noprice," - %price%") & "</td></tr>" & _
								"<tr><td class=""reltabdescription"">%description%</td></tr>"
						end if
					end if
					sSQL="SELECT pId,pSKU,pSection,"&getlangid("pName",1)&","&WSP&"pPrice,pStaticPage,pStaticURL,pDateAdded,pExemptions,pOrder,"&getlangid("pDescription",2)&" FROM products INNER JOIN relatedprods ON products.pId=relatedprods.rpRelProdID WHERE " & IIfVs(ectsiteid<>"", "pSiteID=" & ectsiteid & " AND ") & "pDisplay<>0 AND rpProdID='"&prodid&"'"
					if relatedproductsbothways=TRUE then sSQL=sSQL & " UNION SELECT pId,pSKU,pSection,"&getlangid("pName",1)&","&WSP&"pPrice,pStaticPage,pStaticURL,pDateAdded,pExemptions,pOrder,"&getlangid("pDescription",2)&" FROM products INNER JOIN relatedprods ON products.pId=relatedprods.rpProdID WHERE " & IIfVs(ectsiteid<>"", "pSiteID=" & ectsiteid & " AND ") & "pDisplay<>0 AND rpRelProdID='"&prodid&"'"
					if sSortBy<>"" then sSQL=sSQL&" ORDER BY " & sSortBy & IIfVs(isdesc," DESC")
					rs2.Open sSQL,cnn,0,1
					if rs2.EOF then
						hascontent=FALSE
					else
						do while NOT rs2.EOF
							rpsmallimage="" : rplargeimage=""
							sSQL="SELECT imageSrc,imageType FROM productimages WHERE imageProduct='" & rs2("pId") & "' AND (imageType=0 OR imageType=1) AND imageNumber=0"
							rs3.open sSQL,cnn,0,1
							do while NOT rs3.EOF
								if rs3("imageType")=0 then rpsmallimage=rs3("imageSrc") else rplargeimage=rs3("imageSrc")
								rs3.movenext
							loop
							rs3.close
							thedetailslink=getdetailsurl(rs2("pId"),rs2("pStaticPage"),rs2(getlangid("pName",1)),trim(rs2("pStaticURL")&""),IIfVr(catid<>"" AND catid<>"0" AND int(catid)<>rs2("pSection") AND nocatid<>TRUE,"cat="&catid,""),pathtohere)
							if detailslink<>"" then
								startlink=replace(replace(detailslink,"%largeimage%",rplargeimage),"%pid%", rs2("pId"))
								endlink=detailsendlink
							else
								startlink="<a class=""ectlink"" href="""&htmlspecials(thedetailslink)&""">"
								endlink="</a>"
							end if
							rtc=replace(relatedtabtemplate, "%img%", IIfVr(rpsmallimage<>"", startlink & "<img class=""reltabimage"" src=""" & rpsmallimage & """ style=""border:0"" alt="""&replace(strip_tags2(rs2(getlangid("pName",1))),"""","&quot;")&""" />" & endlink, "&nbsp;"))
							rtc=replace(rtc, "%name%", startlink & rs2(getlangid("pName",1)) & endlink)
							rtc=replace(rtc, "%id%", startlink & rs2("pId") & endlink)
							rt_totprice=rs2("pPrice")
							rtc=replace(rtc, "%price%", IIfVr(rt_totprice=0 AND pricezeromessage<>"",pricezeromessage,FormatEuroCurrency(IIfVr(showtaxinclusive=2 AND (rs2("pExemptions") AND 2)<>2, rt_totprice+(rt_totprice*thetax/100.0), rt_totprice))))
							shortdesc=rs2(getlangid("pDescription",2))&""
							if shortdescriptionlimit<>"" then if nostripshortdescription<>TRUE then shortdesc=strip_tags2(shortdesc) : shortdesc=left(shortdesc, shortdescriptionlimit) & IIfVr(len(shortdesc)>shortdescriptionlimit AND shortdescriptionlimit<>0, "...", "")
							rtc=replace(rtc, "%description%", shortdesc)
							dcontent=dcontent & rtc
							rs2.movenext
						loop
					end if
					rs2.Close
					dcontent=dcontent & IIfVr(usecsslayout, "</div>", "</table>")
				elseif dspecial="ectdescription" then
					isdescriptiontab=TRUE
				end if
				if hascontent then
					if ecttabs="slidingpanel" then
						displaytabs=displaytabs&"<li><a href=""#"" id=""ecttab"&tabindex&""" class=""tab"&IIfVr(tabindex=1,"Active","")&""" title="""&dtitle&""">"
					else
						displaytabs=displaytabs&"<li class=""" & trim(dclass & " TabbedPanelsTab") & """ tabindex=""0""" & IIfVs(did<>""," id="""&did&"""") & ">"
					end if
					if dimage<>"" then
						displaytabs=displaytabs & "<img src="""&dimage&""" alt=""" & htmlspecials(dtitle) & """ "
						if dimageov<>"" then displaytabs=displaytabs & "onmouseover=""this.src='"&dimageov&"'"" onmouseout=""this.src='"&dimage&"'"" "
						displaytabs=displaytabs & "/>"
					else
						displaytabs=displaytabs & replace(dtitle," ","&nbsp;")
					end if
					if ecttabs="slidingpanel" then
						displaytabs=displaytabs&"</a></li>"
						tabcontent=tabcontent&"<div id=""ecttab"&tabindex&"Panel"" class=""tabpanelcontent"""&IIfVs(isdescriptiontab AND NOT noschemamarkup," itemprop=""description""")&">"&dcontent&"</div>"
					else
						displaytabs=displaytabs&"</li>"
						tabcontent=tabcontent&"<div class=""tabpanelcontent"""&IIfVs(isdescriptiontab AND NOT noschemamarkup," itemprop=""description""")&">"&dcontent&"</div>"
					end if
				end if
				tabindex=tabindex+1
			end if
		loop
		if ecttabs="slidingpanel" then
			displaytabs=displaytabs&"</ul></div>"&tabcontent&"</div></div>"
			displaytabs=displaytabs&"<script>var sp2;var quotes;var lastTab=""ecttab1"";"
			displaytabs=displaytabs&"function switchTab(tab){if(tab!=lastTab){document.getElementById(tab).className=(""tabActive"");document.getElementById(lastTab).className=(""tab"");sp2.showPanel(tab+""Panel"");lastTab=tab;}}"
			displaytabs=displaytabs&"Spry.Utils.addLoadListener(function(){"
			displaytabs=displaytabs&"	Spry.$$('.slidingTabPanelWrapper').setStyle('display: block');"
			displaytabs=displaytabs&"	Spry.$$('#ecttab1"
			for i=2 to tabindex-1
				displaytabs=displaytabs&",#ecttab"&i
			next
			displaytabs=displaytabs&"').addEventListener('click', function(){ switchTab(this.id); return false; }, false);"
			displaytabs=displaytabs&"	Spry.$$('#slidingPanel').addClassName('SlidingPanels').setAttribute('tabindex', '0');"
			displaytabs=displaytabs&"	Spry.$$('#slidingPanel > div').addClassName('SlidingPanelsContentGroup');"
			displaytabs=displaytabs&"	Spry.$$('#slidingPanel .SlidingPanelsContentGroup > div').addClassName('SlidingPanelsContent');"
			displaytabs=displaytabs&"	sp2=new Spry.Widget.SlidingPanels('slidingPanel');"
			displaytabs=displaytabs&"});</script>"
		else
			displaytabs=displaytabs&"</ul>"&tabcontent&"</div></div>"
			displaytabs=displaytabs&"<script>var TabbedPanels1=new Spry.Widget.TabbedPanels(""TabbedPanels1"");</script>"
		end if
		displaytabs=">" & replace(displaytabs, "</ecttab>","")
	else
		displaytabs=IIfVs(NOT noschemamarkup," itemprop=""description""") & ">" & thedesc
	end if
end function
if magictoolbox<>"" then magictoolboxjs=magictoolbox : magictoolbox=replace(magictoolbox,"MagicZoomPlus","MagicZoom") : magictool=magictoolbox : giantimageinpopup=FALSE
sub showdetailimages()
	if thumbnailstyle="" then thumbnailstyle="width:75px;padding:3px"
	if isarray(allimages) then
		if magictoolbox<>"" AND (isarray(allgiantimages) OR lcase(magictoolbox)="magic360") then
			print "<script src=""" & lcase(magictoolboxjs) & "/" & lcase(magictoolboxjs) & ".js""></script>" & magictooloptionsjs
			if magicscrollthumbnails then print "<script src=""magicscroll/magicscroll.js""></script>" & magicscrollthumbnailsjs
			if magictoolbox="MagicSlideshow" OR magictoolbox="MagicScroll" then
				print "<div class=""" & magictoolbox & """ "&magictooloptions&">"
				for index=0 to UBOUND(allimages,2)
					if UBOUND(allgiantimages,2)>=index then giantimage=allgiantimages(0,index) else giantimage=""
					print "<img" & IIfVs(NOT noschemamarkup," itemprop=""url""") & " src=""" & allimages(0,index) & """ alt="""" "&IIfVs(giantimage<>"" AND magictoolbox="MagicSlideshow","data-fullscreen-image="""&giantimage&""" ")&"/>"
				next
				print "</div>"
			elseif lcase(magictoolbox)="magic360" then
				magictoolbox=replace(magictoolbox,"magic360","Magic360")
				if magic360images="" then magic360images=18
				imgpattern=replace(allimages(0,0),"01","{col}")
				if instr(imgpattern,"/")>0 then imgpattern=right(imgpattern,len(imgpattern)-instrrev(imgpattern,"/"))
				if isarray(allgiantimages) then giantimage=allgiantimages(0,0) else giantimage="#"
				print "<a href="""&giantimage&""" class="""&magictoolbox&""" data-magic360-options=""columns:"&magic360images&";filename:"&imgpattern&";""><img" & IIfVs(NOT noschemamarkup," itemprop=""url""") & " src="""&allimages(0,0)&""" alt="""" /></a>"
			elseif magictoolbox="MagicZoom" OR magictoolbox="MagicZoomPlus" OR magictoolbox="MagicThumb" then
				if UBOUND(allimages,2)>0 AND NOT usecsslayout then print "<table class=""detailimage allprodimages"" border=""0"" cellspacing=""1"" cellpadding=""1""><tr><td class=""mainimage"">"
				if magictoolbox="MagicThumb" then magictooloptions=replace(magictooloptions,"data-options=","rel=") else magictooloptions=replace(magictooloptions,"rel=","data-options=")
				print "<a href=""" & allgiantimages(0,0) & """ class=""" & magictoolbox & """ " & magictooloptions & " id=""zoom1""><img" & IIfVs(NOT noschemamarkup," itemprop=""url""") & " id=""prodimage"&Count&""" class=""detailimage allprodimages"" src=""" & allimages(0,0) & """ style=""border:0"" alt="""&replace(strip_tags2(rs(getlangid("pName",1))),"""","&quot;")&""" /></a>"
				if magictoolbox="MagicThumb" then relid=" rel=""thumb-id:zoom1""" else relid=""
				if magictoolbox="MagicZoom" OR magictoolbox="MagicZoomPlus" then relid=" data-zoom-id=""zoom1"""
				if UBOUND(allimages,2)>0 then
					if usecsslayout then print "<div class=""thumbnailimage detailthumbnailimage"">" else print "</td></tr><tr><td class=""thumbnailimage detailthumbnailimage"" align=""center"">"
					if magicscrollthumbnails then print "<div class=""MagicScroll"" "&magicscrollthumbnailoptions&">"
					for index=0 to UBOUND(allimages,2)
						if UBOUND(allgiantimages,2)>=index then print "<a href=""" & allgiantimages(0,index) & """ rev=""" & allimages(0,index) & """" & relid & "><img src=""" & allimages(0,index) & """ style=""" & thumbnailstyle & """ alt="""&replace(strip_tags2(rs(getlangid("pName",1))),"""","&quot;")&" #" & (index+1) & """ /></a>"
					next
					if magicscrollthumbnails then print "</div>"
					if usecsslayout then print "</div>" else print "</td></tr></table>"
				end if
			else
				print "Magic Toolbox Option Not Recognized : " & magictoolbox & "<br />"
			end if
		else
			if (UBOUND(allimages,2)>0 OR isarray(allgiantimages)) AND NOT usecsslayout then print "<table class=""detailimage allprodimages"" border=""0"" cellspacing=""1"" cellpadding=""1""><tr><td class=""mainimage"">"
			print "<img" & IIfVs(NOT noschemamarkup," itemprop=""url""") & " id=""prodimage"&Count&""" class=""detailimage allprodimages"" src=""" & allimages(0,0) & """ alt="""&replace(strip_tags2(rs(getlangid("pName",1))),"""","&quot;")&""" />"
			if isarray(allgiantimages) then showimglink="<span class=""extraimage extraimgnumof""><a class=""ectlink"" href=""javascript:showgiantimage()"">"&xxEnlrge&"</a></span>" else showimglink=""
			if UBOUND(allimages,2)>0 OR isarray(allgiantimages) then print IIfVr(usecsslayout, "<div class=""imagenavigator detailimagenavigator"">", "</td></tr><tr><td class=""imagenavigator detailimagenavigator"" align=""center"">") & IIfVr(UBOUND(allimages,2)>0,imageorbutton(imgdetailprevimg,xxPrImTx,"detailprevimg","updateprodimage("&Count&",false)",TRUE),"&nbsp;")&" "&IIfVr(UBOUND(allimages,2)>0, "<span class=""extraimage extraimagenum"" id=""extraimcnt"&Count&""">1</span> <span class=""extraimage"">"&xxOf&" "&extraimages&"</span> ", "") & showimglink & " " & IIfVr(UBOUND(allimages,2)>0,imageorbutton(imgdetailnextimg,xxNeImTx,"detailnextimg","updateprodimage("&Count&",true)",TRUE),"&nbsp;") & IIfVr(usecsslayout, "</div>", "</td></tr></table>")
		end if
	elseif psmallimage<>"" then
		print "<img" & IIfVs(NOT noschemamarkup," itemprop=""url""") & " id=""prodimage"&Count&""" class=""detailimage allprodimages"" src=""" & psmallimage & """ alt="""&replace(strip_tags2(rs(getlangid("pName",1))),"""","&quot;")&""" />"
	else
		print "&nbsp;"
	end if
end sub
sub writepreviousnextlinks()
	currcat=int(IIfVr(thecatid<>"", thecatid, catid))
	if previousid<>"" then print "<a class=""ectlink"" href="""&getdetailsurl(previousid,previousidstatic,previousidname,trim(previousstaticurl&""),IIfVs(previousidcat<>currcat AND nocatid<>TRUE,"cat="&currcat),pathtohere)&""">"
	print "<strong>&laquo; "&xxPrev&"</strong>"
	if previousid<>"" then print "</a>"
	print " | "
	if nextid<>"" then print "<a class=""ectlink"" href="""&getdetailsurl(nextid,nextidstatic,nextidname,trim(nextstaticurl&""),IIfVs(nextidcat<>currcat AND nocatid<>TRUE,"cat="&currcat),pathtohere)&""">"
	print "<strong>"&xxNext&" &raquo;</strong>"
	if nextid<>"" then print "</a>"
end sub
function detailpageurl(params)
	detailpageurl=""
	if prodid<>giftcertificateid AND prodid<>donationid then detailpageurl=getdetailsurl(prodid,hasstaticpage,rs(getlangid("pName",1)),rs("pStaticURL"),params,pathtohere)
end function
function showreviews(theSQL,showall)
	savenoschemamarkup=noschemamarkup
	if isarticle then noschemamarkup=TRUE
	numreviews=0 : totrating=0
	totSQL="SELECT COUNT(*) as numreviews, SUM(rtRating) AS totrating FROM ratings WHERE rtApproved<>0 AND rtProdID='"&escape_string(prodid)&"'"
	' if ratingslanguages<>"" then totSQL=totSQL & " AND rtLanguage+1 IN ("&ratingslanguages&")" else if languageid<>"" then totSQL=totSQL & " AND rtLanguage="&(int(languageid)-1) else totSQL=totSQL & " AND rtLanguage=0"
	rs2.open totSQL,cnn,0,1
	if NOT isnull(rs2("numreviews")) AND NOT isnull(rs2("totrating")) then
		numreviews=clng(rs2("numreviews"))
		totrating=clng(rs2("totrating"))
	end if
	rs2.close
	showreviews=IIfVr(usecsslayout, "<div", "<tr><td") & " class=""reviews"" id=""reviews"">"
	showreviews=showreviews & "<div class=""reviewtotals""><span class=""numreviews""" & IIfVs(numreviews<>0 AND NOT noschemamarkup," itemprop=""aggregateRating"" itemscope itemtype=""http://schema.org/AggregateRating""") & ">" & IIfVs(numreviews<>0,"<span class=""count""" & IIfVs(NOT noschemamarkup," itemprop=""ratingCount""") & ">"&numreviews&"</span> ") & xxRvPrRe
	if numreviews > 0 then
		showreviews=showreviews & " - "&xxRvAvRa&" <span class=""rating average""" & IIfVs(NOT noschemamarkup," itemprop=""ratingValue""") & ">"&vsround((totrating/numreviews)/2,1)&"</span> / 5"
	end if
	showreviews=showreviews & "</span><span class=""showallreview"">"
	if showall then
		showreviews=showreviews & " (<a class=""ectlink readreview"" rel=""nofollow"" href="""&detailpageurl("review=all"&IIfVr(thecatid<>"","&amp;cat="&thecatid,"")&"&amp;ro=1")&""">"&xxRvBest&"</a>"
		showreviews=showreviews & " | <a class=""ectlink readreview"" rel=""nofollow"" href="""&detailpageurl("review=all"&IIfVr(thecatid<>"","&amp;cat="&thecatid,"")&"&amp;ro=2")&""">"&xxRvWors&"</a>"
		showreviews=showreviews & " | <a class=""ectlink readreview"" rel=""nofollow"" href="""&detailpageurl("review=all"&IIfVr(thecatid<>"","&amp;cat="&thecatid,""))&""">"&xxRvRece&"</a>"
		showreviews=showreviews & " | <a class=""ectlink readreview"" rel=""nofollow"" href="""&detailpageurl("review=all"&IIfVr(thecatid<>"","&amp;cat="&thecatid,"")&"&amp;ro=3")&""">"&xxRvOld&"</a>)"
	elseif numreviews > 0 then
		showreviews=showreviews & " (<a class=""ectlink showallreview"" rel=""nofollow"" href="""&detailpageurl("review=all"&IIfVr(thecatid<>"","&amp;cat="&thecatid,""))&""">"&xxShoAll&"</a>)"
	end if
	showreviews=showreviews & "</span></div>"
	rs2.CursorLocation=3 ' adUseClient
	if allreviewspagesize="" then allreviewspagesize=30
	if showall then thepagesize=allreviewspagesize else thepagesize=numcustomerratings
	rs2.CacheSize=thepagesize
	rs2.open theSQL,cnn,0,1
	if NOT rs2.EOF then
		rs2.MoveFirst
		rs2.PageSize=thepagesize
		if getget("pg")="" then CurPage=1 else CurPage=int(getget("pg"))
		iNumOfPages=Int((rs2.RecordCount + (thepagesize-1)) / thepagesize)
		rs2.AbsolutePage=CurPage
	end if
	recordcount=0
	if NOT rs2.EOF then
		if NOT (onlyclientratings AND SESSION("clientID")="") then showreviews=showreviews & "<div class=""clickreview"">" & imageorbuttontag(imgclicktoreview,xxClkRev,"clickreview",detailpageurl("review=true"&IIfVr(thecatid<>"","&amp;cat="&thecatid,"")), FALSE) & "</div>"
		do while NOT rs2.EOF AND recordcount < rs2.PageSize
			showreviews=showreviews & "<div class=""ecthreview""" & IIfVs(NOT noschemamarkup," itemprop=""review"" itemscope itemtype=""http://schema.org/Review""") & ">"
				showreviews=showreviews & "<div class=""reviewstarsheader largereviewstars""><span class=""rating""" & IIfVs(NOT noschemamarkup," itemprop=""reviewRating"" itemscope itemtype=""http://schema.org/Rating""") & ">" & IIfVs(NOT noschemamarkup,"<meta itemprop=""worstRating"" content=""1"" /><meta itemprop=""bestRating"" content=""5"" /><meta itemprop=""ratingValue"" content="""&vsround(cint(rs2("rtRating"))/2,0)&""" />")
				if imgreviewcart="" then call displayreviewicons()
				for index=1 to int(cint(rs2("rtRating")) / 2)
					if imgreviewcart<>"" then
						showreviews=showreviews & "<img class=""detailreviewstars"" src=""images/"&imgreviewcart&""" alt="""" />"
					else
						showreviews=showreviews & "<svg viewBox=""0 0 24 24"" class=""icon"" style=""max-width:30px""><use xlink:href=""#review-icon-full""></use></svg>"
					end if
				next
				ratingover=cint(rs2("rtRating"))
				if ratingover / 2 > int(ratingover / 2) then
					if imgreviewcart<>"" then
						showreviews=showreviews & "<img class=""detailreviewstars"" src=""images/"&replace(imgreviewcart,".","hg.")&""" alt="""" />"
					else
						showreviews=showreviews & "<svg viewBox=""0 0 24 24"" class=""icon"" style=""max-width:30px""><use xlink:href=""#review-icon-half""></use></svg>"
					end if
					ratingover=ratingover + 1
				end if
				for index=int(ratingover / 2) + 1 to 5
					if imgreviewcart<>"" then
						showreviews=showreviews & "<img class=""detailreviewstars"" src=""images/"&replace(imgreviewcart,".","g.")&""" alt="""" />"
					else
						showreviews=showreviews & "<svg viewBox=""0 0 24 24"" class=""icon"" style=""max-width:30px""><use xlink:href=""#review-icon-empty""></use></svg>"
					end if
				next
				showreviews=showreviews & "</span> <span class=""reviewheader""" & IIfVs(NOT noschemamarkup," itemprop=""name""") & ">" & htmlspecials(rs2("rtHeader")) & "</span></div>"
				showreviews=showreviews & "<div class=""reviewname""><span class=""reviewer""" & IIfVs(NOT noschemamarkup," itemprop=""author"" itemscope itemtype=""http://schema.org/Person""") & ">" & IIfVs(NOT noschemamarkup,"<meta itemprop=""name"" content=""" & htmlspecials(rs2("rtPosterName")) & """ />") & htmlspecials(rs2("rtPosterName")) & "</span> - <span class=""dtreviewed"">" & rs2("rtDate") & IIfVs(NOT noschemamarkup,"<meta itemprop=""datePublished"" content="""&iso8601date(rs2("rtDate"))&""" />") & "</span></div>"
				thecomments=rs2("rtComments")
				if NOT allowhtmlinreviews then thecomments=htmlspecials(thecomments)
				hasextracomments=FALSE
				if NOT showall then
					if customerratinglength="" then customerratinglength=255
					if len(thecomments)>customerratinglength then
						thecomments=left(thecomments, customerratinglength) & "<span class=""extracommentsdots"" id=""extracommentsdots" & recordcount & """>" & xxReaMor & "</span><span id=""extracomments" & recordcount & """ style=""display:none"">" & right(thecomments, len(thecomments) - customerratinglength) & "</span>"
						hasextracomments=TRUE
					end if
				end if
				showreviews=showreviews & "<div class=""reviewcomments""" & IIfVs(NOT noschemamarkup," itemprop=""reviewBody""") & IIfVs(hasextracomments," onclick=""ectexpandreview(" & recordcount & ");this.style.cursor='auto'"" style=""cursor:pointer""") & ">" & replace(thecomments, vbCrLf, "<br />") & "</div>"
			showreviews=showreviews & "</div>"
			recordcount=recordcount + 1
			rs2.movenext
		loop
	else
		showreviews=showreviews & "<span class=""noreview"">" & xxRvNone & "</span><br /><hr class=""review"" />"
	end if
	rs2.close
	if NOT (onlyclientratings AND SESSION("clientID")="") then showreviews=showreviews & "<div class=""clickreview"">" & imageorbuttontag(imgclicktoreview,xxClkRev,"clickreview",detailpageurl("review=true"&IIfVr(thecatid<>"","&amp;cat="&thecatid,"")), FALSE) & "</div>"
	showreviews=showreviews & IIfVr(usecsslayout, "</div>", "</td></tr>")
	pblink=""
	for each objQS in request.querystring
		if objQS<>"id" AND objQS<>"pg" AND NOT (objQS="prod" AND seodetailurls) then pblink=pblink & urlencode(objQS) & "=" & urlencode(getget(objQS)) & "&amp;"
	next
	pblink="<a class=""ectlink"" href="""&detailpageurl(pblink & "pg=")
	if showall AND iNumOfPages > 1 then showreviews=showreviews & IIfVr(usecsslayout, "<div", "<tr><td align=""center""") & " class=""pagenumbers"">" & writepagebar(CurPage,iNumOfPages,xxPrev,xxNext,pblink,TRUE) & IIfVr(usecsslayout, "</div>", "</td></tr>")
	noschemamarkup=savenoschemamarkup
end function
sub schemaconditionavail()
	if setschemacondition AND NOT noschemamarkup then print "<meta itemprop=""itemCondition"" itemscope itemtype=""http://schema.org/OfferItemCondition"" content=""http://schema.org/"&IIfVr(schemaitemcondition<>"",schemaitemcondition,"NewCondition")&""" />"
	if setschemaavailability AND NOT noschemamarkup then
		isinstock=cint(rs("pSell"))<>0
		if useStockManagement then
			if cint(rs("pStockByOpts"))<>0 then isinstock=optionshavestock else isinstock=rs("pInStock")>rs("pMinQuant")
		end if
		print "<meta itemprop=""availability"" content=""http://schema.org/"&IIfVr(isinstock,"InStock","OutOfStock")&""" />"
	end if
end sub
sub pddsearchwords()
	if trim(rs(getlangid("pSearchParams",4194304))&"")<>"" then
		searchprms=split(rs(getlangid("pSearchParams",4194304)),IIfVr(searchwordsseparator<>"",searchwordsseparator," "))
		print "<div class=""searchwords"">"
		if searchwordsheading<>"" then print "<div class=""searchwordsheading"">" & searchwordsheading & "</div>"
		for indexsw=0 to UBOUND(searchprms)
			print IIfVs(indexsw<>0," ") & "<a class=""ectlink searchwords"" href=""search"&extension&"?pg=1&amp;stext="&urlencode(searchprms(indexsw))&IIfVs(searchwordsnobox,"&amp;nobox=true")&""">"&htmlspecials(searchprms(indexsw))&"</a>"
		next
		print "</div>"
	end if
end sub
function pddquantitypricing()
	pddquantitypricing=""
	sSQL="SELECT " & WSP & "pPrice,pbQuantity,"&IIfVs(WSP<>"","pbWholesalePercent AS ")&"pbPercent FROM pricebreaks WHERE pbProdID='" & escape_string(rs("pId")) & "' ORDER BY pbQuantity"
	rs2.open sSQL,cnn,0,1
	if NOT rs2.EOF then
		pddquantitypricing="<div class=""detailquantpricingwrap""><div class=""detailquantpricing"" style=""display:table"">"
		if xxQuaPri<>"" then pddquantitypricing=pddquantitypricing & "<div class=""detailqpheading"" style=""display:table-caption"">" & xxQuaPri & "</div>"
		pddquantitypricing=pddquantitypricing & "<div class=""detailqpheaders"" style=""display:table-row""><div class=""detailqpheadquant"" style=""display:table-cell"">" & xxQuanti & "</div><div class=""detailqpheadprice"" style=""display:table-cell"">" & xxPriQua & "</div></div>"
		quantpricearray=rs2.getrows()
		for index=0 to UBOUND(quantpricearray,2)
			if index<UBOUND(quantpricearray,2) then
				nextquant=quantpricearray(1,index+1)-1
				nextquant=IIfVs(nextquant>quantpricearray(1,index),"-" & nextquant)
			else
				nextquant="+"
			end if
			pddquantitypricing=pddquantitypricing & "<div class=""detailqprow"" style=""display:table-row""><div class=""detailqpquant"" style=""display:table-cell"">" & quantpricearray(1,index) & nextquant & "</div><div class=""detailqpprice"" style=""display:table-cell"">" & FormatEuroCurrency(IIfVr(clng(quantpricearray(2,index))<>0,rs("pPrice")-((rs("pPrice")*quantpricearray(0,index))/100),quantpricearray(0,index))) & "</div></div>"
		next
		pddquantitypricing=pddquantitypricing & "</div></div>"
	end if
	rs2.close
end function
sub pddreviewstars(issmall)
	if rs("pNumRatings")>0 then
		if imgreviewcart="" then call displayreviewicons()
		print "<div class=""" & IIfVr(issmall,"small","large") & "reviewstars detailreviewstars""><a href="""&detailpageurl("")&"#reviews"">"
		therating=cint(rs("pTotRating")/rs("pNumRatings"))
		for index=1 to int(therating / 2)
			if imgreviewcart<>"" then
				print "<img class=""detailreviewstars"" src=""images/"&IIfVs(issmall,"s")&imgreviewcart&""" alt="""" />"
			else
				print "<svg viewBox=""0 0 24 24"" class=""icon"" style=""max-width:30px""><use xlink:href=""#review-icon-full""></use></svg>"
			end if
		next
		ratingover=therating
		if ratingover / 2 > int(ratingover / 2) then
			if imgreviewcart<>"" then
				print "<img class=""detailreviewstars"" src=""images/"&IIfVs(issmall,"s")&replace(imgreviewcart,".","hg.")&""" alt="""" />"
			else
				print "<svg viewBox=""0 0 24 24"" class=""icon"" style=""max-width:30px""><use xlink:href=""#review-icon-half""></use></svg>"
			end if
			ratingover=ratingover + 1
		end if
		for index=int(ratingover / 2) + 1 to 5
			if imgreviewcart<>"" then
				print "<img class=""detailreviewstars"" src=""images/"&IIfVs(issmall,"s")&replace(imgreviewcart,".","g.")&""" alt="""" />"
			else
				print "<svg viewBox=""0 0 24 24"" class=""icon"" style=""max-width:30px""><use xlink:href=""#review-icon-empty""></use></svg>"
			end if
		next
		print "</a><span class=""prodratingtext"">"
		if detailreviewstarstext<>"" then print replace(replace(replace(detailreviewstarstext,"%numratings%",rs("pNumRatings")),"%totrating%",vsround(rs("pTotRating")/rs("pNumRatings")/2,1)),"%reviewlink%",detailpageurl("")&"#reviews")
		print "</span></div>"
	else
		print detailreviewnoratings
	end if
end sub
sub pddsreviews()
	if rs("pNumRatings")>0 then print showproductreviews(1,"detailrating")	
end sub
sub pddreviews()
	if getpost("review")="true" OR getget("review")="all" then
		' Do nothing
	elseif enablecustomerratings AND getget("review")="true" then
		if onlyclientratings AND SESSION("clientID")="" then
			print "<tr><td align=""center"">Only logged in customers can review products.</td></tr>"
		else
			if NOT usecsslayout then print "<tr><td>" %>
	<script>
	/* <![CDATA[ */
	function checkratingform(frm){
	if(frm.ratingstars.selectedIndex==0){
		alert("<%=jscheck(xxRvPlsS)%>.");
		frm.ratingstars.focus();
		return(false);
	}
	if(frm.reviewposter.value==""){
		alert("<%=jscheck(xxPlsEntr&" """&xxRvPosb)%>\".");
		frm.reviewposter.focus();
		return(false);
	}
	if(frm.reviewheading.value==""){
		alert("<%=jscheck(xxPlsEntr&" """&xxRvHead)%>\".");
		frm.reviewheading.focus();
		return(false);
	}
<%	if reviewcommentsminlength<>"" then %>
	if(frm.reviewcomments.value.length<<%=reviewcommentsminlength%>){
		alert("<%=jscheck(replace(xxMinLen,"%s",reviewcommentsminlength)&" """&xxRvComm)%>\".");
		frm.reviewcomments.focus();
		return(false);
	}
<%	end if
	if recaptchaenabled(32) then print "if(!reviewcaptchaok){ alert(""" & jscheck(xxRecapt) & """);return(false); }" %>
	document.getElementById('rfsectgrp1').value=document.getElementById('ratingstars')[document.getElementById('ratingstars').selectedIndex].value;
	document.getElementById('rfsectgrp2').value=document.getElementById('reviewposter').value.length;
	return(true);
	}
	/* ]]> */
	</script>
		<form method="post" action="<%=detailpageurl(IIfVr(thecatid<>"","cat="&thecatid,""))%>" style="margin:0px; padding:0px;" onsubmit="return checkratingform(this)">
		<input type="hidden" name="review" value="true" />
		<input type="hidden" name="rfsectgrp1" id="rfsectgrp1" value="6344" />
		<input type="hidden" name="rfsectgrp2" id="rfsectgrp2" value="923" />
		<div class="reviewformblock">
			<div class="ectformline reviewformline"><div class="reviewlabels"><%=redstar & xxRvRati%>:</div><div class="reviewfields"><select size="1" name="ratingstars" id="ratingstars" class="reviewform"><option value=""><%=xxPlsSel%></option><%
				for index=1 to 5
					print "<option value="""&index&""">"&index&" "&xxStars&"</option>"
				next %></select></div></div>
			<div class="ectformline reviewformline"><div class="reviewlabels"><%=redstar & xxRvPosb%>:</div><div class="reviewfields"><input type="text" size="20" name="reviewposter" id="reviewposter" maxlength="64" value="<%=htmlspecials(SESSION("clientUser"))%>" class="reviewform" /></div></div>
			<div class="ectformline reviewformline"><div class="reviewlabels"><%=redstar & xxRvHead%>:</div><div class="reviewfields"><input type="text" size="40" name="reviewheading" maxlength="253" class="reviewform" /></div></div>
			<div class="ectformline reviewformline"><div class="reviewlabels"><% if reviewcommentsminlength<>"" then print redstar %><%=xxRvComm%>:</div><div class="reviewfields"><textarea name="reviewcomments" cols="38" rows="8" class="reviewform"></textarea></div></div>
<%			if recaptchaenabled(32) then %>
				<div class="ectformline reviewformline"><div class="reviewlabels">&nbsp;</div><div class="reviewfields"><% call displayrecaptchajs("reviewcaptcha",TRUE,FALSE) %><div id="reviewcaptcha"></div></div></div>
<%			end if %>
			<div class="ectformline reviewformline"><div class="reviewlabels"></div><div class="reviewfields"><input type="submit" value="<%=xxSubmt%>" class="ectbutton reviewsubmit" /></div></div>
		</div>
		</form>
<%			if NOT usecsslayout then print "</td></tr>"
		end if
	elseif enablecustomerratings then
		sSQL="SELECT rtID,rtRating,rtPosterName,rtHeader,rtDate,rtComments FROM ratings WHERE rtApproved<>0 AND rtProdID='"&escape_string(prodid)&"'"
		if ratingslanguages<>"" then sSQL=sSQL & " AND rtLanguage+1 IN ("&ratingslanguages&")" else if languageid<>"" then sSQL=sSQL & " AND rtLanguage="&(int(languageid)-1) else sSQL=sSQL & " AND rtLanguage=0"
		sSQL=sSQL & " ORDER BY rtDate DESC,rtRating DESC"
		if NOT reviewsshown AND productindb then print showreviews(sSQL,FALSE)
	end if
end sub
sub pddprodnavigation()
	if NOT (nobuyorcheckout OR nocheckoutbutton) then print "<div class=""catnavandcheckout catnavdetail"">"
	print "<div class=""catnavigation catnavdetail"">" & tslist & "</div>" & vbCrLf
	if NOT (nobuyorcheckout OR nocheckoutbutton) then print "<div class=""catnavcheckout"">" & imageorbutton(imgcheckoutbutton,xxCOTxt,"checkoutbutton","cart"&extension, FALSE) & "</div></div>" & vbCrLf
end sub
sub pddcheckoutbutton()
	' Not used now
end sub
sub pddproductimage()
	if NOT usecsslayout then print "<table width=""100%"" border=""0"" cellspacing=""3"" cellpadding=""3""><tr>"
	print IIfVr(usecsslayout, "<div", "<td width=""30%"" align=""center""") & " itemprop=""image"" itemscope itemtype=""https://schema.org/ImageObject"" class=""detailimage allprodimages" & IIfVs(magictoolbox=""," ectnomagicimage") & """>"
	showdetailimages()
	print IIfVr(usecsslayout, "</div>", "</td>")
end sub
sub pddproductid()
	if NOT usecsslayout then print "<td>&nbsp;</td><td width=""70%"" valign=""top"" class=""detail"">"
	if showproductid OR hascustomlayout then print "<div class=""detailid"">" & IIfVs(xxPrId<>"","<span class=""prodidlabel detailidlabel"">" & xxPrId & "</span> ") & IIfVs(NOT noschemamarkup,"<span itemprop=""productID"">") & rs("pID") & IIfVs(NOT noschemamarkup,"</span>") & "</div>"
end sub
sub pddmanufacturer(islinked)
	startlink="" : endlink=""
	if NOT IsNull(rs(getlangid("scName",131072))) then
		if islinked then
			if trim(rs("scURL")&"")<>"" then
				newloc=getcatid(rs("scURL"),rs("scURL"),seomanufacturerpattern)
			else
				newloc=IIfVs(NOT seocategoryurls, "products" & extension & "?man=") & getcatid(rs("scID"),rs(getlangid("scName",131072)),seomanufacturerpattern)
			end if
			startlink="<a href=""" & newloc & """>"
			endlink="</a>"
		end if
		print "<div class=""detailmanufacturer"">" & IIfVs(xxManLab<>"","<span class=""detailmanufacturerlabel"">" & xxManLab & "</span> ") & IIfVs(NOT noschemamarkup,"<span itemprop=""manufacturer"">") & startlink & rs(getlangid("scName",131072)) & endlink & IIfVs(NOT noschemamarkup,"</span>") & "</div>"
	end if
end sub
sub pddsku()
	if (showproductsku<>"" OR hascustomlayout) AND trim(rs("pSKU")&"")<>"" then print "<div class=""detailsku"">" & IIfVs(showproductsku<>"","<span class=""prodskulabel detailskulabel"">" & showproductsku & "</span> ") & IIfVs(NOT noschemamarkup,"<span itemprop=""" & skuschemaidentifier & """>") & rs("pSKU") & IIfVs(NOT noschemamarkup,"</span>") & "</div>"
end sub
sub pddcustom(custid,customlabel)
	if trim(rs("pCustom"&custid)&"")<>"" then print "<div class=""detailcustom"&custid&""">" & customlabel & rs("pCustom"&custid) & "</div>"
end sub
sub pdddateadded()
	if NOT isnull(rs("pDateAdded")) then print "<div class=""detaildateadded"">" & IIfVs(xxDatLab<>"","<div class=""detaildateaddedlabel"">" & xxDatLab & "</div>") & "<div class=""detaildateaddeddate"">" & IIfVs(NOT noschemamarkup AND isarticle,"<meta itemprop=""datePublished dateModified"" content=""" & iso8601date(rs("pDateAdded")) & """>") & FormatDateTime(rs("pDateAdded"),0) & "</div></div>"
end sub
sub pdddetailname()
	print sstrong & "<div class=""detailname""><h1"&IIfVs(NOT noschemamarkup," itemprop=""name" & IIfVs(isarticle," headline") & """") & ">" & rs(getlangid("pName",1))&"</h1>"&xxDot
	if alldiscounts<>"" then print "<div class=""discountsapply detaildiscountsapply"">"&xxDsApp&"</div>"
	print "</div>" & estrong
end sub
sub pdddiscounts()
	if alldiscounts<>"" then print "<div class=""detaildiscounts"">" & alldiscounts & "</div>"
end sub
sub pddinstock()
	if useStockManagement AND (showinstock OR hascustomlayout) AND (rs("pInStock")<=clng(stockdisplaythreshold) OR stockdisplaythreshold="") then if cint(rs("pStockByOpts"))=0 then print "<div class=""detailinstock"">" & IIfVs(xxInStoc<>"","<span class=""detailinstocklabel"">" & xxInStoc & "</span> ") & vrmax(0,rs("pInStock")) & "</div>"
end sub
sub pddshortdescription()
	print "<div class=""detailshortdescription"">"&shortdesc&"</div>"
end sub
sub pdddescription()
	if NOT usecsslayout then print "<br />"
	longdesc=longdesc&ectextralongdescription
	if usedetailbodyformat=3 then
	elseif longdesc<>"" then
		print "<div class=""detaildescription"""&displaytabs(longdesc)&"</div>"
	elseif shortdesc<>"" then
		print "<div class=""detaildescription""" & IIfVs(NOT noschemamarkup," itemprop=""description""") & ">"&shortdesc&"</div>"
	end if
end sub
sub pddlistprice()
	if noprice=TRUE then
		print "&nbsp;"
	elseif cdbl(rs("pListPrice"))<>0.0 then
		plistprice=IIfVr(showtaxinclusive=2 AND (rs("pExemptions") AND 2)<>2,rs("pListPrice")+(rs("pListPrice")*thetax/100.0), rs("pListPrice"))
		yousaveprice=plistprice-IIfVr(showtaxinclusive=2 AND (rs("pExemptions") AND 2)<>2,rs("pPrice")+(rs("pPrice")*thetax/100.0),rs("pPrice"))
		print "<div class=""detaillistprice"" id=""listdivec" & Count & """" & IIfVs(yousaveprice<=0," style=""display:none""") & ">" & Replace(xxListPrice, "%s", FormatEuroCurrency(plistprice)) & IIfVs(yousavetext<>"" AND yousaveprice>0,replace(yousavetext,"%s",FormatEuroCurrency(yousaveprice))) & "</div>"
	end if
end sub
sub pddprice()
	if noprice<>TRUE then
		separatetaxinc=showtaxinclusive=1 AND (rs("pExemptions") AND 2)<>2
		displayprice=IIfVr(showtaxinclusive=2 AND (rs("pExemptions") AND 2)<>2, totprice+(totprice*thetax/100.0), totprice)
		print "<div class=""detailprice"""&IIfVs(totprice<>0 AND NOT noschemamarkup," itemprop=""offers"" itemscope itemtype=""http://schema.org/Offer""><meta itemprop=""priceCurrency"" content="""&countryCurrency&"""")&"><strong>" & xxPrice&IIfVs(xxPrice<>"",":") & "</strong> <span class=""price"" id=""pricediv" & Count & """"&IIfVs(totprice<>0 AND NOT noschemamarkup AND NOT separatetaxinc," itemprop=""price"" content="""&FormatNumberUS(displayprice,2,-1,0,0)&"""")&">" & IIfVr(totprice=0 AND pricezeromessage<>"",pricezeromessage,FormatEuroCurrency(displayprice)) & "</span><link itemprop=""url"" href=""" & detailpageurl("") & """> "
		if separatetaxinc then print "<span class=""inctax"" id=""taxmsg" & Count & """" & IIfVs(totprice=0, " style=""display:none""") & ">" & Replace(ssIncTax,"%s", "<span id=""pricedivti" & Count & """" & IIfVs(NOT noschemamarkup," itemprop=""price"" content="""&FormatNumberUS(totprice+(totprice*thetax/100.0),2,-1,0,0)&"""") & ">" & IIfVr(totprice=0, "-", FormatEuroCurrency(totprice+(totprice*thetax/100.0))) & "</span> ") & "</span>"
		call schemaconditionavail()
		print "</div>"
	end if
end sub
sub pddextracurrency()
	if noprice<>TRUE OR hascustomlayout then
		extracurr=""
		if currRate1<>0 AND currSymbol1<>"" then extracurr=replace(currFormat1, "%s", FormatNumber(totprice*currRate1, checkDPs(currSymbol1))) & currencyseparator
		if currRate2<>0 AND currSymbol2<>"" then extracurr=extracurr & replace(currFormat2, "%s", FormatNumber(totprice*currRate2, checkDPs(currSymbol2))) & currencyseparator
		if currRate3<>0 AND currSymbol3<>"" then extracurr=extracurr & replace(currFormat3, "%s", FormatNumber(totprice*currRate3, checkDPs(currSymbol3)))
		if extracurr<>"" then print "<div class=""detailcurrency""><span class=""extracurr"" id=""pricedivec" & Count & """>" & IIfVr(totprice=0, "", extracurr) & "</span></div>"
		if showquantitypricing AND NOT hascustomlayout then print pddquantitypricing()
		if NOT usecsslayout then print "<hr width=""80%"" class=""detailhr detailcurrencyhr currencyhr"" />"
	end if
end sub
sub pddquantity(wantquanttext)
	if hascustomlayout then
		if (isinstock OR isbackorder) AND nobuyorcheckout<>TRUE AND (showquantondetail=TRUE OR IsEmpty(showquantondetail)) AND hasmultipurchase=0 then
			print "<div class=""detailquantity"">" & IIfVs(wantquanttext, "<div class=""detailquantitytext"">" & xxQuant & "</div>") & quantitymarkup(FALSE,Count,TRUE,"",FALSE) & "</div>"
		end if
	end if
end sub
sub pddoptions()
	call displayformvalidator()
	if optjs<>"" then
		print "<script>/* <![CDATA[ */"&optjs&"/* ]]> */</script>"
		if mustincludepopcalendar AND NOT hasincludedpopcalendar then print "<script>var ectpopcalisproducts=1;" & ectpopcalendarjs & IIfVs(storelang<>"en" AND storelang<>"","var ectpopcallang='"&storelang&"'") & vbLf & "</script><script src=""vsadmin/popcalendar.js""></script>" : hasincludedpopcalendar=TRUE
	end if
	if isarray(prodoptions) then
		if abs(prodoptions(1,0))=4 then thestyle="" else thestyle=" width=""100%"""
		if optionshtml<>"" then optionshtml="<div class=""detailoptions""" & IIfVs(NOT usecsslayout," style=""float:left;width:98%;""") & ">" & optionshtml & "</div>"
		if NOT hascustomlayout AND (isinstock OR isbackorder) AND nobuyorcheckout<>TRUE AND (showquantondetail=TRUE OR IsEmpty(showquantondetail)) AND hasmultipurchase=0 then
			optionshtml=optionshtml & "<div class=""detailquantity""><div class=""detailquantitytext"">" & xxQuant & ":" & "</div>" & quantitymarkup(FALSE,Count,TRUE,"",FALSE) & "</div>"
		end if
	elseif NOT hascustomlayout then
		if (isinstock OR isbackorder) AND nobuyorcheckout<>TRUE AND (showquantondetail=TRUE OR IsEmpty(showquantondetail)) then
			optionshtml=optionshtml & IIfVr(usecsslayout, "<div class=""detailquantity""><div class=""detailquantitytext"">", "<table border=""0"" cellspacing=""1"" cellpadding=""1"" width=""100%""><tr><td align=""right"">")
			optionshtml=optionshtml & xxQuant & ":"
			optionshtml=optionshtml & IIfVr(usecsslayout, "</div>", "</td><td>")
			optionshtml=optionshtml & quantitymarkup(FALSE,Count,TRUE,"",FALSE) & IIfVr(usecsslayout, "</div>", "</td></tr></table>")
		end if
	end if
end sub
sub pddcatcontentregion()
	sSQL="SELECT contentID,contentName,"&getlangid("contentData",32768)&" FROM contentregions WHERE contentName='catcontentregion"&escape_string(catid)&"'"
	rs2.open sSQL,cnn,0,1
	if NOT rs2.EOF then
		print "<div class=""detailcontentregion detailregioncatid" & catid & """>" & rs2(getlangid("contentData",32768)) & "</div>" & vbLf
	end if
	rs2.close
end sub
sub pddcontentregion(contentid)
	if is_numeric(contentid) then
		sSQL="SELECT " & getlangid("contentData",32768) & " FROM contentregions WHERE contentID='" & escape_string(contentid) & "'"
		rs2.open sSQL,cnn,0,1
		if NOT rs2.EOF then
			print "<div class=""detailcontentregion detailregionid" & contentid & """>" & rs2(getlangid("contentData",32768)) & "</div>" & vbLf
		end if
		rs2.close
	end if
end sub
isfirstaddtocart=TRUE
sub pddaddtocart()
	atcmu=IIfVs(NOT usecsslayout, "<p align=""center"">")
	if nobuyorcheckout=TRUE then
		atcmu=atcmu & "&nbsp;"
	else
		if totprice=0 AND nosellzeroprice=TRUE then
			atcmu=atcmu & "&nbsp;"
		elseif isinstock OR isbackorder then
			if isfirstaddtocart then
				call writehiddenvar("id", rs("pID"))
				call writehiddenvar("mode", "add")
				if wishlistondetail then call writehiddenvar("listid", "")
			end if
			if usecsslayout then atcmu=atcmu & "<div class=""addtocart detailaddtocart"">"
			if isbackorder then
				if usehardaddtocart then atcmu=atcmu & imageorsubmit(imgbackorderbutton,xxBakOrd,"buybutton backorder detailbuybutton detailbackorder") else atcmu=atcmu & imageorbuttontag(imgbackorderbutton,xxBakOrd,"buybutton backorder detailbuybutton detailbackorder","subformid("&Count&",'','')",TRUE)
			else
				if custombuybutton<>"" then
					atcmu=atcmu & custombuybutton
				else
					if usehardaddtocart then atcmu=atcmu & imageorsubmit(imgbuybutton,xxAddToC,"buybutton detailbuybutton") else atcmu=atcmu & imageorbuttontag(imgbuybutton,xxAddToC,"buybutton detailbuybutton"" id=""ectaddcart"&Count,"subformid("&Count&",'','')",TRUE)
				end if
			end if
			if wishlistondetail then atcmu=atcmu & "<div class=""wishlistcontainer detailwishlist"">" & imageorbuttontag(imgaddtolist,xxAddLis,"detailwishlist","gtid="&Count&";return displaysavelist(this,event,window)",TRUE) & "</div>"
			if usecsslayout then atcmu=atcmu & "</div>"
		else
			if usecsslayout then atcmu=atcmu & "<div class=""addtocart detailaddtocart detailoutofstock"">"
			if notifybackinstock then
				atcmu=atcmu & imageorbuttontag(imgnotifyinstock,xxNotBaS,"notifystock detailnotifystock","return notifyinstock(false,'"&replace(rs("pID"),"'","\'")&"','"&replace(rs("pID"),"'","\'")&"',"&IIfVr(cint(rs("pStockByOpts"))<>0 AND NOT optionshavestock,"-1","0")&")", TRUE)
			else
				atcmu=atcmu & "<button class=""ectbutton outofstock detailoutofstock"" type=""button"" disabled=""disabled"">" & xxOutStok & "</button>"
			end if
			if usecsslayout then atcmu=atcmu & "</div>" else atcmu=atcmu & "<br />"
		end if
	end if
	isfirstaddtocart=FALSE
end sub
sub pddpreviousnext()
	if previousid<>"" OR nextid<>"" then
		print IIfVr(usecsslayout, "<div class=""previousnext"">", "</p><p class=""pagenumbers"" align=""center"">")
		call writepreviousnextlinks()
		print IIfVr(usecsslayout, "</div>", "<br />")
	end if
end sub
sub pddemailfriend()
	if usedetailbodyformat=3 AND emailfriendoraskquestion then print "<br />" : call pddsocialmedia()
	if usedetailbodyformat=4 AND emailfriendoraskquestion then print "<div class=""emailfriend"">" : call pddsocialmedia() : print "</div>"
	if NOT usecsslayout then print "</p><hr width=""80%"" class=""detailhr detailhrbottom"" />"
	if usedetailbodyformat=2 AND emailfriendoraskquestion then print "<p align=""center"">" : call pddsocialmedia() : print "</p>"
	if NOT usecsslayout then print "</td></tr>"
	if usedetailbodyformat=2 OR usedetailbodyformat=4 then
	elseif longdesc<>"" then
		print IIfVs(NOT usecsslayout, "<tr><td colspan=""3"" class=""detaildescription"">") & "<div class=""detaildescription"""&displaytabs(longdesc)&"</div>" & IIfVs(NOT usecsslayout, "</td></tr>")
	elseif shortdesc<>"" then
		print IIfVs(NOT usecsslayout, "<tr><td colspan=""3"" class=""detaildescription"">") & "<div class=""detaildescription""" & IIfVs(NOT noschemamarkup," itemprop=""description""") & ">"&shortdesc&"</div>" & IIfVs(NOT usecsslayout, "</td></tr>")
	end if
	if NOT usecsslayout then print "</table>"
end sub
sub pddsocialmedia()
	socialmediaarray=split(lcase(replace(socialmediabuttons," ","")),",")
	if isempty(smallsocialbuttons) then smallsocialbuttons=TRUE
	buttonlang="en_US"
	if storelang="fr" then buttonlang="fr_FR"
	if storelang="es" then buttonlang="es_ES"
	if storelang="de" then buttonlang="de_DE"
	if storelang="es" then buttonlang="es_ES"
	if storelang="dk" then buttonlang="da_DK"
	if storelang="it" then buttonlang="it_IT"
	if storelang="nl" then buttonlang="nl_NL"
	if storelang="pt" then buttonlang="pt_BR"
	addand="" : newqs=""
	for each objitem in request.querystring
		if objitem<>"prod" AND objitem<>"cat" then newqs=newqs&addand&objitem&"="&urlencode(getget(objitem)) : addand="&amp;"
	next
	thisurl=getfullurl(getdetailsurl(rs("pId"),rs("pStaticPage"),rs(getlangid("pName",1)),rs("pStaticURL"),newqs,pathtohere))
	print "<div class=""socialmediabuttons"">"
	hasdisplayedone=FALSE
	for each layoutoption in socialmediaarray
		if hasdisplayedone AND socialmediaseparator<>"" then print socialmediaseparator
		hasdisplayedone=TRUE
		if layoutoption="facebook" then
		print "<div class=""socialmediabutton smfacebook"">"
		print "<div id=""fb-root""></div><script>(function(d,s,id){var js, fjs = d.getElementsByTagName(s)[0];if (d.getElementById(id)) return;js = d.createElement(s); js.id = id;js.src = ""//connect.facebook.net/" & buttonlang & "/sdk.js#xfbml=1&version=v2.7"";fjs.parentNode.insertBefore(js, fjs);}(document, 'script', 'facebook-jssdk'));</script>"
		if sbfacebook<>"" then print replace(sbfacebook,"%pageurl%",thisurl) else print "<div class=""fb-like"" data-href=""" & thisurl & """ data-layout=""button_count"" data-action=""like"" data-size=""" & IIfVr(smallsocialbuttons,"small","large") & """ data-show-faces=""false"" data-share=""true""></div>"
		print "</div>"
		elseif layoutoption="linkedin" then
		print "<div class=""socialmediabutton smlinkedin"">" & IIfVs(NOT smallsocialbuttons,"<div class=""smlinkedininner"">")
		print "<script src=""//platform.linkedin.com/in.js"">lang:" & buttonlang & "</script>"
		if sblinkedin<>"" then print replace(sblinkedin,"%pageurl%",thisurl) else print "<script type=""IN/Share"" data-url=""" & thisurl & """ data-counter=""right""></script>"
		print "</div>" & IIfVs(NOT smallsocialbuttons,"</div>")
		elseif layoutoption="twitter" then
		print "<div class=""socialmediabutton smtwitter"">"
		print "<script async src=""https://platform.twitter.com/widgets.js""></script>"
		if sbtwitter<>"" then print replace(sbtwitter,"%pageurl%",thisurl) else print "<a class=""twitter-share-button"" " & IIfVs(storelang<>"en","lang="""&storelang&""" ") & "href=""https://twitter.com/intent/tweet"" data-size=""" & IIfVr(smallsocialbuttons,"default","large") & """ data-url=""" & thisurl & """>Tweet</a>"
		print "</div>"
		elseif layoutoption="pinterest" then
		print "<div class=""socialmediabutton smpinterest"">"
		print "<a data-pin-do=""buttonBookmark""" & IIfVs(NOT smallsocialbuttons," data-pin-tall=""28px""") & " data-pin-save=""true"" href=""https://www.pinterest.com/pin/create/button/""></a>"
		print "<script async defer src=""//assets.pinterest.com/js/pinit.js""></script>"
		print "</div>"
		elseif layoutoption="askaquestion" then
		print "<div class=""socialmediabutton smaskaquestion"">"
		print imageorbutton(imgaskaquestion,xxAskQue,IIfVr(smallsocialbuttons,"sm","lg") & "askaquestion","openEFWindow('"&urlencode(prodid)&"',true)",TRUE)
		print "</div>"
		elseif layoutoption="emailfriend" then
		print "<div class=""socialmediabutton smemailfriend"">"
		print imageorbutton(imgemailfriend,xxEmFrnd,IIfVr(smallsocialbuttons,"sm","lg") & "emailfriend","openEFWindow('"&urlencode(prodid)&"',false)",TRUE)
		print "</div>"
		elseif layoutoption="custom" then
		if sbcustom<>"" then print "<div class=""socialmediabutton socialcustom"">" & sbcustom & "</div>"
		end if
	next
	print "</div>"
end sub
set rs=Server.CreateObject("ADODB.RecordSet")
set rs2=Server.CreateObject("ADODB.RecordSet")
set rs3=Server.CreateObject("ADODB.RecordSet")
set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
thesessionid=getsessionid()
hasextracurrency=FALSE : wantmanufacturer=FALSE
for each layoutoption in customlayoutarray
	if lcase(trim(layoutoption))="currency" then hasextracurrency=TRUE
	if lcase(trim(layoutoption))="manufacturer" then wantmanufacturer=TRUE
	if lcase(trim(layoutoption))="instock" then showinstock=TRUE
next
if hasextracurrency then
	call checkCurrencyRates(currConvUser,currConvPw,currLastUpdate,currRate1,currSymbol1,currRate2,currSymbol2,currRate3,currSymbol3)
else
	currRate1=0 : currRate2=0 : currRate3=0
end if
get_wholesaleprice_sql()
disabledsection=FALSE
psmallimage=""
allimages=""
allgiantimages=""
numallimages=0
numallgiantimages=0
pCustomCSS=""
pSchemaType=0
sSQL="SELECT pId,pSKU,"&getlangid("pName",1)&","&WSP&"pPrice,pSection,pListPrice,pSell,pStockByOpts,pStaticPage,pStaticURL,pInStock,pBackOrder,pExemptions,"&IIfVr(detailslink<>"","'' AS ","")&"pTax,pTotRating,pNumRatings,pOrder,pDateAdded,"&getlangid("pSearchParams",4194304)&",pMinQuant,pCustomCSS,pSchemaType,pCustom1,pCustom2,pCustom3,"&IIfVs(wantmanufacturer,getlangid("scName",131072)&",scID,scURL,")&getlangid("pDescription",2)&","&getlangid("pLongDescription",4)&" FROM products "&IIfVr(wantmanufacturer,"LEFT OUTER JOIN searchcriteria on products.pManufacturer=searchcriteria.scID ","")&"WHERE " & IIfVs(ectsiteid<>"","pSiteID=" & ectsiteid & " AND ") & "pDisplay<>0 AND ("&IIfVr(usepnamefordetaillinks AND trim(explicitid)="",getlangid("pName",1),"pId")&"='"&escape_string(prodid)&"'"&IIfVs(seodetailurls," OR pStaticURL='"&escape_string(prodid)&"'")&")"
rs.open sSQL,cnn,0,1
productindb=NOT rs.EOF
disabledsection=FALSE
if productindb then
	shortdesc=trim(rs(getlangid("pDescription",2))&"")
	longdesc=trim(rs(getlangid("pLongDescription",4))&"")
	origprodid=prodid
	sectionid=rs("pSection")
	prodid=rs("pId")
	pCustomCSS=trim(rs("pCustomCSS")&"")
	pSchemaType=rs("pSchemaType")
	sSQL="SELECT sectionDisabled,topSection FROM sections WHERE sectionID=" & sectionid
	rs2.open sSQL,cnn,0,1
	if NOT rs2.EOF then
		if rs2("sectionDisabled")>minloglevel then disabledsection=TRUE
	end if
	rs2.close
end if
prodlist="'" & escape_string(prodid) & "'"
sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 1 ")&"imageSrc FROM productimages WHERE imageType=0 AND imageProduct='"&escape_string(prodid)&"' ORDER BY imageNumber"&IIfVs(mysqlserver=TRUE," LIMIT 0,1")
rs2.Open sSQL,cnn,0,1
if NOT rs2.EOF then psmallimage=rs2("imageSrc")
rs2.close
sSQL="SELECT imageSrc FROM productimages WHERE imageType=1 AND imageProduct='"&escape_string(prodid)&"' ORDER BY imageNumber"
rs2.Open sSQL,cnn,0,1
if NOT rs2.EOF then allimages=rs2.getrows() : numallimages=UBOUND(allimages,2)+1
rs2.close
sSQL="SELECT imageSrc FROM productimages WHERE imageType=2 AND imageProduct='"&escape_string(prodid)&"' ORDER BY imageNumber"
rs2.Open sSQL,cnn,0,1
if NOT rs2.EOF then allgiantimages=rs2.getrows() : numallgiantimages=UBOUND(allgiantimages,2)+1
rs2.close
if (rs.EOF AND prodid<>giftcertificateid AND prodid<>donationid) OR disabledsection then ' {
	print "<div class=""prodnoexist"" style=""text-align:center;margin:50px"">"&xxSryNA&"</div>"
	if prodid<>"" AND NOT disabledsection then
		if usepnamefordetaillinks AND getget("prod")<>"" then
			sSQL="SELECT "&getlangid("pName",1)&",pStaticPage,pStaticURL FROM products WHERE pID='"&escape_string(getget("prod"))&"'"
			rs2.open sSQL,cnn,0,1
			if NOT rs2.EOF then
				addand="" : newqs=""
				for each objitem in request.querystring
					if objitem<>"prod" then newqs=newqs&addand&objitem&"="&urlencode(getget(objitem)) : addand="&"
				next
				newloc=getfullurl(getdetailsurl(getget("prod"),rs2("pStaticPage"),rs2(getlangid("pName",1)),rs2("pStaticURL"),newqs,pathtohere))
				response.status="301 Moved Permanently"
				response.addheader "Location", newloc
				response.end
			end if
			rs2.close
		end if
		sSQL="SELECT pId,pSection,"&getlangid("sectionName",256)&","&getlangid("sectionurl",2048)&",rootSection,sectionImage FROM products INNER JOIN sections ON products.pSection=sections.sectionID WHERE ("&IIfVr(usepnamefordetaillinks AND trim(explicitid)="",getlangid("pName",1),"pId")&"='"&escape_string(prodid)&"'"&IIfVs(seodetailurls," OR pStaticURL='"&escape_string(prodid)&"'")&")"
		rs2.open sSQL,cnn,0,1
		if NOT rs2.EOF then
			caturl=getcategoryurl(rs2("pSection"),rs2(getlangid("sectionName",256)),rs2(getlangid("sectionurl",2048)),rs2("rootSection"))
			print "<div class=""prodnoexistcat"" style=""text-align:center;margin:50px""><a class=""ectlink"" href=""" & caturl & """>" & xxSimPro & "</a></div>"
			if trim(rs2("sectionImage")&"")<>"" then
				print "<div class=""prodnoexistcatimg"" style=""text-align:center;margin:50px""><a class=""ectlink"" href=""" & caturl & """><img src=""" & rs2("sectionImage") & """ alt=""" & rs2(getlangid("sectionName",256)) & """ /></a></div>"
			end if
		end if
		rs2.close
	end if
	response.status="404 Not Found"
else ' }{
	prodoptions=""
	if pSchemaType=1 then
		isarticle=TRUE
		skuschemaidentifier="author"
		hascustomlayout=TRUE
		if articlepagelayout="" then detailpagelayout="productimage,sku,productname,description,dateadded,previousnext" else hascustomlayout=TRUE : detailpagelayout=articlepagelayout
		customlayoutarray=split(detailpagelayout,",")
		enablecustomerratings=instr(detailpagelayout,"review")>0
		if NOT isempty(ecttabsspecialsarticle) then ecttabsspecials=ecttabsspecialsarticle
		if NOT isempty(defaultdescriptiontabarticle) then defaultdescriptiontab=defaultdescriptiontabarticle
	end if
	if prodid<>giftcertificateid AND prodid<>donationid then ' {
		HTTP_X_ORIGINAL_URL=trim(split(request.servervariables("HTTP_X_ORIGINAL_URL")&"?","?")(0))
		if HTTP_X_ORIGINAL_URL="" then HTTP_X_ORIGINAL_URL=trim(split(request.servervariables("HTTP_X_REWRITE_URL")&"?","?")(0))
		if getget("prod")<>"" AND seodetailurls AND seourlsthrow301 AND (HTTP_X_ORIGINAL_URL="" OR (origprodid<>trim(rs("pStaticURL")&"") AND trim(rs("pStaticURL")&"")<>"")) then
			addand="" : newqs=""
			for each objitem in request.querystring
				if objitem<>"prod" then newqs=newqs&addand&objitem&"="&urlencode(getget(objitem)) : addand="&"
			next
			newloc=getfullurl(getdetailsurl(rs("pId"),rs("pStaticPage"),rs(getlangid("pName",1)),rs("pStaticURL"),newqs,pathtohere))
			response.status="301 Moved Permanently"
			response.addheader "Location", newloc
			response.end
		end if
		if getget("prod")<>"" AND cint(rs("pStaticPage"))<>0 AND redirecttostatic=TRUE then
			response.status="301 Moved Permanently"
			response.addheader "Location", cleanforurl(rs(getlangid("pName",1)))&extension
			response.end
		end if
		hasstaticpage=cint(rs("pStaticPage"))<>0
		tslist=""
		if IsNull(rs("pSection")) then catid=0 else catid=rs("pSection")
		if ectsiteid<>"" AND SESSION("savecatid")<>"" then
			if instr(SESSION("saveprodlist"),rs("pId")&" ")>0 then catid=SESSION("savecatid")
		end if
		if is_numeric(getget("cat")) AND getget("cat")<>"0" then catid=getget("cat")
		if is_numeric(getget("cat")) AND getget("cat")<>"0" then thecatid=getget("cat") else thecatid=""
		thetopts=catid
		topsectionids=catid
		isrootsection=FALSE
		currcaturl=""
		for index=0 to 10
			if cstr(thetopts)=cstr(catalogroot) then
				caturl=storehomeurl
				sSQL="SELECT sectionID,topSection,"&getlangid("sectionName",256)&",rootSection,sectionDisabled,"&IIfVs(languageid<>1,getlangid("sectionurl",2048)&" AS ")&"sectionurl FROM sections WHERE sectionID=" & catalogroot
				rs2.Open sSQL,cnn,0,1
				if NOT rs2.EOF then
					xxHome=rs2(getlangid("sectionName",256))
					if trim(rs2("sectionurl")&"")<>"" then caturl=rs2("sectionurl")
				end if
				rs2.Close
				tslist="<a class=""ectlink"" href="""&caturl&""">"&xxHome&"</a>" & tslist
				exit for
			elseif index=10 then
				tslist="<strong>Loop</strong>" & tslist
			else
				sSQL="SELECT sectionID,topSection,"&getlangid("sectionName",256)&",rootSection,"&IIfVs(languageid<>1,getlangid("sectionurl",2048)&" AS ")&"sectionurl FROM sections WHERE sectionID=" & thetopts
				rs2.Open sSQL,cnn,0,1
				if NOT rs2.EOF then
					if dynamicbreadcrumbs then
						tslist="<div class=""ectbreadcrumb"">&raquo; " & breadcrumbselect(rs2("sectionID"),rs2("topSection")) & "</div>" & tslist
					else
						tslist="<div class=""ectbreadcrumb"">&raquo; <a class=""ectlink"" href=""" & getcategoryurl(rs2("sectionID"),rs2(getlangid("sectionName",256)),rs2("sectionurl"),rs2("rootSection")) & """>" & rs2(getlangid("sectionName",256)) & "</a></div>" & tslist
					end if
					thetopts=rs2("topSection")
					topsectionids=topsectionids & "," & thetopts
				else
					if ectsiteid<>"" then
						thetopts=catalogroot
						tslist=""
					else
						tslist="Top Section Deleted" & tslist
						rs2.Close
					end if
					exit for
				end if
				rs2.Close
			end if
		next
		if dynamicbreadcrumbs AND currcaturl<>"" AND xxAlProd<>"" then tslist=tslist&"<div class=""ectbreadcrumb"">&raquo; <a class=""ectlink"" href=""" & htmlspecials(currcaturl) & """>" & xxAlProd & "</a></div>"
		nextid=""
		previousid=""
		sectionids=getsectionids(catid, false)
		if SESSION("sortby")<>"" then dosortby=SESSION("sortby")
		if dosortby=2 OR dosortby=12 OR dosortby=5 then
		elseif dosortby=14 OR dosortby=15 then
			sSortBy="pSKU"
			sSortValue="'"&escape_string(rs("pSKU"))&"'"
		elseif dosortby=3 OR dosortby=4 then
			sSortBy=TWSP
			sSortValue=rs("pPrice")
		elseif dosortby=6 OR dosortby=7 then
			sSortBy="pOrder"
			sSortValue=rs("pOrder")
		elseif dosortby=8 OR dosortby=9 then
			sSortBy="pDateAdded"
			sSortValue=vsusdate(rs("pDateAdded"))
		else
			sSortBy=getlangid("pName",1)
			sSortValue="'"&escape_string(rs(getlangid("pName",1)))&"'"
		end if
		if dosortby=4 OR dosortby=7 OR dosortby=9 OR dosortby=11 OR dosortby=12 OR dosortby=15 then isdesc=TRUE else isdesc=FALSE
		if nopreviousnextlinks<>TRUE then
			session.LCID=1033
			sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 1 ")&"products.pId,"&getlangid("pName",1)&",pStaticPage,pStaticURL,products.pSection FROM products LEFT JOIN multisections ON products.pId=multisections.pId WHERE (products.pSection IN (" & sectionids & ") OR multisections.pSection IN (" & sectionids & "))"&IIfVr(useStockManagement AND noshowoutofstock, " AND (pInStock>pMinQuant OR pStockByOpts<>0)", "") & IIfVs(ectsiteid<>"", " AND pSiteID=" & ectsiteid) & " AND pDisplay<>0 AND ("&IIfVr(sSortBy<>"","(("&sSortBy&"="&sSortValue&" AND products.pId > '"&escape_string(prodid)&"') OR "&sSortBy&IIfVr(isdesc,"<",">")&sSortValue&")","products.pId "&IIfVr(isdesc,"<",">")&" '"&escape_string(prodid)&"'")&") AND products.pId NOT IN ('"&escape_string(giftcertificateid)&"','"&escape_string(donationid)&"') ORDER BY "&IIfVr(sSortBy<>"",sSortBy&IIfVr(isdesc," DESC,"," ASC,"),"")&"products.pId"&IIfVs(isdesc," DESC")&IIfVs(mysqlserver=TRUE," LIMIT 0,1")
			rs2.Open sSQL,cnn,0,1
			if NOT rs2.EOF then
				nextid=IIfVr(usepnamefordetaillinks,replace(rs2(getlangid("pName",1))," ",detlinkspacechar),rs2("pId"))
				nextidname=rs2(getlangid("pName",1))
				nextidstatic=(cint(rs2("pStaticPage"))<>0)
				nextstaticurl=rs2("pStaticURL")
				nextidcat=rs2("pSection")
			end if
			rs2.Close
			sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 1 ")&"products.pId,"&getlangid("pName",1)&",pStaticPage,pStaticURL,products.pSection FROM products LEFT JOIN multisections ON products.pId=multisections.pId WHERE (products.pSection IN (" & sectionids & ") OR multisections.pSection IN (" & sectionids & "))"&IIfVr(useStockManagement AND noshowoutofstock, " AND (pInStock>pMinQuant OR pStockByOpts<>0)", "") & IIfVs(ectsiteid<>"", " AND pSiteID=" & ectsiteid) & " AND pDisplay<>0 AND ("&IIfVr(sSortBy<>"","(("&sSortBy&"="&sSortValue&" AND products.pId < '"&escape_string(prodid)&"') OR "&sSortBy&IIfVr(isdesc,">","<")&sSortValue&")","products.pId "&IIfVr(isdesc,">","<")&" '"&escape_string(prodid)&"'")&") AND products.pId NOT IN ('"&escape_string(giftcertificateid)&"','"&escape_string(donationid)&"') ORDER BY "&IIfVr(sSortBy<>"",sSortBy&IIfVr(isdesc," ASC,"," DESC,"),"")&"products.pId"&IIfVs(NOT isdesc," DESC")&IIfVs(mysqlserver=TRUE," LIMIT 0,1")
			rs2.Open sSQL,cnn,0,1
			if NOT rs2.EOF then
				previousid=IIfVr(usepnamefordetaillinks,replace(rs2(getlangid("pName",1))," ",detlinkspacechar),rs2("pId"))
				previousidname=rs2(getlangid("pName",1))
				previousidstatic=(cint(rs2("pStaticPage"))<>0)
				previousstaticurl=rs2("pStaticURL")
				previousidcat=rs2("pSection")
			end if
			rs2.Close
			session.LCID=saveLCID
		end if
		extraimages=0
		giantimages=0
		if currencyseparator="" then currencyseparator=" "
		call productdisplayscript(TRUE,TRUE)
		if perproducttaxrate=TRUE AND NOT IsNull(rs("pTax")) then thetax=rs("pTax") else thetax=countryTaxRate
		call updatepricescript()
		if emailfriendoraskquestion then call emailfriendjavascript()
		if magictoolbox="" AND isarray(allgiantimages) AND getget("review")<>"true" then %>
<script>
/* <![CDATA[ */<%
			print "pIX[999]=0;pIM[999]=["
			for index=0 to numallgiantimages-1
				print IIfVs(index>0,",")&"'"&encodeimage(allgiantimages(0,index))&"'"
			next
			print "];" %>
function showgiantimage(){
<%			if giantimageinpopup then %>
document.getElementById('giantimgspan').style.display='';
document.getElementById('mainbodyspan').style.display='none';
displayprodimagenum(999,pIX[0]?pIX[0]:0);
document.getElementById('prodimage999').style.width='auto';
document.getElementById('prodimage999').style.maxWidth='100%';
<%			else %>
document.getElementById('giantimgspan').style.display='';
document.getElementById('mainbodyspan').style.display='none';
displayprodimagenum(999,pIX[0]?pIX[0]:0);
<%			end if %>
}
function hidegiantimage(){
document.getElementById('giantimgspan').style.display='none';
document.getElementById('mainbodyspan').style.display='';
return(false);
}
function showgiantrightleft(doshow){
	document.getElementById('giantimgleft').style.display=document.getElementById('giantimgright').style.display=doshow?'':'none';
}
function displayprodimagenum(theitem,imagenum){
var imlist=pIM[theitem];
pIX[theitem]=imagenum;
if(document.getElementById("prodimage"+theitem)){document.getElementById("prodimage"+theitem).src='';document.getElementById("prodimage"+theitem).src=vsdecimg(imlist[pIX[theitem]]);}
document.getElementById("extraimcnt"+theitem).innerHTML=pIX[theitem]+1;
return false;
}
/* ]]> */
</script>
<%			if NOT nogiantimagepopup then %>
<div id="giantimgspan" style="position:fixed;text-align:center;display:none;padding:0px;background-color:rgba(140,140,150,0.5);width:100%;height:100%;z-index:1200;top:0px;left:0px">
	<div class="giantimgdiv" style="display:flex;background-color:#FFFFFF;width:98%;height:97%;margin:1%;border-radius:8px;box-shadow:5px 5px 2px #666;overflow-x:auto">
<%				if numallgiantimages>1 AND NOT mobilebrowser then %>
		<div id="giantthumbcontainer" class="giantthumbcontainer" style="padding:0px 10px 10px 10px;width:15%">
			<div style="margin:13px"><span class="extraimage extraimagenum" id="extraimcnt999">1</span> <span class="extraimage"><%=xxOf & " " & numallgiantimages%></span></div>
			<div class="giantthumbs" style="border:1px solid grey">
				<div style="text-align:right;margin:4px 4px -5px 0"><img src="images/close.gif" style="cursor:pointer" onclick="document.getElementById('giantthumbcontainer').style.display='none'" /></div>
<%					for index=0 to UBOUND(allgiantimages,2)
						if index<numallimages then imgsrc=allimages(0,index) else imgsrc="images/ectquestionmark.png"
						if index<numallgiantimages then print "<div class=""giantthumb"" style=""margin:10px""><img class=""giantthumb"" style=""width:100px;box-shadow:5px 5px 2px #999"" src=""" & imgsrc & """ alt="""" onclick=""displayprodimagenum(999,"&index&")"" /></div>"
					next %>
			</div>
		</div>
<%				end if %>
		<div class="giantimg" style="margin:3px 10px 4px 4px;flex-grow:1">
			<div class="giantimgheader">
				<div class="giantimgname" style="text-align:center;margin:8px;font-size:1.3em"><a href="#" onclick="return hidegiantimage()"><%=rs(getlangid("pName",1))%></a></div>
				<div class="giantimgclose" style="float:right;position:absolute;top:20px;right:2%"><a href="#" onclick="return hidegiantimage()"><img src="images/close.gif" style="margin:5px 0" alt="<%=xxClsWin%>" /></a></div>
			</div>
			<div style="position:relative;display:inline-block">
				<img class="giantimage allprodimages" id="prodimage999" src="" alt="<%=replace(strip_tags2(rs(getlangid("pName",1))),"""","&quot;")%>" <% if numallgiantimages>1 then print "onmouseover=""showgiantrightleft(true)"" onmouseout=""showgiantrightleft(false)"" " %>/>
<%				if numallgiantimages>1 then %>
				<img id="giantimgright" src="images/ectimageright.png" style="display:none;position:absolute;bottom:40%;right:0;height:20%" onmouseover="this.style.cursor='pointer';showgiantrightleft(true)" onmouseout="showgiantrightleft(false)" onclick="return updateprodimage(999,true);" alt="PREV" />
				<img id="giantimgleft" src="images/ectimageleft.png" style="display:none;position:absolute;bottom:40%;left:0;height:20%" onmouseover="this.style.cursor='pointer';showgiantrightleft(true)" onmouseout="showgiantrightleft(false)" onclick="return updateprodimage(999,false);" alt="NEXT" />
<%				end if %>
			</div>
		</div>
		<div style="clear:both;margin-bottom:10px"></div>
	</div>
</div>
<%			else %>
<div id="giantimgspan" style="width:98%;text-align:center;display:none">
	<div><span class="giantimgname detailname"><%=rs(getlangid("pName",1)) & " </span> <span class=""giantimgback""><a class=""ectlink"" href=""" & detailpageurl(IIfVs(thecatid<>"","cat="&thecatid))&""" onclick=""javascript:return hidegiantimage();"" >" & xxRvBack & "</a></span>" %></div>
	<div class="giantimg" style="margin:0 auto;display:inline-block">
<%				if numallgiantimages>1 then %>
		<div style="text-align:center;margin-bottom:3px"><%=imageorbutton(imgdetailprevimg,xxPrImTx,"giantprevimg","updateprodimage(999,false)",TRUE)%> <span class="extraimage extraimagenum" id="extraimcnt999">1</span> <span class="extraimage"><%=xxOf & " " & numallgiantimages%></span> <%=imageorbutton(imgdetailnextimg,xxNeImTx,"giantnextimg","updateprodimage(999,true)",TRUE)%></div>
<%				end if %>
		<div style="text-align:center"><img id="prodimage999" class="giantimage allprodimages" src="" alt="<%=replace(strip_tags2(rs(getlangid("pName",1))),"""","&quot;")%>" <% if numallgiantimages>1 then print "onclick=""return updateprodimage(999,true);"" onmouseover=""this.style.cursor='pointer'""" %> style="margin:0px;" /></div>
<%				if numallgiantimages>1 then %>
		<div><%=imageorbutton(imgdetailprevimg,xxPrImTx,"giantprevimg","updateprodimage(999,false)",TRUE) & "&nbsp;" & imageorbutton(imgdetailnextimg,xxNeImTx,"giantnextimg","updateprodimage(999,true)",TRUE)%></div>
<%				end if %>
	</div>
</div>
<%			end if
		end if
	else
		proddetailtopbuybutton=FALSE
	end if ' }
	optionshavestock=TRUE
	optjs=""
	if isarray(prodoptions) AND request("review")="" then
		if usedetailbodyformat=1 OR usedetailbodyformat="" then
			optionshtml=displayproductoptions("<strong><span class=""detailoption"">","</span></strong>",optdiff,thetax,TRUE,hasmultipurchase,optjs)
		else
			optionshtml=displayproductoptions("<span class=""detailoption"">","</span>",optdiff,thetax,TRUE,hasmultipurchase,optjs)
		end if
	end if
	if prodid=giftcertificateid OR prodid=donationid then
		isinstock=TRUE : isbackorder=FALSE
	else
		if useStockManagement then
			if cint(rs("pStockByOpts"))<>0 then isinstock=optionshavestock else isinstock=rs("pInStock")>rs("pMinQuant")
		else
			isinstock=cint(rs("pSell"))<>0
		end if
		isbackorder=NOT isinstock AND cint(rs("pBackOrder"))<>0
	end if
	theuagent=lcase(request.servervariables("HTTP_USER_AGENT"))
	iswebbot=FALSE
	if instr(theuagent,"baiduspider")>0 OR instr(theuagent,"bingbot")>0 OR instr(theuagent,"crawler")>0 OR instr(theuagent,"duckduckbot")>0 OR instr(theuagent,"exabot")>0 OR instr(theuagent,"ezooms")>0 OR instr(theuagent,"facebook")>0 OR instr(theuagent,"googlebot")>0 OR instr(theuagent,"gulliver")>0 OR instr(theuagent,"ia_archiver")>0 OR instr(theuagent,"infoseek")>0 OR instr(theuagent,"inktomi")>0 OR instr(theuagent,"mj12bot")>0 OR instr(theuagent,"scooter")>0 OR instr(theuagent,"sogou")>0 OR instr(theuagent,"speedy spider")>0 OR instr(theuagent,"yahoo!")>0 OR instr(theuagent,"yandexbot")>0 then iswebbot=TRUE
	if iswebbot then
		recentlyviewed=FALSE
	else
		sSQL="UPDATE products SET pPopularity=pPopularity+1 WHERE pID='"&escape_string(prodid)&"'"
		ect_query(sSQL)
	end if
	if recentlyviewed=TRUE AND NOT (prodid=giftcertificateid OR prodid=donationid) then
		tcnt=NULL
		if numrecentlyviewed="" then numrecentlyviewed=6
		sSQL="DELETE FROM recentlyviewed WHERE rvDate<" & vsusdate(Date()-3)
		ect_query(sSQL)
		sSQL="SELECT rvID FROM recentlyviewed WHERE rvProdID='"&escape_string(prodid)&"' AND " & IIfVr(SESSION("clientID")<>"", "rvCustomerID="&replace(SESSION("clientID"),"'",""), "(rvCustomerID=0 AND rvSessionID='"&thesessionid&"')")
		rs2.open sSQL,cnn,0,1
		if rs2.EOF then
			sSQL="INSERT INTO recentlyviewed (rvProdID,rvProdName,rvProdSection,rvProdURL,rvSessionID,rvCustomerID,rvDate) VALUES ('"&escape_string(prodid)&"','"&escape_string(rs(getlangid("pName",1)))&"',"&IIfVr(catid<>"",catid,"0")&",'"&escape_string(detailpageurl(IIfVr(thecatid<>"","cat="&thecatid,"")))&"','"&thesessionid&"',"&IIfVr(SESSION("clientID")<>"",SESSION("clientID"), 0)&","&vsusdatetime(Now())&")"
			ect_query(sSQL)
		else
			sSQL="UPDATE recentlyviewed SET rvDate="&vsusdatetime(Now())&" WHERE rvID="&rs2("rvID")
			ect_query(sSQL)
		end if
		rs2.Close
		sSQL="SELECT COUNT(*) AS tcnt FROM recentlyviewed WHERE " & IIfVr(SESSION("clientID")<>"", "rvCustomerID="&replace(SESSION("clientID"),"'",""), "(rvCustomerID=0 AND rvSessionID='"&thesessionid&"')")
		rs2.open sSQL,cnn,0,1
		if NOT rs2.EOF then tcnt=rs2("tcnt")
		rs2.Close
		if NOT isnull(tcnt) then
			if tcnt>numrecentlyviewed then
				sSQL="SELECT rvID,MIN(rvDate) FROM recentlyviewed WHERE " & IIfVr(SESSION("clientID")<>"", "rvCustomerID="&replace(SESSION("clientID"),"'",""), "(rvCustomerID=0 AND rvSessionID='"&thesessionid&"')")&" GROUP BY rvID"
				rs2.open sSQL,cnn,0,1
				if NOT rs2.EOF then
					ect_query("DELETE FROM recentlyviewed WHERE rvID="&rs2("rvID"))
				end if
				rs2.close
			end if
		end if
	end if
	if usecsslayout then print "<div id=""mainbodyspan"" class=""proddetail " & trim(pCustomCSS & " " & prodid) & IIfVr(getget("review")="true"," detailaddreview""",IIfVr(noschemamarkup,"""",""" itemscope itemtype=""http://schema.org/" & IIfVr(isarticle,"Article","Product") & """")) & "><link itemprop=""mainEntityOfPage"" href=""" & detailpageurl("") & """>" else print "<table id=""mainbodyspan"" class=""proddetail"" border=""0"" cellspacing=""0"" cellpadding=""0"" width=""98%"" align=""center""" & IIfVs(NOT noschemamarkup,""" itemscope itemtype=""http://schema.org/" & IIfVr(isarticle,"Article","Product") & """") & "><tr><td width=""100%"">"
	if getget("review")<>"true" then print "<form method=""post"" id=""ectform0"" action="""&IIfVr(prodid=giftcertificateid OR prodid=donationid,request.servervariables("URL")&IIfVr(request.servervariables("QUERY_STRING")<>"", "?" & replace(strip_tags2(request.servervariables("QUERY_STRING")),"""",""), ""), "cart"&extension)&""" onsubmit=""return formvalidator"&Count&"(this)"">"
	if NOT hascustomlayout AND (isempty(showcategories) OR showcategories=TRUE) AND NOT (prodid=giftcertificateid OR prodid=donationid) then
		call pddprodnavigation()
		call pddcheckoutbutton()
	end if
	alldiscounts=""
	if nowholesalediscounts=TRUE AND SESSION("clientUser")<>"" then
		if ((SESSION("clientActions") AND 8)=8) OR ((SESSION("clientActions") AND 16)=16) then noshowdiscounts=TRUE
	end if
	if noshowdiscounts<>TRUE AND prodid<>giftcertificateid AND prodid<>donationid then
		if (rs("pExemptions") AND 64)<>64 then
			Session.LCID=1033
			tdt=Date()
			attributelist=""
			sSQL="SELECT mSCscID FROM multisearchcriteria WHERE mSCpID='"&escape_string(rs("pID"))&"'"
			rs2.open sSQL,cnn,0,1
			do while NOT rs2.EOF
				attributelist=attributelist&rs2("mSCscID")&" "
				rs2.movenext
			loop
			rs2.close
			attributelist=replace(trim(attributelist)," ","','")
			sSQL="SELECT DISTINCT "&getlangid("cpnName",1024)&" FROM coupons LEFT OUTER JOIN cpnassign ON coupons.cpnID=cpnassign.cpaCpnID WHERE cpnNumAvail>0 AND cpnStartDate<=" & vsusdate(tdt)&" AND cpnEndDate>=" & vsusdate(tdt)&" AND cpnIsCoupon=0 AND " & _
				"((cpnSitewide=1 OR cpnSitewide=2) OR (cpnSitewide=0 AND cpaType=2 AND cpaAssignment='"&escape_string(rs("pID"))&"') " & _
				"OR ((cpnSitewide=0 OR cpnSitewide=3) AND ((cpaType=1 AND cpaAssignment IN ('"&Replace(topsectionids,",","','")&"'))" & IIfVs(attributelist<>""," OR (cpaType=3 AND cpaAssignment IN ('"&attributelist&"'))") & ")))" & _
				" AND ((cpnLoginLevel>=0 AND cpnLoginLevel<="&minloglevel&") OR (cpnLoginLevel<0 AND -1-cpnLoginLevel="&minloglevel&"))"
			if (rs("pExemptions") AND 16)=16 then sSQL=sSQL&" AND cpnType<>0"
			Session.LCID=saveLCID
			rs2.Open sSQL,cnn,0,1
			do while NOT rs2.EOF
				alldiscounts=alldiscounts & "<div>" & rs2(getlangid("cpnName",1024)) & "</div>"
				rs2.MoveNext
			loop
			rs2.Close
		end if
	end if
	if enablecustomerratings AND getpost("review")="true" then ' {
		hitlimit=FALSE
		print "<table border=""0"" cellspacing=""2"" cellpadding=""2"" width=""100%"" align=""center"">"
		sSQL="SELECT COUNT(*) as thecount FROM ratings WHERE rtDate=" & vsusdate(Date) & " AND rtIPAddress='" & left(request.servervariables("REMOTE_ADDR"), 32) & "'"
		rs2.open sSQL,cnn,0,1
		if NOT rs2.EOF then
			if dailyratinglimit="" then dailyratinglimit=10
			if NOT isnull(rs2("thecount")) then
				if rs2("thecount")>dailyratinglimit then hitlimit=TRUE
			end if
		end if
		rs2.Close
		theip=trim(replace(left(request.servervariables("REMOTE_ADDR"), 48),"'",""))
		if theip="" then theip="none"
		if theip="none" then
			sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 1 ")&"dcid FROM ipblocking"&IIfVs(mysqlserver=TRUE," LIMIT 0,1")
		else
			sSQL="SELECT dcid FROM ipblocking WHERE (dcip1=" & ip2long(theip) & " AND dcip2=0) OR (dcip1 <= " & ip2long(theip) & " AND " & ip2long(theip) & "<=dcip2 AND dcip2<>0)"
		end if
		rs2.Open sSQL,cnn,0,1
		if NOT rs2.EOF then hitlimit=TRUE
		rs2.Close
		referer=request.servervariables("HTTP_REFERER")
		host=request.servervariables("HTTP_HOST")
		if instr(referer, host)=0 then
			print "<tr><td align=""center"">Sorry but your review could not be sent at this time.</td></tr>"
		elseif hitlimit then
			print "<tr><td><div class=""nosearchresults"">"&xxRvLim&"</div></td></tr>"
		elseif onlyclientratings AND SESSION("clientID")="" then
			print "<tr><td align=""center"">Only logged in customers can review products.</td></tr>"
		elseif is_numeric(getpost("ratingstars")) AND is_numeric(getpost("rfsectgrp1")) AND is_numeric(getpost("rfsectgrp2")) AND getpost("ratingstars")<>"" AND getpost("reviewposter")<>"" AND getpost("reviewheading")<>"" then
			if int(getpost("ratingstars"))=int(getpost("rfsectgrp1")) AND int(getpost("rfsectgrp2"))=len(request.form("reviewposter")) then
				sSQL="INSERT INTO ratings (rtProdID,rtRating,rtPosterName,rtHeader,rtIPAddress,rtApproved,rtLanguage,rtDate,rtPosterLoginID,rtComments) VALUES (" & _
					"'" & escape_string(strip_tags2(prodid)) & "'," & IIfVr(is_numeric(getpost("ratingstars")), int(getpost("ratingstars"))*2, 0) & ",'" & escape_string(strip_tags2(getpost("reviewposter"))) & "','" & escape_string(strip_tags2(getpost("reviewheading"))) & "','" & escape_string(strip_tags2(left(request.servervariables("REMOTE_ADDR"), 32))) & "',0,"
				if languageid<>"" then sSQL=sSQL & (int(languageid)-1) & "," else sSQL=sSQL & "0,"
				sSQL=sSQL & vsusdate(Date) & "," & IIfVr(SESSION("clientID")<>"", SESSION("clientID"), 0) & ",'" & escape_string(strip_tags2(getpost("reviewcomments"))) & "')"
				ect_query(sSQL)
				if (adminEmailConfirm AND 8)=8 then
					emailmessage="There has been a new customer review at your store: " & emlNl & _
						"Product ID: " & strip_tags2(prodid) & emlNl & _
						"Rating: " & strip_tags2(getpost("ratingstars")) & emlNl & _
						"Poster: " & strip_tags2(getpost("reviewposter")) & emlNl & _
						"IP: " & strip_tags2(left(request.servervariables("REMOTE_ADDR"), 32)) & emlNl & _
						"Heading: " & strip_tags2(getpost("reviewheading")) & emlNl & _
						"Comments: " & strip_tags2(getpost("reviewcomments")) & emlNl
					call DoSendEmailEO(emailAddr,emailAddr,"","New Customer Review",emailmessage,emailObject,themailhost,theuser,thepass)
				end if
			else
				xxRvThks="Error, I'm sorry but your review could not be recorded at this time."
			end if
			print "<tr><td align=""center"">&nbsp;<br />&nbsp;<br />"&xxRvThks&"<br />&nbsp;<br />&nbsp;"
			print xxRvRet&" <a class=""ectlink"" href="""&detailpageurl(IIfVr(thecatid<>"","cat="&thecatid,""))&""">" & xxClkHere & "</a>"
			print "<br />&nbsp;<br />&nbsp;"
			print "<script>setTimeout(function(){document.location='" & jsescape(detailpageurl(IIfVr(thecatid<>"","cat="&thecatid,""))) & "'},10000)</script>"
			print "</td></tr>"
		end if
		print "</table>"
	elseif enablecustomerratings AND getget("review")="all" then ' }{
		print IIfVr(usecsslayout, "<div class=""reviews"">", "<table border=""0"" cellspacing=""2"" cellpadding=""2"" width=""100%"" align=""center""><tr><td>")
		if psmallimage<>"" then
			print "<img align=""middle"" id=""prodimage0"" class=""prodimage detailreviewimage allprodimages"" src="""&replace(psmallimage,"%s","")&""" alt="""&replace(strip_tags2(rs(getlangid("pName",1))),"""","&quot;")&""" />&nbsp;"
		end if
		print "<span class=""reviewsforprod"">"&xxRvRevP&" - </span><span class=""reviewprod""" & IIfVs(NOT noschemamarkup," itemprop=""name""") & ">" & rs(getlangid("pName",1)) & "</span> <span class=""reviewback"">(<a class=""ectlink"" href="""&detailpageurl(IIfVr(thecatid<>"","cat="&thecatid,""))&""">" & xxRvBack & "</a>)</span><br />&nbsp;</td></tr>"
		sSQL="SELECT rtID,rtRating,rtPosterName,rtHeader,rtDate,rtComments FROM ratings WHERE rtApproved<>0 AND rtProdID='"&escape_string(prodid)&"'"
		if ratingslanguages<>"" then sSQL=sSQL & " AND rtLanguage+1 IN ("&ratingslanguages&")" else if languageid<>"" then sSQL=sSQL & " AND rtLanguage="&(int(languageid)-1) else sSQL=sSQL & " AND rtLanguage=0"
		if getget("ro")="1" then
			sSQL=sSQL & " ORDER BY rtRating DESC"
		elseif getget("ro")="2" then
			sSQL=sSQL & " ORDER BY rtRating"
		elseif getget("ro")="3" then
			sSQL=sSQL & " ORDER BY rtDate"
		else
			sSQL=sSQL & " ORDER BY rtDate DESC"
		end if
		print showreviews(sSQL,TRUE)
		print IIfVr(usecsslayout, "</div>", "</table>")
	elseif enablecustomerratings AND getget("review")="true" then ' }{
		print IIfVr(usecsslayout, "<div class=""reviewprod"">", "<table border=""0"" cellspacing=""2"" cellpadding=""2"" width=""100%"" align=""center""><tr><td>")
		print "<span class=""reviewing"">"&xxRvAreR&" - </span><span class=""reviewprod""" & IIfVs(NOT noschemamarkup," itemprop=""name""") & ">" & rs(getlangid("pName",1)) & "</span> <span class=""reviewback"">(<a class=""ectlink"" href="""&detailpageurl(IIfVr(thecatid<>"","cat="&thecatid,""))&""">" & xxRvBack & "</a>)</span>"
		print IIfVr(usecsslayout, "</div>", "<br />&nbsp;</td></tr></table>")
	elseif prodid=giftcertificateid OR prodid=donationid then
		isincluded=TRUE %>
<!--#include file="incspecials.asp"-->
<%	elseif usedetailbodyformat=1 OR usedetailbodyformat="" then ' }{
%>          <table width="100%" border="0" cellspacing="3" cellpadding="3">
              <tr> 
                <td width="100%" colspan="4" class="detail"> 
<%		if showproductid=TRUE then print "<div class=""detailid""><strong>" & xxPrId & ":</strong> " & IIfVs(NOT noschemamarkup,"<span itemprop=""productID"">") & rs("pID") & IIfVs(NOT noschemamarkup,"</span>") & "</div>"
		if xxManLab<>"" then
			if NOT IsNull(rs(getlangid("scName",131072))) then print "<div class=""detailmanufacturer""><strong>" & xxManLab & ":</strong> " & IIfVs(NOT noschemamarkup,"<span itemprop=""manufacturer"">") & rs(getlangid("scName",131072)) & IIfVs(NOT noschemamarkup,"</span>") & "</div>"
		end if
		if showproductsku<>"" AND trim(rs("pSKU")&"")<>"" then print "<div class=""detailsku""><strong>" & showproductsku & ":</strong> " & IIfVs(NOT noschemamarkup,"<span itemprop=""" & skuschemaidentifier & """>") & rs("pSKU") & IIfVs(NOT noschemamarkup,"</span>") & "</div>"
		print sstrong & "<div class=""detailname""><h1"&IIfVs(NOT noschemamarkup," itemprop=""name""") & ">" & rs(getlangid("pName",1))&"</h1>"&xxDot
		if alldiscounts<>"" then print " <span class=""discountsapply detaildiscountsapply"">"&xxDsApp&"</span></div>"&estrong&"<div class=""detaildiscounts"">" & alldiscounts & "</div>" else print "</div>" & estrong
		if useStockManagement AND showinstock=TRUE AND (rs("pInStock")<=clng(stockdisplaythreshold) OR stockdisplaythreshold="") then if cint(rs("pStockByOpts"))=0 then print "<div class=""detailinstock""><strong>" & xxInStoc & ":</strong> " & vrmax(0,rs("pInStock")) & "</div>" %>
                </td>
              </tr>
              <tr><td width="100%" colspan="4" align="center" class="detailimage allprodimages"><% showdetailimages() %></td></tr>
              <tr> 
                <td width="100%" colspan="4" class="detaildescription"><%
		if longdesc<>"" then
			print "<div class=""detaildescription"""&displaytabs(longdesc)&"</div>"
		elseif shortdesc<>"" then
			print "<div class=""detaildescription""" & IIfVs(NOT noschemamarkup," itemprop=""description""") & ">"&shortdesc&"</div>"
		else
			print "&nbsp;"
		end if
		print "&nbsp;<br />"
		totprice=rs("pPrice")
		if isarray(prodoptions) then
			totprice=totprice + optdiff
			if optionshtml<>"" then print "<div class=""detailoptions"">" & optionshtml &"</div>"
		end if
		call displayformvalidator()
		if optjs<>"" then
			print "<script>/* <![CDATA[ */"&optjs&"/* ]]> */</script>"
			if mustincludepopcalendar AND NOT hasincludedpopcalendar then print "<script>var ectpopcalisproducts=1;" & ectpopcalendarjs & IIfVs(storelang<>"en" AND storelang<>"","var ectpopcallang='"&storelang&"'") & vbLf & "</script><script src=""vsadmin/popcalendar.js""></script>" : hasincludedpopcalendar=TRUE
		end if %></td>
              </tr>
			</table>
			<table width="100%" border="0" cellspacing="3" cellpadding="3">
              <tr>
			    <td><% if socialmediabuttons<>"" then call pddsocialmedia() else print "&nbsp;" %></td>
                <td align="center"><%
		if noprice=TRUE then
			print "&nbsp;"
		else
			if cdbl(rs("pListPrice"))<>0.0 then
				plistprice=IIfVr(showtaxinclusive=2 AND (rs("pExemptions") AND 2)<>2,rs("pListPrice")+(rs("pListPrice")*thetax/100.0), rs("pListPrice"))
				if yousavetext<>"" then yousaveprice=plistprice-IIfVr(showtaxinclusive=2 AND (rs("pExemptions") AND 2)<>2,rs("pPrice")+(rs("pPrice")*thetax/100.0),rs("pPrice")) else yousaveprice=0
				print "<div class=""detaillistprice"" id=""listdivec" & Count & """" & IIfVs(yousaveprice<=0," style=""display:none""") & ">" & Replace(xxListPrice, "%s", FormatEuroCurrency(plistprice)) & IIfVs(yousavetext<>"" AND yousaveprice>0,replace(yousavetext,"%s",FormatEuroCurrency(yousaveprice))) & "</div>"
			end if
			displayprice=IIfVr(showtaxinclusive=2 AND (rs("pExemptions") AND 2)<>2, totprice+(totprice*thetax/100.0), totprice)
			print "<div class=""detailprice""" & IIfVs(NOT noschemamarkup," itemprop=""offers"" itemscope itemtype=""http://schema.org/Offer""><meta itemprop=""priceCurrency"" content="""&countryCurrency&""">") & "<strong>" & xxPrice&IIfVs(xxPrice<>"",":") & "</strong> <span class=""price"" id=""pricediv" & Count & """"&IIfVs(totprice<>0 AND NOT noschemamarkup," itemprop=""price"" content="""&FormatNumberUS(displayprice,2,-1,0,0)&"""")&">" & IIfVr(totprice=0 AND pricezeromessage<>"",pricezeromessage,FormatEuroCurrency(displayprice)) & "</span><link itemprop=""url"" href=""" & detailpageurl("") & """> "
			if showtaxinclusive=1 AND (rs("pExemptions") AND 2)<>2 then print "<span class=""inctax"" id=""taxmsg" & Count & """" & IIfVs(totprice=0, " style=""display:none""") & ">" & Replace(ssIncTax,"%s", "<span id=""pricedivti" & Count & """>" & IIfVr(totprice=0, "-", FormatEuroCurrency(totprice+(totprice*thetax/100.0))) & "</span> ") & "</span>"
			call schemaconditionavail()
			print "</div>"
			extracurr=""
			if currRate1<>0 AND currSymbol1<>"" then extracurr=replace(currFormat1, "%s", FormatNumber(totprice*currRate1, checkDPs(currSymbol1))) & currencyseparator
			if currRate2<>0 AND currSymbol2<>"" then extracurr=extracurr & replace(currFormat2, "%s", FormatNumber(totprice*currRate2, checkDPs(currSymbol2))) & currencyseparator
			if currRate3<>0 AND currSymbol3<>"" then extracurr=extracurr & replace(currFormat3, "%s", FormatNumber(totprice*currRate3, checkDPs(currSymbol3)))
			if showquantitypricing AND NOT hascustomlayout then print pddquantitypricing()
			if extracurr<>"" then print "<div class=""detailcurrency""><span class=""extracurr"" id=""pricedivec" & Count & """>" & IIfVr(totprice=0, "", extracurr) & "</span></div>"
		end if %>
				</td> 
                <td align="right">
<%		if nobuyorcheckout=TRUE then
			print "&nbsp;"
		else
			if totprice=0 AND nosellzeroprice=TRUE then
				print "&nbsp;"
			elseif isinstock OR isbackorder then
				call writehiddenvar("id", rs("pID"))
				call writehiddenvar("mode", "add")
				if wishlistondetail then call writehiddenvar("listid", "")
				if showquantondetail AND hasmultipurchase=0 then print "<table><tr><td align=""center"">" & quantitymarkup(FALSE,Count,TRUE,"",FALSE) & "</td><td align=""center"">"
				if isbackorder then
					if usehardaddtocart then print imageorsubmit(imgbackorderbutton,xxBakOrd,"buybutton backorder detailbuybutton detailbackorder") else print imageorbutton(imgbackorderbutton,xxBakOrd,"buybutton backorder detailbuybutton detailbackorder","subformid("&Count&",'','')",TRUE)
				else
					if custombuybutton<>"" then
						print custombuybutton
					else
						if usehardaddtocart then print imageorsubmit(imgbuybutton,xxAddToC,"buybutton detailbuybutton") else print imageorbutton(imgbuybutton,xxAddToC,"buybutton detailbuybutton","subformid("&Count&",'','')",TRUE)
					end if
				end if
				if wishlistondetail then print "<div class=""wishlistcontainer detailwishlist"">" & imageorlink(imgaddtolist,xxAddLis,"","gtid="&Count&";return displaysavelist(this,event,window)",TRUE) & "</div>"
				if showquantondetail AND hasmultipurchase=0 then print "</td></tr></table>"
			else
				if notifybackinstock then
					print "<div class=""notifystock detailnotifystock"">" & imageorbutton(imgnotifyinstock,xxNotBaS,"notifystock detailnotifystock","return notifyinstock(false,'"&replace(rs("pID"),"'","\'")&"','"&replace(rs("pID"),"'","\'")&"',"&IIfVr(cint(rs("pStockByOpts"))<>0 AND NOT optionshavestock,"-1","0")&")", TRUE) & "</div>"
				else
					print "<div class=""outofstock detailoutofstock"">" & sstrong & xxOutStok & estrong & "</div>"
				end if
			end if
		end if %></td>
			  </tr>
<%		if previousid<>"" OR nextid<>"" then
			print "<tr><td align=""center"" colspan=""4"" class=""pagenumbers""><p class=""pagenumbers"">&nbsp;<br />"
			call writepreviousnextlinks()
			print "</p></td></tr>"
		end if %> 
			</table>
<%	else ' }{ if usedetailbodyformat=2/3/4
		totprice=rs("pPrice")
		if isarray(prodoptions) then
			totprice=totprice + optdiff
		end if
		hasformvalidator=FALSE
		atcmu=""
		call pddoptions()
		call pddaddtocart()
		for each layoutoption in customlayoutarray
			layoutoption=lcase(trim(layoutoption))
			if layoutoption="minquantity" then
				if rs("pMinQuant")>0 then print "<div class=""detailminquant"">" & replace(xxMinQua,"%quant%",rs("pMinQuant")+1) & "</div>"
			elseif layoutoption="navigation" then
				call pddprodnavigation()
			elseif layoutoption="checkoutbutton" then
				call pddcheckoutbutton()
			elseif layoutoption="productimage" then
				call pddproductimage()
			elseif layoutoption="productid" then
				call pddproductid()
			elseif layoutoption="manufacturer" then
				call pddmanufacturer(FALSE)
			elseif layoutoption="manufacturerlink" then
				call pddmanufacturer(TRUE)
			elseif layoutoption="sku" then
				call pddsku()
			elseif layoutoption="productname" then
				call pdddetailname()
			elseif layoutoption="discounts" then
				call pdddiscounts()
			elseif layoutoption="instock" then
				call pddinstock()
			elseif layoutoption="shortdescription" then
				call pddshortdescription()
			elseif layoutoption="description" then
				call pdddescription()
			elseif left(layoutoption,13)="contentregion" then
				call pddcontentregion(right(layoutoption,len(layoutoption)-13))
			elseif layoutoption="catcontentregion" then
				call pddcatcontentregion()
			elseif layoutoption="listprice" then
				call pddlistprice()
			elseif layoutoption="price" then
				call pddprice()
			elseif layoutoption="currency" then
				call pddextracurrency()
			elseif layoutoption="options" then
				hasformvalidator=TRUE
				print optionshtml
			elseif layoutoption="addtocartquant" then
				print "<div class=""addtocartquant detailaddtocartquant"">"
				call pddquantity(FALSE)
				print atcmu & "</div>"
			elseif layoutoption="quantity" then
				call pddquantity(TRUE)
			elseif layoutoption="addtocart" then
				print atcmu
			elseif layoutoption="previousnext" then
				call pddpreviousnext()
			elseif layoutoption="emailfriend" then
				call pddemailfriend()
			elseif layoutoption="reviews" then
				call pddreviews()
			elseif layoutoption="sreviews" then
				call pddsreviews()
			elseif layoutoption="reviewstars" then
				call pddreviewstars(TRUE)
			elseif layoutoption="reviewstarslarge" then
				call pddreviewstars(FALSE)
			elseif layoutoption="searchwords" then
				call pddsearchwords()
			elseif layoutoption="quantitypricing" then
				print pddquantitypricing()
			elseif layoutoption="custom1" then
				call pddcustom(1,detailcustomlabel1)
			elseif layoutoption="custom2" then
				call pddcustom(2,detailcustomlabel2)
			elseif layoutoption="custom3" then
				call pddcustom(3,detailcustomlabel3)
			elseif layoutoption="dateadded" then
				call pdddateadded()
			elseif layoutoption="socialmedia" then
				call pddsocialmedia()
			elseif left(layoutoption,1)="<" then
				print layoutoption
			else
				print "UNKNOWN LAYOUT OPTION:"&layoutoption&"<br />"
			end if
		next
		if isarticle then print publisher
		if NOT hasformvalidator then
			prodoptions="" : optjs="" : defimagejs=""
			totprice=rs("pPrice")
			call displayformvalidator()
			if optjs<>"" then
				print "<script>/* <![CDATA[ */"&optjs&"/* ]]> */</script>"
				if mustincludepopcalendar AND NOT hasincludedpopcalendar then print "<script>var ectpopcalisproducts=1;" & ectpopcalendarjs & IIfVs(storelang<>"en" AND storelang<>"","var ectpopcallang='"&storelang&"'") & vbLf & "</script><script src=""vsadmin/popcalendar.js""></script>" : hasincludedpopcalendar=TRUE
			end if
		end if
	end if ' } usedetailbodyformat
	if getget("review")<>"true" then print "</form>"
	if NOT usecsslayout then print "</td></tr>"
	if NOT hascustomlayout OR request("review")<>"" then call pddreviews()
	if usecsslayout then print "</div>" else print "</table>"
end if ' } rs.EOF
rs.close
set rs=nothing
set rs2=nothing
set rs3=nothing
cnn.close
set cnn=nothing
if defimagejs<>"" then print "<script>"&defimagejs&"</script>"
%>