<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
Dim sSQL,rs,alldata,success,cnn,rowcounter,allsections,errmsg,prodoptions,aFields(6),dorefresh,thecat
set prregexp=new RegExp
prregexp.ignorecase=TRUE
prregexp.global=TRUE
if maxprodoptions="" then maxprodoptions=15
allcoupon="" : pidlist="" : rid="" : currentattribute="" : currentoption="" : currentdiscount="" : currentsection=""
resultcounter=0
dynamicadminmenus=TRUE
if lcase(adminencoding)="iso-8859-1" then raquo="»" else raquo=">"
if admincustomlabel1="" then admincustomlabel1="Custom 1"
if admincustomlabel2="" then admincustomlabel2="Custom 2"
if admincustomlabel3="" then admincustomlabel3="Custom 3"
if seodetailurls then usepnamefordetaillinks=TRUE
sub writemenulevel(id,itlevel)
	Dim wmlindex
	if itlevel<10 then
		for wmlindex=0 TO ubound(alldata,2)
			if alldata(2,wmlindex)=id then
				print "<option value='"&alldata(0,wmlindex)&"'"
				if thecat=alldata(0,wmlindex) then print " selected=""selected"">" else print ">"
				for index=0 to itlevel-2
					print raquo & " "
				next
				print htmldisplay(alldata(1,wmlindex))&"</option>" & vbCrLf
				if alldata(3,wmlindex)=0 then call writemenulevel(alldata(0,wmlindex),itlevel+1)
			end if
		next
	end if
end sub
success=true
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs2=Server.CreateObject("ADODB.RecordSet")
Set rs3=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
Session.LCID=1033
' usesshipweight=(shipType=2 OR shipType=3 OR shipType=4 OR shipType=6 OR shipType=7 OR adminIntShipping=2 OR adminIntShipping=3 OR adminIntShipping=4 OR adminIntShipping=6 OR adminIntShipping=7)
usesshipweight=TRUE
usesflatrate=(shipType=1 OR adminIntShipping=1)
dorefresh=FALSE
if htmlemails=TRUE then emlNl="<br />"&vbCrLf else emlNl=vbCrLf
sub dodeleteprod(pid)
	sSQL="DELETE FROM pricebreaks WHERE pbProdID='" & escape_string(pid) & "'"
	ect_query(sSQL)
	sSQL="DELETE FROM cpnassign WHERE cpaType=2 AND cpaAssignment='" & escape_string(pid) & "'"
	ect_query(sSQL)
	sSQL="DELETE FROM products WHERE pID='" & escape_string(pid) & "'"
	ect_query(sSQL)
	sSQL="DELETE FROM prodoptions WHERE poProdID='" & escape_string(pid) & "'"
	ect_query(sSQL)
	sSQL="DELETE FROM multisections WHERE pID='" & escape_string(pid) & "'"
	ect_query(sSQL)
	sSQL="DELETE FROM multisearchcriteria WHERE mSCpID='" & escape_string(pid) & "'"
	ect_query(sSQL)
	sSQL="DELETE FROM ratings WHERE rtProdID='" & escape_string(pid) & "'"
	ect_query(sSQL)
	sSQL="DELETE FROM relatedprods WHERE rpProdID='" & escape_string(pid) & "' OR rpRelProdID='" & escape_string(pid) & "'"
	ect_query(sSQL)
	sSQL="DELETE FROM notifyinstock WHERE nsProdID='" & escape_string(pid) & "' OR nsTriggerProdID='" & escape_string(pid) & "'"
	ect_query(sSQL)
	sSQL="DELETE FROM productimages WHERE imageProduct='" & escape_string(pid) & "'"
	ect_query(sSQL)
	sSQL="DELETE FROM productpackages WHERE pID='" & escape_string(pid) & "'"
	ect_query(sSQL)
end sub
sub notifyallstock()
	allprods=""
	sSQL="SELECT DISTINCT nsTriggerProdID FROM notifyinstock INNER JOIN products ON notifyinstock.nsTriggerProdID=products.pID WHERE " & IIfVr (useStockManagement,"pInStock>0","pSell<>0") & " AND nsOptID=0"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		allprods=rs.getrows()
	end if
	rs.close
	if isarray(allprods) then
		for index=0 to UBOUND(allprods,2)
			call checknotifystock(allprods(0,index))
		next
	end if
	if useStockManagement then
		allprods=""
		sSQL="SELECT DISTINCT nsOptID FROM notifyinstock INNER JOIN options ON notifyinstock.nsOptID=options.optID WHERE optStock>0 AND nsOptID<>0"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			allprods=rs.getrows()
		end if
		rs.close
		if isarray(allprods) then
			for index=0 to UBOUND(allprods,2)
				call checknotifystockoption(allprods(0,index))
			next
		end if
	end if
end sub
sub checknotifystockoption(theoid)
	if useStockManagement AND notifybackinstock then
		sSQL="SELECT "&getlangid("notifystocksubject",4096)&","&getlangid("notifystockemail",4096)&" FROM emailmessages WHERE emailID=1"
		rs.open sSQL,cnn,0,1
		oemailsubject=trim(rs(getlangid("notifystocksubject",4096))&"")
		oemailmessage=rs(getlangid("notifystockemail",4096))&""
		rs.close
		idlist=""
		if mysqlserver then
			sSQL="SELECT DISTINCT nsProdID FROM notifyinstock INNER JOIN prodoptions ON notifyinstock.nsProdID=prodoptions.poProdID INNER JOIN options ON prodoptions.poOptionGroup=options.optGroup WHERE nsOptID=-1 AND optID="&theoid
		else
			sSQL="SELECT DISTINCT nsProdID FROM notifyinstock INNER JOIN (prodoptions INNER JOIN options ON prodoptions.poOptionGroup=options.optGroup) ON notifyinstock.nsProdID=prodoptions.poProdID WHERE nsOptID=-1 AND optID="&theoid
		end if
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			gotall=TRUE
			sSQL="SELECT poOptionGroup FROM prodoptions INNER JOIN optiongroup ON prodoptions.poOptionGroup=optiongroup.optGrpID WHERE poProdID='"&escape_string(rs("nsProdID"))&"'"
			rs2.Open sSQL,cnn,0,1
			do while NOT rs2.EOF
				sSQL="SELECT optID FROM options WHERE optStock>0 AND optGroup="&rs2("poOptionGroup")
				rs3.Open sSQL,cnn,0,1
				if rs3.EOF then gotall=FALSE
				rs3.Close
				rs2.movenext
			loop
			rs2.Close
			if gotall then idlist=idlist&"'"&escape_string(rs("nsProdID"))&"',"
			rs.movenext
		loop
		rs.close
		if idlist<>"" then idlist=left(idlist,len(idlist)-1)

		pStockByOpts=0
		sSQL="SELECT pId,pName,pStockByOpts,pStaticPage,pStaticURL,pInStock,nsEmail FROM products INNER JOIN notifyinstock ON products.pID=notifyinstock.nsProdID WHERE nsOptId="&theoid
		if idlist<>"" then sSQL=sSQL & " OR (nsOptID=-1 AND nsProdID IN ("&idlist&"))"
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			nspid=rs("pId")
			pName=trim(rs("pName"))
			pStockByOpts=rs("pStockByOpts")
			pStaticPage=rs("pStaticPage")
			pStaticURL=rs("pStaticURL")
			pInStock=rs("pInStock")
			theemail=rs("nsEmail")
			thelink=storeurl & getdetailsurl(nspid,pStaticPage,pName,trim(pStaticURL&""),"","")
			if htmlemails=TRUE AND thelink<>"" then thelink="<a href=""" & thelink & """>" & thelink & "</a>"
			emailsubject=replace(oemailsubject,"%pid%",trim(nspid))
			emailsubject=replace(emailsubject,"%pname%",pName)
			emailmessage=replace(oemailmessage,"%pid%",trim(nspid))
			emailmessage=replace(emailmessage,"%pname%",pName)
			emailmessage=replace(emailmessage,"%link%",thelink)
			emailmessage=replace(emailmessage,"%storeurl%",storeurl)
			emailmessage=replace(emailmessage, "<br />", emlNl)
			emailmessage=replace(emailmessage, "%nl%", emlNl)
			call DoSendEmailEO(theemail,emailAddr,"",emailsubject,emailmessage,emailObject,themailhost,theuser,thepass)
			rs.movenext
		loop
		rs.close
		sSQL="DELETE FROM notifyinstock WHERE nsOptId="&theoid
		if idlist<>"" then sSQL=sSQL & " OR (nsOptID=-1 AND nsProdID IN ("&idlist&"))"
		ect_query(sSQL)
	end if
end sub
sub checknotifystock(thepid)
	if notifybackinstock then
		pStockByOpts=1
		pInStock=0
		sSQL="SELECT pStockByOpts,pInStock,pSell FROM products WHERE pID='"&escape_string(thepid)&"'"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			pStockByOpts=rs("pStockByOpts")
			if NOT useStockManagement then pInStock=IIfVr(rs("pSell")<>0,1,0) : pStockByOpts=0 else pInStock=rs("pInStock")
		end if
		rs.close
		if pStockByOpts=0 AND pInStock>0 then
			sSQL="SELECT "&getlangid("notifystocksubject",4096)&","&getlangid("notifystockemail",4096)&" FROM emailmessages WHERE emailID=1"
			rs.open sSQL,cnn,0,1
			emailsubject=trim(rs(getlangid("notifystocksubject",4096))&"")
			emailmessage=rs(getlangid("notifystockemail",4096))&""
			rs.close
			sSQL="SELECT nsEmail,nsProdId,nsTriggerProdId,pName,pStaticPage,pStaticURL FROM notifyinstock INNER JOIN products ON notifyinstock.nsProdId=products.pId WHERE nsTriggerProdID='"&escape_string(thepid)&"'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				nspid=rs("nsProdId")
				pName=trim(rs("pName"))
				pStaticPage=rs("pStaticPage")
				pStaticURL=rs("pStaticURL")
				thelink=storeurl & getdetailsurl(nspid,pStaticPage,pName,trim(pStaticURL&""),"","")
				if htmlemails=TRUE AND thelink<>"" then thelink="<a href=""" & thelink & """>" & thelink & "</a>"
				emailsubject=replace(emailsubject,"%pid%",trim(nspid))
				emailsubject=replace(emailsubject,"%pname%",pName)
				emailmessage=replace(emailmessage,"%pid%",trim(nspid))
				emailmessage=replace(emailmessage,"%pname%",pName)
				emailmessage=replace(emailmessage,"%link%",thelink)
				emailmessage=replace(emailmessage,"%storeurl%",storeurl)
				emailmessage=replace(emailmessage, "<br />", emlNl)
				emailmessage=replace(emailmessage, "%nl%", emlNl)
				do while NOT rs.EOF
					call DoSendEmailEO(rs("nsEmail"),emailAddr,"",emailsubject,emailmessage,emailObject,themailhost,theuser,thepass)
					rs.movenext
				loop
			end if
			rs.close
			ect_query("DELETE FROM notifyinstock WHERE nsTriggerProdID='"&escape_string(thepid)&"'")
		end if
	end if
end sub
function getstaticprodcurl(prodid,prodname,forcelower,spacereplace,removepunctuation,addextension)
	if getpost("addprodid")="prepend" then prodname=prodid&" "&prodname
	if getpost("addprodid")="append" then prodname=prodname&" "&prodid
	prodname=replaceaccents(prodname)
	prodname=strip_tags2(prodname)
	if forcelower then prodname=lcase(prodname)
	if spacereplace="remove" then spacereplace=""
	prodname=replace(prodname," ",spacereplace)
	if removepunctuation then
		prregexp.pattern="&(?:[a-z\d]+|#\d+|#x[a-f\d]+);"
		prodname=prregexp.replace(prodname,"")
		prregexp.pattern="[$&+,/:;=?@""'<>#%{}|\\^~\[\]`]"
		prodname=prregexp.replace(prodname,"")
	end if
	if spacereplace<>"" then
		prregexp.pattern="["&spacereplace&"]{2,}"
		prodname=prregexp.replace(prodname,cstr(spacereplace))
	end if
	if addextension then prodname=prodname&".asp"
	getstaticprodcurl=prodname
end function
sub docheckpackage(newid)
	sSQL="SELECT packageID FROM productpackages WHERE pID='"&escape_string(newid)&"'"
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		sSQL="SELECT pID,quantity FROM productpackages WHERE packageID='"&escape_string(rs("packageID"))&"'"
		rs2.open sSQL,cnn,0,1
		sumprice=0 : sumwsprice=0 : sumlistprice=0 : sumweight=0 : stockavailable=100000
		do while NOT rs2.EOF
			sSQL="SELECT pPrice,pWholesalePrice,pListPrice,pWeight,pInStock FROM products WHERE pID='"&escape_string(rs2("pID"))&"'"
			rs3.open sSQL,cnn,0,1
			if NOT rs3.EOF then
				sumprice=sumprice+(rs3("pPrice")*rs2("quantity"))
				sumwsprice=sumwsprice+(rs3("pWholesalePrice")*rs2("quantity"))
				sumlistprice=sumlistprice+(rs3("pListPrice")*rs2("quantity"))
				sumweight=sumweight+(rs3("pWeight")*rs2("quantity"))
				if int(rs3("pInStock")/rs2("quantity"))<stockavailable then stockavailable=int(rs3("pInStock")/rs2("quantity"))
			end if
			rs3.close
			rs2.movenext
		loop
		rs2.close
		sSQL="UPDATE products SET pPrice="&sumprice&",pWholesalePrice="&sumwsprice&",pListPrice="&sumlistprice&",pWeight="&sumweight&",pInStock="&stockavailable&" WHERE pID='"&escape_string(rs("packageID"))&"'"
		cnn.execute(sSQL)
		rs.movenext
	loop
	rs.close
end sub
if defaultprodimages="" then defaultprodimages="prodimages/"
if getpost("posted")="1" then
	pExemptions=0
	newid=getpost("newid")
	if getpost("pExemptions")<>"" then
		pExemptArray=split(getpost("pExemptions"), ",")
		for each pExemptObj in pExemptArray
			pExemptions=pExemptions + pExemptObj
		next
	end if
	if getpost("act")="recalcsalesrank" then
		ect_query("UPDATE products SET pNumSales=0")
		sSQL="SELECT cartProdID,COUNT(*) AS thecount FROM cart WHERE cartCompleted<>0"
		if is_numeric(getpost("id")) then
			sSQL=sSQL&" AND cartDateAdded BETWEEN " & vsusdate(date()-(30*int(getpost("id")))) & " AND " & vsusdate(date()+1)
		end if
		sSQL=sSQL&" GROUP BY cartProdID"
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			sSQL="UPDATE products SET pNumSales=" & rs("thecount") & " WHERE pID='"&escape_string(rs("cartProdID"))&"'"
			ect_query(sSQL)
			rs.movenext
		loop
		rs.close
		dorefresh=TRUE
	elseif getpost("act")="dotablechecks" then
		if getpost("subact")="manattr" OR getpost("subact")="fixall" then
			sSQL="SELECT mSCpID,scID FROM searchcriteria INNER JOIN multisearchcriteria ON searchcriteria.scid=multisearchcriteria.mSCscID INNER JOIN products on multisearchcriteria.mSCpID=products.pID WHERE scGroup=0 AND mSCscID<>pManufacturer"
			rs.open sSQL,cnn,0,1
			do while NOT rs.EOF
				sSQL="DELETE FROM multisearchcriteria WHERE mSCscID="&rs("scID")&" AND mSCpID='"&escape_string(rs("mSCpID"))&"'"
				ect_query(sSQL)
				rs.movenext
			loop
			rs.close
		end if
		if getpost("subact")="mannoexist" OR getpost("subact")="fixall" then
			sSQL="SELECT pID FROM products LEFT JOIN searchcriteria ON products.pManufacturer=searchcriteria.scid WHERE pManufacturer<>0 AND scName IS NULL"
			rs.open sSQL,cnn,0,1
			do while NOT rs.EOF
				sSQL="UPDATE products SET pManufacturer=0 WHERE pID='"&escape_string(rs("pID"))&"'"
				ect_query(sSQL)
				rs.movenext
			loop
			rs.close

			sSQL="SELECT pID FROM products INNER JOIN searchcriteria ON products.pManufacturer=searchcriteria.scid WHERE scGroup<>0"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				sSQL="UPDATE products SET pManufacturer=0 WHERE pID='"&escape_string(rs("pID"))&"'"
				ect_query(sSQL)
			end if
			rs.close
		end if
		dorefresh=TRUE
	elseif getpost("act")="allstk" then
		call notifyallstock()
		dorefresh=TRUE
	elseif getpost("act")="delete" then
		call dodeleteprod(getpost("id"))
		dorefresh=TRUE
	elseif getpost("act")="updatepackages" then
		pid=getpost("pid")
		for each objItem in request.form
			if left(objItem,4)="updq" then
				theprodid=right(objItem,len(objItem)-4)
				sSQL="DELETE FROM productpackages WHERE (packageID='"&escape_string(pid)&"' AND pID='"&escape_string(getpost(objItem))&"')"
				ect_query(sSQL)
				if getpost("updr"&theprodid)="1" AND is_numeric(getpost("pqa"&theprodid)) then
					if int(getpost("pqa"&theprodid))>=1 then
						sSQL="INSERT INTO productpackages (packageID,pID,quantity) VALUES ('"&escape_string(pid)&"','"&escape_string(getpost(objItem))&"',"&int(getpost("pqa"&theprodid))&")"
						ect_query(sSQL)
						docheckpackage(getpost(objItem))
					end if
				end if
			end if
		next
		dorefresh=TRUE
	elseif getpost("act")="updaterelations" then
		rid=getpost("rid")
		for each objItem in request.form
			if left(objItem,4)="updq" then
				theprodid=right(objItem,len(objItem)-4)
				sSQL="DELETE FROM relatedprods WHERE (rpProdID='"&escape_string(rid)&"' AND rpRelProdID='"&escape_string(getpost(objItem))&"')"
				if relatedproductsbothways=TRUE then sSQL=sSQL & " OR (rpRelProdID='"&escape_string(rid)&"' AND rpProdID='"&escape_string(getpost(objItem))&"')"
				ect_query(sSQL)
				if getpost("updr"&theprodid)="1" then
					sSQL="INSERT INTO relatedprods (rpProdID,rpRelProdID) VALUES ('"&escape_string(rid)&"','"&escape_string(getpost(objItem))&"')"
					ect_query(sSQL)
				end if
			end if
		next
		dorefresh=TRUE
	elseif getpost("act")="quickupdate" AND getpost("wholedb")<>"" then
		if getpost("wholedb")="clear" then
			ect_query("UPDATE products SET pStaticURL=''")
		else
			rs.open "SELECT pID,pName FROM products",cnn,0,1
			do while NOT rs.EOF
				ect_query("UPDATE products SET pStaticURL='"&escape_string(getstaticprodcurl(rs("pID"),rs("pName"),getpost("lcase")="yes",request.form("space"),getpost("punctuation")="remove",getpost("extension")="yes"))&"' WHERE pID='"&escape_string(rs("pID"))&"'")
				rs.movenext
			loop
			rs.close
		end if
		dorefresh=TRUE
	elseif getpost("act")="quickupdate" then
		attrgroup=-1
		checkpackage=FALSE
		if getpost("currentattribute")<>"" then
			rs.open "SELECT scGroup FROM searchcriteria WHERE scID="&getpost("currentattribute"),cnn,0,1
			attrgroup=rs("scGroup")
			rs.close
		end if
		for each objItem in request.form
			if left(objItem, 4)="pra_" then
				origid=right(objItem, len(objItem)-4)
				theid=getpost("pid"&origid)
				theval=getpost(objItem)
				pract=getpost("pract")
				sSQL=""
				if pract="prn" then
					if trim(theval)<>"" then sSQL="UPDATE products SET pName='" & escape_string(theval) & "'"
				elseif pract="prn2" then
					if trim(theval)<>"" then sSQL="UPDATE products SET pName2='" & escape_string(theval) & "'"
				elseif pract="prn3" then
					if trim(theval)<>"" then sSQL="UPDATE products SET pName3='" & escape_string(theval) & "'"
				elseif pract="pra" AND getpost("currentattribute")<>"" then
					if getpost("prb_" & origid)="1" then
						if attrgroup=0 then
							rs.open "SELECT mSCscID FROM multisearchcriteria INNER JOIN searchcriteria ON multisearchcriteria.mSCscID=searchcriteria.scID WHERE scGroup=0 AND mSCpID='"&escape_string(theid)&"'",cnn,0,1
							do while NOT rs.EOF
								ect_query("DELETE FROM multisearchcriteria WHERE mSCpID='"&escape_string(theid)&"' AND mSCscID="&rs("mSCscID"))
								rs.movenext
							loop
							rs.close
						end if
						on error resume next
						ect_query("INSERT INTO multisearchcriteria (mSCpID,mSCscID) VALUES ('"&escape_string(theid)&"',"&getpost("currentattribute")&")")
						on error goto 0
						if attrgroup=0 then ect_query("UPDATE products SET pManufacturer="&getpost("currentattribute")&" WHERE pID='"&escape_string(theid)&"'")
					else
						ect_query("DELETE FROM multisearchcriteria WHERE mSCpID='"&escape_string(theid)&"' AND mSCscID="&getpost("currentattribute"))
						if attrgroup=0 then ect_query("UPDATE products SET pManufacturer=0 WHERE pManufacturer="&getpost("currentattribute")&" AND pID='"&escape_string(theid)&"'")
					end if
				elseif pract="prp" AND getpost("currentoption")<>"" then
					if getpost("prb_" & origid)="1" then
						on error resume next
						ect_query("INSERT INTO prodoptions (poProdID,poOptionGroup) VALUES ('"&escape_string(theid)&"',"&getpost("currentoption")&")")
						on error goto 0
					else
						ect_query("DELETE FROM prodoptions WHERE poProdID='"&escape_string(theid)&"' AND poOptionGroup="&getpost("currentoption"))
					end if
				elseif pract="dis" AND getpost("currentdiscount")<>"" then
					ect_query("DELETE FROM cpnassign WHERE cpaType=2 AND cpaAssignment='"&escape_string(theid)&"' AND cpaCpnID="&getpost("currentdiscount"))
					if getpost("prb_" & origid)="1" then
						ect_query("INSERT INTO cpnassign (cpaType,cpaAssignment,cpaCpnID) VALUES (2,'"&escape_string(theid)&"',"&getpost("currentdiscount")&")")
					end if
				elseif pract="ads" AND getpost("currentsection")<>"" then
					ect_query("DELETE FROM multisections WHERE pID='"&escape_string(theid)&"' AND pSection="&getpost("currentsection"))
					if getpost("prb_" & origid)="1" then
						rs.open "SELECT pID FROM products WHERE pID='"&escape_string(theid)&"' AND pSection="&getpost("currentsection"),cnn,0,1
						if rs.EOF then ect_query("INSERT INTO multisections (pID,pSection) VALUES ('"&escape_string(theid)&"',"&getpost("currentsection")&")")
						rs.close
					end if
				elseif pract="sec" then
					if is_numeric(theval) then sSQL="UPDATE products SET pSection=" & theval
				elseif pract="psp" then
					sSQL="UPDATE products SET pSearchParams='" & escape_string(theval) & "'"
				elseif pract="psp2" then
					sSQL="UPDATE products SET pSearchParams2='" & escape_string(theval) & "'"
				elseif pract="psp3" then
					sSQL="UPDATE products SET pSearchParams3='" & escape_string(theval) & "'"
				elseif pract="pti" then
					sSQL="UPDATE products SET pTitle='" & escape_string(theval) & "'"
				elseif pract="pti2" then
					sSQL="UPDATE products SET pTitle2='" & escape_string(theval) & "'"
				elseif pract="pti3" then
					sSQL="UPDATE products SET pTitle3='" & escape_string(theval) & "'"
				elseif pract="pmd" then
					sSQL="UPDATE products SET pMetaDesc='" & escape_string(theval) & "'"
				elseif pract="pmd2" then
					sSQL="UPDATE products SET pMetaDesc2='" & escape_string(theval) & "'"
				elseif pract="pmd3" then
					sSQL="UPDATE products SET pMetaDesc3='" & escape_string(theval) & "'"
				elseif pract="pri" then
					if is_numeric(theval) then sSQL="UPDATE products SET pPrice=" & theval
					checkpackage=TRUE
				elseif pract="wpr" then
					if is_numeric(theval) then sSQL="UPDATE products SET pWholesalePrice=" & theval
					checkpackage=TRUE
				elseif pract="lpr" then
					if is_numeric(theval) then sSQL="UPDATE products SET pListPrice=" & theval
					checkpackage=TRUE
				elseif pract="sid" then
					if is_numeric(theval) then sSQL="UPDATE products SET pSiteID=" & theval
				elseif pract="stk" then
					if is_numeric(theval) then sSQL="UPDATE products SET pInStock=" & theval
					checkpackage=TRUE
				elseif pract="pop" then
					if is_numeric(theval) then sSQL="UPDATE products SET pPopularity=" & theval
				elseif pract="sal" then
					if is_numeric(theval) then sSQL="UPDATE products SET pNumSales=" & theval
				elseif pract="mnq" then
					if is_numeric(theval) then sSQL="UPDATE products SET pMinQuant=" & IIfVr(cint(theval)<=0,0,cint(theval)-1)
				elseif pract="sta" then
					if is_numeric(theval) then sSQL="UPDATE products SET pInStock=pInStock+" & theval
					checkpackage=TRUE
				elseif pract="del" then
					if theval="del" then call dodeleteprod(theid)
					sSQL=""
				elseif pract="pru" then
					sSQL="UPDATE products SET pUpload=" & IIfVr(getpost("prb_" & origid)="1","1","0")
				elseif pract="prw" then
					if is_numeric(theval) then sSQL="UPDATE products SET pWeight=" & theval
					checkpackage=TRUE
				elseif pract="dip" then
					sSQL="UPDATE products SET pDisplay=" & IIfVr(getpost("prb_" & origid)="1","1","0")
				elseif pract="stp" then
					sSQL="UPDATE products SET pStaticPage=" & IIfVr(getpost("prb_" & origid)="1","1","0")
				elseif pract="stu" then
					sSQL="UPDATE products SET pStaticURL='" & escape_string(theval) & "'"
				elseif pract="css" then
					sSQL="UPDATE products SET pCustomCSS='" & escape_string(theval) & "'"
				elseif pract="rec" then
					sSQL="UPDATE products SET pRecommend=" & IIfVr(getpost("prb_" & origid)="1","1","0")
				elseif pract="gwr" then
					sSQL="UPDATE products SET pGiftWrap=" & IIfVr(getpost("prb_" & origid)="1","1","0")
				elseif pract="isa" then
					sSQL="UPDATE products SET pSchemaType=" & IIfVr(getpost("prb_" & origid)="1","1","0")
				elseif pract="bak" then
					sSQL="UPDATE products SET pBackOrder=" & IIfVr(getpost("prb_" & origid)="1","1","0")
				elseif pract="sku" then
					sSQL="UPDATE products SET pSKU='" & escape_string(theval) & "'"
				elseif pract="pro" then
					if is_numeric(theval) then sSQL="UPDATE products SET pOrder=" & theval
				elseif pract="ppt" then
					if is_numeric(theval) then sSQL="UPDATE products SET pTax=" & theval
				elseif pract="sel" then
					theval=IIfVr(getpost("prb_" & origid)="1",1,0)
					sSQL="UPDATE products SET pSell=" & theval
				elseif pract="dld" then
					sSQL="UPDATE products SET pDownload='" & escape_string(theval) & "'"
				elseif pract="frs" then
					ship1=IIfVr(is_numeric(theval), theval, 0)
					ship2=IIfVr(is_numeric(getpost("prb_" & origid)), getpost("prb_" & origid), 0)
					sSQL="UPDATE products SET pShipping=" & ship1 & ", pShipping2=" & ship2
				elseif pract="daa" then
					if NOT isdate(theval) then theval=date()
					sSQL="UPDATE products SET pDateAdded=" & vsusdate(theval)
				elseif pract="ste" OR pract="cte" OR pract="she" OR pract="hae" OR pract="fse" OR pract="pte" OR pract="pde" then
					fieldnum=1
					if pract="cte" then fieldnum=2
					if pract="she" then fieldnum=4
					if pract="hae" then fieldnum=8
					if pract="fse" then fieldnum=16
					if pract="pte" then fieldnum=32
					if pract="pde" then fieldnum=64
					rs.open "SELECT pExemptions FROM products WHERE pID='"&escape_string(theid)&"'",cnn,0,1
					if NOT rs.EOF then theval=rs("pExemptions") else theval=0
					rs.close
					if getpost("prb_" & origid)="1" then
						if (theval AND fieldnum)=0 then theval=theval + fieldnum
					else
						if (theval AND fieldnum)<>0 then theval=theval - fieldnum
					end if
					sSQL="UPDATE products SET pExemptions=" & theval
				elseif pract="csu" then
					rs.open "SELECT pID,pName FROM products WHERE pID='"&escape_string(theid)&"'",cnn,0,1
					if NOT rs.EOF then
						sSQL="UPDATE products SET pStaticURL='"&escape_string(getstaticprodcurl(rs("pID"),rs("pName"),getpost("lcase")="yes",request.form("space"),getpost("punctuation")="remove",getpost("extension")="yes"))&"'"
					end if
					rs.close
				elseif pract<>"" AND customquickupdate<>"" then
					cqupdatearr=split(customquickupdate,",")
					for index=0 to UBOUND(cqupdatearr)
						cqitemarr=split(cqupdatearr(index),":")
						if pract=cqitemarr(0) then
							fieldtype="text"
							if UBOUND(cqitemarr)>=2 then
								fieldtype=lcase(cqitemarr(2))
								if lcase(cqitemarr(2))="num" then
									if NOT is_numeric(theval) then theval=0
								end if
							end if
							if fieldtype="num" then
								if is_numeric(theval) then sSQL="UPDATE products SET "&pract&"=" & escape_string(theval)
							elseif fieldtype="check" then
								sSQL="UPDATE products SET "&pract&"=" & IIfVr(getpost("prb_" & origid)="1","1","0")
							else
								sSQL="UPDATE products SET "&pract&"='" & escape_string(theval) & "'"
							end if
						end if
					next
				end if
				if sSQL<>"" then
					sSQL=sSQL & " WHERE pID='"&escape_string(theid)&"'"
					ect_query(sSQL)
					if checkpackage then docheckpackage(theid)
				end if
				if (pract="stk" OR pract="sta" OR (NOT useStockManagement AND pract="sel")) AND trim(theval)<>"" then
					if int(theval)>0 then call checknotifystock(theid)
				end if
			end if
		next
		dorefresh=TRUE
	elseif getpost("act")="domodify" then
		if newid<>getpost("id") then
			if lcase(newid)=lcase(getpost("id")) then
				success=TRUE
			else
				sSQL="SELECT * FROM products WHERE pID='"&escape_string(newid)&"'"
				rs.open sSQL,cnn,0,1
				success=rs.EOF
				rs.close
				if success then
					ect_query("UPDATE cpnassign SET cpaAssignment='"&escape_string(newid)&"' WHERE cpaType=2 AND cpaAssignment='"&escape_string(getpost("id"))&"'")
					ect_query("UPDATE multisections SET pID='"&escape_string(newid)&"' WHERE pID='"&escape_string(getpost("id"))&"'")
					ect_query("UPDATE multisearchcriteria SET mSCpID='"&escape_string(newid)&"' WHERE mSCpID='"&escape_string(getpost("id"))&"'")
					ect_query("UPDATE notifyinstock SET nsProdID='"&escape_string(newid)&"' WHERE nsProdID='"&escape_string(getpost("id"))&"'")
					ect_query("UPDATE notifyinstock SET nsTriggerProdID='"&escape_string(newid)&"' WHERE nsTriggerProdID='"&escape_string(getpost("id"))&"'")
					ect_query("UPDATE pricebreaks SET pbProdID='"&escape_string(newid)&"' WHERE pbProdID='"&escape_string(getpost("id"))&"'")
					ect_query("UPDATE productpackages SET pID='"&escape_string(newid)&"' WHERE pID='"&escape_string(getpost("id"))&"'")
					ect_query("UPDATE productpackages SET packageID='"&escape_string(newid)&"' WHERE packageID='"&escape_string(getpost("id"))&"'")
					ect_query("UPDATE ratings SET rtProdID='"&escape_string(newid)&"' WHERE rtProdID='"&escape_string(getpost("id"))&"'")
					ect_query("UPDATE relatedprods SET rpProdID='"&escape_string(newid)&"' WHERE rpProdID='"&escape_string(getpost("id"))&"'")
					ect_query("UPDATE relatedprods SET rpRelProdID='"&escape_string(newid)&"' WHERE rpRelProdID='"&escape_string(getpost("id"))&"'")
				end if
			end if
		end if
		if success then
			pOrder=getpost("pOrder")
			if NOT is_numeric(pOrder) then pOrder=0
			sSQL="UPDATE products SET " & _
				"pID='"& escape_string(newid) &"', " & _
				"pName='"& escape_string(getpost("pName")) &"', " & _
				"pSection="& getpost("psection") &", " & _
				"pDropship="& getpost("pDropship") &", " & _
				"pManufacturer="& getpost("pManufacturer") &", " & _
				"pSKU='"& escape_string(getpost("pSKU")) &"', " & _
				"pOrder="& pOrder &", " & _
				"pExemptions="& pExemptions &", " & _
				"pSearchParams='"& escape_string(getpost("pSearchParams")) &"', " & _
				"pTitle='"& escape_string(getpost("pTitle")) &"', " & _
				"pMetaDesc='"& escape_string(getpost("pMetaDesc")) &"', " & _
				"pDescription='"& escape_string(getpost("pDescription")) &"', " & _
				"pLongDescription='"& escape_string(getpost("pLongDescription")) &"', "
			for index=2 to adminlanguages+1
				if (adminlangsettings AND 1)=1 then sSQL=sSQL & "pName"&index&"='"& escape_string(getpost("pName"&index)) &"', "
				if (adminlangsettings AND 2)=2 then sSQL=sSQL & "pDescription"&index&"='"& escape_string(getpost("pDescription"&index)) &"', "
				if (adminlangsettings AND 4)=4 then sSQL=sSQL & "pLongDescription"&index&"='"& escape_string(getpost("pLongDescription"&index)) &"', "
				if (adminlangsettings AND 2097152)=2097152 then sSQL=sSQL & "pTitle"&index&"='"& escape_string(getpost("pTitle"&index)) &"', "
				if (adminlangsettings AND 2097152)=2097152 then sSQL=sSQL & "pMetaDesc"&index&"='"& escape_string(getpost("pMetaDesc"&index)) &"', "
				if (adminlangsettings AND 4194304)=4194304 then sSQL=sSQL & "pSearchParams"&index&"='"& escape_string(getpost("pSearchParams"&index)) &"', "
			next
			sSQL=sSQL & "pDisplay="&IIfVr(getpost("pDisplay")="ON",1,0)&","
			if perproducttaxrate=true then sSQL=sSQL & "pTax=" & getpost("pTax") & ","
			if is_numeric(getpost("inStock")) AND getpost("stocksetting")="1" then sSQL=sSQL & "pInStock=" & getpost("inStock")&","
			sSQL=sSQL & "pStockByOpts=" & IIfVr(getpost("pStockByOpts")="1", 1, 0) & "," & _
				"pStaticPage=" & IIfVr(getpost("pStaticPage")="1", 1, 0) & "," & _
				"pStaticURL='" & escape_string(getpost("pStaticURL")) & "'," & _
				"pRecommend=" & IIfVr(getpost("pRecommend")="1", 1, 0) & "," & _
				"pGiftWrap=" & IIfVr(getpost("pGiftWrap")="1", 1, 0) & "," & _
				"pBackOrder=" & IIfVr(getpost("pBackOrder")="1", 1, 0) & "," & _
				"pSell=" & IIfVr(getpost("pSell")="ON", 1, 0) & ","
				if (adminUnits AND 12) > 0 then sSQL=sSQL & "pDims='" & getpost("plen") & "x" & getpost("pwid") & "x" & getpost("phei") & "',"
				if digidownloads=true then sSQL=sSQL & "pDownload='" & escape_string(getpost("pDownload")) & "',"
				sSQL=sSQL & "pShipping="&IIfVr(is_numeric(getpost("pShipping")),getpost("pShipping"),0)&","
				sSQL=sSQL & "pShipping2="&IIfVr(is_numeric(getpost("pShipping2")),getpost("pShipping2"),0)&","
				sSQL=sSQL & "pWeight="&IIfVr(is_numeric(getpost("pWeight")),getpost("pWeight"),0)&","
				sSQL=sSQL & "pWholesalePrice="&IIfVr(is_numeric(getpost("pWholesalePrice")),getpost("pWholesalePrice"),0)&","
				sSQL=sSQL & "pListPrice="&IIfVr(is_numeric(getpost("pListPrice")),getpost("pListPrice"),0)&","
				if instr(productpagelayout&detailpagelayout,"custom1")>0 then sSQL=sSQL & "pCustom1='"& escape_string(getpost("pCustom1")) &"',"
				if instr(productpagelayout&detailpagelayout,"custom2")>0 then sSQL=sSQL & "pCustom2='"& escape_string(getpost("pCustom2")) &"',"
				if instr(productpagelayout&detailpagelayout,"custom3")>0 then sSQL=sSQL & "pCustom3='"& escape_string(getpost("pCustom3")) &"',"
				session.lcid=saveLCID
				if getpost("pDateAdded")<>"" then
					sSQL=sSQL & "pDateAdded=" & vsusdate(datevalue(getpost("pDateAdded"))) & ","
				else
					sSQL=sSQL & "pDateAdded=" & vsusdate(date()) & ","
				end if
				session.lcid=1033
				sSQL=sSQL & "pPrice="& getpost("pPrice") &" WHERE pID='"&escape_string(getpost("id"))&"'"
			ect_query(sSQL)
			call checknotifystock(newid)
			docheckpackage(newid)
			dorefresh=TRUE
		else
			errmsg=yyPrDup
		end if
	elseif getpost("act")="doaddnew" then
		sSQL="SELECT * FROM products WHERE pID='"&escape_string(newid)&"'"
		rs.open sSQL,cnn,0,1
		success=rs.EOF
		rs.close
		if success then
			pOrder=getpost("pOrder")
			if NOT is_numeric(pOrder) then pOrder=0
			sSQL="INSERT INTO products (pID,pName,pDateAdded,pSection,pDropship,pManufacturer,pSKU,pOrder,pExemptions,pSearchParams,pTitle,pMetaDesc,pCustom1,pCustom2,pCustom3,pDescription,pLongDescription,"
			for index=2 to adminlanguages+1
				if (adminlangsettings AND 1)=1 then sSQL=sSQL & "pName" & index & ","
				if (adminlangsettings AND 2)=2 then sSQL=sSQL & "pDescription" & index & ","
				if (adminlangsettings AND 4)=4 then sSQL=sSQL & "pLongDescription" & index & ","
				if (adminlangsettings AND 2097152)=2097152 then sSQL=sSQL & "pTitle" & index & ","
				if (adminlangsettings AND 2097152)=2097152 then sSQL=sSQL & "pMetaDesc" & index & ","
				if (adminlangsettings AND 4194304)=4194304 then sSQL=sSQL & "pSearchParams" & index & ","
			next
			sSQL=sSQL & "pPrice,pWholesalePrice,pListPrice,pShipping,pShipping2,pDisplay,"
			if perproducttaxrate=true then sSQL=sSQL & "pTax,"
			if is_numeric(getpost("inStock")) then sSQL=sSQL & "pInStock,"
			if (adminUnits AND 12) > 0 then sSQL=sSQL & "pDims,"
			if digidownloads=true then sSQL=sSQL & "pDownload,"
			sSQL=sSQL & "pStockByOpts,pStaticPage,pStaticURL,pRecommend,pGiftWrap,pBackOrder,pSell,pWeight) VALUES (" & _
						"'"&escape_string(newid)&"'," & _
						"'"&escape_string(getpost("pName"))&"'," & _
						vsusdate(date())&"," & _
						getpost("psection")&"," & _
						getpost("pDropship")&"," & _
						getpost("pManufacturer")&"," & _
						"'"&escape_string(getpost("pSKU"))&"'," & _
						pOrder&"," & _
						pExemptions &"," & _
						"'"&escape_string(getpost("pSearchParams"))&"'," & _
						"'"&escape_string(getpost("pTitle"))&"'," & _
						"'"&escape_string(getpost("pMetaDesc"))&"'," & _
						"'"&escape_string(getpost("pCustom1"))&"'," & _
						"'"&escape_string(getpost("pCustom2"))&"'," & _
						"'"&escape_string(getpost("pCustom3"))&"'," & _
						"'"&escape_string(getpost("pDescription"))&"'," & _
						"'"&escape_string(getpost("pLongDescription"))&"',"
						for index=2 to adminlanguages+1
							if (adminlangsettings AND 1)=1 then sSQL=sSQL & "'"& escape_string(getpost("pName"&index)) &"',"
							if (adminlangsettings AND 2)=2 then sSQL=sSQL & "'"& escape_string(getpost("pDescription"&index)) &"',"
							if (adminlangsettings AND 4)=4 then sSQL=sSQL & "'"& escape_string(getpost("pLongDescription"&index)) &"',"
							if (adminlangsettings AND 2097152)=2097152 then sSQL=sSQL & "'"& escape_string(getpost("pTitle"&index)) &"',"
							if (adminlangsettings AND 2097152)=2097152 then sSQL=sSQL & "'"& escape_string(getpost("pMetaDesc"&index)) &"',"
							if (adminlangsettings AND 4194304)=4194304 then sSQL=sSQL & "'"& escape_string(getpost("pSearchParams"&index)) &"',"
						next
						sSQL=sSQL & getpost("pPrice")&","
						if getpost("pWholesalePrice")<>"" then
							sSQL=sSQL & getpost("pWholesalePrice") & ","
						else
							sSQL=sSQL & "0,"
						end if
						if getpost("pListPrice")<>"" then
							sSQL=sSQL & getpost("pListPrice") & ","
						else
							sSQL=sSQL & "0,"
						end if
						if NOT is_numeric(getpost("pShipping")) then
							sSQL=sSQL & "0,"
						else
							sSQL=sSQL & getpost("pShipping")&","
						end if
						if NOT is_numeric(getpost("pShipping2")) then
							sSQL=sSQL & "0,"
						else
							sSQL=sSQL & getpost("pShipping2")&","
						end if
						if getpost("pDisplay")="ON" then
							sSQL=sSQL & "1,"
						else
							sSQL=sSQL & "0,"
						end if
						if perproducttaxrate=true then sSQL=sSQL & getpost("pTax") & ","
						if useStockManagement AND is_numeric(getpost("inStock")) then
							sSQL=sSQL & getpost("inStock")&","
						end if
						if (adminUnits AND 12) > 0 then
							sSQL=sSQL & "'" & getpost("plen") & "x" & getpost("pwid") & "x" & getpost("phei") & "',"
						end if
						if digidownloads=true then
							sSQL=sSQL & "'" & escape_string(getpost("pDownload")) & "',"
						end if
						sSQL=sSQL & IIfVr(getpost("pStockByOpts")="1", 1, 0) & ","
						sSQL=sSQL & IIfVr(getpost("pStaticPage")="1", 1, 0) & "," & _
								"'" & escape_string(getpost("pStaticURL")) & "',"
						sSQL=sSQL & IIfVr(getpost("pRecommend")="1", 1, 0) & ","
						sSQL=sSQL & IIfVr(getpost("pGiftWrap")="1", 1, 0) & ","
						sSQL=sSQL & IIfVr(getpost("pBackOrder")="1", 1, 0) & ","
						sSQL=sSQL & IIfVr(getpost("pSell")="ON", 1, 0) & ","
						if is_numeric(getpost("pWeight")) then sSQL=sSQL & getpost("pWeight") else sSQL=sSQL &"0"
						sSQL=sSQL & ")"
			on error resume next
			ect_query(sSQL)
			if err.number<>0 then
				success=false
				errmsg=errmsg & err.description
			else
				dorefresh=TRUE
			end if
			on error goto 0
		else
			errmsg=yyPrDup
		end if
	elseif getpost("act")="dodiscounts" then
		sSQL="INSERT INTO cpnassign (cpaCpnID,cpaType,cpaAssignment) VALUES ("&getpost("assdisc")&",2,'"&escape_string(getpost("id"))&"')"
		ect_query(sSQL)
		dorefresh=TRUE
	elseif getpost("act")="deletedisc" then
		sSQL="DELETE FROM cpnassign WHERE cpaID="&getpost("id")
		ect_query(sSQL)
		dorefresh=TRUE
	end if
	if success AND (getpost("act")="domodify" OR getpost("act")="doaddnew") then
		maximgindex=int(getpost("maximgindex"))
		if getpost("act")="domodify" then ect_query("DELETE FROM productimages WHERE imageProduct='" & escape_string(getpost("id")) & "'")
		for index=0 to maximgindex
			if getpost("smim" & index)<>"" then ect_query("INSERT INTO productimages (imageProduct,imageSrc,imageNumber,imageType) VALUES ('" & escape_string(newid) & "','" & escape_string(getpost("smim" & index)) & "'," & index & ",0)")
			if getpost("lgim" & index)<>"" then ect_query("INSERT INTO productimages (imageProduct,imageSrc,imageNumber,imageType) VALUES ('" & escape_string(newid) & "','" & escape_string(getpost("lgim" & index)) & "'," & index & ",1)")
			if getpost("gtim" & index)<>"" then ect_query("INSERT INTO productimages (imageProduct,imageSrc,imageNumber,imageType) VALUES ('" & escape_string(newid) & "','" & escape_string(getpost("gtim" & index)) & "'," & index & ",2)")
		next
		ect_query("DELETE FROM prodoptions WHERE poProdID='"&escape_string(getpost("id"))&"' OR poProdID='"&escape_string(newid)&"'")
		ect_query("DELETE FROM multisections WHERE pID='"&escape_string(getpost("id"))&"' OR pID='"&escape_string(newid)&"'")
		ect_query("DELETE FROM multisearchcriteria WHERE mSCpID='"&escape_string(getpost("id"))&"' OR mSCpID='"&escape_string(newid)&"'")
		if getpost("pManufacturer")<>"0" then
			sSQL="INSERT INTO multisearchcriteria (mSCpID,mSCscID) VALUES ('"&escape_string(newid)&"',"&getpost("pManufacturer")&")"
			ect_query(sSQL)
		end if
		for rowcounter=0 to 100
			if getpost("poption"&rowcounter)<>"" AND getpost("poption"&rowcounter)<>"0" then
				sSQL="INSERT INTO prodoptions (poProdID,poOptionGroup) VALUES ('"&escape_string(newid)&"',"&getpost("poption"&rowcounter)&")"
				ect_query(sSQL)
			end if
			if getpost("psection"&rowcounter)<>"" AND getpost("psection"&rowcounter)<>"0" AND getpost("psection")<>getpost("psection"&rowcounter) then
				sSQL="SELECT pID FROM multisections WHERE pID='" & escape_string(newid) & "' AND pSection="&getpost("psection"&rowcounter)
				rs.open sSQL,cnn,0,1
				if rs.EOF then
					sSQL="INSERT INTO multisections (pID,pSection) VALUES ('"&escape_string(newid)&"',"&getpost("psection"&rowcounter)&")"
					ect_query(sSQL)
				end if
				rs.close
			end if
			if getpost("psearch"&rowcounter)<>"" AND getpost("psearch"&rowcounter)<>"0" then
				sSQL="SELECT mSCpID FROM multisearchcriteria WHERE mSCpID='" & escape_string(newid) & "' AND mSCscID="&getpost("psearch"&rowcounter)
				rs.open sSQL,cnn,0,1
				if rs.EOF then
					sSQL="INSERT INTO multisearchcriteria (mSCpID,mSCscID) VALUES ('"&escape_string(newid)&"',"&getpost("psearch"&rowcounter)&")"
					ect_query(sSQL)
				end if
				rs.close
			end if
		next
		' Price Breaks
		ect_query("DELETE FROM pricebreaks WHERE pbProdID='" & escape_string(newid) & "'")
		pricebreakrows=getpost("pricebreakrows")
		for index=1 to pricebreakrows-1
			thequant=getpost("pbquant" & index)
			if NOT is_numeric(thequant) then thequant=0
			price=getpost("pbprice" & index)
			if NOT is_numeric(price) then price=0
			wprice=getpost("pbwholeprice" & index)
			if NOT is_numeric(wprice) then wprice=0
			wpercent=IIfVr(getpost("wpercent"&index)="1","1","0")
			wholesalepercent=IIfVr(getpost("wholesalepercent"&index)="1","1","0")
			if thequant<>0 AND (price<>0 OR wprice<>0) then
				sSQL="INSERT INTO pricebreaks (pbProdID,pbQuantity,pPrice,pWholesalePrice,pbPercent,pbWholesalePercent) VALUES ('" & escape_string(newid) & "',"
				sSQL=sSQL & thequant & ","
				sSQL=sSQL & price & ","
				sSQL = sSQL & wprice & ","
				sSQL = sSQL & wpercent & ","
				sSQL = sSQL & wholesalepercent & ")"
				on error resume next
				ect_query(sSQL)
				on error goto 0
			end if
		next
		dorefresh=TRUE
	end if
	if getpost("act")="modify" OR getpost("act")="clone" OR getpost("act")="addnew" then
		if getpost("act")="modify" OR getpost("act")="clone" then
			sSQL="SELECT poID,poOptionGroup,optGrpWorkingName FROM prodoptions INNER JOIN optiongroup ON prodoptions.poOptionGroup=optiongroup.optGrpID WHERE poProdID='"&escape_string(getpost("id"))&"' ORDER BY poID"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then prodoptions=rs.getrows
			rs.close
			sSQL="SELECT pSection,sectionWorkingName FROM multisections INNER JOIN sections ON multisections.pSection=sections.sectionID WHERE pID='"&escape_string(getpost("id"))&"'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then prodsections=rs.getrows
			rs.close
			sSQL="SELECT scID,scWorkingName FROM multisearchcriteria INNER JOIN searchcriteria ON multisearchcriteria.mSCscID=searchcriteria.scID WHERE scGroup<>0 AND mSCpID='"&escape_string(getpost("id"))&"' ORDER BY scGroup,scOrder"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then prodsearchcriteria=rs.getrows
			rs.close
		end if
		sSQL="SELECT dsID,dsName FROM dropshipper ORDER BY dsName"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then alldropship=rs.getrows
		rs.close
	end if
end if
if getpost("posted")="1" AND getpost("act")="altids" AND getpost("doupdate")="1" then
	dorefresh=TRUE
end if
if dorefresh then
	print "<meta http-equiv=""refresh"" content=""1; url=adminprods.asp"
	print "?sos=" & urlencode(getpost("sos")) & "&rid=" & urlencode(getpost("rid")) & "&pid=" & urlencode(getpost("pid")) & "&disp=" & getpost("disp") & "&stext=" & urlencode(getpost("stext")) & "&sprice=" & urlencode(getpost("sprice")) & "&stype=" & getpost("stype") & "&scat=" & getpost("scat") & "&pg=" & getpost("pg")
	print """>"
end if
  if getpost("posted")="1" AND getpost("act")="tablechecks" then
		sSQL="SELECT COUNT(*) AS tcount FROM (searchcriteria INNER JOIN multisearchcriteria ON searchcriteria.scid=multisearchcriteria.mSCscID) INNER JOIN products on multisearchcriteria.mSCpID=products.pID WHERE scGroup=0 AND mSCscID<>pManufacturer"
		rs.open sSQL,cnn,0,1
		tcount1=rs("tcount")
		rs.close
		
		sSQL="SELECT COUNT(*) AS tcount FROM products LEFT JOIN searchcriteria ON products.pManufacturer=searchcriteria.scid WHERE pManufacturer<>0 AND scName IS NULL"
		rs.open sSQL,cnn,0,1
		tcount2=rs("tcount")
		rs.close
		
		sSQL="SELECT COUNT(*) AS tcount FROM products INNER JOIN searchcriteria ON products.pManufacturer=searchcriteria.scid WHERE scGroup<>0"
		rs.open sSQL,cnn,0,1
		tcount2=tcount2+rs("tcount")
		rs.close
%>
		<form name="mainform" id="mainform" method="post" action="adminprods.asp">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="dotablechecks" />
			<input type="hidden" id="subact" name="subact" value="" />
			<table id="producttable" border="" cellspacing="3" cellpadding="3" style="margin:0 auto;border:1px solid;border-collapse:collapse">
			<tr><td colspan="3" align="center" style="border:1px solid"><div style="font-weight:bold">Products Table Checks</div></td></tr>
			<tr><td style="border:1px solid">Products where manufacturer doesn't match attributes.</td><td style="border:1px solid" align="center"><%=tcount1%></td><td style="border:1px solid" align="center"><% if tcount1>0 then print "<input type=""button"" value=""Fix"" onclick=""document.getElementById('subact').value='manattr';document.getElementById('mainform').submit()"" />" else print "-" %></td></tr>
			<tr><td style="border:1px solid">Products where manufacturer doesn't exist.</td><td style="border:1px solid" align="center"><%=tcount2%></td><td style="border:1px solid" align="center"><% if tcount2>0 then print "<input type=""button"" value=""Fix"" onclick=""document.getElementById('subact').value='mannoexist';document.getElementById('mainform').submit()"" />" else print "-" %></td></tr>
			<tr><td style="border:1px solid" colspan="3" align="center"><input type="button" value="Fix All" onclick="document.getElementById('subact').value='fixall';document.getElementById('mainform').submit()" /> <input type="button" value="Back to Products" onclick="document.location='adminprods.asp'" /></td></tr>
			</table>
		</form>
<%
  elseif getpost("posted")="1" AND getpost("act")="altids" AND getpost("doupdate")="1" then
		originalid=getpost("originalid")
		existingrows=int(getpost("existingrows"))
		newrows=int(getpost("newrows"))
		sSQL="SELECT pID,pName,pName2,pName3,pPrice,pWholesalePrice,pWeight,pInStock,pExemptions,pSection FROM products WHERE pID='" & escape_string(originalid) & "'"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			opname=rs("pName")
			opname2=rs("pName2")
			opname3=rs("pName3")
			opprice=rs("pPrice")
			opwprice=rs("pWholesalePrice")
			opweight=rs("pWeight")
			psection=rs("pSection")
		end if
		rs.close
		for index=0 to existingrows-1
			pid=getpost("xid"&index)
			pname=getpost("xna"&index)
			if pname="" then pname=opname
			pname2=getpost("xnb"&index)
			if pname2="" then pname2=IIfVr(opname2<>"",opname2,opname)
			pname3=getpost("xnc"&index)
			if pname3="" then pname3=IIfVr(opname3<>"",opname3,opname)
			pprice=getpost("xpr"&index)
			if NOT is_numeric(pprice) then pprice=opprice
			pwprice=getpost("xwp"&index)
			if NOT is_numeric(pwprice) then pwprice=opwprice
			pweight=getpost("xwe"&index)
			if NOT is_numeric(pweight) then pweight=opweight
			pinstock=getpost("xsk"&index)
			if NOT is_numeric(pinstock) then pinstock=0
			pexemptions=0
			if getpost("xst"&index)="1" then pexemptions=1
			if getpost("xct"&index)="1" then pexemptions=pexemptions+2
			if getpost("xsh"&index)="1" then pexemptions=pexemptions+4
			if getpost("xha"&index)="1" then pexemptions=pexemptions+8
			if getpost("xfs"&index)="1" then pexemptions=pexemptions+16
			if getpost("xpt"&index)="1" then pexemptions=pexemptions+32
			if getpost("xpd"&index)="1" then pexemptions=pexemptions+64
			pimage=getpost("xsmim"&index)
			plgimage=getpost("xlgim"&index)
			pgtimage=getpost("xgtim"&index)
			cnn.execute("DELETE FROM productimages WHERE imageNumber=0 AND imageProduct='" & escape_string(pid) & "'")
			if getpost("xde"&index)="1" then
				cnn.execute("DELETE FROM productimages WHERE imageProduct='" & escape_string(pid) & "'")
				sSQL="DELETE FROM products WHERE pID='" & escape_string(pid) & "'"
				cnn.execute(sSQL)
			else
				if pimage<>"" then cnn.execute("INSERT INTO productimages (imageProduct,imageSrc,imageNumber,imageType) VALUES ('" & escape_string(pid) & "','" & escape_string(pimage) & "',0,0)")
				if plgimage<>"" then cnn.execute("INSERT INTO productimages (imageProduct,imageSrc,imageNumber,imageType) VALUES ('" & escape_string(pid) & "','" & escape_string(plgimage) & "',0,1)")
				if pgtimage<>"" then cnn.execute("INSERT INTO productimages (imageProduct,imageSrc,imageNumber,imageType) VALUES ('" & escape_string(pid) & "','" & escape_string(pgtimage) & "',0,2)")
				sSQL="UPDATE products SET" & _
					" pName='" & escape_string(pname) & "'"
					if (adminlangsettings AND 1)=1 then
						if adminlanguages>=1 then sSQL=sSQL & ",pName2='" & escape_string(pname2) & "'"
						if adminlanguages>=2 then sSQL=sSQL & ",pName3='" & escape_string(pname3) & "'"
					end if
				sSQL=sSQL & ",pPrice=" & pprice & _
					",pWholesalePrice=" & pwprice & _
					",pWeight=" & pweight & _
					",pInStock=" & pinstock & _
					",pExemptions=" & pexemptions & _
					",pSection=" & psection & _
					",pDisplay=0" & _
					" WHERE pID='" & escape_string(pid) & "'"
				cnn.execute(sSQL)
				call checknotifystock(pid)
			end if
		next
		for index=0 to newrows-1
			pid=getpost("yid"&index)
			pname=getpost("yna"&index)
			if pname="" then pname=opname
			pname2=getpost("ynb"&index)
			pname3=getpost("ync"&index)
			if pname2="" then pname2=IIfVr(opname2<>"",opname2,opname)
			if pname3="" then pname3=IIfVr(opname3<>"",opname3,opname)
			if (adminlangsettings AND 1)<>1 OR adminlanguages<1 then pname2=""
			if (adminlangsettings AND 1)<>1 OR adminlanguages<2 then pname3=""
			pprice=getpost("ypr"&index)
			if NOT is_numeric(pprice) then pprice=opprice
			pwprice=getpost("ywp"&index)
			if NOT is_numeric(pwprice) then pwprice=opwprice
			pweight=getpost("ywe"&index)
			if NOT is_numeric(pweight) then pweight=opweight
			pinstock=getpost("ysk"&index)
			if NOT is_numeric(pinstock) then pinstock=0
			pexemptions=0
			if getpost("yst"&index)="1" then pexemptions=1
			if getpost("yct"&index)="1" then pexemptions=pexemptions+2
			if getpost("ysh"&index)="1" then pexemptions=pexemptions+4
			if getpost("yha"&index)="1" then pexemptions=pexemptions+8
			if getpost("yfs"&index)="1" then pexemptions=pexemptions+16
			if getpost("ypt"&index)="1" then pexemptions=pexemptions+32
			if getpost("ypd"&index)="1" then pexemptions=pexemptions+64
			pimage=getpost("ysmim"&index)
			plgimage=getpost("ylgim"&index)
			pgtimage=getpost("ygtim"&index)
			if getpost("ycr"&index)="1" then
				cnn.execute("DELETE FROM productimages WHERE imageProduct='" & escape_string(pid) & "'")
				if pimage<>"" then cnn.execute("INSERT INTO productimages (imageProduct,imageSrc,imageNumber,imageType) VALUES ('" & escape_string(pid) & "','" & escape_string(pimage) & "',0,0)")
				if plgimage<>"" then cnn.execute("INSERT INTO productimages (imageProduct,imageSrc,imageNumber,imageType) VALUES ('" & escape_string(pid) & "','" & escape_string(plgimage) & "',0,1)")
				if pgtimage<>"" then cnn.execute("INSERT INTO productimages (imageProduct,imageSrc,imageNumber,imageType) VALUES ('" & escape_string(pid) & "','" & escape_string(pgtimage) & "',0,2)")
				sSQL="INSERT INTO products (pID,pName,pName2,pName3,pPrice,pWholesalePrice,pWeight,pInStock,pExemptions,pSection,pDisplay) VALUES (" & _
					"'" & escape_string(pid) & "'" & _
					",'" & escape_string(pname) & "'" & _
					",'" & escape_string(pname2) & "'" & _
					",'" & escape_string(pname3) & "'" & _
					"," & pprice & _
					"," & pwprice & _
					"," & pweight & _
					"," & pinstock & _
					"," & pexemptions & _
					"," & psection & _
					",0)"
				cnn.execute(sSQL)
			end if
		next
%>
      <table border="" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
			<td align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
					<%=yyNoAuto%> <a href="adminprods.asp<%
						print "?rid=" & urlencode(getpost("rid")) & "&pid=" & urlencode(getpost("pid")) & "&disp=" & getpost("disp") & "&stext=" & urlencode(getpost("stext")) & "&sprice=" & urlencode(getpost("sprice")) & "&stype=" & getpost("stype") & "&scat=" & getpost("scat") & "&pg=" & getpost("pg")
					%>"><strong><%=yyClkHer%></strong></a>.<br /><br />&nbsp;<br />&nbsp;
			</td>
        </tr>
      </table>
<%
  elseif getpost("posted")="1" AND getpost("act")="altids" then %>
<script>
/* <![CDATA[ */
	function cr(trow,ischecked){
		for(var index=1;index<=6;index++){
			if(document.getElementById('z'+index+'a'+trow)){
				document.getElementById('z'+index+'a'+trow).style.display=ischecked?'none':'';
				document.getElementById('z'+index+'b'+trow).style.display=ischecked?'':'none';
			}
		}
	}
	function docreateall(telem){
		for(var index=0;index<parseInt(document.getElementById('newrows').value);index++){
			document.getElementById('ycr'+index).checked=telem.checked;
			cr(index,telem.checked);
		}
	}
	function dodeleteall(telem){
		for(var index=0;index<parseInt(document.getElementById('existingrows').value);index++){
			document.getElementById('xde'+index).checked=telem.checked;
		}
	}
	function displaymultilangname(isxy,index){
		if(document.getElementById(isxy+'nb'+index))document.getElementById(isxy+'nb'+index).style.display='block';
		if(document.getElementById(isxy+'nc'+index))document.getElementById(isxy+'nc'+index).style.display='block';
	}
/* ]]> */
</script>  
<%
		idlist=getpost("id")
		existingrows=0
		newrows=0
		sSQL="SELECT poOptionGroup FROM prodoptions WHERE poProdID='" & escape_string(getpost("id")) & "' ORDER BY poID"
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			sSQL="SELECT optGroup,optName,optRegExp FROM options WHERE optGroup=" & rs("poOptionGroup") & " AND optRegExp<>'' AND NOT (optRegExp IS NULL)"
			rs2.open sSQL,cnn,0,1
			if NOT rs2.EOF then
				newids=""
				idarray=split(trim(idlist)," ")
				do while NOT rs2.EOF
					for index=0 TO UBOUND(idarray)
						theid=idarray(index)
						theexp=trim(rs2("optRegExp")&"")
						if theexp<>"" AND left(theexp,1)<>"!" then
							theexp=replace(theexp, "%s", theid)
							if instr(theexp, " ")>0 then ' Search and replace
								exparr=split(theexp, " ", 2)
								theid=replace(theid, exparr(0), exparr(1), 1, 1)
							else
								theid=theexp
							end if
						end if
						newids=newids&theid&" "
					next
					rs2.movenext
				loop
				idlist=newids
			end if
			rs2.close
			rs.movenext
		loop
		rs.close
		if trim(idlist)=getpost("id") OR trim(idlist)="" then
			print "<div style=""text-align:center;margin:50px"">There are no product options with Alternate Product ID's assigned to this product</div>"
			print "<div style=""text-align:center;margin:50px""><input type=""button"" value="""&yyClkBac&""" onclick=""history.go(-1)"" /></div>"
		else
			idarray=split(trim(idlist)," ")
			sSQL="SELECT pID,pName,pName2,pName3,pPrice,pWholesalePrice,pWeight,pInStock,pExemptions FROM products WHERE pID IN ("
			for index=0 to UBOUND(idarray)
				sSQL=sSQL & "'" & escape_string(idarray(index)) & "'"
				if index<>UBOUND(idarray) then sSQL=sSQL & ","
			next
			sSQL=sSQL & ")"
			print "<form method=""post"" action=""adminprods.asp"">"
			print whv("act","altids") & whv("posted",1) & whv("doupdate",1) & whv("originalid",getpost("id"))
			call writehiddenvar("disp", getpost("disp"))
			call writehiddenvar("stext", getpost("stext"))
			call writehiddenvar("sprice", getpost("sprice"))
			call writehiddenvar("scat", getpost("scat"))
			call writehiddenvar("stype", getpost("stype"))
			call writehiddenvar("sos", getpost("sos"))
			call writehiddenvar("pg", getpost("pg"))
			print "<table width=""100%"" class=""admin-table-a"">"
			print "<tr><th>ID</th><th>Product Name</th><th class=""minicell"">"&yyPrPri&" / WS</th><th class=""minicell"" style=""width:5%"">Weight</th>"
			if useStockManagement then print "<th class=""minicell"" style=""width:5%"">Stock</th>"
			print "<th class=""minicell"">Exemptions</th><th style=""width:30%"">Images</th><th class=""minicell"" style=""white-space:nowrap;width:5%""><input type=""checkbox"" title=""Check All"" style=""vertical-align:top"" onclick=""docreateall(this)"" /> Create</th><th class=""minicell"" style=""white-space:nowrap;width:5%""><input type=""checkbox"" title=""Check All"" style=""vertical-align:top"" onclick=""dodeleteall(this)"" /> Delete</th></tr>"
			rs.open sSQL,cnn,0,1
			rowcounter=0
			do while NOT rs.EOF
				image1="" : image2="" : image3=""
				sSQL="SELECT imageSrc,imageType FROM productimages WHERE imageProduct='" & escape_string(rs("pID")) & "' AND imageNumber=0"
				rs2.open sSQL,cnn,0,1
				do while NOT rs2.EOF
					if rs2("imageType")=0 then image1=trim(rs2("imageSrc"))
					if rs2("imageType")=1 then image2=trim(rs2("imageSrc"))
					if rs2("imageType")=2 then image3=trim(rs2("imageSrc"))
					rs2.movenext
				loop
				rs2.close
				pExemptions=rs("pExemptions")
				print "<tr><td>" & whv("xid" & rowcounter,rs("pID")) & rs("pID") & "</td>" & _
					"<td>"
					print "<input type=""text"" name=""xna"&rowcounter&""" value=""" & htmlspecials(rs("pName")) & """ size=""25"" onmouseover=""displaymultilangname('x',"&rowcounter&")"" />"
					if (adminlangsettings AND 1)=1 then
						if adminlanguages>=1 then print "<input type=""text"" style=""display:none;margin-top:2px"" name=""xnb"&rowcounter&""" id=""xnb"&rowcounter&""" size=""25"" placeholder=""Product Name Language 2"" value=""" & htmlspecialsucode(rs("pName2")) & """ />"
						if adminlanguages>=2 then print "<input type=""text"" style=""display:none;margin-top:2px"" name=""xnc"&rowcounter&""" id=""xnc"&rowcounter&""" size=""25"" placeholder=""Product Name Language 3"" value=""" & htmlspecialsucode(rs("pName3")) & """ />"
					end if
					print "</td>" & _
					"<td class=""minicell"" style=""white-space:nowrap"">" & _
						"<input type=""text"" id=""xpr"&rowcounter&""" name=""xpr"&rowcounter&""" value=""" & rs("pPrice") & """ size=""7"" title="""&yyPrPri&""" onfocus=""document.getElementById('xwp"&rowcounter&"').size=4;this.size=7"" onkeyup=""checkrequiredfields()"" />" & _
						" <input type=""text"" id=""xwp"&rowcounter&""" name=""xwp"&rowcounter&""" size=""4"" value=""" & rs("pWholesalePrice") & """ placeholder=""" & yyWhoPri & """ title=""" & yyWhoPri & """ onfocus=""document.getElementById('xpr"&rowcounter&"').size=4;this.size=7"" />" & _
					"</td>" & _
					"<td align=""center""><input type=""text"" name=""xwe"&rowcounter&""" value=""" & rs("pWeight") & """ size=""5"" /></td>"
				if useStockManagement then print "<td align=""center""><input type=""text"" name=""xsk"&rowcounter&""" value=""" & rs("pInStock") & """ size=""4"" /></td>"
				print "<td style=""white-space:nowrap;text-align:center"">" & _
						"<input type=""checkbox"" name=""xst"&rowcounter&""" value=""1"" title="""&yyExStat&""" "&IIfVs((pExemptions AND 1)=1,"checked=""checked"" ")&"/>" & _
						"<input type=""checkbox"" name=""xct"&rowcounter&""" value=""1"" title="""&yyExCoun&""" "&IIfVs((pExemptions AND 2)=2,"checked=""checked"" ")&"/>" & _
						"<input type=""checkbox"" name=""xsh"&rowcounter&""" value=""1"" title="""&yyExShip&""" "&IIfVs((pExemptions AND 4)=4,"checked=""checked"" ")&"/>" & _
						"<input type=""checkbox"" name=""xha"&rowcounter&""" value=""1"" title="""&yyExHand&""" "&IIfVs((pExemptions AND 8)=8,"checked=""checked"" ")&"/>" & _
						"<input type=""checkbox"" name=""xfs"&rowcounter&""" value=""1"" title="""&"Free Shipping Exempt"&""" "&IIfVs((pExemptions AND 16)=16,"checked=""checked"" ")&"/>" & _
						"<input type=""checkbox"" name=""xpt"&rowcounter&""" value=""1"" title="""&"Pack Together Exempt"&""" "&IIfVs((pExemptions AND 32)=32,"checked=""checked"" ")&"/>" & _
						"<input type=""checkbox"" name=""xpd"&rowcounter&""" value=""1"" title="""&"Product Discount Exempt"&""" "&IIfVs((pExemptions AND 64)=64,"checked=""checked"" ")&"/>" & _
					"</td>" & _
					"<td>" & _
						"<div class=""small""><input type=""text"" value="""&htmlspecials(image1)&""" placeholder="""&yyImage&""" name=""xsmim"&rowcounter&""" style=""width:99%"" onmouseover=""document.getElementById('xlgim"&rowcounter&"').style.display='';document.getElementById('xgtim"&rowcounter&"').style.display=''"" /></div>" & _
						"<div class=""small""><input type=""text"" value="""&htmlspecials(image2)&""" placeholder="""&yyLgeImg&""" name=""xlgim"&rowcounter&""" id=""xlgim"&rowcounter&""" style=""display:none;width:99%"" /></div>" & _
						"<div class=""small""><input type=""text"" value="""&htmlspecials(image3)&""" placeholder="""&yyGiaImg&""" name=""xgtim"&rowcounter&""" id=""xgtim"&rowcounter&""" style=""display:none;width:99%"" /></div>" & _
					"</td>" & _
					"<td align=""center"">-</td><td align=""center""><input type=""checkbox"" id=""xde"&rowcounter&""" name=""xde"&rowcounter&""" value=""1"" /></td></tr>"
				for index=0 to UBOUND(idarray)
					if rs("pID")=idarray(index) then idarray(index)=""
				next
				rowcounter=rowcounter+1
				rs.movenext
			loop
			existingrows=rowcounter
			rowcounter=0
			for index=0 to UBOUND(idarray)
				if idarray(index)<>"" then
					print "<tr>" & _
						"<td>" & whv("yid" & rowcounter,idarray(index)) & idarray(index) & "</td>" & _
						"<td><div style=""text-align:center"" id=""z1a"&rowcounter&""">-</div>" & _
						"<div id=""z1b"&rowcounter&""" style=""display:none""><input type=""text"" name=""yna"&rowcounter&""" value="""" size=""25"" onmouseover=""displaymultilangname('y',"&rowcounter&")"" />"
					if (adminlangsettings AND 1)=1 then
						if adminlanguages>=1 then print "<input type=""text"" style=""display:none;margin-top:2px"" name=""ynb"&rowcounter&""" id=""ynb"&rowcounter&""" size=""25"" placeholder=""Product Name Language 2"" value="""" />"
						if adminlanguages>=2 then print "<input type=""text"" style=""display:none;margin-top:2px"" name=""ync"&rowcounter&""" id=""ync"&rowcounter&""" size=""25"" placeholder=""Product Name Language 3"" value="""" />"
					end if
					print "</div></td>" & _
						"<td class=""minicell"" style=""white-space:nowrap""><div style=""text-align:center"" id=""z2a"&rowcounter&""">-</div><div id=""z2b"&rowcounter&""" style=""display:none"">" & _
							"<input type=""text"" id=""ypr"&rowcounter&""" name=""ypr"&rowcounter&""" value="""" size=""7"" title="""&yyPrPri&""" onfocus=""document.getElementById('ywp"&rowcounter&"').size=4;this.size=7"" onkeyup=""checkrequiredfields()"" />" & _
							" <input type=""text"" id=""ywp"&rowcounter&""" name=""ywp"&rowcounter&""" size=""4"" value="""" placeholder=""" & yyWhoPri & """ title=""" & yyWhoPri & """ onfocus=""document.getElementById('ypr"&rowcounter&"').size=4;this.size=7"" />" & _
						"</div></td>" & _
						"<td align=""center""><div style=""text-align:center"" id=""z3a"&rowcounter&""">-</div><div id=""z3b"&rowcounter&""" style=""display:none""><input type=""text"" name=""ywe"&rowcounter&""" value="""" size=""5"" /></div></td>"
					if useStockManagement then print "<td><div style=""text-align:center"" id=""z4a"&rowcounter&""">-</div><div id=""z4b"&rowcounter&""" style=""display:none;text-align:center""><input type=""text"" name=""ysk"&rowcounter&""" value="""" size=""4"" /></div></td>"
					print "<td style=""white-space:nowrap;text-align:center""><div style=""text-align:center"" id=""z5a"&rowcounter&""">-</div><div id=""z5b"&rowcounter&""" style=""display:none"">" & _
							"<input type=""checkbox"" name=""yst"&rowcounter&""" value=""1"" title="""&yyExStat&""" "&IIfVs((pExemptions AND 1)=1,"checked=""checked"" ")&"/>" & _
							"<input type=""checkbox"" name=""yct"&rowcounter&""" value=""1"" title="""&yyExCoun&""" "&IIfVs((pExemptions AND 2)=2,"checked=""checked"" ")&"/>" & _
							"<input type=""checkbox"" name=""ysh"&rowcounter&""" value=""1"" title="""&yyExShip&""" "&IIfVs((pExemptions AND 4)=4,"checked=""checked"" ")&"/>" & _
							"<input type=""checkbox"" name=""yha"&rowcounter&""" value=""1"" title="""&yyExHand&""" "&IIfVs((pExemptions AND 8)=8,"checked=""checked"" ")&"/>" & _
							"<input type=""checkbox"" name=""yfs"&rowcounter&""" value=""1"" title="""&"Free Shipping Exempt"&""" "&IIfVs((pExemptions AND 16)=16,"checked=""checked"" ")&"/>" & _
							"<input type=""checkbox"" name=""ypt"&rowcounter&""" value=""1"" title="""&"Pack Together Exempt"&""" "&IIfVs((pExemptions AND 32)=32,"checked=""checked"" ")&"/>" & _
							"<input type=""checkbox"" name=""ypd"&rowcounter&""" value=""1"" title="""&"Product Discount Exempt"&""" "&IIfVs((pExemptions AND 64)=64,"checked=""checked"" ")&"/>" & _
						"</div></td>" & _
						"<td><div style=""text-align:center"" id=""z6a"&rowcounter&""">-</div><div id=""z6b"&rowcounter&""" style=""display:none"">" & _
							"<div class=""small""><input type=""text"" value="""" name=""ysmim"&rowcounter&""" style=""width:99%"" onmouseover=""document.getElementById('ylgim"&rowcounter&"').style.display='';document.getElementById('ygtim"&rowcounter&"').style.display=''"" placeholder="""&yyImage&""" /></div>" & _
							"<div class=""small""><input type=""text"" value="""" name=""ylgim"&rowcounter&""" id=""ylgim"&rowcounter&""" style=""display:none;width:99%"" placeholder="""&yyLgeImg&""" /></div>" & _
							"<div class=""small""><input type=""text"" value="""" name=""ygtim"&rowcounter&""" id=""ygtim"&rowcounter&""" style=""display:none;width:99%"" placeholder="""&yyGiaImg&""" /></div>" & _
						"</div></td>" & _
						"<td align=""center""><input type=""checkbox"" id=""ycr"&rowcounter&""" name=""ycr"&rowcounter&""" value=""1"" onchange=""cr("&rowcounter&",this.checked)"" /></td><td align=""center"">-</td>" & _
					"</tr>"
					for index2=index+1 to UBOUND(idarray)
						if idarray(index)=idarray(index2) then idarray(index2)=""
					next
					rowcounter=rowcounter+1
				end if
			next
			newrows=rowcounter
			print "</table>"
			print "<div style=""text-align:center""><input type=""submit"" value=""" & yySubmit & """ /></div>"
			call writehiddenidvar("existingrows",existingrows)
			call writehiddenidvar("newrows",newrows)
			print "</form>"
		end if
  elseif getpost("posted")="1" AND (getpost("act")="modify" OR getpost("act")="clone" OR getpost("act")="addnew") then
		Dim pNames(10)
		if htmleditor="ckeditor" then %>
<script src="ckeditor/ckeditor.js"></script>
<%		end if
		maximagenumber=-1
		imageindex=0
		smimgindx=0
		lgimgindx=0
		gtimgindx=0
		numsmimgs=-1
		numlgimgs=-1
		numgtimgs=-1
		sub getnext3images(byref smimg,byref lgimg,byref gtimg)
			smimg="" : lgimg="" : gtimg=""
			if smimgindx<numsmimgs then smimg=allsmimgs(0,smimgindx) : smimgindx=smimgindx+1 else smimg=""
			if lgimgindx>=numlgimgs then
				if gtimgindx>=numgtimgs then gtimg="" else gtimg=allgtimgs(0,gtimgindx) : gtimgindx=gtimgindx+1
			elseif gtimgindx>=numgtimgs then
				if lgimgindx>=numlgimgs then lgimg="" else lgimg=alllgimgs(0,lgimgindx) : lgimgindx=lgimgindx+1
			elseif alllgimgs(1,lgimgindx) > allgtimgs(1,gtimgindx) then
				gtimg=allgtimgs(0,gtimgindx) : gtimgindx=gtimgindx+1
			elseif alllgimgs(1,lgimgindx) < allgtimgs(1,gtimgindx) then
				lgimg=alllgimgs(0,lgimgindx) : lgimgindx=lgimgindx+1
			else
				lgimg=alllgimgs(0,lgimgindx) : lgimgindx=lgimgindx+1
				gtimg=allgtimgs(0,gtimgindx) : gtimgindx=gtimgindx+1
			end if
		end sub
		sub displayimagerow(imgrow,smimg,lgimg,gtimg)
			print "<tr>"
			print "<td style=""white-space:nowrap""><input type=""text"" name=""smim" & imgrow & """ id=""smim" & imgrow & """ value=""" & htmlspecials(smimg) & """ style=""width:85%"" " & IIfVr(imgrow=0, "onchange=""document.getElementById('pImage').value=this.value""", "") & "/>&nbsp;<input type=""button"" value=""..."" onclick=""uploadimage('smim" & imgrow & "')"" /></td>"
			print "<td style=""white-space:nowrap""><input type=""text"" name=""lgim" & imgrow & """ id=""lgim" & imgrow & """ value=""" & htmlspecials(lgimg) & """ style=""width:85%"" " & IIfVr(imgrow=0, "onchange=""document.getElementById('pLargeImage').value=this.value""", "") & "/>&nbsp;<input type=""button"" value=""..."" onclick=""uploadimage('lgim" & imgrow & "')"" /></td>"
			print "<td style=""white-space:nowrap""><input type=""text"" name=""gtim" & imgrow & """ id=""gtim" & imgrow & """ value=""" & htmlspecials(gtimg) & """ style=""width:85%"" " & IIfVr(imgrow=0, "onchange=""document.getElementById('pGiantImage').value=this.value""", "") & "/>&nbsp;<input type=""button"" value=""..."" onclick=""uploadimage('gtim" & imgrow & "')"" /></td>"
			print "</tr>"
		end sub
		doaddnew=TRUE
		if getpost("act")="modify" OR getpost("act")="clone" then
			sSQL="SELECT imageSrc,imageNumber,imageType FROM productimages WHERE imageProduct='" & escape_string(getpost("id")) & "' AND imageType=0 ORDER BY imageNumber"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then allsmimgs=rs.getrows() : numsmimgs=UBOUND(allsmimgs,2)+1 : maximagenumber=vrmax(maximagenumber,numsmimgs)
			rs.close
			sSQL="SELECT imageSrc,imageNumber,imageType FROM productimages WHERE imageProduct='" & escape_string(getpost("id")) & "' AND imageType=1 ORDER BY imageNumber"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then alllgimgs=rs.getrows() : numlgimgs=UBOUND(alllgimgs,2)+1 : maximagenumber=vrmax(maximagenumber,numlgimgs)
			rs.close
			sSQL="SELECT imageSrc,imageNumber,imageType FROM productimages WHERE imageProduct='" & escape_string(getpost("id")) & "' AND imageType=2 ORDER BY imageNumber"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then allgtimgs=rs.getrows() : numgtimgs=UBOUND(allgtimgs,2)+1 : maximagenumber=vrmax(maximagenumber,numgtimgs)
			rs.close
			sSQL="SELECT pId,pName,pName2,pName3,pSection,pPrice,pWholesalePrice,pListPrice,pDisplay,pStaticPage,pStaticURL,pRecommend,pStockByOpts,pSell,pShipping,pShipping2,pWeight,pExemptions,pInStock,pDims,pTax,pDropship,pManufacturer,pSKU,pOrder,pDateAdded,"
			if digidownloads=true then sSQL=sSQL & "pDownload,"
			sSQL=sSQL & "pGiftWrap,pBackOrder,pSearchParams,pTitle,pTitle2,pTitle3,pMetaDesc,pMetaDesc2,pMetaDesc3,pCustom1,pCustom2,pCustom3,pDescription,pLongDescription FROM products WHERE pId='"&escape_string(getpost("id"))&"'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				doaddnew=FALSE
				pName=rs("pName")
				for index=2 to adminlanguages+1
					pNames(index)=rs("pName" & index)
				next
				pID=rs("pID")
				pSection=rs("pSection")
				pPrice=rs("pPrice")
				pWholesalePrice=rs("pWholesalePrice")
				pListPrice=rs("pListPrice")
				pDisplay=rs("pDisplay")
				pStaticPage=rs("pStaticPage")
				pStaticURL=trim(rs("pStaticURL")&"")
				pRecommend=rs("pRecommend")
				pStockByOpts=rs("pStockByOpts")
				pSell=rs("pSell")
				pShipping=rs("pShipping")
				pShipping2=rs("pShipping2")
				pWeight=rs("pWeight")
				pExemptions=rs("pExemptions")
				pInStock=rs("pInStock")
				pDims=rs("pDims")
				if digidownloads=true then pDownload=rs("pDownload")
				pTax=rs("pTax")
				pDropship=rs("pDropship")
				pManufacturer=rs("pManufacturer")
				pSKU=rs("pSKU")
				pOrder=rs("pOrder")
				if getpost("act")="clone" then pDateAdded=Date() else pDateAdded=rs("pDateAdded")
				if isnull(pDateAdded) then pDateAdded=Date()
				pGiftWrap=rs("pGiftWrap")
				pBackOrder=cint(rs("pBackOrder"))
				pSearchParams=rs("pSearchParams")
				pTitle=rs("pTitle")
				pTitle2=rs("pTitle2")
				pTitle3=rs("pTitle3")
				pMetaDesc=rs("pMetaDesc")
				pMetaDesc2=rs("pMetaDesc2")
				pMetaDesc3=rs("pMetaDesc3")
				pCustom1=rs("pCustom1")
				pCustom2=rs("pCustom2")
				pCustom3=rs("pCustom3")
				pDescription=rs("pDescription")
				pLongDescription=rs("pLongDescription")
			end if
			rs.close
		end if
		if doaddnew then
			pID=""
			if getpost("scat")<>"" then pSection=int(getpost("scat")) else pSection=0
			pImage=defaultprodimages
			pPrice=""
			pWholesalePrice=""
			pListPrice=0
			pDisplay=1
			pStaticPage=0
			pStaticURL=""
			pRecommend=0
			pStockByOpts=0
			pSell=1
			pShipping=""
			pShipping2=""
			pLargeImage=defaultprodimages
			pGiantImage=""
			pWeight=""
			pExemptions=0
			pInStock=""
			pDims=""
			pDownload=""
			pTax=0
			pDropship=0
			pManufacturer=0
			pSKU=""
			pOrder=0
			pDateAdded=Date()
			pGiftWrap=0
			pBackOrder=0
			pSearchParams=""
			pTitle=""
			pTitle2=""
			pTitle3=""
			pMetaDesc=""
			pMetaDesc2=""
			pMetaDesc3=""
			pCustom1=""
			pCustom2=""
			pCustom3=""
			pDescription=""
			pLongDescription=""
		end if
%>
<script>
/* <![CDATA[ */
var oAR=new Array();
var sAR=new Array();
var cAR=new Array();
<%	sSQL="SELECT optGrpID,optGrpWorkingName,optType FROM optiongroup ORDER BY optGrpWorkingName"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then alloptions=rs.getrows
	rs.close
	if isarray(alloptions) then
		for rowcounter=0 to UBOUND(alloptions,2)
			print "oAR["&rowcounter&"]=["&alloptions(0,rowcounter)&",'"&jsescape(alloptions(1,rowcounter))&"',"&alloptions(2,rowcounter)&"];"&vbCrLf
		next
	end if
	sSQL="SELECT sectionID,sectionWorkingName FROM sections WHERE rootSection=1 ORDER BY sectionWorkingName"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then allsections=rs.getrows
	rs.close
	if isarray(allsections) then
		for rowcounter=0 to UBOUND(allsections,2)
			print "sAR["&rowcounter&"]=["&allsections(0,rowcounter)&",'"&jsescape(allsections(1,rowcounter))&"'];"&vbCrLf
		next
	end if
	sSQL="SELECT scID,scWorkingName FROM searchcriteria WHERE scGroup<>0 ORDER BY scGroup,scOrder,scName"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then allsearchcriteria=rs.getrows
	rs.close
	if isarray(allsearchcriteria) then
		for rowcounter=0 to UBOUND(allsearchcriteria,2)
			print "cAR["&rowcounter&"]=["&allsearchcriteria(0,rowcounter)&",'"&jsescape(allsearchcriteria(1,rowcounter))&"'];"&vbCrLf
		next
	end if
%>
function checkastring(thestr,validchars){
  for (i=0; i < thestr.length; i++){
    ch=thestr.charAt(i);
    for (j=0;  j < validchars.length;  j++)
      if(ch==validchars.charAt(j))
        break;
    if(j==validchars.length)
	  return(false);
  }
  return(true);
}
function formvalidator(theForm){
  if(theForm.newid.value==""){
    alert("<%=jscheck(yyPlsEntr&" """&yyPrRef)%>\".");
    theForm.newid.focus();
    return(false);
  }
  if(theForm.psection.options[theForm.psection.selectedIndex].value==""){
    alert("<%=jscheck(yyPlsSel&" """&yySection)%>\".");
    theForm.psection.focus();
    return(false);
  }
  if(theForm.pName.value==""){
    alert("<%=jscheck(yyPlsEntr&" """&yyPrNam)%>\".");
    theForm.pName.focus();
    return(false);
  }
<%	for index=2 to adminlanguages+1
		if(adminlangsettings AND 1)=1 then %>
  if(theForm.pName<%=index%>.value==""){
	displaymultilangname('pName');
    alert("<%=jscheck(yyPlsEntr&" """&yyPrNam & " " & index)%>\".");
    theForm.pName<%=index%>.focus();
    return(false);
  }
<%		end if
	next %>
  if(theForm.pPrice.value==""){
    alert("<%=jscheck(yyPlsEntr&" """&yyPrPri)%>\".");
    theForm.pPrice.focus();
    return(false);
  }
  var checkOK="'\" ";
  var checkStr=theForm.newid.value;
  var allValid=true;
  for (i=0;  i < checkStr.length;  i++){
    ch=checkStr.charAt(i);
    for (j=0;  j < checkOK.length;  j++)
      if(ch==checkOK.charAt(j)){
	    allValid=false;
        break;
	  }
  }
  if(!allValid){
    alert("<%=jscheck(yyQuoSpa&" """&yyPrRef)%>\".");
    theForm.newid.focus();
    return(false);
  }
  if(!checkastring(theForm.pPrice.value,"0123456789.")){
    alert("<%=jscheck(yyOnlyDec&" """&yyPrPri)%>\".");
    theForm.pPrice.focus();
    return(false);
  }
  if(!checkastring(theForm.pWholesalePrice.value,"0123456789.")){
    alert("<%=jscheck(yyOnlyDec&" """&yyWhoPri)%>\".");
    theForm.pWholesalePrice.focus();
    return(false);
  }
  if(!checkastring(theForm.pListPrice.value,"0123456789.")){
    alert("<%=jscheck(yyOnlyDec&" """&yyListPr)%>\".");
    theForm.pListPrice.focus();
    return(false);
  }
<%	if(adminUnits AND 12) > 0 then %>
  var checkOK="0123456789.";
  if(!checkastring(theForm.plen.value,checkOK)){
	alert("<%=jscheck(yyOnlyDec&" """&yyDims)%>\".");
	theForm.plen.focus();
	return(false);
  }
  if(!checkastring(theForm.pwid.value,checkOK)){
	alert("<%=jscheck(yyOnlyDec&" """&yyDims)%>\".");
	theForm.pwid.focus();
	return(false);
  }
  if(!checkastring(theForm.phei.value,checkOK)){
	alert("<%=jscheck(yyOnlyDec&" """&yyDims)%>\".");
	theForm.phei.focus();
	return(false);
  }
<%	end if
	if usesshipweight then %>
  var checkOK="0123456789.";
  if(!checkastring(theForm.pWeight.value,checkOK)){
    alert("<%=jscheck(yyOnlyDec&" """&yyPrWght)%>\".");
    theForm.pWeight.focus();
    return(false);
  }
<%	end if
	if usesflatrate then %>
  var checkOK="0123456789.";
  if(!checkastring(theForm.pShipping.value,checkOK)){
    alert("<%=jscheck(yyOnlyDec&" """&yyFlatShp & ": " & yyFirShi)%>\".");
    theForm.pShipping.focus();
    return(false);
  }
  if(!checkastring(theForm.pShipping2.value,"0123456789.")){
    alert("<%=jscheck(yyOnlyDec&" """&yyFlatShp & ": " & yySubShi)%>\".");
    theForm.pShipping2.focus();
    return(false);
  }
<%	end if
	if useStockManagement then %>
  if(!(theForm.pStockByOpts.selectedIndex==1) && theForm.inStock.value==""){
    alert("<%=jscheck(yyPlsEntr&" """&yyInStk)%>\".");
    theForm.inStock.focus();
    return(false);
  }
  if(!(theForm.pStockByOpts.selectedIndex==1) && !checkastring(theForm.inStock.value,"0123456789-")){
    alert("<%=jscheck(yyOnlyNum&" """&yyInStk)%>\".");
    theForm.inStock.focus();
    return(false);
  }
  if(theForm.pStockByOpts.selectedIndex==1 && parseInt(document.getElementById('pnumoptions').value)==0){
    alert("<%=jscheck(yyStkWrn)%>");
    theForm.pStockByOpts.focus();
    return(false);
  }
<%	end if
	if perproducttaxrate=true then %>
  if(theForm.pTax.value==""){
	alert("<%=jscheck(yyPlsEntr&" """&yyTax)%>\".");
	theForm.pTax.focus();
	return(false);
  }
  if(!checkastring(theForm.pTax.value,"0123456789.")){
    alert("<%=jscheck(yyOnlyDec&" """&yyTax)%>\".");
    theForm.pTax.focus();
    return(false);
  }
<%	end if %>
  if(!checkastring(theForm.pOrder.value,"0123456789")){
    alert("<%=jscheck(yyOnlyNum&" """&yyProdOr)%>\".");
    theForm.pOrder.focus();
    return(false);
  }
	nummultioptions=0;
	for(index=0;index<parseInt(document.getElementById('pnumoptions').value);index++){
		var thisOption=document.getElementById('poption'+index);
		if(parseInt(thisOption.selectedIndex)!=0){
			var optval=parseInt(thisOption[thisOption.selectedIndex].value);
			for(var i=0;i<oAR.length;i++){
				if(oAR[i][0]==optval&&Math.abs(oAR[i][2])==4)nummultioptions++;
			}
		}
	}
	if(nummultioptions>1){
		alert("<%=jscheck(yyMBOUni)%>");
		theForm.poption0.focus();
		return(false);
	}
	if(document.getElementById('staticpage').selectedIndex==0)
		document.getElementById('pStaticURL').value='';
	return(true);
}
function populateoptionsselect(oSelect,optsect){
	var insbefore=oSelect.selectedIndex!=0;
	var existingitem=oSelect.options[oSelect.selectedIndex];
	var osarray;
	if(optsect=='option') osarray=oAR; else if(optsect=='search') osarray=cAR; else osarray=sAR;
	for(var i=0;i<osarray.length;i++){
		if(existingitem.value==osarray[i][0]){
			insbefore=false;
		}else{
			var y=document.createElement('option');
			y.innerHTML=osarray[i][1];
			y.value=osarray[i][0];
			if(insbefore){
				try{oSelect.add(y,existingitem);} // FF etc
				catch(ex){oSelect.add(y,oSelect.selectedIndex);} // IE
			}else{
				try{oSelect.add(y,null);} // FF etc
				catch(ex){oSelect.add(y);} // IE
			}
		}
	}
}
function addnewoption(thisindex,optsect){
	var pNumOpts=parseInt(document.getElementById("pnum"+optsect+"s").value);
	if(thisindex==pNumOpts){
		pNumOpts+=1;
		var stable=document.getElementById(optsect+'stable');
		newrow=stable.insertRow(-1);
		newcell=newrow.insertCell(-1);
		newcell.align='right';
		newcell.innerHTML=(pNumOpts+1);
		newcell=newrow.insertCell(-1);
		newcell.innerHTML='<select style="width:180px" size="1" id="p'+optsect+pNumOpts+'" name="p'+optsect+pNumOpts+'" onchange="addnewoption('+pNumOpts+',\''+optsect+'\');"><option value="0"><%=jsescape(yySelect)%></option></select>';
		document.getElementById("pnum"+optsect+"s").value=pNumOpts;
		populateoptionsselect(document.getElementById("p"+optsect+pNumOpts),optsect);
	}
}
function setprodoptions(optsect){
	var pNumOpts=document.getElementById("pnum"+optsect+"s").value;
	for(var numopts=0;numopts<=pNumOpts;numopts++){
		oSelect=document.getElementById("p"+optsect+numopts);
		populateoptionsselect(oSelect,optsect);
	}
}
function setstockcontrols(resctrl){
	if(document.forms.mainform.pStockByOpts.selectedIndex==1){
		document.getElementById('stocksetting').value='';
		document.getElementById('inStock').style.display='none';
		if(document.getElementById('stockbutton'))document.getElementById('stockbutton').style.display='none';
	}else if(resctrl){
		document.getElementById('stocksetting').value='';
		document.getElementById('inStock').style.display='none';
		if(document.getElementById('stockbutton'))document.getElementById('stockbutton').style.display='';
	}else{
		document.getElementById('stocksetting').value='1';
		document.getElementById('inStock').style.display='';
		if(document.getElementById('stockbutton'))document.getElementById('stockbutton').style.display='none';
		document.getElementById('inStock').focus();
	}
}
function setstocktype(){
var si=document.forms.mainform.pStockByOpts.selectedIndex;
document.forms.mainform.inStock.disabled=(si==1);
document.getElementById('setbyopts').style.display=(si==1?'':'none');
<%	if getpost("act")="modify" then %>
setstockcontrols(true);
<%	else %>
document.getElementById('inStock').style.display=(si==1?'none':'');
<%	end if %>
}
function uploadimage(imfield){
	var winwid=400; var winhei=300;
	var prnttext='<html><head><link rel="stylesheet" type="text/css" href="adminstyle.css"/></head><body>\n';
	prnttext+='<form name="mainform" method="post" action="doupload.asp?defimagepath=<%=defaultprodimages%>" enctype="multipart/form-data">';
	prnttext+='<input type="hidden" name="defimagepath" value="<%=defaultprodimages%>" />';
	prnttext+='<input type="hidden" name="imagefield" value="'+imfield+'" />';
	prnttext+='<table border="" cellspacing="1" cellpadding="3" width="100%">';
	prnttext+='<tr><td align="center" colspan="2">&nbsp;<br /><strong><%=replace(yyUplIma, "'", "\'")%></strong><br />&nbsp;</td></tr>';
	prnttext+='<tr><td align="center" colspan="2"><%=replace(yyPlsSUp, "'", "\'")%><br />&nbsp;</td></tr>';
	prnttext+='<tr><td align="right"><%=replace(yyLocIma, "'", "\'")%>:</td><td><input type="file" name="imagefile" /></td></tr>';
	prnttext+='<tr><td colspan="2" align="center">&nbsp;<br /><input type="submit" value="<%=replace(yySubmit, "'", "\'")%>" /></td></tr>';
	prnttext+='</table></form>';
	prnttext+='<p align="center"><a href="javascript:window.close()"><strong><%=replace(yyClsWin, "'", "\'")%></strong></a></p>';
	prnttext+='</body></'+'html>';
	scrwid=screen.width; scrhei=screen.height;
	var newwin=window.open("","uploadimage",'menubar=no,scrollbars=yes,width='+winwid+',height='+winhei+',left='+((scrwid-winwid)/2)+',top=100,directories=no,location=no,resizable=yes,status=no,toolbar=no');
	newwin.document.open();
	newwin.document.write(prnttext);
	newwin.document.close();
	newwin.focus();
}
function imagemanager(){
	if(document.getElementById('extraimages').style.display=='none'){
		document.getElementById('extraimages').style.display='';
		document.getElementById('lessimages').style.display='none';
		document.getElementById('lessimages2').style.display='none';
		document.getElementById('but_pImage').value="<%=yyClose&" "&yyImgMgr%>";
		document.getElementById('pImage').disabled=true;
		document.getElementById('smallimup').style.display='none';
		document.getElementById('moreimages').style.display='';
	}else{
		document.getElementById('extraimages').style.display='none';
		document.getElementById('lessimages').style.display='';
		document.getElementById('lessimages2').style.display='';
		document.getElementById('but_pImage').value="<%=yyImgMgr%>";
		document.getElementById('pImage').disabled=false;
		document.getElementById('smallimup').style.display='';
		document.getElementById('moreimages').style.display='none';
	}
}
function moreimagefn(){
	var thetable=document.getElementById('extraimagetable');
	var currmax=parseInt(document.getElementById('maximgindex').value);
	for(imindx=currmax; imindx<currmax+5; imindx++){
		newrow=thetable.insertRow(-1);
		newcell=newrow.insertCell(0);
		newcell.style.whiteSpace='nowrap';
		newcell.innerHTML='<input type="text" name="smim' + imindx + '" id="smim' + imindx + '" value="" style="width:85%" />&nbsp;<input type="button"" value="..." onclick="uploadimage(\'smim' + imindx + '\')" />';
		newcell=newrow.insertCell(1);
		newcell.style.whiteSpace='nowrap';
		newcell.innerHTML='<input type="text" name="lgim' + imindx + '" id="lgim' + imindx + '" value="" style="width:85%" />&nbsp;<input type="button"" value="..." onclick="uploadimage(\'lgim' + imindx + '\')" />';
		newcell=newrow.insertCell(2);
		newcell.style.whiteSpace='nowrap';
		newcell.innerHTML='<input type="text" name="gtim' + imindx + '" id="gtim' + imindx + '" value="" style="width:85%" />&nbsp;<input type="button"" value="..." onclick="uploadimage(\'gtim' + imindx + '\')" />';
	}
	document.getElementById('maximgindex').value=imindx;
}
function setstatic(setting){
	if(setting==0){
		document.getElementById('staticpagediv').style.display='';
		document.getElementById('staticurldiv').style.display='none';
	}else{
		document.getElementById('staticpagediv').style.display='none'
		document.getElementById('staticurldiv').style.display='';
	}
}
function displaymultilangname(elem){
	for(var index=2;index<=3;index++){
		if(document.getElementById(elem+index))document.getElementById(elem+index).style.display='block';
	}
}
function getectobj(objid){
	return(document.getElementById(objid));
}
function expandckeditor(objtxt,editornumber){
	for(index=1;index<=3;index++){
		if(document.getElementById('editordiv'+objtxt.substr(0,4)+index)){
			document.getElementById('editordiv'+objtxt.substr(0,4)+index).style.display='block';
<%	if htmleditor="froala" then %>
			if(index!=1) eval("dfe_"+objtxt.replace(/pDes/,"pDescription").replace(/pLon/,"pLongDescription")+index+"()");
<%	end if %>
		}
	}
	if(document.getElementById('editordiv'+objtxt.substr(0,4)+editornumber)){
		document.getElementById('editordiv'+objtxt.substr(0,4)+editornumber).style.border='none';
		document.getElementById('editordiv'+objtxt.substr(0,4)+editornumber).style.padding=0;
	}
	getectobj('descshort').style.width=objtxt.substr(0,4)=='pDes'?'60%':'40%';
	getectobj('desclong').style.width=objtxt.substr(0,4)=='pDes'?'40%':'60%';
}
function displaymultilangdescs(islongdesc,thisobj){
	var setobj;
	for(var index=2;index<=3;index++){
		if(document.getElementById('pDescription'+index))document.getElementById('pDescription'+index).style.display='block';
		if(document.getElementById('pLongDescription'+index))document.getElementById('pLongDescription'+index).style.display='block';
	}
	for(var index=1;index<=3;index++){
		if(!islongdesc){
			if(setobj=getectobj('pDescription'+(index==1?'':index)))setobj.style.width='500px';
			if(setobj=getectobj('pLongDescription'+(index==1?'':index)))setobj.style.width='300px';
			if(index==thisobj){
				if(setobj=getectobj('pDescription'+(index==1?'':index)))setobj.style.height='200px';
				if(setobj=getectobj('pLongDescription'+(index==1?'':index)))setobj.style.height='100px';
			}else{
				if(setobj=getectobj('pDescription'+(index==1?'':index)))setobj.style.height='100px';
				if(setobj=getectobj('pLongDescription'+(index==1?'':index)))setobj.style.height='100px';
			}
		}
		if(islongdesc){
			if(setobj=getectobj('pDescription'+(index==1?'':index)))setobj.style.width='300px';
			if(setobj=getectobj('pLongDescription'+(index==1?'':index)))setobj.style.width='500px';
			if(index==thisobj){
				if(setobj=getectobj('pDescription'+(index==1?'':index)))setobj.style.height='100px';
				if(setobj=getectobj('pLongDescription'+(index==1?'':index)))setobj.style.height='200px';
			}else{
				if(setobj=getectobj('pDescription'+(index==1?'':index)))setobj.style.height='100px';
				if(setobj=getectobj('pLongDescription'+(index==1?'':index)))setobj.style.height='100px';
			}
		}
	}
}
function checkrequiredfields(){
	document.getElementById('newid').style.borderColor=(document.getElementById('newid').value.replace(/ /g,'')==''?'red':'');
	document.getElementById('pName').style.borderColor=(document.getElementById('pName').value.replace(/ /g,'')==''?'red':'');
	document.getElementById('pPrice').style.borderColor=(document.getElementById('pPrice').value.replace(/ /g,'')==''?'red':'');
	document.getElementById('psection').style.borderColor=(document.getElementById('psection').selectedIndex==0?'red':'');
}
function createextrapbrow(tnum){
	var rownum=parseInt(document.getElementById('pricebreakrows').value);
	if(rownum==tnum){
		rownum++;
		document.getElementById('pricebreakrows').value=rownum;
		var newdiv=document.createElement('div');
		newdiv.style.display='table-row';
		newdiv.style.fontSize='11px';
		newdiv.innerHTML='<div style="display:table-cell"><input type="text" name="pbquant'+rownum+'" size="4" value="" onchange="createextrapbrow('+rownum+')" /></div>' +
			'<div style="display:table-cell"><input type="text" name="pbprice'+rownum+'" size="4" value="" /></div>' +
			'<div style="display:table-cell"><input type="checkbox" name="wpercent'+rownum+'" title="Percentage" value="1" /></div>' +
			'<div style="display:table-cell"><input type="text" name="pbwholeprice'+rownum+'" size="4" value="" /></div>' +
			'<div style="display:table-cell"><input type="checkbox" name="wholesalepercent'+rownum+'" title="Percentage" value="1" /></div>';
		document.getElementById('pricebreaktable').appendChild(newdiv);
	}
}
/* ]]> */
</script>
<script src="popcalendar.js"></script>
<script>try{languagetext('<%=adminlang%>');}catch(err){}</script>
	<form name="mainform" method="post" action="adminprods.asp" onsubmit="return formvalidator(this)">
			<input type="hidden" name="posted" value="1" />
			<%	if getpost("act")="modify" AND NOT doaddnew then %>
			<input type="hidden" name="act" value="domodify" />
			<input type="hidden" name="id" value="<%=htmlspecials(pID)%>" />
			<%	else %>
			<input type="hidden" name="act" value="doaddnew" />
			<%	end if
				call writehiddenvar("disp", getpost("disp"))
				call writehiddenvar("stext", getpost("stext"))
				call writehiddenvar("sprice", getpost("sprice"))
				call writehiddenvar("scat", getpost("scat"))
				call writehiddenvar("stype", getpost("stype"))
				call writehiddenvar("sos", getpost("sos"))
				call writehiddenvar("pg", getpost("pg"))
				if NOT usesflatrate then
					print "<input type=""hidden"" name=""pShipping"" value="""&pShipping&""" />"
					print "<input type=""hidden"" name=""pShipping2"" value="""&pShipping2&""" />"
				end if
				%>
			<table id="producttable" width="100%" border="" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="4" align="center"><strong><%
					if doaddnew then
						print yyPrUpd
					elseif getpost("act")="modify" then
						print yyYouMod & " &quot;" & pName & "&quot;"
					else
						print yyYouCln & " &quot;" & pName & "&quot;"
					end if
				%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
			    <td align="right"><%=redasterix&yyPrRef%>:</td><td><input type="text" id="newid" name="newid" size="25" value="<%=htmlspecials(pID)%>" onfocus="document.getElementById('pSKU').size=10;this.size=25" onkeyup="checkrequiredfields()" />
				 / <input type="text" id="pSKU" name="pSKU" size="10" value="<%=htmlspecials(pSKU)%>" placeholder="SKU" title="SKU" onfocus="document.getElementById('newid').size=10;this.size=25" />
				</td>
			    <td align="right"><%=redasterix&yySection%>:</td><td><select size="1" name="psection" id="psection" onchange="checkrequiredfields()"><option value=""><%=yySelect%></option><%
					if isarray(allsections) then
						for rowcounter=0 to UBOUND(allsections,2)
							if allsections(0,rowcounter)=psection then print "<option value="""&allsections(0,rowcounter)&""" selected=""selected"">" & htmldisplay(allsections(1,rowcounter)) & "</option>" &vbCrLf
						next
					end if %></select></td>
			  </tr>
			  <tr>
			    <td align="right"><%=redasterix&yyPrNam%>:</td><td><input type="text" name="pName" id="pName" size="40" value="<%=htmlspecialsucode(pName)%>" onfocus="displaymultilangname('pName')" onkeyup="checkrequiredfields()" />
<%				for index=2 to adminlanguages+1
					if (adminlangsettings AND 1)=1 then
			%><input type="text" style="display:none;margin-top:2px" name="pName<%=index%>" id="pName<%=index%>" size="40" placeholder="Product Name Language <%=index%>" value="<%=htmlspecialsucode(pNames(index))%>" /><%
					end if
				next %>
				</td>
			    <td align="right" style="white-space:nowrap"><%=redasterix&yyPrPri%> / WS / List:</td><td style="white-space:nowrap"><input type="text" name="pPrice" id="pPrice" size="10" value="<%=pPrice%>" placeholder="<%=yyPrPri%>" title="<%=yyPrPri%>" onfocus="document.getElementById('pWholesalePrice').size=5;document.getElementById('pListPrice').size=6;this.size=10" onkeyup="checkrequiredfields()" />
				/ <input type="text" id="pWholesalePrice" name="pWholesalePrice" size="5" value="<%=pWholesalePrice%>" placeholder="<%=yyWhoPri%>" title="<%=yyWhoPri%>" onfocus="document.getElementById('pPrice').size=5;document.getElementById('pListPrice').size=6;this.size=10" />
				/ <input type="text" id="pListPrice" name="pListPrice" size="6" value="<% if cdbl(pListPrice)<>0.0 then print pListPrice %>" placeholder="<%=yyListPr%>" title="<%=yyListPr%>" onfocus="document.getElementById('pPrice').size=5;document.getElementById('pWholesalePrice').size=5;this.size=10" />
				</td>
			  </tr>
<%				if useStockManagement OR perproducttaxrate then %>
			  <tr>
<%					if useStockManagement then %>
				<td align="right">
				<input type="hidden" name="stocksetting" id="stocksetting" value="" />
				<select name="pStockByOpts" size="1" onchange="setstocktype()">
				<option value="0">&nbsp;&nbsp;&nbsp;<%=yyInStk%>:</option>
				<option value="1"<% if cint(pStockByOpts)<>0 then print " selected=""selected""" %>><%=yyByOpt%>:</option></select>
				</td><td<% if NOT perproducttaxrate then print " colspan=""3"""%>><%
						if getpost("act")="modify" then %>
				<input type="button" id="stockbutton" value="<%=pInStock%> (Click to Set)" onclick="setstockcontrols(false)" />
<%						end if %>
				<span id="setbyopts" style="display:none">(Set By Product Options)</span><input type="text" name="inStock" id="inStock" size="10" value="<%=pInStock%>" /></td>
<%					else %>
				<td colspan="2"><input type="hidden" name="pStockByOpts" value="<% if cint(pStockByOpts)<>0 then print "1" %>" /></td>
<%					end if
					if perproducttaxrate then %>
					<td align="right"><%=yyTax%>:</td><td><input type="text" style="text-align:right" size="6" name="pTax" value="<%=pTax%>" />%</td>
<%					end if %>
			  </tr>
<%				end if %>
			  <tr>
                <td align="right"><%=yyPrWght%>:</td>
                <td><input type="text" name="pWeight" size="9" value="<%=pWeight%>" /></td>
				<%	if (adminUnits AND 12) > 0 then
						proddims=split(pDims&"", "x") %>
				<td align="right"><%=yyDims%>:</td>
				<td><input type="text" name="plen" size="4" value="<%if UBOUND(proddims)>=0 then print proddims(0)%>" /> <strong>X</strong> 
				<input type="text" name="pwid" size="4" value="<%if UBOUND(proddims)>=1 then print proddims(1)%>" /> <strong>X</strong> 
				<input type="text" name="phei" size="4" value="<%if UBOUND(proddims)>=2 then print proddims(2)%>" /></td>
				<%	else %>
			    <td align="center" colspan="2">&nbsp;</td>
				<%	end if %>
			  </tr>
			  <tr>
			    <td align="right"><span style="color:#BB0000"><%=yyImage%></span>:</td>
				<td style="white-space:nowrap"><table style="border-collapse:collapse" width="100%"><tr><td style="border:0px;padding:0px;margin:0px;width:100%"><input type="text" id="pImage" style="width:99%" value="<% if isarray(allsmimgs) then print htmlspecials(allsmimgs(0,0)) %>" onchange="document.getElementById('smim0').value=this.value" onfocus="this.size=30;document.getElementById('but_pImage').value='IM'" onblur="this.size=16;document.getElementById('but_pImage').value='<%=jsescape(yyImgMgr)%>'" /></td><td style="border:0px;padding:0px;margin:0px;width:40px"><input type="button" style="margin-left:13px" id="smallimup" value="..." onclick="uploadimage('pImage')" />&nbsp;<input type="button" id="but_pImage" value="<%=yyImgMgr%>" onclick="imagemanager()" /></td></tr></table></td>
<%				session.lcid=saveLCID
				themask=cStr(DateSerial(2003,12,11))
				themask=replace(themask,"2003","yyyy")
				themask=replace(themask,"12","mm")
				themask=replace(themask,"11","dd")
%>				<td align="right"><%=yyDateAd%>:</td>
				<td><div style="position:relative;display:inline"><input type="text" size="14" name="pDateAdded" value="<% if pDateAdded<>"" then print DateSerial(year(pDateAdded),month(pDateAdded),day(pDateAdded))%>" style="vertical-align:middle" /> <input type="button" onclick="popUpCalendar(this, document.forms.mainform.pDateAdded, '<%=themask%>', -200)" value="DP" /></div></td>
			<%	session.lcid=1033 %>
			  </tr>
			  <tr id="lessimages">
                <td align="right"><span style="color:#00BB00"><%=yyLgeImg%></span>:</td>
                <td><table style="border-collapse:collapse" width="100%"><tr><td style="border:0px;padding:0px;margin:0px;width:100%"><input type="text" id="pLargeImage" style="width:100%" value="<% if isarray(alllgimgs) then print htmlspecials(alllgimgs(0,0)) %>" onchange="document.getElementById('lgim0').value=this.value" /></td><td style="padding:0px"><input type="button" style="margin-left:15px" value="..." onclick="uploadimage('pLargeImage')" /></td></tr></table></td>
				<td align="right"><%
					sSQL="SELECT scgWorkingName FROM searchcriteriagroup WHERE scgID=0"
					rs.open sSQL,cnn,0,1
					if NOT rs.EOF then yyManuf=rs("scgWorkingName")
					rs.close
					print yyManuf %>:</td>
				<td><select name="pManufacturer" size="1">
				  <option value="0"><%=yyNone%></option><%
					gotmanufacturer=FALSE
					sSQL="SELECT scID,scWorkingName FROM searchcriteria WHERE scGroup=0 ORDER BY scWorkingName"
					rs.open sSQL,cnn,0,1
					do while NOT rs.EOF
						print "<option value="""&rs("scID")&""""
						if rs("scID")=pManufacturer then print " selected=""selected""" : gotmanufacturer=TRUE
						print ">"&rs("scWorkingName")&"</option>"&vbCrLf
						rs.movenext
					loop
					rs.close
					if pManufacturer<>0 AND NOT gotmanufacturer then response.write "<option value=""0"" selected=""selected"">** DELETED **</option>"
%>				  </select>
				</td>
			  </tr>
			  <tr id="lessimages2">
                <td align="right"><span style="color:#0000BB"><%=yyGiaImg%></span>:</td>
                <td style="white-space:nowrap"><table style="border-collapse:collapse" width="100%"><tr><td style="border:0px;padding:0px;margin:0px;width:100%"><input type="text" id="pGiantImage" style="width:100%" value="<% if isarray(allgtimgs) then print htmlspecials(allgtimgs(0,0)) %>" onchange="document.getElementById('gtim0').value=this.value" /></td><td style="padding:0px"><input type="button" style="margin-left:15px" value="..." onclick="uploadimage('pGiantImage')" /></td></tr></table></td>
				<td align="right"><%=yyDrSppr%>:</td>
				<td><select name="pDropship" size="1">
				  <option value="0"><%=yyNone%></option>
				<%	if IsArray(alldropship) then
						for index=0 to UBOUND(alldropship, 2)
							print "<option value="""&alldropship(0,index)&""""
							if alldropship(0,index)=pDropship then print " selected=""selected"""
							print ">"&alldropship(1,index)&"</option>"&vbCrLf
						next
					end if %>
				  </select></td>
			  </tr>
			  <tr id="extraimages" style="display:none">
				<td colspan="4" align="center">
				  <table id="extraimagetable" style="border:1px;border-color:#555;border-style:solid;padding:3px;width:90%">
					<tr><td align="left" height="30"><input type="button" id="moreimages" value="<%=yyMorImg%>" onclick="moreimagefn()" style="margin-right:5px" /><span style="color:#BB0000"><%=yyImage%></span></td><td align="center"><span style="color:#00BB00"><%=yyLgeImg%></span></td><td align="center"><span style="color:#0000BB"><%=yyGiaImg%></span></td></tr>
<%				if NOT doaddnew then
					for imageindex=0 to maximagenumber-1
						call getnext3images(smallimg,largeimg,giantimg)
						call displayimagerow(imageindex,smallimg,largeimg,giantimg)
					next
				end if
				for maximgindex=imageindex to vrmax(5,imageindex+2)
					call displayimagerow(maximgindex,"","","")
				next
%>
				  </table>
				  <input type="hidden" name="maximgindex" id="maximgindex" value="<%=maximgindex%>" />
				</td>
			  </tr>
			  <tr>
				<td align="right"><select size="1" id="staticpage" onchange="setstatic(this.selectedIndex)"><option value=""><%=yyStatPg%></option><option value="1"<%=IIfVs(pStaticURL<>""," selected=""selected""")%>>Has Static URL</option></select></td>
                <td><div id="staticpagediv"<%=IIfVs(pStaticURL<>""," style=""display:none""")%>><input type="checkbox" name="pStaticPage" value="1"<% if int(pStaticPage)<>0 then print " checked=""checked""" %> /></div>
				<div id="staticurldiv"<%=IIfVs(pStaticURL=""," style=""display:none""")%>><input type="text" name="pStaticURL" id="pStaticURL" size="40" value="<%=htmlspecials(pStaticURL)%>" /></div></td>
				<td align="right"><%=yyProdOr%>:</td>
                <td><input type="text" name="pOrder" size="10" value="<%=pOrder%>" /></td>
			  </tr>
<%				if usesflatrate then %>
			  <tr>
                <td align="right"><%=yyFlatShp & ":<br />" & yyFirShi%>:</td>
                <td><input type="text" name="pShipping" size="15" value="<%=pShipping%>" /></td>
                <td align="right"><%=yyFlatShp & ":<br />" & yySubShi%></td>
                <td><input type="text" name="pShipping2" size="15" value="<%=pShipping2%>" /></td>
			  </tr>
<%				end if %>
			  <tr>
				<td align="right"><%=yyExemp%>:<br /><span style="font-size:10px">&lt;Ctrl>+Click&nbsp;</span></td><td>
					<select name="pExemptions" size="5" multiple="multiple">
					<option value="1"<%if (pExemptions AND 1)=1 then print " selected=""selected"""%>><%=yyExStat%></option>
					<option value="2"<%if (pExemptions AND 2)=2 then print " selected=""selected"""%>><%=yyExCoun%></option>
					<option value="4"<%if (pExemptions AND 4)=4 then print " selected=""selected"""%>><%=yyExShip%></option>
					<option value="8"<%if (pExemptions AND 8)=8 then print " selected=""selected"""%>><%=yyExHand%></option>
					<option value="16"<%if (pExemptions AND 16)=16 then print " selected=""selected"""%>>Free Shipping Exempt</option>
					<option value="32"<%if (pExemptions AND 32)=32 then print " selected=""selected"""%>>Pack Together Exempt</option>
					<option value="64"<%if (pExemptions AND 64)=64 then print " selected=""selected"""%>>Product Discount Exempt</option>
					</select>
				</td>
				<td colspan="2" width="50%">
<%				if useStockManagement then %>
					<input type="hidden" name="pSell" value="<% if int(pSell)<>0 then print "ON" %>" />
<%				end if %>
					<div class="separator">Flags</div>
					<div style="max-width:400px">
<%				if NOT useStockManagement then %>
						<div style="float:left;padding:5px"><div style="float:left"><%=yySellBut%>:</div><div style="float:left"><input type="checkbox" name="pSell" value="ON"<% if int(pSell)<>0 then print " checked=""checked""" %> /></div></div>
<%				end if %>
						<div style="float:left;padding:5px"><div style="float:left"><%=yyDisPro%>:</div><div style="float:left"><input type="checkbox" name="pDisplay" value="ON"<% if cint(pDisplay)<>0 then print " checked=""checked""" %> /></div></div>
						<div style="float:left;padding:5px"><div style="float:left"><%=yyRecomd%>:</div><div style="float:left"><input type="checkbox" name="pRecommend" value="1"<% if int(pRecommend)<>0 then print " checked=""checked""" %> /></div></div>
						<div style="float:left;padding:5px"><div style="float:left"><%=yyGifWra%>:</div><div style="float:left"><input type="checkbox" name="pGiftWrap" value="1"<% if int(pGiftWrap)<>0 then print " checked=""checked""" %> /></div></div>
						<div style="float:left;padding:5px"><div style="float:left"><%=yyBakOrd%>:</div><div style="float:left"><input type="checkbox" name="pBackOrder" value="1"<% if int(pBackOrder)<>0 then print " checked=""checked""" %> /></div></div>
					</div>
				</td>
			  </tr>
			  <tr>
				<td align="right"><%=yyAddSrP%>:</td>
                <td colspan="3"><input type="text" name="pSearchParams" style="width:80%" onfocus="displaymultilangname('pSearchParams')" value="<%=htmlspecials(pSearchParams)%>" />
<%				for index=2 to adminlanguages+1
					if (adminlangsettings AND 4194304)=4194304 then
			%><input type="text" style="width:80%;margin-top:2px" name="pSearchParams<%=index%>" id="pSearchParams<%=index%>" placeholder="<%=yyAddSrP%> Language <%=index%>" value="<%=htmlspecialsucode(IIfVr(index=2,pSearchParams2,pSearchParams3))%>" /><%
					end if
				next %></td>
			  </tr>
			  <tr>
				<td align="right">Page Title Tag:</td>
                <td colspan="3"><input type="text" name="pTitle" style="width:80%" onfocus="displaymultilangname('pTitle')" value="<%=htmlspecials(pTitle)%>" maxlength="255" />
<%				for index=2 to adminlanguages+1
					if (adminlangsettings AND 2097152)=2097152 then
			%><input type="text" style="width:80%;margin-top:2px" name="pTitle<%=index%>" id="pTitle<%=index%>" placeholder="Page Title Language <%=index%>" value="<%=htmlspecialsucode(IIfVr(index=2,pTitle2,pTitle3))%>" maxlength="255" /><%
					end if
				next %></td>
			  </tr>
			  <tr>
				<td align="right">Meta Description:</td>
                <td colspan="3"><input type="text" name="pMetaDesc" style="width:80%" onfocus="displaymultilangname('pMetaDesc')" value="<%=htmlspecials(pMetaDesc)%>" maxlength="255" />
<%				for index=2 to adminlanguages+1
					if (adminlangsettings AND 2097152)=2097152 then
			%><input type="text" style="width:80%;margin-top:2px" name="pMetaDesc<%=index%>" id="pMetaDesc<%=index%>" placeholder="Meta Description Language <%=index%>" value="<%=htmlspecialsucode(IIfVr(index=2,pMetaDesc2,pMetaDesc3))%>" maxlength="255" /><%
					end if
				next %></td>
			  </tr>
<%		if instr(productpagelayout&detailpagelayout,"custom1")>0 then %>
			  <tr>
				<td align="right"><%=admincustomlabel1%>:</td>
                <td colspan="3"><input type="text" name="pCustom1" style="width:80%" value="<%=htmlspecials(pCustom1)%>" maxlength="255" /></td>
			  </tr>
<%		end if
		if instr(productpagelayout&detailpagelayout,"custom2")>0 then %>
			  <tr>
				<td align="right"><%=admincustomlabel2%>:</td>
                <td colspan="3"><input type="text" name="pCustom2" style="width:80%" value="<%=htmlspecials(pCustom2)%>" maxlength="255" /></td>
			  </tr>
<%		end if
		if instr(productpagelayout&detailpagelayout,"custom3")>0 then %>
			  <tr>
				<td align="right"><%=admincustomlabel3%>:</td>
                <td colspan="3"><input type="text" name="pCustom3" style="width:80%" value="<%=htmlspecials(pCustom3)%>" maxlength="255" /></td>
			  </tr>
<%		end if
		if digidownloads=true then %>
			  <tr>
                <td align="right"><%=yyDownl%>:</td>
                <td colspan="3"><input type="text" size="30" name="pDownload" value="<%=pDownload%>" maxlength="255" /></td>
			  </tr>
<%		end if %>
			  <tr>
				<td colspan="4">
			<table width="100%">
			  <tr>
				<td width="25%" align="center">Product Options</td><td width="25%" align="center"><%=yyAddSec%></td><td width="25%" align="center"><%=yySeaCri%></td><td width="25%" align="center">Quantity Pricing</td>
			  </tr>
			  <tr>
				<td align="center" valign="top">
				  <table id="optionstable">
<%		rowcounter=0
		if isarray(alloptions) then
			if isarray(prodoptions) then
				for rowcounter=0 to UBOUND(prodoptions,2)
					print "<tr><td>"&(rowcounter+1)&"</td><td><select style=""width:180px"" size=""1"" id=""poption"&rowcounter&""" name=""poption"&rowcounter&"""><option value=""0"">"&yyDelete&"...</option><option value="""&prodoptions(1,rowcounter)&""" selected=""selected"">"&prodoptions(2,rowcounter)&"</option></select></td></tr>"&vbCrLf
				next
			end if
		end if %>
					<tr><td><%=(rowcounter+1)%></td><td><select style="width:180px" size="1" id="poption<%=rowcounter%>" name="poption<%=rowcounter%>" onchange="addnewoption(<%=rowcounter%>,'option');"><option value="0"><%=yySelect%></option></select></td></tr>
				  </table>
				  <input type="hidden" id="pnumoptions" value="0" />
				</td>
				<td align="center" valign="top">
				  <table id="sectionstable">
<%		rowcounter=0
		if isarray(allsections) then
			if isarray(prodsections) then
				for rowcounter=0 to UBOUND(prodsections,2)
					print "<tr><td>"&(rowcounter+1)&"</td><td><select style=""width:180px"" size=""1"" id=""psection"&rowcounter&""" name=""psection"&rowcounter&"""><option value=""0"">"&yyDelete&"...</option><option value="""&prodsections(0,rowcounter)&""" selected=""selected"">"&prodsections(1,rowcounter)&"</option></select></td></tr>"&vbCrLf
				next
			end if
		end if %>
					<tr><td><%=(rowcounter+1)%></td><td><select style="width:180px" size="1" id="psection<%=rowcounter%>" name="psection<%=rowcounter%>" onchange="addnewoption(<%=rowcounter%>,'section');"><option value="0"><%=yySelect%></option></select></td></tr>
				  </table>
				  <input type="hidden" id="pnumsections" value="0" />
				</td>
				<td align="center" valign="top">
				  <table id="searchstable">
<%		rowcounter=0
		if isarray(allsearchcriteria) then
			if isarray(prodsearchcriteria) then
				for rowcounter=0 to UBOUND(prodsearchcriteria,2)
					print "<tr><td>"&(rowcounter+1)&"</td><td><select style=""width:180px"" size=""1"" id=""psearch"&rowcounter&""" name=""psearch"&rowcounter&"""><option value=""0"">"&yyDelete&"...</option><option value="""&prodsearchcriteria(0,rowcounter)&""" selected=""selected"">"&prodsearchcriteria(1,rowcounter)&"</option></select></td></tr>"&vbCrLf
				next
			end if
		end if %>
					<tr><td><%=(rowcounter+1)%></td><td><select style="width:180px" size="1" id="psearch<%=rowcounter%>" name="psearch<%=rowcounter%>" onchange="addnewoption(<%=rowcounter%>,'search');"><option value="0"><%=yySelect%></option></select></td></tr>
				  </table>
				  <input type="hidden" id="pnumsearchs" value="0" />
				</td>
				<td align="center" valign="top"><%
		rowcounter=1
		print "<div style=""display:table"" id=""pricebreaktable"">"
		print "<div style=""display:table-row""><div style=""display:table-cell;font-size:11px;text-align:center"">" & "Quant" & "</div><div style=""display:table-cell;font-size:11px;text-align:center"">" & "Price" & "</div><div style=""display:table-cell;font-size:11px;text-align:center"">%</div><div style=""display:table-cell;font-size:11px;text-align:center"">" & "WS" & "</div><div style=""display:table-cell;font-size:11px;text-align:center"">%</div></div>"
		sSQL="SELECT pPrice,pWholesalePrice,pbQuantity,pbPercent,pbWholesalePercent FROM pricebreaks WHERE pbProdID='"&escape_string(pId)&"' ORDER BY pbQuantity"
		rs2.open sSQL,cnn,0,1
		do while NOT rs2.EOF
			print "<div style=""display:table-row;font-size:11px"">" & _
				"<div style=""display:table-cell""><input type=""text"" name=""pbquant"&rowcounter&""" size=""4"" value=""" & rs2("pbQuantity") & """ title=""" & yyQuant & """ /></div>" & _
				"<div style=""display:table-cell""><input type=""text"" name=""pbprice"&rowcounter&""" size=""4"" value=""" & rs2("pPrice") & """ title=""" & yyPrPri & """ /></div>" & _
				"<div style=""display:table-cell""><input type=""checkbox"" name=""wpercent"&rowcounter&""" title=""Percentage"" value=""1"" " & IIfVs(rs2("pbPercent")<>0,"checked=""checked"" ") & "/></div>" & _
				"<div style=""display:table-cell""><input type=""text"" name=""pbwholeprice"&rowcounter&""" size=""4"" value=""" & rs2("pWholesalePrice") & """ title=""" & yyWhoPri & """ /></div>" & _
				"<div style=""display:table-cell""><input type=""checkbox"" name=""wholesalepercent"&rowcounter&""" title=""Percentage"" value=""1"" " & IIfVs(rs2("pbWholesalePercent")<>0,"checked=""checked"" ") & "/></div>" & _
			"</div>"
			rowcounter=rowcounter+1
			rs2.movenext
		loop
		rs2.close
		print "<div style=""display:table-row;font-size:11px"">" & _
			"<div style=""display:table-cell""><input type=""text"" name=""pbquant"&rowcounter&""" size=""4"" value="""" onchange=""createextrapbrow("&rowcounter&")"" title=""" & yyQuant & """ /></div>" & _
			"<div style=""display:table-cell""><input type=""text"" name=""pbprice"&rowcounter&""" size=""4"" value="""" title=""" & yyPrPri & """ /></div>" & _
			"<div style=""display:table-cell""><input type=""checkbox"" name=""wpercent"&rowcounter&""" title=""Percentage"" value=""1"" /></div>" & _
			"<div style=""display:table-cell""><input type=""text"" name=""pbwholeprice"&rowcounter&""" size=""4"" value="""" title=""" & yyWhoPri & """ /></div>" & _
			"<div style=""display:table-cell""><input type=""checkbox"" name=""wholesalepercent"&rowcounter&""" title=""Percentage"" value=""1"" /></div>" & _
		"</div>"
		print "</div>"
		call writehiddenidvar("pricebreakrows",rowcounter)
				%></td>
			  </tr>
			</table>
				</td>
			  </tr>
			  <tr>
				<td colspan="4">
			<table width="100%">
			  <tr> 
                <td width="50%" align="center" id="descshort"><%=yyDesc%></td>
                <td width="50%" align="center" id="desclong"><%=yyLnDesc%></td>
			  </tr>
			  <tr> 
                <td align="center" valign="top">
<%				if htmleditor="froala" then print "<div id=""editordivpDes1"" class=""htmleditorcontainer"">" %>
			<textarea onfocus="displaymultilangdescs(false,1)" name="pDescription" id="pDescription" cols="45" rows="8" placeholder="Product Description"><%=htmlspecialsucode(pDescription)%></textarea>
<%				if htmleditor="froala" then print "</div>"
				for index=2 to adminlanguages+1
					if (adminlangsettings AND 2)=2 then
						if NOT doaddnew then
							sSQL="SELECT pDescription" & index & " FROM products WHERE pId='"&escape_string(getpost("id"))&"'"
							rs2.Open sSQL,cnn,0,1
							thedescription=rs2("pDescription" & index)
							rs2.Close
						end if
						if htmleditor="ckeditor" OR htmleditor="froala" then print "<div id=""editordivpDes"&index&""" class=""htmleditorcontainer"" style="""&IIfVs(htmleditor="froala","display:none;")&"margin-top:20px"">" %>
			<textarea onfocus="displaymultilangdescs(false,<%=index%>)" <%=IIfVs(htmleditor="froala","style=""display:none"" ")%>id="pDescription<%=index%>" name="pDescription<%=index%>" cols="45" rows="8" placeholder="Description for Language <%=index%>"><%=htmlspecialsucode(thedescription)%></textarea>
<%						if htmleditor="ckeditor" OR htmleditor="froala" then print "</div>"
					end if
				next %>
				</td>
                <td align="center">
<%				if htmleditor="froala" then print "<div id=""editordivpLon1"" class=""htmleditorcontainer"">" %>
			<textarea onfocus="displaymultilangdescs(true,1)" name="pLongDescription" id="pLongDescription" cols="55" rows="9" placeholder="Product Long Description"><%=htmlspecialsucode(pLongDescription)%></textarea>
<%				if htmleditor="froala" then print "</div>"
				for index=2 to adminlanguages+1
					if (adminlangsettings AND 4)=4 then
						if NOT doaddnew then
							sSQL="SELECT pLongDescription" & index & " FROM products WHERE pId='"&escape_string(getpost("id"))&"'"
							rs2.Open sSQL,cnn,0,1
							thedescription=rs2("pLongDescription" & index)
							rs2.Close
						end if
						if htmleditor="ckeditor" OR htmleditor="froala" then print "<div id=""editordivpLon"&index&""" class=""htmleditorcontainer"" style="""&IIfVs(htmleditor="froala","display:none;")&"margin-top:20px"">" %>
			<textarea onfocus="displaymultilangdescs(true,<%=index%>)" <%=IIfVs(htmleditor="froala","style=""display:none"" ")%>id="pLongDescription<%=index%>" name="pLongDescription<%=index%>" cols="55" rows="9" placeholder="Long Description for Language <%=index%>"><%=htmlspecialsucode(thedescription)%></textarea>
<%						if htmleditor="ckeditor" OR htmleditor="froala" then print "</div>"
					end if
				next %>
				</td>
			  </tr>
			</table>
				</td>
			  </tr>
			</table>
			<table width="100%" border="" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="4" align="center">&nbsp;<br />
					<input type="submit" value="<%=yySubmit%>" />&nbsp;&nbsp;<input type="reset" value="<%=yyReset%>" />
                </td>
			  </tr>
            </table>
	</form>
<%	if htmleditor="ckeditor" AND NOT issubproduct then
		pathtovsadmin=request.servervariables("URL")
		slashpos=instrrev(pathtovsadmin, "/")
		if slashpos>0 then pathtovsadmin=left(pathtovsadmin, slashpos-1)
		print "<script>function loadeditors(){"
		streditor="var pDescription=CKEDITOR.replace('pDescription',{extraPlugins : 'stylesheetparser,autogrow',autoGrow_maxHeight : 800,removePlugins : 'resize', toolbarStartupExpanded : false, toolbar : 'Basic', filebrowserBrowseUrl :'ckeditor/filemanager/browser/default/browser.html?Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserImageBrowseUrl : 'ckeditor/filemanager/browser/default/browser.html?Type=Image&Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserFlashBrowseUrl :'ckeditor/filemanager/browser/default/browser.html?Type=Flash&Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=File',filebrowserImageUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=Image',filebrowserFlashUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=Flash'});" & vbLf
		streditor=streditor & "pDescription.on('instanceReady',function(event){var myToolbar='Basic';event.editor.on( 'beforeMaximize', function(){if(event.editor.getCommand('maximize').state==CKEDITOR.TRISTATE_ON && myToolbar!='Basic'){pDescription.setToolbar('Basic');myToolbar='Basic';pDescription.execCommand('toolbarCollapse');}else if(event.editor.getCommand('maximize').state==CKEDITOR.TRISTATE_OFF && myToolbar!='Full'){pDescription.setToolbar('Full');myToolbar='Full';pDescription.execCommand('toolbarCollapse');}});event.editor.on('contentDom', function(e){event.editor.document.on('blur', function(){if(!pDescription.isToolbarCollapsed){pDescription.execCommand('toolbarCollapse');pDescription.isToolbarCollapsed=true;}});event.editor.document.on('focus',function(){expandckeditor(event.editor.name,ECTEDITORNUMBER);if(pDescription.isToolbarCollapsed){pDescription.execCommand('toolbarCollapse');pDescription.isToolbarCollapsed=false;}});});pDescription.fire('contentDom');pDescription.isToolbarCollapsed=true;});" &vbLf
		print replace(streditor,"ECTEDITORNUMBER",1)
		for index=2 to adminlanguages+1
			if (adminlangsettings AND 2)=2 then print replace(replace(streditor,"pDescription","pDescription"&index),"ECTEDITORNUMBER",index)
		next
		print replace(replace(streditor,"pDescription","pLongDescription"),"ECTEDITORNUMBER",1)
		for index=2 to adminlanguages+1
			if (adminlangsettings AND 4)=4 then print replace(replace(streditor,"pDescription","pLongDescription"&index),"ECTEDITORNUMBER",index)
		next
		print "}window.onload=function(){loadeditors();init();}</script>"
	elseif htmleditor="froala" then
		call displayfroalaeditor("pDescription",yyDesc,".on('froalaEditor.focus',function(){expandckeditor(""pDes"",1);})",FALSE,FALSE,1,FALSE)
		call displayfroalaeditor("pLongDescription",yyLnDesc,".on('froalaEditor.focus',function(){expandckeditor(""pLon"",1);})",FALSE,FALSE,1,FALSE)
		for index=2 to adminlanguages+1
			if (adminlangsettings AND 2)=2 then call displayfroalaeditor("pDescription"&index,yyDesc&" (language "&index&")",".on('froalaEditor.focus',function(){expandckeditor(""pDes"","&index&");})",FALSE,FALSE,1,TRUE)
			if (adminlangsettings AND 4)=4 then call displayfroalaeditor("pLongDescription"&index,yyLnDesc&" (language "&index&")",".on('froalaEditor.focus',function(){expandckeditor(""pLon"","&index&");})",FALSE,FALSE,1,TRUE)
		next
	end if
%>
<script>
/* <![CDATA[ */
checkrequiredfields();
document.getElementById("pnumoptions").value=<% if IsArray(prodoptions) then print (UBOUND(prodoptions,2)+1) else print "0" %>;
document.getElementById("pnumsections").value=<% if IsArray(prodsections) then print (UBOUND(prodsections,2)+1) else print "0" %>;
document.getElementById("pnumsearchs").value=<% if IsArray(prodsearchcriteria) then print (UBOUND(prodsearchcriteria,2)+1) else print "0" %>;
setprodoptions("option");
setprodoptions("section");
setprodoptions("search");
<% 	if useStockManagement then %>
setstocktype();
<%	end if %>
populateoptionsselect(document.getElementById('psection'),'section');
/* ]]> */
</script>
<%
   elseif getpost("act")="discounts" then 
		sSQL="SELECT pName FROM products WHERE pID='"&escape_string(getpost("id"))&"'"
		rs.open sSQL,cnn,0,1
		thisname=rs("pName")
		rs.close
		alldata=""
		sSQL="SELECT cpaID,cpaCpnID,cpnWorkingName,cpnSitewide,cpnEndDate,cpnType FROM cpnassign INNER JOIN coupons ON cpnassign.cpaCpnID=coupons.cpnID WHERE cpaType=2 AND cpaAssignment='" & escape_string(getpost("id")) & "'"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then alldata=rs.GetRows
		rs.close
		alldata2=""
		tdt=Date()
		sSQL="SELECT cpnID,cpnWorkingName,cpnSitewide FROM coupons WHERE cpnSitewide=0 AND cpnEndDate >=" & vsusdate(tdt)
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then alldata2=rs.GetRows
		rs.close
%>
<script>
/* <![CDATA[ */
function drk(id){
if(confirm("<%=jscheck(yyConAss)%>\n")){
	document.mainform.id.value=id;
	document.mainform.act.value="deletedisc";
	document.mainform.submit();
}
}
/* ]]> */
</script>
        <tr>
		<form name="mainform" method="post" action="adminprods.asp">
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="dodiscounts" />
			<input type="hidden" name="id" value="<%=htmlspecials(getpost("id"))%>" />
<%				call writehiddenvar("disp", getpost("disp"))
				call writehiddenvar("stext", getpost("stext"))
				call writehiddenvar("sprice", getpost("sprice"))
				call writehiddenvar("scat", getpost("scat"))
				call writehiddenvar("stype", getpost("stype"))
				call writehiddenvar("sos", getpost("sos"))
				call writehiddenvar("pg", getpost("pg")) %>
            <table width="100%" border="" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="4" align="center"><strong><%=yyAssPrd%> &quot;<%=thisname%>&quot;.</strong><br />&nbsp;</td>
			  </tr>
<%	gotone=false
	if IsArray(alldata2) then
		thestr="<tr><td colspan='4' align='center'>"&yyAsDsCp&": <select name='assdisc' size='1'>"
		for index=0 to UBOUND(alldata2,2)
			alreadyassign=false
			if IsArray(alldata) then
				for index2=0 to UBOUND(alldata,2)
					if alldata2(0,index)=alldata(1,index2) then alreadyassign=true
				next
			end if
			if NOT alreadyassign then
				thestr=thestr & "<option value='"&alldata2(0,index)&"'>"&alldata2(1,index)&"</option>" & vbCrLf
				gotone=true
			end if
		next
		thestr=thestr & "</select> <input type='submit' value='Go' /></td></tr>"
	end if
	if gotone then
		print thestr
	else
%>			  <tr> 
                <td width="100%" colspan="4" align="center"><br /><strong><%=yyNoDis%></strong></td>
			  </tr>
<%	end if
	if IsArray(alldata) then
%>			  <tr> 
                <td width="100%" colspan="4" align="center"><br /><strong><%=yyCurDis%> &quot;<%=thisname%>&quot;.</strong><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td><strong><%=yyWrkNam%></strong></td>
				<td><strong><%=yyDisTyp%></strong></td>
				<td><strong><%=yyExpire%></strong></td>
				<td align="center"><strong><%=yyDelete%></strong></td>
			  </tr>
<%		for index=0 to UBOUND(alldata,2)
			prefont=""
			postfont=""
			if alldata(3,index)=1 OR alldata(4,index)-Date() < 0 then
				prefont="<span style=""color:#FF0000"">"
				postfont="</span>"
			end if
%>			  <tr> 
                <td><%=prefont & alldata(2,index) & postfont %></td>
				<td><%	if alldata(5,index)=0 then
							print prefont & yyFrSShp & postfont
						elseif alldata(5,index)=1 then
							print prefont & yyFlatDs & postfont
						elseif alldata(5,index)=2 then
							print prefont & yyPerDis & postfont
						end if %></td>
				<td><%	if alldata(4,index)=DateSerial(3000,1,1) then
							print yyNever
						elseif alldata(4,index)-Date() < 0 then
							print "<span style=""color:#FF0000"">"&yyExpird&"</span>"
						else
							print prefont & alldata(4,index) & postfont
						end if %></td>
				<td align="center"><input type="button" value="Delete Assignment" onclick="drk('<%=alldata(0,index)%>')" /></td>
			  </tr>
<%		next
	else
%>			  <tr> 
                <td width="100%" colspan="4" align="center"><br /><strong><%=yyNoAss%></strong></td>
			  </tr>
<%	end if %>
			  <tr><td width="100%" colspan="4" align="center"><br />&nbsp;</td></tr>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br /><a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
			  </tr>
            </table></td>
		  </form>
        </tr>
<% elseif getpost("posted")="1" AND success then %>
      <table border="" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <a href="adminprods.asp<%
							print "?sos=" & urlencode(getpost("sos")) & "&rid=" & urlencode(getpost("rid")) & "&pid=" & urlencode(getpost("pid")) & "&disp=" & getpost("disp") & "&stext=" & urlencode(getpost("stext")) & "&sprice=" & urlencode(getpost("sprice")) & "&stype=" & getpost("stype") & "&scat=" & getpost("scat") & "&pg=" & getpost("pg")
						%>"><strong><%=yyClkHer%></strong></a>.<br /><br />&nbsp;<br />&nbsp;
                </td>
			  </tr>
			</table></td>
        </tr>
      </table>
<% elseif getpost("posted")="1" then %>
      <table border="" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" align="center"><br /><span style="color:#FF0000;font-weight:bold"><%=yyOpFai%></span><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table></td>
        </tr>
      </table>
<% elseif getget("act")="stknot" then %>
	<form method="post" action="adminprods.asp">
	<input type="hidden" name="posted" value="1" />
	<input type="hidden" name="act" value="allstk" />
      <table border="" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td align="center">
			<table class="admin-table-b" border="" cellspacing="3" cellpadding="3">
			<thead>
			  <tr> 
                <th scope="col" style="white-space:nowrap">&nbsp;<%=yyPrId%>&nbsp;</th>
				<th scope="col" style="white-space:nowrap">&nbsp;<%=yyPrName%>&nbsp;</th>
				<th scope="col" style="white-space:nowrap">&nbsp;<%=yyPOName%>&nbsp;</th>
				<th scope="col" style="white-space:nowrap">&nbsp;<%=yyQuant%>&nbsp;</th>
				<th scope="col" style="white-space:nowrap">&nbsp;<%=yyDelete%>&nbsp;</th>
			  </tr>
			</thead>
<%		if getget("pid")<>"" AND getget("oid")<>"" then
			sSQL="DELETE FROM notifyinstock WHERE nsProdID='"&escape_string(getget("pid"))&"' AND nsOptID="&getget("oid")
			ect_query(sSQL)
		end if
		sSQL="SELECT nsProdID,nsTriggerProdID,pName,nsOptID,COUNT(*) AS tcnt FROM notifyinstock LEFT JOIN products on notifyinstock.nsProdID=products.pID GROUP BY nsProdID,nsTriggerProdID,pName,nsOptID ORDER BY "&IIfVr(mysqlserver=TRUE,"tcnt","COUNT(*)")&" DESC"
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			optname=""
			if rs("nsOptID")<>0 then
				sSQL="SELECT optName FROM options WHERE optID="&rs("nsOptID")
				rs2.open sSQL,cnn,0,1
				if NOT rs2.EOF then optname=rs2("optName")
				rs2.close
			end if
			pname=trim(rs("pName")&"")
			if pname="" then pname="**DELETED**"
			prodid=trim(rs("nsProdID")&"")
			if trim(rs("nsTriggerProdID")&"")<>prodid then prodid=prodid&" / "&rs("nsTriggerProdID")
			print "<tr><td style=""white-space:nowrap"">"&prodid&"</td><td style=""white-space:nowrap"">"&pname&"</td><td style=""white-space:nowrap"">"&IIfVr(optname<>"",optname,"-")&"</td><td style=""white-space:nowrap"">"&rs("tcnt")&"</td><td style=""white-space:nowrap""><input type=""button"" value="""&yyDelete&""" onclick=""document.location='adminprods.asp?act=stknot&pid="&rs("nsProdID")&"&oid="&rs("nsOptID")&"'"" /></td></tr>"
			rs.movenext
		loop
		rs.close
%>			  <tr> 
                <td colspan="5" align="center">&nbsp;<br /><input type="submit" value="Send Stock Notifications (Where stock available)" /> <input type="button" onclick="document.location='adminprods.asp'" value="<%=yyClkBac%>" /></td>
			  </tr>
			</table></td>
        </tr>
      </table>
	</form>
<% else
		pract=request.cookies("pract")
		modclone=request.cookies("modclone")
		sortorder=request.cookies("psort")
		catorman=request.cookies("pcatorman") %>
<script>
/* <![CDATA[ */
function setCookie(c_name,value,expiredays){
	var exdate=new Date();
	exdate.setDate(exdate.getDate()+expiredays);
	document.cookie=c_name+ "=" +escape(value)+((expiredays==null) ? "" : ";expires="+exdate.toGMTString());
}
function mr(id){
	document.mainform.action="adminprods.asp";
	document.mainform.pid.value='';
	document.mainform.rid.value='';
	document.mainform.id.value=id;
	document.mainform.act.value="modify";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function cr(id){
	document.mainform.action="adminprods.asp";
	document.mainform.pid.value='';
	document.mainform.rid.value='';
	document.mainform.id.value=id;
	document.mainform.act.value="clone";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function al(id){
	document.mainform.action="adminprods.asp";
	document.mainform.pid.value='';
	document.mainform.rid.value='';
	document.mainform.id.value=id;
	document.mainform.act.value="altids";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function rel(id,relorpak){
	document.mainform.action="adminprods.asp?"+relorpak+"=go";
	relorpak=='package'?document.mainform.pid.value=id:document.mainform.rid.value=id;
	document.mainform.act.value="search";
	document.mainform.posted.value="";
	document.mainform.submit();
}
function updaterelations(relorpack){
	document.mainform.action="adminprods.asp";
	document.mainform.act.value="update"+relorpack;
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function newrec(id){
	document.mainform.action="adminprods.asp";
	document.mainform.pid.value='';
	document.mainform.rid.value='';
	document.mainform.id.value=id;
	document.mainform.act.value="addnew";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function quickupdate(){
	if(document.mainform.pract.value=="del"){
		if(!confirm("<%=jscheck(yyConDel)%>\n"))
			return;
	}
	document.mainform.action="adminprods.asp";
	document.mainform.act.value="quickupdate";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function dsc(id){
	document.mainform.action="adminprods.asp";
	document.mainform.id.value=id;
	document.mainform.act.value="discounts";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function startsearch(tact){
	document.mainform.action="adminprods.asp";
	document.mainform.act.value=tact;
	document.mainform.posted.value="";
	document.mainform.submit();
}
function recalcsalesrank(){
	var tpm=document.getElementById('timeperiod');
	document.mainform.action="adminprods.asp";
	document.mainform.id.value=tpm[tpm.selectedIndex].value;
	document.mainform.act.value="recalcsalesrank";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function inventorymenu(tmen){
	themenuitem=tmen.options[tmen.selectedIndex].value;
	tmen.selectedIndex=0;
	if(themenuitem=="7"){
		document.getElementById('inventorymenu').style.display='none';
		document.getElementById('calcsalesrank').style.display='';
		return
	}else if(themenuitem=="5"){
		document.mainform.action="adminprods.asp";
		document.mainform.act.value="tablechecks";
		document.mainform.posted.value="1";
	}else{
		if(themenuitem=="") return;
		if(themenuitem=="1") document.mainform.act.value="stockinventory";
		if(themenuitem=="2") document.mainform.act.value="fullinventory";
		if(themenuitem=="3") document.mainform.act.value="dump2COinventory";
		if(themenuitem=="4") document.mainform.act.value=<% if getrequest("scat")<>"" then print "confirm(""Would you like to filter by the current category / manufacturer?\n\nOK=Yes, filter by category.\nCancel=All Images."")?""filteredimages"":"%>"productimages";
		if(themenuitem=="6") document.mainform.act.value="filteredinventory";
		if(themenuitem=="8") document.mainform.act.value="filteredstock";
		document.mainform.action="dumporders.asp";
	}
	document.mainform.submit();
}
function dr(id){
if(confirm("<%=jscheck(yyConDel)%>\n")){
	document.mainform.action="adminprods.asp";
	document.mainform.id.value=id;
	document.mainform.act.value="delete";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
}
function changepract(obj){
	setCookie('pract',obj[obj.selectedIndex].value,600);
	startsearch("search");
}
function switchcatorman(obj){
	setCookie('pcatorman',obj[obj.selectedIndex].value,600);
	startsearch("<%	if getpost("act")="search" OR getget("pg")<>"" then	print "search"%>");
}
function changesortorder(obj){
	setCookie('psort',obj[obj.selectedIndex].value,600);
	startsearch("<%	if getpost("act")="search" OR getget("pg")<>"" then	print "search"%>");
}
function setto(){
	maxitems=document.getElementById("resultcounter").value;
	amnt=document.getElementById("txtadd").value;
	for(index=0;index<maxitems;index++){
		if(document.getElementById("chkbx"+index)){
			document.getElementById("chkbx"+index).value=amnt;
			document.getElementById("chkbx"+index).onchange();
		}
	}
}
function addto(){
	maxitems=document.getElementById("resultcounter").value;
	amnt=document.getElementById("txtadd").value;
	if(amnt.indexOf("%") > 0) ispercent=true; else ispercent=false;
	amnt.replace(/%/g, "");
	amnt=parseFloat(amnt);
	if(! isNaN(amnt)){
		for(index=0;index<maxitems;index++){
			if(document.getElementById("chkbx"+index)){
				theval=parseFloat(document.getElementById("chkbx"+index).value);
				if(! isNaN(theval))
					document.getElementById("chkbx"+index).value=ispercent?theval+((amnt*theval)/100.0):theval+amnt;
				document.getElementById("chkbx"+index).onchange();
			}
		}
	}
}
function checkboxes(docheck){
	maxitems=document.getElementById("resultcounter").value;
	for(index=0;index<maxitems;index++){
		var thiselem=document.getElementById("chkbx"+index);
		if(thiselem.checked!=docheck&&!thiselem.disabled){
			thiselem.checked=docheck;
			if(thiselem.onchange) thiselem.onchange();
		}
	}
}
function changemodclone(modclone){
	setCookie('modclone',modclone[modclone.selectedIndex].value,600);
	startsearch("search");
}
function tqn(objid,pidind){
	var ttr=document.getElementById('tr'+pidind);
	ttr.cells[5].innerHTML=objid.checked?'<input type="text" name="pqa'+pidind+'" value="'+(pa[pidind][2]==''?'1':pa[pidind][2])+'" size="3" />':'-';
}
function changeso(){
	var sos=1;
	if(!document.getElementById('sos1').checked) sos+=2;
	if(!document.getElementById('sos2').checked) sos+=4;
	if(!document.getElementById('sos3').checked) sos+=8;
	if(!document.getElementById('sos4').checked) sos+=16;
	if(!document.getElementById('sos5').checked) sos+=32;
	document.getElementById('sos').value=sos;
}
/* ]]> */
</script>
<h2><%=yyPrdAdm%></h2>
<%	pid=trim(request("pid"))
	rid=trim(request("rid"))
	sos=trim(request("sos"))
	if is_numeric(sos) then sos=int(sos) else sos=49
	ridarr=""
	if pid<>"" then
		sSQL="SELECT pID,quantity FROM productpackages WHERE packageID='" & escape_string(pid) & "'"
		rs.open sSQL,cnn,0,1
		if NOT rs.eof then ridarr=rs.getrows
		rs.close
	elseif rid<>"" then
		sSQL="SELECT rpRelProdID FROM relatedprods WHERE rpProdID='" & escape_string(rid) & "'"
		if relatedproductsbothways=TRUE then sSQL=sSQL & " UNION SELECT rpProdID FROM relatedprods WHERE rpRelProdID='" & escape_string(rid) & "'"
		rs.open sSQL,cnn,0,1
		if NOT rs.eof then ridarr=rs.getrows
		rs.close
	end if
	if getpost("disp")<>"" then
		response.cookies("pdisp")=getpost("disp")
		response.cookies("pdisp").Expires=Date()+365
		if request.servervariables("HTTPS")="on" then response.cookies("pdisp").secure=TRUE
	end if
	if request("disp")<>"" then productdisplay=request("disp") else productdisplay=request.cookies("pdisp")
	if (getget("related")="go" OR getget("package")="go") then SESSION("savesearch")= "disp="&getpost("disp")&"&stext=" & urlencode(getpost("stext")) & "&sprice=" & urlencode(getpost("sprice")) & "&stype=" & getpost("stype") & "&scat=" & getpost("scat") & "&sos=" & getpost("sos") & "&pg=" & getpost("pg")
%>
		<form name="mainform" method="post" action="adminprods.asp">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="xxxxx" />
			<input type="hidden" name="id" value="xxxxx" />
			<input type="hidden" name="pid" value="<%=pid%>" />
			<input type="hidden" name="rid" value="<%=rid%>" />
			<input type="hidden" id="sos" name="sos" value="<%=sos%>" />
			<input type="hidden" name="pg" value="<%=IIfVr(getpost("act")="search", "1", getget("pg"))%>" />
<%		if is_numeric(request("scat")) then thecat=int(request("scat")) else thecat=""
		sSQL="SELECT payProvEnabled,payProvData1 FROM payprovider WHERE payProvID=2"
		rs.open sSQL,cnn,0,1
		if rs("payProvEnabled")=1 AND trim(rs("payProvData1")&"")<>"" then twocoinventory=TRUE else twocoinventory=FALSE
		rs.close
%>			<table class="cobtbl" width="100%" border="" cellspacing="1" cellpadding="3">
<%			if pid<>"" OR rid<>"" then %>
			  <tr><td class="cobhl" align="center" colspan="4" height="22"><strong> <%=IIfVr(pid<>"","Products included in package "&pid,"Products related to "&rid) %></strong> </td></tr>
<%			end if %>
			  <tr> 
				<td class="cobhl" width="25%" align="right"><select name="disp" size="1">
					<option value="5"><%=yySearch%> Visible Prods</option>
					<option value="1"<% if productdisplay="1" then print " selected=""selected"""%>><%=yySearch%> All Prods</option>
					<option value="2"<% if productdisplay="2" then print " selected=""selected"""%>><%=yySearch%> Hidden Prods</option>
<%					if useStockManagement then print "<option value=""3"""&IIfVr(productdisplay="3", " selected=""selected""", "")&">"&yySearch&" "&yyOOStoc&"</option>" %>
					<option value="4"<% if productdisplay="4" then print " selected=""selected"""%>><%=yySearch%> Orphan Prods</option>
					<option value="6"<% if productdisplay="6" then print " selected=""selected"""%>> = Back Order</option>
					<option value="7"<% if productdisplay="7" then print " selected=""selected"""%>> &#8800; Back Order</option>
					<option value="8"<% if productdisplay="8" then print " selected=""selected"""%>> = Gift Wrap</option>
					<option value="9"<% if productdisplay="9" then print " selected=""selected"""%>> &#8800; Gift Wrap</option>
					<option value="10"<% if productdisplay="10" then print " selected=""selected"""%>> = Recommended</option>
					<option value="11"<% if productdisplay="11" then print " selected=""selected"""%>> &#8800; Recommended</option>
					<option value="12"<% if productdisplay="12" then print " selected=""selected"""%>> = Static Page</option>
					<option value="13"<% if productdisplay="13" then print " selected=""selected"""%>> &#8800; Static Page</option>
				</select></td>
				<td class="cobll" width="25%"><input type="text" name="stext" size="20" value="<%=htmlspecials(request("stext"))%>" /></td>
				<td class="cobhl" width="25%" align="right"><%=yySrchMx%>:</td>
				<td class="cobll" width="25%"><input type="text" name="sprice" size="10" value="<%=htmlspecials(request("sprice"))%>" title="Eg. 50-100 ... -50 ... 50-" /></td>
			  </tr>
			  <tr>
				<td class="cobhl" width="25%" align="right"><%=yySrchTp%>:</td>
				<td class="cobll" width="25%"><select name="stype" size="1">
					<option value=""><%=yySrchAl%></option>
					<option value="any"<% if request("stype")="any" then print " selected=""selected"""%>><%=yySrchAn%></option>
					<option value="exact"<% if request("stype")="exact" then print " selected=""selected"""%>><%=yySrchEx%></option>
					</select>
					<input type="checkbox" id="sos1" title="ID / SKU" onchange="changeso()" style="vertical-align:middle;" <% if (sos AND 2)<>2 then print "checked=""checked"" "%>/>
					<input type="checkbox" id="sos2" title="Name" onchange="changeso()" style="vertical-align:middle;" <% if (sos AND 4)<>4 then print "checked=""checked"" "%>/>
					<input type="checkbox" id="sos3" title="Description" onchange="changeso()" style="vertical-align:middle;" <% if (sos AND 8)<>8 then print "checked=""checked"" "%>/>
					<input type="checkbox" id="sos4" title="Long Description" onchange="changeso()" style="vertical-align:middle;" <% if (sos AND 16)<>16 then print "checked=""checked"" "%>/>
					<input type="checkbox" id="sos5" title="<%=yyAddSrP%>" onchange="changeso()" style="vertical-align:middle;" <% if (sos AND 32)<>32 then print "checked=""checked"" "%>/>
				</td>
				<td class="cobhl" width="25%" align="right"><select size="1" name="catorman" onchange="switchcatorman(this)">
					<option value="cat"><%=yySrchCt%></option>
					<option value="man"<% if catorman="man" then print " selected=""selected"""%>><%=yySeaCri%></option>
					<option value="dis"<% if catorman="dis" then print " selected=""selected"""%>>Discounts Assigned</option>
					<option value="non"<% if catorman="non" then print " selected=""selected"""%>><%=yyNone%></option>
					</select></td>
				<td class="cobll" width="25%">
<%				if catorman="non" then
					print "&nbsp;"
				else %>
				  <select name="scat" size="1">
				  <option value=""><%=IIfVr(catorman="man",yySeaCri,IIfVr(catorman="dis","All Discounts",yySrchAC))%></option>
<%					if catorman="dis" then
						sSQL="SELECT cpnID,cpnWorkingName FROM coupons WHERE cpnSitewide=0 OR cpnSitewide=3 ORDER BY cpnWorkingName"
						rs.open sSQL,cnn,0,1
						do while NOT rs.EOF
							print "<option value='"&rs("cpnID")&"'"
							if rs("cpnID")=thecat then print " selected=""selected"""
							print ">" & htmldisplay(rs("cpnWorkingName")) & "</option>" &vbCrLf
							rs.movenext
						loop
						rs.close
					elseif catorman="man" then
						adminonlysubcats=TRUE
						currgroup=-1
						sSQL="SELECT scID,scName,scGroup,scgTitle FROM searchcriteria INNER JOIN searchcriteriagroup ON searchcriteria.scGroup=searchcriteriagroup.scgID ORDER BY scGroup,scName"
						rs.open sSQL,cnn,0,1
						do while NOT rs.EOF
							if currgroup<>rs("scGroup") then currgroup=rs("scGroup") : print "<option value="""&rs("scID")&""" style=""font-weight:bold;color:#000"" disabled=""disabled"">== " & htmldisplay(rs("scgTitle")) & " ==</option>" &vbCrLf
							print "<option value='"&rs("scID")&"'"
							if rs("scID")=thecat then print " selected=""selected"""
							print ">" & htmldisplay(rs("scName")) & "</option>" &vbCrLf
							rs.movenext
						loop
						rs.close
					elseif noadmincategorysearch<>TRUE AND catorman<>"non" then
						sSQL="SELECT sectionID,sectionWorkingName,topSection,rootSection FROM sections " & IIfVs(adminonlysubcats, "WHERE rootSection=1 ") & "ORDER BY sectionWorkingName"
						rs.open sSQL,cnn,0,1
						if NOT rs.eof then alldata=rs.getrows
						rs.close
						if isarray(alldata) then
							if adminonlysubcats=true then
								for rowcounter=0 to UBOUND(alldata,2)
									print "<option value='"&alldata(0,rowcounter)&"'"
									if alldata(0,rowcounter)=thecat then print " selected=""selected"""
									print ">" & htmldisplay(alldata(1,rowcounter)) & "</option>" &vbCrLf
								next
							else
								call writemenulevel(0,1)
							end if
						end if
					end if %>
				  </select>
<%				end if %></td>
			  </tr>
			  <tr>
				<td class="cobhl" align="center">
					<div id="calcsalesrank" style="display:none">
						<div style="padding-bottom:5px;font-weight:bold">Recalculate Sales Rank</div>
						<div style="padding-bottom:7px">Sales over last:
							<select id="timeperiod" size="1">
								<option value="">All Time</option>
<%				for index=1 to 12
					print "<option value="""&index&""">"&index&" month(s)</option>"
				next %>
								<option value="18">18 month(s)</option>
								<option value="24">2 years</option>
								<option value="36">3 years</option>
							</select>
						</div>
						<input type="button" value="Go" onclick="recalcsalesrank()" /> &nbsp; <input type="button" value="Cancel" onclick="document.getElementById('inventorymenu').style.display='';document.getElementById('calcsalesrank').style.display='none';" /> 
					</div>
					<div id="inventorymenu"><%
				if pid="" AND rid="" then %>
					<select name="inventoryselect" size="1" onchange="inventorymenu(this)">
						<option value="">Select Action...</option>
<%					if useStockManagement then print "<option value=""1"">"&yyStkInv&"</option><option value=""8"">Filtered Stock Inventory</option>" %>
						<option value="2"><%=yyFulInv%></option>
						<option value="6">Filtered Inventory</option>
<%					if twocoinventory then print "<option value=""3"">2Checkout Inventory</option>" %>
						<option value="4">Product Images</option>
						<option value="5">Table Checks</option>
						<option value="7">Recalculate Sales Rank</option>
					</select>
<%				end if
				if getpost("act")="search" OR getget("pg")<>"" then
					isacheckbox=FALSE
					isaninteger=FALSE
					if customquickupdate<>"" then
						cqupdatearr=split(customquickupdate,",")
						for index=0 to UBOUND(cqupdatearr)
							cqitemarr=split(cqupdatearr(index),":")
							if pract=cqitemarr(0) AND UBOUND(cqitemarr)>=2 then
								if lcase(cqitemarr(2))="check" then isacheckbox=TRUE
								if lcase(cqitemarr(2))="num" then isaninteger=TRUE
							end if
						next
					end if
					if pid<>"" OR rid<>"" OR isacheckbox OR pract="ads" OR pract="bak" OR pract="csu" OR pract="cte" OR pract="del" OR pract="dip" OR pract="dis" OR pract="fse" OR pract="gwr" OR pract="hae" OR pract="isa" OR pract="pde" OR pract="pra" OR pract="prp" OR pract="pru" OR pract="pte" OR pract="rec" OR pract="sel" OR pract="she" OR pract="ste" OR pract="stp" then %>
					<div style="margin-top:2px"><input type="button" value="<%=yyCheckA%>" onclick="checkboxes(true);" /> <input type="button" value="<%=yyUCheck%>" onclick="checkboxes(false);" /></div>
<%					elseif isaninteger OR pract="pri" OR pract="wpr" OR pract="lpr" OR pract="stk" OR pract="mnq" OR pract="prw" OR pract="pro" then %>
					<div style="margin-top:2px"><input type="text" name="txtadd" id="txtadd" size="5" value="0" style="vertical-align:middle" /> <input type="button" value="Add" onclick="addto()" /> <input type="button" value="Set" onclick="setto()" /></div>
<%					else %>
					<div style="margin-top:2px"><input type="text" name="txtadd" id="txtadd" size="5" value="" style="vertical-align:middle" /> <input type="button" value="Set" onclick="setto()" /></div>
<%					end if
				elseif pid="" AND rid="" then
					if notifybackinstock then
						sSQL="SELECT COUNT(*) AS tcnt FROM notifyinstock"
						rs.open sSQL,cnn,0,1
						if NOT rs.EOF then
							if rs("tcnt")>0 then print "<div style=""margin-top:2px""><input type=""button"" value="""&yyStkNot&" ("&rs("tcnt")&")"&""" onclick=""document.location='adminprods.asp?act=stknot'"" /></div>"
						end if
						rs.close
					end if
				end if %></div></td>
				<td class="cobll" colspan="3"><table width="100%" cellspacing="0" cellpadding="0" border="">
					<tr>
					  <td class="cobll" align="center" style="white-space:nowrap">
						<select name="sort" size="1" onchange="changesortorder(this)" style="vertical-align:middle">
						<option value="ida"<% if sortorder="ida" then print " selected=""selected"""%>>Sort - ID ASC</option>
						<option value="idd"<% if sortorder="idd" then print " selected=""selected"""%>>Sort - ID DESC</option>
						<option value=""<% if sortorder="" then print " selected=""selected"""%>>Sort - Name ASC</option>
						<option value="nad"<% if sortorder="nad" then print " selected=""selected"""%>>Sort - Name DESC</option>
						<option value="pra"<% if sortorder="pra" then print " selected=""selected"""%>>Sort - Price ASC</option>
						<option value="prd"<% if sortorder="prd" then print " selected=""selected"""%>>Sort - Price DESC</option>
						<option value="daa"<% if sortorder="daa" then print " selected=""selected"""%>>Sort - Date ASC</option>
						<option value="dad"<% if sortorder="dad" then print " selected=""selected"""%>>Sort - Date DESC</option>
						<option value="poa"<% if sortorder="poa" then print " selected=""selected"""%>>Sort - pOrder ASC</option>
						<option value="pod"<% if sortorder="pod" then print " selected=""selected"""%>>Sort - pOrder DESC</option>
<%					if useStockManagement then print "<option value=""sta"""&IIfVr(sortorder="sta", " selected=""selected""", "")&">Sort - Stock ASC</option><option value=""std"""&IIfVr(sortorder="std", " selected=""selected""", "")&">Sort - Stock DESC</option>"
					for index=2 to adminlanguages+1
						if (adminlangsettings AND 1)=1 then %>
						<option value="na<%=index%>"<% if sortorder="na"&index then print " selected=""selected"""%>>Sort - Name<%=" " & index%></option>
<%						end if
					next %>
						<option value="ska"<% if sortorder="ska" then print " selected=""selected"""%>>Sort - SKU ASC</option>
						<option value="skd"<% if sortorder="skd" then print " selected=""selected"""%>>Sort - SKU DESC</option>
						<option value="nsa"<% if sortorder="nsa" then print " selected=""selected"""%>>Sort - Sales Rank ASC</option>
						<option value="nsd"<% if sortorder="nsd" then print " selected=""selected"""%>>Sort - Sales Rank DESC</option>
						<option value="pla"<% if sortorder="pla" then print " selected=""selected"""%>>Sort - Popularity ASC</option>
						<option value="pld"<% if sortorder="pld" then print " selected=""selected"""%>>Sort - Popularity DESC</option>
						<option value="nsf"<% if sortorder="nsf" then print " selected=""selected"""%>>No Sort (Fastest)</option>
						</select>
						<input type="submit" value="<%=yyListPd%>" onclick="startsearch('search')" />
<%					if pid<>"" OR rid<>"" then %>
						<strong>&raquo;</strong> <input type="button" value="<%=yyBckLis%>" onclick="document.location='adminprods.asp?<%=SESSION("savesearch")%>'">
<%					else %>
						<input type="button" value="<%=yyNewPr%>" onclick="newrec()" />
<%					end if %>
					  </td>
					  <td class="cobll" height="26" width="20%" align="right" style="white-space:nowrap">
<%					if pid<>"" OR rid<>"" then %>
						<input type="button" value="<%=IIfVr(pid<>"","Update Packages",yyUpdRel)%>" onclick="updaterelations('<%=IIfVr(pid<>"","packages","relations")%>')">
<%					else %>
					
<%					end if %></td>
					</tr>
				  </table></td>
			  </tr>
			</table>
<br />
            <table width="100%" class="stackable admin-table-a sta-white" id="prodstable">
<%	jscript="" : qetype="" : qesize=""
	columnlist="products.pID,pName,pName2,pName3,pDisplay,pSell,pExemptions,pShipping,pShipping2,pInStock,rootSection,pStockByOpts,pPrice,pWholesalePrice,pListPrice,pOrder,pRecommend,pGiftWrap,pSchemaType,pBackOrder,pStaticPage,pStaticURL,pSKU,pWeight,products.pSection,pDateAdded,pTax,pMinQuant,pSearchParams,pSearchParams2,pSearchParams3,pTitle,pTitle2,pTitle3,pMetaDesc,pMetaDesc2,pMetaDesc3,pPopularity,pNumSales,pCustomCSS,pUpload,pSiteID"
	if digidownloads then columnlist=columnlist&",pDownload"
	if customquickupdate<>"" then
		cqupdatearr=split(customquickupdate,",")
		for index=0 to UBOUND(cqupdatearr)
			cqitemarr=split(cqupdatearr(index),":")
			columnlist=columnlist & "," & cqitemarr(0) 
		next
	end if
	if getpost("act")="search" OR getget("pg")<>"" then
		sub displayprodrow(xrs)
			stockbyoptions=FALSE
			hascoupon="0"
			if useStockManagement then
				if cint(xrs("pStockByOpts"))<>0 then stockbyoptions=TRUE
			end if
			jscript=jscript&"pa["&resultcounter&"]=["
			%><tr id="tr<%=resultcounter%>"><td><%
				print "-"
				if pid<>"" OR rid<>"" then
					jscript=jscript&"''"
				elseif pract="prn" then
					jscript=jscript&"'"&jsspecials(xrs("pName"))&"'"
					qetype="text"
					qesize="18"
				elseif pract="prn2" then
					jscript=jscript&"'"&jsspecials(xrs("pName2"))&"'"
					qetype="text"
					qesize="18"
				elseif pract="prn3" then
					jscript=jscript&"'"&jsspecials(xrs("pName3"))&"'"
					qetype="text"
					qesize="18"
				elseif pract="sec" then
					jscript=jscript&"'"&jsspecials(xrs("pSection"))&"'"
					qetype="section"
				elseif pract="psp" then
					jscript=jscript&"'"&jsspecials(xrs("pSearchParams"))&"'"
					qetype="text"
					qesize="18"
				elseif pract="psp2" then
					jscript=jscript&"'"&jsspecials(xrs("pSearchParams2"))&"'"
					qetype="text"
					qesize="18"
				elseif pract="psp3" then
					jscript=jscript&"'"&jsspecials(xrs("pSearchParams3"))&"'"
					qetype="text"
					qesize="18"
				elseif pract="pti" then
					jscript=jscript&"'"&jsspecials(xrs("pTitle"))&"'"
					qetype="text"
					qesize="18"
				elseif pract="pti2" then
					jscript=jscript&"'"&jsspecials(xrs("pTitle2"))&"'"
					qetype="text"
					qesize="18"
				elseif pract="pti3" then
					jscript=jscript&"'"&jsspecials(xrs("pTitle3"))&"'"
					qetype="text"
					qesize="18"
				elseif pract="pmd" then
					jscript=jscript&"'"&jsspecials(xrs("pMetaDesc"))&"'"
					qetype="text"
					qesize="18"
				elseif pract="pmd2" then
					jscript=jscript&"'"&jsspecials(xrs("pMetaDesc2"))&"'"
					qetype="text"
					qesize="18"
				elseif pract="pmd3" then
					jscript=jscript&"'"&jsspecials(xrs("pMetaDesc3"))&"'"
					qetype="text"
					qesize="18"
				elseif pract="pri" then
					jscript=jscript&"'"&jsspecials(xrs("pPrice"))&"'"
					qetype="text"
					qesize="5"
				elseif pract="wpr" then
					jscript=jscript&"'"&jsspecials(xrs("pWholesalePrice"))&"'"
					qetype="text"
					qesize="5"
				elseif pract="lpr" then
					jscript=jscript&"'"&jsspecials(xrs("pListPrice"))&"'"
					qetype="text"
					qesize="5"
				elseif pract="sid" then
					jscript=jscript&"'"&xrs("pSiteID")&"'"
					qetype="text"
					qesize="5"
				elseif pract="stk" then
					jscript=jscript&IIfVr(stockbyoptions,"false","'"&jsspecials(xrs("pInStock"))&"'")
					qetype="text"
					qesize="5"
				elseif pract="pop" then
					jscript=jscript&"'"&xrs("pPopularity")&"'"
					qetype="text"
					qesize="5"
				elseif pract="sal" then
					jscript=jscript&"'"&xrs("pNumSales")&"'"
					qetype="text"
					qesize="5"
				elseif pract="mnq" then
					jscript=jscript&IIfVr(stockbyoptions,"false","'"&jsspecials(xrs("pMinQuant")+1)&"'")
					qetype="text"
					qesize="5"
				elseif pract="css" then
					jscript=jscript&"'"&jsspecials(xrs("pCustomCSS"))&"'"
					qetype="text"
					qesize="16"
				elseif pract="sta" then
					jscript=jscript&IIfVr(stockbyoptions,"false","''")
					qetype="text"
					qesize="5"
				elseif pract="del" then
					jscript=jscript&"'del'"
					qetype="delbox"
				elseif pract="pru" then
					jscript=jscript&IIfVr(xrs("pUpload")<>0,1,0)
					qetype="checkbox"
				elseif pract="prw" then
					jscript=jscript&"'"&jsspecials(xrs("pWeight"))&"'"
					qetype="text"
					qesize="5"
				elseif pract="pra" AND currentattribute<>"" then
					sSQL="SELECT mSCscID FROM multisearchcriteria WHERE mSCpID='"&escape_string(xrs("pID"))&"' AND mSCscID="&currentattribute
					rs2.open sSQL,cnn,0,1
					jscript=jscript&IIfVr(rs2.EOF,0,1)
					rs2.close
					qetype="checkbox"
				elseif pract="prp" AND currentoption<>"" then
					sSQL="SELECT poID FROM prodoptions WHERE poProdID='"&escape_string(xrs("pID"))&"' AND poOptionGroup="&currentoption
					rs2.open sSQL,cnn,0,1
					jscript=jscript&IIfVr(rs2.EOF,0,1)
					rs2.close
					qetype="checkbox"
				elseif pract="dis" AND currentdiscount<>"" then
					sSQL="SELECT cpaID FROM cpnassign WHERE cpaType=2 AND cpaAssignment='"&escape_string(xrs("pID"))&"' AND cpaCpnID="&currentdiscount
					rs2.open sSQL,cnn,0,1
					jscript=jscript&IIfVr(rs2.EOF,0,1)
					rs2.close
					qetype="checkbox"
				elseif pract="ads" AND currentsection<>"" then
					sSQL="SELECT pID FROM multisections WHERE pID='"&escape_string(xrs("pID"))&"' AND pSection="&currentsection
					rs2.open sSQL,cnn,0,1
					jscript=jscript&IIfVr(rs2.EOF,0,1)
					rs2.close
					qetype="checkbox"
				elseif pract="dip" then
					jscript=jscript&IIfVr(cint(xrs("pDisplay"))<>0,1,0)
					qetype="checkbox"
				elseif pract="stp" then
					jscript=jscript&IIfVr(cint(xrs("pStaticPage"))<>0,1,0)
					qetype="checkbox"
				elseif pract="stu" then
					jscript=jscript&"'"&jsspecials(xrs("pStaticURL"))&"'"
					qetype="text"
					qesize="18"
				elseif pract="rec" then
					jscript=jscript&IIfVr(cint(xrs("pRecommend"))<>0,1,0)
					qetype="checkbox"
				elseif pract="gwr" then
					jscript=jscript&IIfVr(cint(xrs("pGiftWrap"))<>0,1,0)
					qetype="checkbox"
				elseif pract="isa" then
					jscript=jscript&IIfVr(cint(xrs("pSchemaType"))<>0,1,0)
					qetype="checkbox"
				elseif pract="bak" then
					jscript=jscript&IIfVr(cint(xrs("pBackOrder"))<>0,1,0)
					qetype="checkbox"
				elseif pract="sku" then
					jscript=jscript&"'"&jsspecials(xrs("pSKU"))&"'"
					qetype="text"
					qesize="10"
				elseif pract="pro" then
					jscript=jscript&"'"&jsspecials(xrs("pOrder"))&"'"
					qetype="text"
					qesize="5"
				elseif pract="ppt" then
					jscript=jscript&"'"&jsspecials(xrs("pTax"))&"'"
					qetype="text"
					qesize="5"
				elseif pract="sel" then
					jscript=jscript&IIfVr(cint(xrs("pSell"))<>0,1,0)
					qetype="checkbox"
				elseif pract="ste" OR pract="cte" OR pract="she" OR pract="hae" OR pract="fse" OR pract="pte" OR pract="pde" then
					fieldnum=1
					if pract="cte" then fieldnum=2
					if pract="she" then fieldnum=4
					if pract="hae" then fieldnum=8
					if pract="fse" then fieldnum=16
					if pract="pte" then fieldnum=32
					if pract="pde" then fieldnum=64
					jscript=jscript&IIfVr((xrs("pExemptions") AND fieldnum)<>0,1,0)
					qetype="checkbox"
				elseif pract="dld" then
					jscript=jscript&"'"&jsspecials(xrs("pDownload"))&"'"
					qetype="text"
					qesize="18"
				elseif pract="frs" then
					jscript=jscript&"'"&jsspecials(xrs("pShipping"))&"'"
					qetype="text"
					qesize="5"
				elseif pract="daa" then
					jscript=jscript&"'"&jsspecials(xrs("pDateAdded"))&"'"
					qetype="text"
					qesize="8"
				elseif pract="csu" then
					jscript=jscript&"0"
					qetype="checkbox"
				elseif pract="vis" OR pract="vil" OR pract="vig" then
					imagetype=0
					if pract="vil" then imagetype=1
					if pract="vig" then imagetype=2
					sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 1 ")&"imageSrc FROM productimages WHERE imageProduct='"&escape_string(xrs("pID"))&"' AND imageType="&imagetype&" ORDER BY imageNumber"&IIfVs(mysqlserver=TRUE," LIMIT 0,1")
					rs2.open sSQL,cnn,0,1
					if rs2.EOF then
						jscript=jscript&"''"
					else
						jscript=jscript&"'"&IIfVs(lcase(left(rs2("imageSrc"),5))<>"http:" AND lcase(left(rs2("imageSrc"),6))<>"https:" AND left(rs2("imageSrc"),1)<>"/","../") & jsspecials(rs2("imageSrc"))&"'"
					end if
					rs2.close
					qetype="image"
				else
					hascustom=FALSE
					if customquickupdate<>"" then
						cqupdatearr=split(customquickupdate,",")
						for index=0 to UBOUND(cqupdatearr)
							cqitemarr=split(cqupdatearr(index),":")
							if pract=cqitemarr(0) then
								qetype="text"
								hascustom=TRUE
								if UBOUND(cqitemarr)>=2 then
									if lcase(cqitemarr(2))="check" then qetype="checkbox"
								end if
								if qetype="checkbox" then
									tval=xrs(cqitemarr(0))
									if isnull(tval) then tval=0
									jscript=jscript&IIfVr(cint(tval)<>0,1,0)
								else
									jscript=jscript&"'"&jsspecials(xrs(cqitemarr(0)))&"'"
									qesize="18"
								end if
							end if
						next
					end if
					if NOT hascustom then print "-"
				end if
			%></td><td>-</td><td><%
					if noautocheckorphans=TRUE AND request("disp")<>"4" then
						' nothing
					elseif IsNull(xrs("rootSection")) then
						print redasterix&" "
						haveerrprods=true
					elseif cint(xrs("rootSection"))<>1 then
						print redasterix&" "
						haveerrprods=true
					end if
					hasstock=TRUE
					if cint(xrs("pDisplay"))=0 OR ((useStockManagement AND xrs("pInStock") <= 0 AND NOT stockbyoptions) OR (NOT useStockManagement AND cint(xrs("pSell"))=0)) then hasstock=FALSE
					if NOT hasstock then print "<span style=""color:#FF0000"">"
					if cint(xrs("pDisplay"))=0 then print "<strike>"
					print xrs("pName"&IIfVr(sortorder="na2",2,IIfVs(sortorder="na3",3)))
					if cint(xrs("pDisplay"))=0 then print "</strike>"
					if NOT hasstock then print "</span>"
					if useStockManagement then
						if stockbyoptions then print " (-)" else print " (" & xrs("pInStock") & ")"
					end if %></td><td>-</td><%
				if pid<>"" OR rid<>"" then
			%><td><input type="hidden" name="updq<%=resultcounter%>" value="<%=htmlspecials(xrs("pID"))%>"><input type="checkbox" name="updr<%=resultcounter%>" id="chkbx<%=resultcounter%>" value="1" <%
					hascoupon="0"
					if pid=xrs("pID") OR rid=xrs("pID") then
						print "disabled "
					else
						if pid<>"" then print "onchange=""tqn(this,"&resultcounter&")"" "
						if isarray(ridarr) then
							for index=0 to UBOUND(ridarr,2)
								if pid<>"" then
									if ridarr(0,index)=xrs("pID") then print "checked=""checked"" " : hascoupon=ridarr(1,index) : exit for
								else
									if ridarr(0,index)=xrs("pID") then print "checked=""checked"" " : exit for
								end if
							next
						end if
					end if %>/></td><%
				else
					hascoupon="0"
					if isarray(allcoupon) then
						for index=0 to UBOUND(allcoupon,2)
							if trim(allcoupon(0,index))=xrs("pID") then
								hascoupon="1"
								exit for
							end if
						next
					end if
			%><td>-</td><%
				end if
			if pid="" AND rid="" then print "<td>-</td>"
			%><td>-</td><td>-</td></tr>
<%			jscript=jscript&",'"&jsspecials(xrs("pID"))&"'," & hascoupon & "];"&vbCrLf
			resultcounter=resultcounter + 1
		end sub
		sub displayheaderrow() %>
			<tr>
				<th class="small minicell">
<%			if pid="" AND rid="" then %>
					<select name="pract" id="pract" size="1" onchange="changepract(this)" style="width:150px">
					<option value="none">Quick Entry...</option>
					<option value="ads"<% if pract="ads" then print " selected=""selected"""%>><%=yyAddSec%> / Categories</option>
					<option value="bak"<% if pract="bak" then print " selected=""selected"""%>><%=yyBakOrd%></option>
					<option value="css"<% if pract="css" then print " selected=""selected"""%>><%="Custom CSS Class"%></option>
					<option value="daa"<% if pract="daa" then print " selected=""selected"""%>><%=yyDateAd%></option>
<%			if digidownloads then %>
					<option value="dld"<% if pract="dld" then print " selected=""selected"""%>>Digital Download</option>
<%			end if %>
					<option value="dis"<% if pract="dis" then print " selected=""selected"""%>><%=yyDiscnt%></option>
					<option value="dip"<% if pract="dip" then print " selected=""selected"""%>><%=yyDisPro%></option>
					<option value="frs"<% if pract="frs" then print " selected=""selected"""%>>Flat Rate Shipping</option>
					<option value="gwr"<% if pract="gwr" then print " selected=""selected"""%>><%=yyGifWra%></option>
					<option value="pru"<% if pract="pru" then print " selected=""selected"""%>><%="Image Upload"%></option>
					<option value="isa"<% if pract="isa" then print " selected=""selected"""%>><%="Is Article"%></option>
					<option value="lpr"<% if pract="lpr" then print " selected=""selected"""%>><%=yyListPr%></option>
					<option value="mnq"<% if pract="mnq" then print " selected=""selected"""%>><%="Minimum Quantity"%></option>
					<option value="pop"<% if pract="pop" then print " selected=""selected"""%>><%="Popularity"%></option>
					<option value="pri"<% if pract="pri" then print " selected=""selected"""%>><%=yyPrPri%></option>
					<option value="pra"<% if pract="pra" then print " selected=""selected"""%>><%=yySeaCri%></option>
					<option value="prn"<% if pract="prn" then print " selected=""selected"""%>><%=yyPrName%></option>
<%			for index=2 to adminlanguages+1
				if (adminlangsettings AND 1)=1 then print "<option value=""prn"&index&""""&IIfVr(pract=("prn"&index)," selected=""selected""","")&">"&yyPrName&" "&index&"</option>"
			next %>
					<option value="prp"<% if pract="prp" then print " selected=""selected"""%>>Product Options</option>
					<option value="pro"<% if pract="pro" then print " selected=""selected"""%>><%=yyProdOr%></option>
					<option value="prw"<% if pract="prw" then print " selected=""selected"""%>><%=yyPrWght%></option>
					<option value="rec"<% if pract="rec" then print " selected=""selected"""%>><%=yyRecomd%></option>
					<option value="sal"<% if pract="sal" then print " selected=""selected"""%>><%="Sales Rank"%></option>
					<option value="sec"<% if pract="sec" then print " selected=""selected"""%>><%=yySection%> / Category</option>
<%			if NOT useStockManagement then %>
					<option value="sel"<% if pract="sel" then print " selected=""selected"""%>><%=yySellBut%></option>
<%			end if %>
					<option value="sku"<% if pract="sku" then print " selected=""selected"""%>>SKU</option>
					<option value="sid"<% if pract="sid" then print " selected=""selected"""%>>Site ID</option>
					<option value="stk"<% if pract="stk" then print " selected=""selected"""%>><%=yyStck%></option>
					<option value="sta"<% if pract="sta" then print " selected=""selected"""%>><%=yyStck%> Add</option>
<%			if perproducttaxrate then %>
					<option value="ppt"<% if pract="ppt" then print " selected=""selected"""%>><%=yyTax%></option>
<%			end if %>
					<option value="wpr"<% if pract="wpr" then print " selected=""selected"""%>><%=yyWhoPri%></option>
					<option value="" disabled="disabled">---------------------</option>
					<option value="pti"<% if pract="pti" then print " selected=""selected"""%>><%="Page Title Tag"%></option>
<%	for index=2 to adminlanguages+1
		if (adminlangsettings AND 2097152)=2097152 then print "<option value=""pti" & index & """" & IIfVs(pract=("pti"&index)," selected=""selected""") & ">" & "Page Title Tag" & " " & index & "</option>"
	next %>
					<option value="pmd"<% if pract="pmd" then print " selected=""selected"""%>><%="Meta Description"%></option>
<%	for index=2 to adminlanguages+1
		if (adminlangsettings AND 2097152)=2097152 then print "<option value=""pmd" & index & """" & IIfVs(pract=("pmd"&index)," selected=""selected""") & ">" & "Meta Description" & " " & index & "</option>"
	next %>
					<option value="psp"<% if pract="psp" then print " selected=""selected"""%>><%=yyAddSrP%></option>
<%	for index=2 to adminlanguages+1
		if (adminlangsettings AND 4194304)=4194304 then print "<option value=""psp" & index & """" & IIfVs(pract=("psp"&index)," selected=""selected""") & ">" & yyAddSrP & " " & index & "</option>"
	next %>
					<option value="" disabled="disabled">---------------------</option>
					<option value="cte"<% if pract="cte" then print " selected=""selected"""%>>Country Tax Exempt</option>
					<option value="fse"<% if pract="fse" then print " selected=""selected"""%>>Free Shipping Exempt</option>
					<option value="hae"<% if pract="hae" then print " selected=""selected"""%>>Handling Exempt</option>
					<option value="pte"<% if pract="pte" then print " selected=""selected"""%>>Pack Together Exempt</option>
					<option value="pde"<% if pract="pde" then print " selected=""selected"""%>>Product Discount Exempt</option>
					<option value="she"<% if pract="she" then print " selected=""selected"""%>>Shipping Exempt</option>
					<option value="ste"<% if pract="ste" then print " selected=""selected"""%>>State Tax Exempt</option>
					<option value="" disabled="disabled">---------------------</option>
					<option value="stp"<% if pract="stp" then print " selected=""selected"""%>><%=yyStatPg%></option>
					<option value="stu"<% if pract="stu" then print " selected=""selected"""%>>Static URL</option>
					<option value="csu"<% if pract="csu" then print " selected=""selected"""%>>Create Static URL</option>
					<option value="" disabled="disabled">---------------------</option>
					<option value="vis"<% if pract="vis" then print " selected=""selected"""%>>View Small Image</option>
					<option value="vil"<% if pract="vil" then print " selected=""selected"""%>>View Large Image</option>
					<option value="vig"<% if pract="vig" then print " selected=""selected"""%>>View Giant Image</option>
					<option value="" disabled="disabled">---------------------</option>
					<option value="del"<% if pract="del" then print " selected=""selected"""%>><%=yyDelete%></option>
<%			if customquickupdate<>"" then
				cqupdatearr=split(customquickupdate,",")
				print "<option value="""" disabled=""disabled"">---------------------</option>"
				for index=0 to UBOUND(cqupdatearr)
					cqitemarr=split(cqupdatearr(index),":")
					print "<option value="""&cqitemarr(0)&""""
					if pract=cqitemarr(0) then print " selected=""selected"""
					print ">"
					if UBOUND(cqitemarr)>0 then print cqitemarr(1) else print cqitemarr(0)
					print "</option>"
				next
			end if %>
					</select><%
			end if
			if pid<>"" OR rid<>"" then
				print "-"
			elseif pract="csu" then
				print "<div style=""margin-top:6px;margin-left:6px;text-align:left"">"
				print "<div style=""text-align:center"" id=""staticurlshow""><input type=""button"" value=""Show Options"" onclick=""document.getElementById('staticurlshow').style.display='none';document.getElementById('staticurloptions').style.display='';"" /></div>"
				print "<div id=""staticurloptions"" style=""display:none""><select size=""1"" name=""extension"" title=""Has Extension (.asp)"" onchange=""setCookie('incpextension',this[this.selectedIndex].value,365)""><option value=""yes"">Extension (.asp)</option><option value=""no"""&IIfVs(seodetailurls OR request.cookies("incpextension")="no"," selected=""selected""")&">Extensionless</option></select>"
				print "<select size=""1"" name=""space"" title=""Space Replacement"" onchange=""setCookie('incpspace',this[this.selectedIndex].value,365)"">" & _
					"<option value="" "">No Space Replacement</option>" & _
					"<option value=""_"""&IIfVs(request.cookies("incpspace")="_" AND detlinkspacechar<>"_"," selected=""selected""")&IIfVs(detlinkspacechar="_"," disabled=""disabled""")&">Underscore"&IIfVs(detlinkspacechar="_"," (detlinkspacechar)")&"</option>" & _
					"<option value=""-"""&IIfVs(request.cookies("incpspace")="-" AND detlinkspacechar<>"-"," selected=""selected""")&IIfVs(detlinkspacechar="-"," disabled=""disabled""")&">Dash"&IIfVs(detlinkspacechar="-"," (detlinkspacechar)")&"</option>" & _
					"<option value=""remove"""&IIfVs(request.cookies("incpspace")="remove"," selected=""selected""")&">Remove</option></select>"
				print "<select size=""1"" name=""lcase"" title=""Lower Case"" onchange=""setCookie('incplcase',this[this.selectedIndex].value,365)""><option value=""no"">Keep Original Case</option><option value=""yes"""&IIfVs(request.cookies("incplcase")="yes"," selected=""selected""")&">Force Lower Case</option></select>"
				print "<select size=""1"" name=""punctuation"" title=""Remove Punctuation"" onchange=""setCookie('incpunctuation',this[this.selectedIndex].value,365)""><option value="""">Keep Punctuation</option><option value=""remove"""&IIfVs(request.cookies("incpunctuation")="remove"," selected=""selected""")&">Remove Punctuation</option></select>"
				print "<select size=""1"" name=""wholedb"" title=""Create Static URL's for All Products""><option value="""">Selected Items Only</option><option value=""clear"">Clear All Static URL's</option><option value=""set"">Create All Static URL's</option></select>"
				print "<select size=""1"" name=""addprodid"" title=""Include Product ID""><option value="""">Don't Include Product ID</option><option value=""prepend"">Prepend Product ID</option><option value=""append"">Append Product ID</option></select></div>"
				print "</div>"
			elseif pract="pra" then
				if is_numeric(request.cookies("currattr")) then
					currentattribute=int(request.cookies("currattr"))
					rs2.open  "SELECT scID FROM searchcriteria WHERE scID="&currentattribute,cnn,0,1
					if rs2.EOF then currentattribute=""
					rs2.close
				else
					currentattribute=""
				end if
				currentgroupid=-1
				print "<div style=""margin-top:2px""><select style=""width:150px"" name=""currentattribute"" size=""1"" onchange=""setCookie('currattr',this[this.selectedIndex].value,600);changepract(document.getElementById('pract'))"">"
				sSQL="SELECT scID,scWorkingName,scgID,scgWorkingName FROM searchcriteria INNER JOIN searchcriteriagroup ON searchcriteria.scGroup=searchcriteriagroup.scgID ORDER BY scgWorkingName,scOrder"
				rs2.open sSQL,cnn,0,1
				if rs2.EOF then print "<option value="""" disabled=""disabled"">== No Attributes Defined ==</option>" & vbCrLf
				do while NOT rs2.EOF
					if currentgroupid<>rs2("scgID") then
						print "<option value="""" disabled=""disabled"">== " & rs2("scgWorkingName") & " ==</option>" & vbCrLf
						currentgroupid=rs2("scgID")
					end if
					print "<option value=""" & rs2("scID") & """" & IIfVs(currentattribute=rs2("scID")," selected=""selected""") & ">" & rs2("scWorkingName") & "</option>" & vbCrLf
					if currentattribute="" then currentattribute=rs2("scID")
					rs2.movenext
				loop
				rs2.close
				print "</select></div>"
			elseif pract="prp" then
				if is_numeric(request.cookies("curroptn")) then
					currentoption=int(request.cookies("curroptn"))
					rs2.open "SELECT optGrpID FROM optiongroup WHERE optGrpID="&currentoption,cnn,0,1
					if rs2.EOF then currentoption=""
					rs2.close
				else
					currentoption=""
				end if
				print "<div style=""margin-top:2px""><select style=""width:150px"" name=""currentoption"" size=""1"" onchange=""setCookie('curroptn',this[this.selectedIndex].value,600);changepract(document.getElementById('pract'))"">"
				sSQL="SELECT optGrpID,optGrpWorkingName FROM optiongroup ORDER BY optGrpWorkingName"
				rs2.open sSQL,cnn,0,1
				if rs2.EOF then print "<option value="""" disabled=""disabled"">== No Options Defined ==</option>" & vbCrLf
				do while NOT rs2.EOF
					print "<option value=""" & rs2("optGrpID") & """" & IIfVs(currentoption=rs2("optGrpID")," selected=""selected""") & ">" & rs2("optGrpWorkingName") & "</option>" & vbCrLf
					if currentoption="" then currentoption=rs2("optGrpID")
					rs2.movenext
				loop
				rs2.close
				print "</select></div>"
			elseif pract="dis" then
				if is_numeric(request.cookies("currdisc")) then
					currentdiscount=int(request.cookies("currdisc"))
					rs2.open "SELECT cpnID FROM coupons WHERE cpnSitewide=0 AND cpnID="&currentdiscount,cnn,0,1
					if rs2.EOF then currentdiscount=""
					rs2.close
				else
					currentdiscount=""
				end if
				print "<div style=""margin-top:2px""><select style=""width:150px"" name=""currentdiscount"" size=""1"" onchange=""setCookie('currdisc',this[this.selectedIndex].value,600);changepract(document.getElementById('pract'))"">"
				sSQL="SELECT cpnID,cpnWorkingName FROM coupons WHERE cpnSitewide=0 ORDER BY cpnWorkingName"
				rs2.open sSQL,cnn,0,1
				if rs2.EOF then print "<option value="""" disabled=""disabled"">== No Assignable Discounts Defined ==</option>" & vbCrLf
				do while NOT rs2.EOF
					print "<option value=""" & rs2("cpnID") & """" & IIfVs(currentdiscount=rs2("cpnID")," selected=""selected""") & ">" & rs2("cpnWorkingName") & "</option>" & vbCrLf
					if currentdiscount="" then currentdiscount=rs2("cpnID")
					rs2.movenext
				loop
				rs2.close
				print "</select></div>"
			elseif pract="ads" then
				if is_numeric(request.cookies("currsec")) then
					currentsection=int(request.cookies("currsec"))
					rs2.open "SELECT sectionID FROM sections WHERE rootSection=1 AND sectionID="&currentsection,cnn,0,1
					if rs2.EOF then currentsection=""
					rs2.close
				else
					currentsection=""
				end if
				print "<div style=""margin-top:2px""><select style=""width:150px"" name=""currentsection"" size=""1"" onchange=""setCookie('currsec',this[this.selectedIndex].value,600);changepract(document.getElementById('pract'))"">"
				sSQL="SELECT sectionID,sectionWorkingName FROM sections WHERE rootSection=1 ORDER BY sectionWorkingName"
				rs2.open sSQL,cnn,0,1
				if rs2.EOF then print "<option value="""" disabled=""disabled"">== No Categories Defined ==</option>" & vbCrLf
				do while NOT rs2.EOF
					print "<option value=""" & rs2("sectionID") & """" & IIfVs(currentsection=rs2("sectionID")," selected=""selected""") & ">" & htmlspecials(rs2("sectionWorkingName")) & "</option>" & vbCrLf
					if currentsection="" then currentsection=rs2("sectionID")
					rs2.movenext
				loop
				rs2.close
				print "</select></div>"
			end if %></th>
				<th style="width:20%"><strong><%=yyPrId%></strong></th>
				<th style="width:30%"><strong><%=yyPrName%></strong></th>
				<th style="width:5%;text-align:center" class="small"><%=yyDiscnt%></th>
				<th style="width:5%;text-align:center" class="small"><%=IIfVr(pid<>"","Package",yyRelate)%></th>
				<th style="width:5%;text-align:center" class="small"><%=IIfVr(pid<>"","Quantity","Package")%></th>
<%			if pid="" AND rid="" then %>
				<th style="width:5%;text-align:center" class="small">Alt IDs</th>
<%			end if %>
				<th style="width:5%;text-align:center" class="small"><%=yyModify%></th>
			</tr>
<%		end sub
		sSQL="SELECT DISTINCT cpaAssignment FROM cpnassign WHERE cpaType=2"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then allcoupon=rs.getrows
		rs.close
		if getget("package")="go" then
			sSQL="SELECT DISTINCT " & columnlist & " FROM productpackages INNER JOIN (products LEFT OUTER JOIN sections ON products.pSection=sections.sectionID) ON products.pId=productpackages.pId WHERE packageID='"&escape_string(pid)&"'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				displayheaderrow()
				do while NOT rs.EOF
					displayprodrow(rs)
					rs.MoveNext
					Count=Count + 1
				loop
			else
				yyPrNoPk="There are currently no products included in this package."
				print "<tr><td width=""100%"" colspan=""6"" align=""center""><p>&nbsp;</p><p>"&yyPrNoPk&"</p><p>"&yyPrReSe&"</p><p>"&yyPrReLs&"</p>&nbsp;</td></tr>"
			end if
			rs.close
		elseif getget("related")="go" then
			if mysqlserver=TRUE then
				sSQL="SELECT DISTINCT " & columnlist & " FROM relatedprods INNER JOIN products ON products.pId=relatedprods.rpRelProdId LEFT OUTER JOIN sections ON products.pSection=sections.sectionID WHERE rpProdId='"&escape_string(rid)&"'"
				if relatedproductsbothways=TRUE then sSQL=sSQL & "UNION SELECT DISTINCT " & columnlist & " FROM relatedprods INNER JOIN products ON products.pId=relatedprods.rpProdId LEFT OUTER JOIN sections ON products.pSection=sections.sectionID WHERE rpRelProdId='"&escape_string(rid)&"'"
			else
				sSQL="SELECT DISTINCT " & columnlist & " FROM relatedprods INNER JOIN (products LEFT OUTER JOIN sections ON products.pSection=sections.sectionID) ON products.pId=relatedprods.rpRelProdId WHERE rpProdId='"&escape_string(rid)&"'"
				if relatedproductsbothways=TRUE then sSQL=sSQL & "UNION SELECT DISTINCT " & columnlist & " FROM relatedprods INNER JOIN (products LEFT OUTER JOIN sections ON products.pSection=sections.sectionID) ON products.pId=relatedprods.rpProdId WHERE rpRelProdId='"&escape_string(rid)&"'"
			end if
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				displayheaderrow()
				do while NOT rs.EOF
					displayprodrow(rs)
					rs.MoveNext
					Count=Count + 1
				loop
			else
				print "<tr><td width=""100%"" colspan=""6"" align=""center""><p>&nbsp;</p><p>"&yyPrNoRe&"</p><p>"&yyPrReSe&"</p><p>"&yyPrReLs&"</p>&nbsp;</td></tr>"
			end if
			rs.close
		else
			whereand=" WHERE "
			if thecat="" OR sortorder="nsf" then
				sSQL="SELECT " & columnlist & " FROM products LEFT OUTER JOIN sections ON products.pSection=sections.sectionID"
			elseif mysqlserver=true then
				sSQL="SELECT DISTINCT " & columnlist & " FROM multisections RIGHT JOIN products ON products.pId=multisections.pId LEFT OUTER JOIN sections ON products.pSection=sections.sectionID"
			else
				sSQL="SELECT DISTINCT " & columnlist & " FROM " & IIfVs(thecat<>"" AND (catorman="man" OR catorman="dis"),"(") & "multisections RIGHT JOIN (products LEFT OUTER JOIN sections ON products.pSection=sections.sectionID) ON products.pId=multisections.pId" & IIfVs(thecat<>"" AND (catorman="man" OR catorman="dis"),")")
			end if
			if thecat<>"" then
				if catorman="dis" then
					sSQL=sSQL & " INNER JOIN cpnassign ON products.pID=cpnassign.cpaAssignment" & whereand & "cpnassign.cpaCpnID=" & thecat : whereand=" AND "
				elseif catorman="man" then
					sSQL=sSQL & " INNER JOIN multisearchcriteria ON products.pID=multisearchcriteria.mSCpID" & whereand & "multisearchcriteria.mSCscID=" & thecat : whereand=" AND "
				else
					sectionids=getsectionids(thecat, TRUE)
					if sectionids<>"" then
						if sortorder="nsf" then
							sSQL=sSQL & whereand & "products.pSection IN (" & sectionids & ") "
						else
							sSQL=sSQL & whereand & "(products.pSection IN (" & sectionids & ") OR multisections.pSection IN (" & sectionids & "))"
						end if
						whereand=" AND "
					end if
				end if
			end if
			if noautocheckorphans=TRUE AND request("disp")<>"4" then
				sSQL=replace(sSQL,"rootSection,","")
				sSQL=replace(sSQL,"LEFT OUTER JOIN sections ON products.pSection=sections.sectionID","")
			end if
			sprice=trim(request("sprice"))
			if sprice<>"" then
				if instr(sprice, "-") > 0 then
					pricearr=split(sprice, "-")
					if NOT is_numeric(pricearr(0)) then pricearr(0)=0
					if NOT is_numeric(pricearr(1)) then pricearr(1)=10000000
					sSQL=sSQL & whereand & "pPrice BETWEEN "&cdbl(replace(pricearr(0),"$",""))&" AND "&cdbl(replace(pricearr(1),"$",""))
					whereand=" AND "
				elseif is_numeric(sprice) then
					sSQL=sSQL & whereand & "pPrice="&cdbl(replace(sprice,"$",""))&" "
					whereand=" AND "
				end if
			end if
			if trim(request("stext"))<>"" AND sos<>63 then
				sText=escape_string(request("stext"))
				aText=split(sText)
				if (sos AND 2)=2 then aFields(0)="" else aFields(0)="products.pId"
				if (sos AND 2)=2 then aFields(1)="" else aFields(1)="pSKU"
				if (sos AND 4)=4 then aFields(2)="" else aFields(2)=getlangid("pName",1)
				if (sos AND 8)=8 then aFields(3)="" else aFields(3)=getlangid("pDescription",2)
				if (sos AND 16)=16 then aFields(4)="" else aFields(4)=getlangid("pLongDescription",2)
				if (sos AND 32)=32 then aFields(5)="" else aFields(5)=getlangid("pSearchParams",4194304)
				if request("stype")="exact" then
					sSQL=sSQL & whereand & "("
					if (sos AND 2)<>2 then sSQL=sSQL & "products.pId LIKE '%"&sText&"%' OR pSKU LIKE '%"&sText&"%' OR "
					if (sos AND 4)<>4 then sSQL=sSQL & getlangid("pName",1)&" LIKE '%"&sText&"%' OR "
					if (sos AND 8)<>8 then sSQL=sSQL & getlangid("pDescription",2)&" LIKE '%"&sText&"%' OR "
					if (sos AND 16)<>16 then sSQL=sSQL & getlangid("pLongDescription",2)&" LIKE '%"&sText&"%' OR"
					sSQL=left(sSQL,len(sSQL)-3) & ") "
					whereand=" AND "
				else
					sJoin="AND "
					if request("stype")="any" then sJoin="OR "
					sSQL=sSQL & whereand&"("
					whereand=" AND "
					for index=0 to 5
						if aFields(index)<>"" then
							sSQL=sSQL & "("
							for rowcounter=0 to UBOUND(aText)
								sSQL=sSQL & aFields(index) & " LIKE '%"&aText(rowcounter)&"%' " & sJoin
							next
							sSQL=left(sSQL,len(sSQL)-len(sJoin)) & ") OR "
						end if
					next
					sSQL=left(sSQL,len(sSQL)-4) & ") "
				end if
			end if
			if request("disp")="6" then sSQL=sSQL & whereand & "pBackOrder<>0" : whereand=" AND "
			if request("disp")="7" then sSQL=sSQL & whereand & "pBackOrder=0" : whereand=" AND "
			if request("disp")="8" then sSQL=sSQL & whereand & "pGiftWrap<>0" : whereand=" AND "
			if request("disp")="9" then sSQL=sSQL & whereand & "pGiftWrap=0" : whereand=" AND "
			if request("disp")="10" then sSQL=sSQL & whereand & "pRecommend<>0" : whereand=" AND "
			if request("disp")="11" then sSQL=sSQL & whereand & "pRecommend=0" : whereand=" AND "
			if request("disp")="12" then sSQL=sSQL & whereand & "pStaticPage<>0" : whereand=" AND "
			if request("disp")="13" then sSQL=sSQL & whereand & "pStaticPage=0" : whereand=" AND "
			if request("disp")="4" then sSQL=sSQL & whereand & "(rootSection IS NULL OR rootSection=0)" : whereand=" AND "
			if request("disp")="3" then sSQL=sSQL & whereand & "(pInStock<=0 AND pStockByOpts=0)" : whereand=" AND "
			if request("disp")="" OR request("disp")="5" then sSQL=sSQL & whereand & "pDisplay<>0" : whereand=" AND "
			if request("disp")="2" then sSQL=sSQL & whereand & "pDisplay=0" : whereand=" AND "
			if sortorder="ida" then
				sSQL=sSQL & " ORDER BY products.pid"
			elseif sortorder="idd" then
				sSQL=sSQL & " ORDER BY products.pid DESC"
			elseif sortorder="" then
				sSQL=sSQL & " ORDER BY pName"
			elseif sortorder="na2" then
				sSQL=sSQL & " ORDER BY pName2"
			elseif sortorder="na3" then
				sSQL=sSQL & " ORDER BY pName3"
			elseif sortorder="nad" then
				sSQL=sSQL & " ORDER BY pName DESC"
			elseif sortorder="pra" then
				sSQL=sSQL & " ORDER BY pPrice"
			elseif sortorder="prd" then
				sSQL=sSQL & " ORDER BY pPrice DESC"
			elseif sortorder="daa" then
				sSQL=sSQL & " ORDER BY pDateAdded"
			elseif sortorder="dad" then
				sSQL=sSQL & " ORDER BY pDateAdded DESC"
			elseif sortorder="poa" then
				sSQL=sSQL & " ORDER BY pOrder"
			elseif sortorder="pod" then
				sSQL=sSQL & " ORDER BY pOrder DESC"
			elseif sortorder="sta" then
				sSQL=sSQL & " ORDER BY products.pInStock"
			elseif sortorder="std" then
				sSQL=sSQL & " ORDER BY products.pInStock DESC"
			elseif sortorder="ska" OR sortorder="skd" then
				sSQL=sSQL & " ORDER BY products.pSKU" & IIfVs(sortorder="skd"," DESC")
			elseif sortorder="nsa" OR sortorder="nsd" then
				sSQL=sSQL & " ORDER BY products.pNumSales" & IIfVs(sortorder="nsd"," DESC")
			elseif sortorder="pla" OR sortorder="pld" then
				sSQL=sSQL & " ORDER BY products.pPopularity" & IIfVs(sortorder="pld"," DESC")
			end if
			if adminproductsperpage="" then adminproductsperpage=200
			rs.CursorLocation=3 ' adUseClient
			rs.CacheSize=adminproductsperpage
			rs.open sSQL, cnn
			if rs.eof or rs.bof then
				success=false
				iNumOfPages=0
			else
				success=true
				rs.MoveFirst
				rs.PageSize=adminproductsperpage
				CurPage=1
				if is_numeric(getget("pg")) then CurPage=int(getget("pg"))
				iNumOfPages=Int((rs.RecordCount + (adminproductsperpage-1)) / adminproductsperpage)
				rs.AbsolutePage=CurPage
			end if
			Count=0
			haveerrprods=FALSE
			if NOT rs.EOF then
				pblink="<a href=""adminprods.asp?"&IIfVs(request("sos")<>"","sos="&request("sos")&"&")&IIfVs(request("pid")<>"","pid="&request("pid")&"&")&IIfVs(request("rid")<>"","rid="&request("rid")&"&")&"disp="&request("disp")&"&scat="&request("scat")&"&stext="&urlencode(request("stext"))&"&stype="&request("stype")&"&sprice="&urlencode(sprice)&"&pg="
				if iNumOfPages > 1 then print "<tr><td colspan=""8"" align=""center"">" & writepagebar(CurPage,iNumOfPages,yyPrev,yyNext,pblink,FALSE) & "</td></tr>"
				displayheaderrow()
				addcomma=""
				do while NOT rs.EOF AND Count < rs.PageSize
					displayprodrow(rs)
					pidlist=pidlist&addcomma&"'"&escape_string(rs("pID"))&"'"
					addcomma=","
					rs.MoveNext
					Count=Count + 1
				loop
				if haveerrprods then print "<tr><td width=""100%"" colspan=""6""><br /><strong>"&redasterix&"</strong>"&yySeePr&"</td></tr>"
				if iNumOfPages > 1 then print "<tr><td colspan=""8"" align=""center"">" & writepagebar(CurPage,iNumOfPages,yyPrev,yyNext,pblink,FALSE) & "</td></tr>"
			else
				print "<tr><td width=""100%"" colspan=""8"" align=""center""><br />"&yyPrNone&"<br />&nbsp;</td></tr>"
			end if
			rs.close
		end if
	else
		if trim(detlinkspacechar)<>"" then
			sSQL="SELECT "&columnlist&" FROM products LEFT OUTER JOIN sections ON products.pSection=sections.sectionID WHERE pStaticURL LIKE '%"&IIfVr(mysqlserver,"\"&escape_string(detlinkspacechar),"["&escape_string(detlinkspacechar)&"]")&"%'" & IIfVs(usepnamefordetaillinks," OR pName LIKE '%"&IIfVr(mysqlserver,"\"&escape_string(detlinkspacechar),"["&escape_string(detlinkspacechar)&"]")&"%'")
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				print "<tr><td colspan=""8"" style=""color:#FF0000"">You have the detlinkspacechar parameter set as &quot;" & detlinkspacechar & "&quot; but have products where the Static URL " & IIfVs(usepnamefordetaillinks,"or Product Name ") & "uses this character and these will not display properly. Consider removing the detlinkspacechar parameter, or replacing it with a space in the Static URL for these products.</td></tr>"
				displayheaderrow()
				do while NOT rs.EOF
					displayprodrow(rs)
					rs.MoveNext
				loop
			end if
			rs.close
		end if
		numitems=0
		sSQL="SELECT COUNT(*) as totcount FROM products"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			numitems=rs("totcount")
		end if
		rs.close
		print "<tr><td colspan=""8""><div class=""itemsdefine"">You have " & numitems & " products defined.</div></td></tr>"
	end if %>
			  <tr>
				<td align="center" style="white-space:nowrap"><% if resultcounter>0 AND ((pract<>"" AND pract<>"none" AND pract<>"vis" AND pract<>"vil" AND pract<>"vig") OR pid<>"" OR rid<>"") then print "<input type=""hidden"" name=""resultcounter"" id=""resultcounter"" value="""&resultcounter&""" /><input type=""button"" value="""&yyUpdate&""" onclick=""quickupdate()"" /> <input type=""reset"" value="""&yyReset&""" />" else print "&nbsp;"%></td>
                <td width="100%" colspan="7" align="center"><br /><a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;<br /></td>
			  </tr>
            </table>
		</form>
<script>
var pa=[];
<%	if qetype="section" then
		print " var pq=[],ps=[" & vbCrLf
		sSQL="SELECT sectionID,sectionWorkingName FROM sections WHERE rootSection=1 ORDER BY sectionWorkingName"
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			print "[" & rs("sectionID") & ",'" & jsspecials(rs("sectionWorkingName")) & "']" & vbCrLf
			rs.movenext
			if NOT rs.EOF then print ","
		loop
		rs.close
		print "];" & vbCrLf
%>
	for(var pidind in ps){
		pq[ps[pidind][0]]=ps[pidind][1];
	}
	function popsection(tmenu){
		var foundthis=false;
		tmenu.onmouseover=null;
		var menucursel=parseInt(tmenu[tmenu.selectedIndex].value);
		for(var idind=0;idind<ps.length;idind++){
			var y=document.createElement('option');
			y.text=ps[idind][1];
			y.value=ps[idind][0];
			if(ps[idind][0]==menucursel)
				foundthis=true;
			else if(!foundthis){
				var sel=tmenu.options[0];
				tmenu.add(y, 0+idind);
			}else{
				try{ tmenu.add(y, null);} // FF etc
				catch(ex){ tmenu.add(y);} // IE
			}
		}
	}
	function createsection(pid,sid){
		var optionsMU='';
		return('<select size="1" id="sec'+pid+'" style="width:165px" onmouseover="popsection(this)" onchange="this.name=\'pra_'+patch_pid(pidind)+'\'"><option value="'+sid+'">'+(pq[sid]?pq[sid]:'**SECTION DELETED**')+'</option></select>');
	}
<%	elseif qetype="image" then %>
function imageisvisible(img){
    var rect=img.getBoundingClientRect();
    return(rect.top<=(window.innerHeight||document.documentElement.clientHeight));
}
function checkvisibleimages(){
	var lzimgs=document.getElementById('prodstable').getElementsByClassName('lazyload');
	var tarray=[];
	if(lzimgs.length==0){
		removeEventListener('scroll',setcheckvisibletimeout);
	}else{
		for(var lzi=0; lzi<lzimgs.length; lzi++){
			var telem=lzimgs[lzi];
			if(imageisvisible(telem)){
				telem.src=telem.getAttribute('data-src');
				tarray.push(telem.id);
			}
		}
		for(var lzi=0; lzi<tarray.length; lzi++){
			document.getElementById(tarray[lzi]).className='lazydoneload';
		}
	}
}
var checkvisibletimeout='';
function setcheckvisibletimeout(){
	if(checkvisibletimeout!='') clearTimeout(checkvisibletimeout);
	checkvisibletimeout=setTimeout(checkvisibleimages,300)
}
addEventListener('scroll',setcheckvisibletimeout);
addEventListener('load',checkvisibleimages);
<%	end if %>
<%=jscript%>
	function patch_pid(pid){
		document.getElementById('pid'+pid).name='pid'+pid;
		document.getElementById('pid'+pid).value=pa[pid][1];
		return pid;
	}
	for(var pidind in pa){
		var ttr=document.getElementById('tr'+pidind);
		ttr.cells[0].className='minicell';
		ttr.cells[3].style.textAlign=ttr.cells[4].style.textAlign=ttr.cells[5].style.textAlign=ttr.cells[6].style.textAlign='center';
		ttr.cells[1].innerHTML='<input type="hidden" id="pid'+pidind+'" value="" />'+pa[pidind][1];
<%	if pid<>"" then %>
		if(pa[pidind][2]!='0'){
			ttr.cells[5].innerHTML='<input type="text" name="pqa'+pidind+'" value="'+pa[pidind][2]+'" size="3" />';
		}
<%	elseif rid="" then %>
		ttr.cells[7].style.textAlign='center';
		ttr.cells[7].style.whiteSpace='nowrap';
		ttr.cells[4].innerHTML='<input type="button" id="rel'+pa[pidind][1]+'" value="<%=jsescape(htmlspecials("Rel"))%>" onclick="rel(\''+pa[pidind][1]+'\',\'related\')" title="<%=jsescape(yyRelate)%>" style="width:40px" />';
		ttr.cells[5].innerHTML='<input type="button" id="pak'+pa[pidind][1]+'" value="<%=jsescape(htmlspecials("Pak"))%>" onclick="rel(\''+pa[pidind][1]+'\',\'package\')" title="<%="Package"%>" style="width:40px" />';
		ttr.cells[6].innerHTML='<input type="button" value="Alt" onclick="al(\''+pa[pidind][1]+'\')" title="<%=jsescape("ALT IDs")%>" style="width:40px" />';
		ttr.cells[7].innerHTML='<input type="button" value="M" style="width:30px;margin-right:4px" onclick="mr(\''+pa[pidind][1]+'\')" title="<%=jsescape(htmlspecials(yyModify))%>" />' +
			'<input type="button" value="C" style="width:30px;margin-right:4px" onclick="cr(\''+pa[pidind][1]+'\')" title="<%=jsescape(htmlspecials(yyClone))%>" />' +
			'<input type="button" value="X" style="width:30px" onclick="dr(\''+pa[pidind][1]+'\')" title="<%=jsescape(htmlspecials(yyDelete))%>" />';
		ttr.cells[0].innerHTML=
<%		if qetype="text" then %>
	pa[pidind][0]===false?'-':'<input type="text" id="chkbx'+pidind+'" size="<%=qesize%>" onchange="this.name=\'pra_'+patch_pid(pidind)+'\'" value="'+pa[pidind][0].replace('"','&quot;')+'" tabindex="'+(pidind+1)+'" />';
<%		elseif qetype="delbox" then %>
	'<input type="checkbox" id="chkbx'+pidind+'" onchange="this.name=\'pra_'+patch_pid(pidind)+'\'" value="del" tabindex="'+(pidind+1)+'" />';
<%		elseif qetype="checkbox" then %>
	'<input type="hidden" id="pra_'+patch_pid(pidind)+'" value="1" /><input type="checkbox" id="chkbx'+pidind+'" onchange="this.name=\'prb_'+patch_pid(pidind)+'\';document.getElementById(\'pra_'+patch_pid(pidind)+'\').name=\'pra_'+patch_pid(pidind)+'\'" value="1" '+(pa[pidind][0]==1?'checked="checked" ':'')+'tabindex="'+(pidind+1)+'" />';
<%		elseif qetype="section" then %>
	createsection(pidind,pa[pidind][0]);
<%		elseif qetype="image" then %>
	(pa[pidind][0]==''?'-':'<img class="lazyload" id="lazyimg'+pidind+'" src="adminimages/imageload.png" data-src="'+pa[pidind][0]+'" style="max-width:80px;cursor:pointer" alt="" onclick="mr(\''+pa[pidind][1]+'\')" />');
<%		else %>
	'&nbsp;';
<%		end if %>
	ttr.cells[3].innerHTML='<input type="button" '+(pa[pidind][2]?' class="ectset"':'')+' value="<%=jsescape(htmlspecials(yyAssign))%>" onclick="dsc(\''+pa[pidind][1]+'\')" />';
<%	end if %>
	}
<%
	if pidlist<>"" AND pid="" AND rid="" then
		print vbCrLf & "function setcl(tid){if(document.getElementById('rel'+tid))document.getElementById('rel'+tid).classList.add('ectset');}" & vbCrLf
		sSQL="SELECT DISTINCT rpProdId FROM relatedprods WHERE rpProdId IN ("&pidlist&")"
		if relatedproductsbothways=TRUE then sSQL=sSQL & "UNION SELECT DISTINCT rpRelProdId FROM relatedprods WHERE rpRelProdId IN ("&pidlist&")"
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			print "setcl('"&rs("rpProdId")&"');" & vbCrLf
			rs.MoveNext
		loop
		rs.close
	end if
%>
</script>
<%
end if
set prregexp=nothing
cnn.Close
set rs=nothing
set rs2=nothing
set rs3=nothing
set cnn=nothing
%>
