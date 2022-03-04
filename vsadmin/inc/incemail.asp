<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if xxNoWtIn="" then xxNoWtIn=" (Shipping Insurance Declined)"
if digidownloadsecret="" then digidownloadsecret="this is some secret text"
emailheader=""
receiptheader=""
recpt=""
hasimageupload=FALSE
sub order_success(sorderid,sEmail,sendstoreemail)
	call do_order_success(sorderid,sEmail,sendstoreemail,TRUE,TRUE,TRUE,TRUE)
end sub
function getemailrecpt(oid)
	recpt=""
	sSQL="SELECT ordID,ordName,ordLastName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordShipName,ordShipLastName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,ordShipPhone,ordPayProvider,ordAuthNumber,ordTotal,ordDate,ordStateTax,ordCountryTax,ordHSTTax,ordHandling,ordShipping,ordAffiliate,ordShipType,ordShipCarrier,ordDiscount,ordDiscountText,ordComLoc,ordExtra1,ordExtra2,ordShipExtra1,ordShipExtra2,ordCheckoutExtra1,ordCheckoutExtra2,ordSessionID,payProvID,ordAddInfo FROM orders INNER JOIN payprovider ON payprovider.payProvID=orders.ordPayProvider WHERE ordAuthNumber<>'' AND ordID="&replace(oid,"'","")
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		ordShipCarrier=rs("ordShipCarrier")
		recpt=recpt & "<table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""3"">"
		recpt=recpt & "  <tr>"
		recpt=recpt & "	<td valign=""top"" colspan=""5"">"
		recpt=recpt & "	  <table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""3"">"
		if trim(rs("ordShipAddress")&"")<>"" then hasshipaddress=TRUE else hasshipaddress=FALSE
		recpt=recpt & "		<tr>"
		recpt=recpt & "		  <td valign=""top""><strong>"&xxOrdId&":</strong></td>"
		recpt=recpt & "		  <td valign=""top"">"&rs("ordID")&"</td>"
		recpt=recpt & "		  <td>&nbsp;</td><td>&nbsp;</td>"
		recpt=recpt & "		</tr>"&vbCrLf
		if trim(extraorderfield1)<>"" then
			recpt=recpt & "	<tr>"
			recpt=recpt & "	  <td valign=""top""><strong>"&extraorderfield1&":</strong></td>"
			recpt=recpt & "	  <td valign=""top"">"&rs("ordExtra1")&"</td>"
			if hasshipaddress then
				recpt=recpt & " <td valign=""top""><strong>"&extraorderfield1&":</strong></td>"
				recpt=recpt & " <td valign=""top"">"&rs("ordShipExtra1")&"</td>"
			end if
			recpt=recpt & "	</tr>"&vbCrLf
		end if
		recpt=recpt & "		<tr>"
		recpt=recpt & "		  <td valign=""top"" width=""20%""><strong>"&xxBilAdd&":</strong></td>"
		recpt=recpt & "		  <td valign=""top"">"
		recpt=recpt & trim(rs("ordName")&" "&rs("ordLastName")) & "<br />"
		recpt=recpt & rs("ordAddress") & "<br />"
		if trim(rs("ordAddress2")&"")<>"" then recpt=recpt & rs("ordAddress2") & "<br />"
		recpt=recpt & rs("ordCity") & ", " & rs("ordState") & "<br />"
		recpt=recpt & rs("ordZip") & "<br />"
		recpt=recpt & rs("ordCountry") & "<br />"
		recpt=recpt & "		  </td>"
		if hasshipaddress then
			recpt=recpt & "	  <td valign=""top"" width=""20%""><strong>"&xxShpAdd&":</strong></td>"
			recpt=recpt & "	  <td valign=""top"">"
			recpt=recpt & trim(rs("ordShipName")&" "&rs("ordShipLastName")) & "<br />"
			recpt=recpt & rs("ordShipAddress") & "<br />"
			if trim(rs("ordShipAddress2")&"")<>"" then recpt=recpt & rs("ordShipAddress2") & "<br />"
			recpt=recpt & rs("ordShipCity") & ", " & rs("ordShipState") & "<br />"
			recpt=recpt & rs("ordShipZip") & "<br />"
			recpt=recpt & rs("ordShipCountry") & "<br />"
			recpt=recpt & "	  </td>"
		end if
		recpt=recpt & "		</tr>"&vbCrLf
		if trim(extraorderfield2)<>"" then
			recpt=recpt & "	<tr>"
			recpt=recpt & "	  <td valign=""top""><strong>"&extraorderfield2&":</strong></td>"
			recpt=recpt & "	  <td valign=""top"">"&rs("ordExtra2")&"</td>"
			if hasshipaddress then
				recpt=recpt & "  <td valign=""top""><strong>"&extraorderfield2&":</strong></td>"
				recpt=recpt & "  <td valign=""top"">"&rs("ordShipExtra2")&"</td>"
			end if
			recpt=recpt & "	</tr>"&vbCrLf
		end if
		recpt=recpt & "		<tr>"
		recpt=recpt & "		  <td valign=""top""><strong>"&xxPhone&":</strong></td>"
		recpt=recpt & "		  <td valign=""top"">"&rs("ordPhone")&"</td>"
		if hasshipaddress then
			recpt=recpt & "	  <td valign=""top""><strong>"&xxPhone&":</strong></td>"
			recpt=recpt & "	  <td valign=""top"">"&rs("ordShipPhone")&"</td>"
		end if
		recpt=recpt & "		</tr>"&vbCrLf
		recpt=recpt & "		<tr>"
		recpt=recpt & "		  <td valign=""top""><strong>"&xxEmail&":</strong></td>"
		recpt=recpt & "		  <td valign=""top"">"&rs("ordEmail")&"</td>"
		ordShipType=rs("ordShipType")
		if ordShipType<>"" then
			shiptext="<td valign=""top""><strong>" & xxShpMet & ":</strong></td><td valign=""top"">" & ordShipType
			if willpickuptext<>ordShipType then
				if (rs("ordComLoc") AND 2)=2 then shiptext=shiptext & xxWtIns else if forceinsuranceselection then shiptext=shiptext & xxNoWtIn
			end if
			shiptext=shiptext & "<br />"
			if (rs("ordComLoc") AND 1)=1 then shiptext=shiptext & xxCerCLo & "<br />"
			if (rs("ordComLoc") AND 4)=4 then shiptext=shiptext & xxSatDeR & "<br />"
			shiptext=shiptext & "</td>"
		else
			shiptext=""
		end if
		if hasshipaddress then
			if shiptext="" then recpt=recpt & "<td>&nbsp;</td><td>&nbsp;</td>" else recpt=recpt & shiptext
		end if
		recpt=recpt & "		</tr>"&vbCrLf
		if NOT hasshipaddress then
			if shiptext<>"" then recpt=recpt & "<tr>" & shiptext & "</tr>"
		end if
		if trim(extracheckoutfield1)<>"" AND trim(rs("ordCheckoutExtra1")&"")<>"" then
			recpt=recpt & "	<tr>"
			recpt=recpt & "	  <td valign=""top""><strong>"&extracheckoutfield1&":</strong></td>"
			recpt=recpt & "	  <td valign=""top"""&IIfVr(hasshipaddress," colspan=""3""","")&">"&rs("ordCheckoutExtra1")&"</td>"
			recpt=recpt & "	</tr>"&vbCrLf
		end if
		if trim(extracheckoutfield2)<>"" AND trim(rs("ordCheckoutExtra2")&"")<>"" then
			recpt=recpt & "	<tr>"
			recpt=recpt & "	  <td valign=""top""><strong>"&extracheckoutfield2&":</strong></td>"
			recpt=recpt & "	  <td valign=""top"""&IIfVr(hasshipaddress," colspan=""3""","")&">"&rs("ordCheckoutExtra2")&"</td>"
			recpt=recpt & "	</tr>"&vbCrLf
		end if
		if loyaltypoints<>"" then recpt=recpt & "<!--%loyaltypointplaceholder%-->"
		ordAddInfo=trim(rs("ordAddInfo"))
		if ordAddInfo<>"" then
			recpt=recpt & "	<tr>"
			recpt=recpt & "	  <td valign=""top""><strong>"&xxAddInf&":</strong></td>"
			recpt=recpt & "	  <td valign=""top"" colspan=""3"">"&replace(replace(ordAddInfo,vbLf,"<br />"),vbCr,"")&"</td>"
			recpt=recpt & "	</tr>"&vbCrLf
		end if
		if digidownloads=TRUE then recpt=recpt & "<!--%digidownloadplaceholder%-->"
		recpt=recpt & "	  </table>"
		recpt=recpt & "	</td>"
		recpt=recpt & "  </tr>"&vbCrLf
		if digidownloads=TRUE then recpt=recpt & "<!--%digidownloaditems%-->"
		recpt=recpt & "  <tr><td align=""center"" colspan=""5""><hr class=""receipthr"" width=""80%""></td></tr>"
		recpt=recpt & "  <tr>"
		recpt=recpt & "	<td class=""receiptheading"" width=""15%"" height=""25"" align=""left""><strong>"&xxCODets&"</strong></td>"
		recpt=recpt & "	<td class=""receiptheading"" width=""33%"" height=""25"" align=""left""><strong>"&xxCOName&"</strong></td>"
		if NOT nopriceanywhere then recpt=recpt & "	<td class=""receiptheading"" width=""14%"" height=""25"" align=""right""><strong>"&xxCOUPri&"</strong></td>"
		recpt=recpt & "	<td class=""receiptheading"" width=""14%"" height=""25"" align=""right""><strong>"&xxQuant&"</strong></td>"
		if NOT nopriceanywhere then recpt=recpt & "	<td class=""receiptheading"" width=""14%"" height=""25"" align=""right""><strong>"&xxTotal&"</strong></td>"
		recpt=recpt & "  </tr>"&vbCrLf
	end if
	rs.close
	sSQL="SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity,cartGiftWrap FROM cart WHERE cartOrderID="&replace(oid,"'","")&" ORDER BY cartID"
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		theoptions=""
		theoptionspricediff=0
		isoutofstock=FALSE
		sSQL="SELECT coID,coOptGroup,coCartOption,coPriceDiff,coWeightDiff,optAltImage FROM cartoptions LEFT JOIN options ON cartoptions.coOptID=options.optID WHERE coCartID="&rs("cartID")& " ORDER BY coID"
		rs2.Open sSQL,cnn,0,1
		do while NOT rs2.EOF
			theoptionspricediff=theoptionspricediff + rs2("coPriceDiff")
			theoptions=theoptions & "<tr>"
			theoptions=theoptions & "<td class=""receiptoption"" align=""right""><span style=""font-size:0.82em""><strong>"&rs2("coOptGroup")&":</strong></span></td>"
			theoptions=theoptions & "<td class=""receiptoption"" align=""left""><span style=""font-size:0.82em"">" & IIfVs(instr(rs2("coCartOption")&"",vbLf)=0,"&nbsp;- ") & replace(replace(strip_tags2(rs2("coCartOption")&""),vbLf,"<br />"),vbCr,"") & "</span></td>"
			if NOT nopriceanywhere then theoptions=theoptions & "<td class=""receiptoption"" align=""right""><span style=""font-size:0.82em"">" & IIfVr(rs2("coPriceDiff")=0 OR hideoptpricediffs=true,"- ", FormatEuroCurrency(rs2("coPriceDiff"))) & "</span></td>"
			theoptions=theoptions & "<td class=""receiptoption"">&nbsp;</td>"
			if NOT nopriceanywhere then theoptions=theoptions & "<td class=""receiptoption"" align=""right""><span style=""font-size:0.82em"">" & IIfVr(rs2("coPriceDiff")=0 OR hideoptpricediffs=true,"- ", FormatEuroCurrency(rs2("coPriceDiff")*rs("cartQuantity"))) & "</span></td>"
			theoptions=theoptions & "</tr>"&vbCrLf
			rs2.MoveNext
		loop
		rs2.Close
		recpt=recpt & "  <tr>"
		recpt=recpt & "	<td class=""cobhl receipthl"" align=""left"" height=""25""><strong>" & rs("cartProdID") & "</strong></td>"
		recpt=recpt & "	<td class=""cobhl receipthl"" align=""left"">"&rs("cartProdName")
		if rs("cartGiftWrap")<>0 then recpt=recpt & "<div class=""giftwrap"">" & xxGWrSel & "</div>"
		sSQL="SELECT productpackages.pID,quantity,pName,quantity FROM productpackages INNER JOIN products on productpackages.pID=products.pID WHERE packageID='"&escape_string(rs("cartProdID"))&"'"
		rs2.open sSQL,cnn,0,1
		if NOT rs2.EOF then
			recpt=recpt & "<table class=""receiptpackage"" style=""font-size:10px"">"
			do while NOT rs2.EOF
				recpt=recpt & "<tr><td> &gt; " & rs2("pID") & ": </td><td>" & rs2("pName") & "</td><td>" & rs2("quantity") & "</td></tr>"
				rs2.movenext
			loop
			recpt=recpt & "</table>"
		end if
		rs2.close
		if NOT nopriceanywhere then recpt=recpt & "</td><td class=""cobhl receipthl"" align=""right"">"&IIfVr(hideoptpricediffs=true,FormatEuroCurrency(rs("cartProdPrice")+theoptionspricediff),FormatEuroCurrency(rs("cartProdPrice")))&"</td>"
		recpt=recpt & "	<td class=""cobhl receipthl"" align=""right"">"&rs("cartQuantity")&"</td>"
		if NOT nopriceanywhere then recpt=recpt & "	<td class=""cobhl receipthl"" align=""right"">"&IIfVr(hideoptpricediffs=true,FormatEuroCurrency((rs("cartProdPrice")+theoptionspricediff)*rs("cartQuantity")),FormatEuroCurrency(rs("cartProdPrice")*rs("cartQuantity")))&"</td>"
		recpt=recpt & "  </tr>"&vbCrLf
		recpt=recpt & theoptions
		rs.MoveNext
	loop
	rs.close
	if NOT nopriceanywhere then
		recpt=recpt & "	  <tr>"
		recpt=recpt & "		<td colspan=""3"">&nbsp;</td>"
		recpt=recpt & "		<td align=""right""><strong>"&xxSubTot&":</strong></td>"
		recpt=recpt & "		<td align=""right"">"&FormatEuroCurrency(ordTotal)&"</td>"
		recpt=recpt & "	  </tr>"&vbCrLf
		if ordDiscount>0 then
			recpt=recpt & " <tr>"
			recpt=recpt & "	<td colspan=""3"">&nbsp;</td>"
			recpt=recpt & "	<td align=""right""><strong>"&xxDscnts&":</strong></td>"
			recpt=recpt & "	<td class=""recptdiscount"" align=""right"" style=""color:#FF0000"">"&FormatEuroCurrency(ordDiscount)&"</td>"
			recpt=recpt & " </tr>"&vbCrLf
		end if
		if ordShipCarrier=0 AND ((ordShipping+ordHandling)=0) then
			' Do nothing
		elseif combineshippinghandling=TRUE then
			recpt=recpt & " <tr>"
			recpt=recpt & "	<td colspan=""2"">&nbsp;</td>"
			recpt=recpt & "	<td colspan=""2"" align=""right""><strong>"&xxShipHa&":</strong></td>"
			recpt=recpt & "	<td align=""right"">"& IIfVr((ordShipping+ordHandling)=0,"<p align=""center""><span style=""color:#FF0000;font-weight:bold"">" & xxFree & "</span></p>",FormatEuroCurrency(ordShipping+ordHandling))&"</td>"
			recpt=recpt & " </tr>"&vbCrLf
		else
			if ordShipping>0 then
				recpt=recpt & "  <tr>"
				recpt=recpt & "	<td colspan=""3"">&nbsp;</td>"
				recpt=recpt & "	<td align=""right""><strong>"&xxShippg&":</strong></td>"
				recpt=recpt & "	<td align=""right"">"&FormatEuroCurrency(ordShipping)&"</td>"
				recpt=recpt & "  </tr>"&vbCrLf
			end if
			if ordHandling>0 then
				recpt=recpt & "	  <tr>"
				recpt=recpt & "		<td colspan=""3"">&nbsp;</td>"
				recpt=recpt & "		<td align=""right""><strong>"&xxHndlg&":</strong></td>"
				recpt=recpt & "		<td align=""right"">"&FormatEuroCurrency(ordHandling)&"</td>"
				recpt=recpt & "	  </tr>"&vbCrLf
			end if
		end if
		if ordStateTax>0 then
			recpt=recpt & "  <tr>"
			recpt=recpt & "	<td colspan=""3"">&nbsp;</td>"
			recpt=recpt & "	<td align=""right""><strong>"&xxStaTax&":</strong></td>"
			recpt=recpt & "	<td align=""right"">"&FormatEuroCurrency(ordStateTax)&"</td>"
			recpt=recpt & "  </tr>"&vbCrLf
		end if
		if ordHSTTax>0 then
			recpt=recpt & "  <tr>"
			recpt=recpt & "	<td colspan=""3"">&nbsp;</td>"
			recpt=recpt & "	<td align=""right""><strong>"&xxHST&":</strong></td>"
			recpt=recpt & "	<td align=""right"">"&FormatEuroCurrency(ordHSTTax)&"</td>"
			recpt=recpt & "  </tr>"&vbCrLf
		end if
		if ordCountryTax>0 OR alwaysdisplaycountrytax then
			recpt=recpt & "  <tr>"
			recpt=recpt & "	<td colspan=""3"">&nbsp;</td>"
			recpt=recpt & "	<td align=""right""><strong>"&xxCntTax&":</strong></td>"
			recpt=recpt & "	<td align=""right"">"&FormatEuroCurrency(ordCountryTax)&"</td>"
			recpt=recpt & "  </tr>"&vbCrLf
		end if
		recpt=recpt & "	  <tr>"
		recpt=recpt & "		<td colspan=""3"">&nbsp;</td>"
		recpt=recpt & "		<td class=""cobhl receipthl"" align=""right""><strong>"&xxGndTot&":</strong></td>"
		recpt=recpt & "		<td class=""cobhl receipthl"" align=""right"">"&FormatEuroCurrency(ordGrandTotal)&"</td>"
		recpt=recpt & "	  </tr>"
		recpt=recpt & "	  <tr>"
		recpt=recpt & "		<td align=""center"" colspan=""5"">&nbsp;</td>"
		recpt=recpt & "	  </tr>"
	end if
	recpt=recpt & "	</table>"&vbCrLf
	getemailrecpt=recpt
end function
function getrecpt(oid)
	recpt=""
	sSQL="SELECT ordID,ordName,ordLastName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordShipName,ordShipLastName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,ordShipPhone,ordPayProvider,ordAuthNumber,ordTotal,ordDate,ordStateTax,ordCountryTax,ordHSTTax,ordHandling,ordShipping,ordAffiliate,ordShipType,ordShipCarrier,ordDiscount,ordDiscountText,ordComLoc,ordExtra1,ordExtra2,ordShipExtra1,ordShipExtra2,ordCheckoutExtra1,ordCheckoutExtra2,ordSessionID,payProvID,ordAddInfo FROM orders INNER JOIN payprovider ON payprovider.payProvID=orders.ordPayProvider WHERE ordAuthNumber<>'' AND ordID="&replace(oid,"'","")
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		ordShipCarrier=rs("ordShipCarrier")
		hasshipaddress=(trim(rs("ordShipAddress")&"")<>"")

		' START : addressess {
		recpt=recpt & "<div class=""receiptaddresses"">"
		' START : receiptbillingaddress {
		recpt=recpt & "		<div class=""receiptaddress receiptbillingaddress"">"
		recpt=recpt & "		  <div class=""receiptsectionhead"">"&xxBilAdd&"</div>"
		if trim(extraorderfield1)<>"" AND trim(rs("ordExtra1"))<>"" then
			recpt=recpt & "		  <div class=""receiptcontainer rcontextra1""><div class=""receiptleft"">"&extraorderfield1&"</div><div class=""receiptright"">" & rs("ordExtra1") & "</div></div>"
		end if
		recpt=recpt & "		  <div class=""receiptcontainer rcontname""><div class=""receiptleft"">" & xxName & "</div><div class=""receiptright"">" & trim(rs("ordName")&" "&rs("ordLastName")) & "</div></div>"
		recpt=recpt & "		  <div class=""receiptcontainer rcontaddress""><div class=""receiptleft"">" & xxAddress & "</div><div class=""receiptright"">"
			recpt=recpt & "<div>" & rs("ordAddress") & "</div>"
			if trim(rs("ordAddress2")&"")<>"" then recpt=recpt & "<div>" & rs("ordAddress2") & "</div>"
			recpt=recpt & "<div>" & rs("ordCity") & ", " & rs("ordState") & "</div>"
			recpt=recpt & "<div>" & rs("ordZip") & "</div>"
			recpt=recpt & "<div>" & rs("ordCountry") & "</div>"
		recpt=recpt & "		  </div></div>"
		recpt=recpt & "		  <div class=""receiptcontainer rcontphone""><div class=""receiptleft"">" & xxPhone & "</div><div class=""receiptright"">" & rs("ordPhone") & "</div></div>"
		if trim(extraorderfield2)<>"" AND trim(rs("ordExtra2"))<>"" then
			recpt=recpt & "		  <div class=""receiptcontainer rcontextra2""><div class=""receiptleft"">"&extraorderfield2&"</div><div class=""receiptright"">" & rs("ordExtra2") & "</div></div>"
		end if
		recpt=recpt & "</div>"
		' END : receiptbillingaddress }
		
		if hasshipaddress then
			' START : receiptshippingaddress {
			recpt=recpt & "<div class=""receiptaddress receiptshippingaddress"">"
			recpt=recpt & "		  <div class=""receiptsectionhead"">"&xxShpAdd&"</div>"
			if trim(extraorderfield1)<>"" AND trim(rs("ordShipExtra1"))<>"" then
				recpt=recpt & "		  <div class=""receiptcontainer rcontextra1""><div class=""receiptleft"">"&extraorderfield1&"</div><div class=""receiptright"">" & rs("ordShipExtra1") & "</div></div>"
			end if
			recpt=recpt & "		  <div class=""receiptcontainer rcontname""><div class=""receiptleft"">" & xxName & "</div><div class=""receiptright"">" & trim(rs("ordShipName")&" "&rs("ordShipLastName")) & "</div></div>"
			recpt=recpt & "		  <div class=""receiptcontainer rcontaddress""><div class=""receiptleft"">" & xxAddress & "</div><div class=""receiptright"">"
			recpt=recpt & "<div>" & rs("ordShipAddress") & "</div>"
			if trim(rs("ordShipAddress2")&"")<>"" then recpt=recpt & "<div>" & rs("ordShipAddress2") & "</div>"
			recpt=recpt & "<div>" & rs("ordShipCity") & ", " & rs("ordShipState") & "</div>"
			recpt=recpt & "<div>" & rs("ordShipZip") & "</div>"
			recpt=recpt & "<div>" & rs("ordShipCountry") & "</div>"
			recpt=recpt & "	  </div></div>"
			recpt=recpt & "		  <div class=""receiptcontainer rcontphone""><div class=""receiptleft"">" & xxPhone & "</div><div class=""receiptright"">" & rs("ordShipPhone") & "</div></div>"
			if trim(extraorderfield2)<>"" AND trim(rs("ordShipExtra2"))<>"" then
				recpt=recpt & "		  <div class=""receiptcontainer rcontextra2""><div class=""receiptleft"">"&extraorderfield2&"</div><div class=""receiptright"">" & rs("ordShipExtra2") & "</div></div>"
			end if
			recpt=recpt & "</div>"
			' END : receiptshippingaddress }
		end if
		recpt=recpt & "	</div>"&vbCrLf
		' END : addressess }

		' START : receiptextra {
		recpt=recpt & "		<div class=""receiptextra"">"
		recpt=recpt & "		  <div class=""receiptcontainer rcontorderid""><div class=""receiptleft"">" & xxOrdId & "</div><div class=""receiptright"">" & rs("ordID") & "</div></div>"
		recpt=recpt & "		  <div class=""receiptcontainer rcontemail""><div class=""receiptleft"">" & xxEmail & "</div><div class=""receiptright"">" & rs("ordEmail") & "</div></div>"

		ordShipType=rs("ordShipType")
		if ordShipType<>"" then
			shiptext="		  <div class=""receiptcontainer rcontshipmethod""><div class=""receiptleft"">" & xxShpMet & "</div><div class=""receiptright""><div>" & ordShipType & "</div>"
			if willpickuptext<>ordShipType then
				if (rs("ordComLoc") AND 2)=2 then shiptext=shiptext & "<div>" & xxWtIns & "</div>" else if forceinsuranceselection then shiptext=shiptext & "<div>" & xxNoWtIn & "</div>"
			end if
			if (rs("ordComLoc") AND 1)=1 then shiptext=shiptext & "<div>" & xxCerCLo & "</div>"
			if (rs("ordComLoc") AND 4)=4 then shiptext=shiptext & "<div>" & xxSatDeR & "</div>"
			shiptext=shiptext & "</div></div>" & vbCrLf
		else
			shiptext=""
		end if
		recpt=recpt & shiptext
		if trim(extracheckoutfield1)<>"" AND trim(rs("ordCheckoutExtra1")&"")<>"" then
			recpt=recpt & "		  <div class=""receiptcontainer rcontextracheckout1""><div class=""receiptleft"">" & extracheckoutfield1 & "</div><div class=""receiptright"">" & rs("ordCheckoutExtra1") & "</div></div>" & vbCrLf
		end if
		if trim(extracheckoutfield2)<>"" AND trim(rs("ordCheckoutExtra2")&"")<>"" then
			recpt=recpt & "		  <div class=""receiptcontainer rcontextracheckout2""><div class=""receiptleft"">" & extracheckoutfield2 & "</div><div class=""receiptright"">" & rs("ordCheckoutExtra2") & "</div></div>" & vbCrLf
		end if
		if loyaltypoints<>"" then recpt=recpt & "<!--%loyaltypointplaceholder%-->"
		ordAddInfo=trim(rs("ordAddInfo"))
		if ordAddInfo<>"" then
			recpt=recpt & "		  <div class=""receiptcontainer rcontaddinfo""><div class=""receiptleft"">" & xxAddInf & "</div><div class=""receiptright"">" & replace(replace(ordAddInfo,vbLf,"<br />"),vbCr,"") & "</div></div>" & vbCrLf
		end if
		if digidownloads=TRUE then recpt=recpt & "<!--%digidownloadplaceholder%-->"
		recpt=recpt & "	</div>"&vbCrLf
		' END : receiptextra }
		
		if digidownloads=TRUE then recpt=recpt & "<!--%digidownloaditems%-->"
		recpt=recpt & "<div class=""receiptlist"">" & _
					"	<div class=""receiptheadrow"">"
		recpt=recpt & "		<div class=""receiptprodid"">"&xxCODets&"</div>"
		recpt=recpt & "		<div class=""receiptprodname"">"&xxCOName&"</div>"
		if NOT nopriceanywhere then recpt=recpt & "	<div class=""receiptunitprice"">"&xxCOUPri&"</div>"
		recpt=recpt & "		<div class=""receiptquantity"">"&xxQuant&"</div>"
		if NOT nopriceanywhere then recpt=recpt & "	<div class=""receipttotal"">"&xxTotal&"</div>"
		recpt=recpt & "  </div>"&vbCrLf
	end if
	rs.close
	sSQL="SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity,cartGiftWrap,pUpload FROM cart LEFT JOIN products ON cart.cartProdID=products.pID WHERE cartOrderID="&replace(oid,"'","")&" ORDER BY cartID"
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		theoptions=""
		theoptionspricediff=0
		isoutofstock=FALSE
		sSQL="SELECT coID,coOptGroup,coCartOption,coPriceDiff,coWeightDiff,optAltImage FROM cartoptions LEFT JOIN options ON cartoptions.coOptID=options.optID WHERE coCartID="&rs("cartID")& " ORDER BY coID"
		rs2.Open sSQL,cnn,0,1
		do while NOT rs2.EOF
			theoptionspricediff=theoptionspricediff + rs2("coPriceDiff")
			theoptions=theoptions & "<div class=""receiptlistrow receiptoptionrow"">"
			theoptions=theoptions & "<div class=""receiptoption receiptoptgroup"">"&rs2("coOptGroup")&"</div>"
			theoptions=theoptions & "<div class=""receiptoption receiptoptname"">" & IIfVs(instr(rs2("coCartOption")&"",vbLf)=0,"&nbsp;- ") & replace(replace(strip_tags2(rs2("coCartOption")&""),vbLf,"<br />"),vbCr,"") & "</div>"
			if NOT nopriceanywhere then theoptions=theoptions & "<div class=""receiptoption receiptoptunitprice"">" & IIfVr(rs2("coPriceDiff")=0 OR hideoptpricediffs,"- ", FormatEuroCurrency(rs2("coPriceDiff"))) & "</div>"
			theoptions=theoptions & "<div class=""receiptoption receiptoptquantity"">&nbsp;</div>"
			if NOT nopriceanywhere then theoptions=theoptions & "<div class=""receiptoption receiptopttotal"">" & IIfVr(rs2("coPriceDiff")=0 OR hideoptpricediffs,"- ", FormatEuroCurrency(rs2("coPriceDiff")*rs("cartQuantity"))) & "</div>"
			theoptions=theoptions & "</div>"&vbCrLf
			rs2.MoveNext
		loop
		rs2.Close
		recpt=recpt & "  <div class=""receiptlistrow receiptitemrow"">"
		recpt=recpt & "		<div class=""receiptprodid"">"
		thankspageimage=""
		if imgonthankspage then
			sSQL="SELECT imageSrc FROM productimages WHERE imageProduct='" & escape_string(rs("cartProdID")) & "' AND (imageType=0 OR imageType=1) ORDER BY imageType,imageNumber"
			rs2.open sSQL,cnn,0,1
			thankspageimage=defaultthankspageimage
			if NOT rs2.EOF then thankspageimage=rs2("imageSrc")
			rs2.close
			if thankspageimage<>"" then recpt=recpt & "<div class=""thankspageimg""><img class=""thankspageimg"" src=""" & thankspageimage & """ alt="""" /><div>"
		end if
		recpt=recpt & rs("cartProdID") & "</div>"
		if thankspageimage<>"" then recpt=recpt & "</div></div>"
		recpt=recpt & "		<div class=""receiptprodname"">"&rs("cartProdName") & IIfVs(rs("cartGiftWrap")<>0,"<div class=""giftwrap"">" & xxGWrSel & "</div>")
			sSQL="SELECT productpackages.pID,quantity,pName,quantity FROM productpackages INNER JOIN products on productpackages.pID=products.pID WHERE packageID='"&escape_string(rs("cartProdID"))&"'"
			rs2.open sSQL,cnn,0,1
			if NOT rs2.EOF then
				recpt=recpt & "<div class=""receiptpackage"">"
				do while NOT rs2.EOF
					recpt=recpt & "<div class=""receiptpackagerow""><div class=""ectleft"">" & rs2("pID") & "</div><div class=""ectright"">" & rs2("pName") & "</div><div>" & rs2("quantity") & "</div></div>"
					rs2.movenext
				loop
				recpt=recpt & "</div>"
			end if
			rs2.close
			if rs("pUpload") then hasimageupload=TRUE
		recpt=recpt & "</div>"
		if NOT nopriceanywhere then recpt=recpt & "<div class=""receiptunitprice"">"&IIfVr(hideoptpricediffs,FormatEuroCurrency(rs("cartProdPrice")+theoptionspricediff),FormatEuroCurrency(rs("cartProdPrice")))&"</div>"
		recpt=recpt & "		<div class=""receiptquantity"">"&rs("cartQuantity")&"</div>"
		if NOT nopriceanywhere then recpt=recpt & "<div class=""receipttotal"">"&IIfVr(hideoptpricediffs,FormatEuroCurrency((rs("cartProdPrice")+theoptionspricediff)*rs("cartQuantity")),FormatEuroCurrency(rs("cartProdPrice")*rs("cartQuantity")))&"</div>"
		recpt=recpt & "  </div>"&vbCrLf
		recpt=recpt & theoptions
		rs.MoveNext
	loop
	rs.close
	recpt=recpt & "  </div>"&vbCrLf
	if hasimageupload then recpt=recpt & "<div id=""imageuploadbutton"" class=""imageuploadbutton no-print""" & IIfVs(NOT imageuploadbybutton," style=""display:none""") & "><button class=""ectbutton"" onclick=""document.getElementById('imageuploadbutton').style.display='none';document.getElementById('imageuploadopdiv').style.display='';return false"">" & xxImgUpl & "</button></div>"
	if NOT nopriceanywhere then
		recpt=recpt & "<div class=""receipttotalscolumn"">" & _
						"<div class=""receipttotalstable"">"
		recpt=recpt & "<div class=""receipttotalsrow rectotsubtotal""><div class=""ectleft"">"&xxSubTot&":</div><div class=""ectright"">"&FormatEuroCurrency(ordTotal)&"</div></div>"&vbCrLf
		if ordDiscount>0 then
			recpt=recpt & "<div class=""receipttotalsrow rectotdiscounts""><div class=""ectleft"">"&xxDscnts&":</div><div class=""ectright"">"&FormatEuroCurrency(ordDiscount)&"</div></div>"&vbCrLf
		end if
		if ordShipCarrier=0 AND ((ordShipping+ordHandling)=0) then
			' Do nothing
		elseif combineshippinghandling then
			recpt=recpt & "<div class=""receipttotalsrow rectotshiphand""><div class=""ectleft"">"&xxShipHa&":</div><div class=""ectright"">"& IIfVr((ordShipping+ordHandling)=0,"<p align=""center""><span style=""color:#FF0000;font-weight:bold"">" & xxFree & "</span></p>",FormatEuroCurrency(ordShipping+ordHandling))&"</div></div>"&vbCrLf
		else
			if ordShipping>0 then
				recpt=recpt & "<div class=""receipttotalsrow rectotshipping""><div class=""ectleft"">"&xxShippg&":</div><div class=""ectright"">"&FormatEuroCurrency(ordShipping)&"</div></div>"&vbCrLf
			end if
			if ordHandling>0 then
				recpt=recpt & "<div class=""receipttotalsrow rectothandling""><div class=""ectleft"">"&xxHndlg&":</div><div class=""ectright"">"&FormatEuroCurrency(ordHandling)&"</div></div>"&vbCrLf
			end if
		end if
		if ordStateTax>0 then
			recpt=recpt & "<div class=""receipttotalsrow rectotstatetax""><div class=""ectleft"">"&xxStaTax&":</div><div class=""ectright"">"&FormatEuroCurrency(ordStateTax)&"</div></div>"&vbCrLf
		end if
		if ordHSTTax>0 then
			recpt=recpt & "<div class=""receipttotalsrow rectothsttax""><div class=""ectleft"">"&xxHST&":</div><div class=""ectright"">"&FormatEuroCurrency(ordHSTTax)&"</div></div>"&vbCrLf
		end if
		if ordCountryTax>0 OR alwaysdisplaycountrytax then
			recpt=recpt & "<div class=""receipttotalsrow rectotcountrytax""><div class=""ectleft"">"&xxCntTax&":</div><div class=""ectright"">"&FormatEuroCurrency(ordCountryTax)&"</div></div>"&vbCrLf
		end if
		recpt=recpt & "<div class=""receipttotalsrow rectotgrandtotal""><div class=""ectleft"">"&xxGndTot&":</div><div class=""ectright"">"&FormatEuroCurrency(ordGrandTotal)&"</div></div>"
		recpt=recpt & "</div>" & _
				"</div>"&vbCrLf
	end if
	getrecpt=recpt
end function
Sub do_order_success(sorderid,sEmail,sendstoreemail,doshowhtml,sendcustemail,sendaffilemail,sendmanufemail)
Dim custEmail,ordAddInfo,affilID,dropShippers()
Redim dropShippers(2,10)
if htmlemails=true then
	emlNl="<br />"&vbCrLf
	xxThkYou=""
else
	emlNl=vbCrLf
end if
affilID=""
if NOT is_numeric(sorderid) then
	print "&nbsp;<br />&nbsp;<br />&nbsp;<br /><p align=""center"">Illegal Order ID</p><br />&nbsp;"
	exit sub
end if
ordID=sorderid
hasdownload=FALSE
orderloyaltypoints=0
ordClientID=0
savelangid=languageid
sSQL="SELECT ordID,ordName,ordLastName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordShipName,ordShipLastName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,ordShipPhone,ordPayProvider,ordAuthNumber,ordTotal,ordDate,ordStateTax,ordCountryTax,ordHSTTax,ordHandling,ordShipping,ordAffiliate,ordShipType,ordDiscount,ordDiscountText,ordComLoc,ordExtra1,ordExtra2,ordShipExtra1,ordShipExtra2,ordCheckoutExtra1,ordCheckoutExtra2,ordSessionID,ordClientID,ordLang,loyaltyPoints,payProvID,ordAddInfo FROM orders INNER JOIN payprovider ON payprovider.payProvID=orders.ordPayProvider WHERE ordAuthNumber<>'' AND ordID="&replace(sorderid,"'","")
rs.open sSQL,cnn,0,1
if NOT rs.EOF then
	orderloyaltypoints=rs("loyaltyPoints")
	ordClientID=rs("ordClientID")
	orderText=""
	saveHeader=""
	success=TRUE
	languageid=rs("ordLang")+1
	ordAuthNumber=rs("ordAuthNumber")
	ordSessionID=rs("ordSessionID")
	payprovid=rs("payProvID")
	ordName=trim(rs("ordName")&" "&rs("ordLastName"))
	ordDate=rs("ordDate")
	if now()-ordDate>7 AND NOT (SESSION("loggedon")<>"" OR SESSION("clientID")=rs("ordClientID")) then
		print "<div style=""text-align:center;padding:50px"">Please contact our customer services department for details of your order</div>"
		exit sub
	end if
	if rs("ordShipType")="MODWARNOPEN" then
		print "<div style=""font-weight:bold;text-align:center"">&nbsp;<br />&nbsp;<br />&nbsp;<br />" & xxManRev & "&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;<br />"
		response.write imageorbutton(imgcontinueshopping,"&nbsp;"&xxCntShp&"&nbsp;","continueshopping",IIfVr(thankspagecontinue<>"",thankspagecontinue,storeurl),IIfVr(thankspagecontinue="javascript:history.go(-1)",TRUE,FALSE))&"&nbsp;"
		print "</div>"
		success=FALSE
		rs.close
		exit sub
	end if
	sSQL="SELECT "&getlangid("emailsubject",4096)&","&getlangid("emailheaders",4096)&","&getlangid("receiptheaders",4096)&" FROM emailmessages WHERE emailID=1"
	rs2.Open sSQL,cnn,0,1
		emailsubject=trim(rs2(getlangid("emailsubject",4096))&"")
		emailheader=trim(rs2(getlangid("emailheaders",4096))&"")
		emailheader=replace(emailheader,"%emailmessage%","%messagebody%")
		receiptheader=trim(rs2(getlangid("receiptheaders",4096))&"")
		if instr(emailheader, "%messagebody%")=0 then emailheader=emailheader & "%messagebody%"
		if instr(receiptheader, "%messagebody%")=0 then receiptheader=receiptheader & "%messagebody%"
	rs2.Close
	sSQL="SELECT "&getlangid("dropshipsubject",4096)&","&getlangid("dropshipheaders",4096)&" FROM emailmessages WHERE emailID=1"
	rs2.Open sSQL,cnn,0,1
		dropshipsubject=trim(rs2(getlangid("dropshipsubject",4096))&"")
		dropshipheader=trim(rs2(getlangid("dropshipheaders",4096))&"")
		dropshipheader=replace(dropshipheader,"%emailmessage%","%messagebody%")
		if instr(dropshipheader, "%messagebody%")=0 then dropshipheader=dropshipheader & "%messagebody%"
	rs2.Close
	sSQL="SELECT "&getlangid("pProvHeaders",4096)&" FROM payprovider WHERE payProvID=" & payprovid
	rs2.Open sSQL,cnn,0,1
		payprovheader=trim(rs2(getlangid("pProvHeaders",4096))&"")
		payprovheader=replace(payprovheader,"%emailmessage%","%messagebody%")
		if instr(payprovheader, "%messagebody%")=0 then payprovheader=payprovheader & "%messagebody%"
	rs2.Close
	sSQL="SELECT "&getlangid("pProvDropShipHeaders",4096)&" FROM payprovider WHERE payProvID=" & payprovid
	rs2.Open sSQL,cnn,0,1
		payprovdropshipheader=trim(rs2(getlangid("pProvDropShipHeaders",4096))&"")
		payprovdropshipheader=replace(payprovdropshipheader,"%emailmessage%","%messagebody%")
		if instr(payprovdropshipheader, "%messagebody%")=0 then payprovdropshipheader=payprovdropshipheader & "%messagebody%"
	rs2.Close
	emailheader=replace(emailheader, "%messagebody%", payprovheader)
	dropshipheader=replace(dropshipheader, "%messagebody%", payprovdropshipheader)
	emailheader=replace(replace(emailheader, "%ordername%", ordName), "%nl%", "<br />")
	emailheader=replace(emailheader, "%orderdate%", FormatDateTime(ordDate, 1) & " " & FormatDateTime(ordDate, 4))
	receiptheader=replace(replace(receiptheader, "%ordername%", ordName), "%nl%", "<br />")
	receiptheader=replace(receiptheader, "%orderdate%", FormatDateTime(ordDate, 1) & " " & FormatDateTime(ordDate, 4))
	dropshipheader=replace(replace(dropshipheader, "%ordername%", ordName), "%nl%", "<br />")
	dropshipheader=replace(dropshipheader, "%orderdate%", FormatDateTime(ordDate, 1) & " " & FormatDateTime(ordDate, 4))

	orderText=orderText & xxOrdId & ": " & rs("ordID") & "<br />"
	if thereference<>"" then orderText=orderText & "Transaction Ref" & ": " & thereference & "<br />"
	orderText=orderText & xxCusDet & ": " & "<br />"
	if trim(extraorderfield1)<>"" then orderText=orderText & extraorderfield1 & ": " & rs("ordExtra1") & "<br />"
	orderText=orderText & ordName & "<br />"
	orderText=orderText & rs("ordAddress") & "<br />"
	if trim(rs("ordAddress2"))<>"" then orderText=orderText & rs("ordAddress2") & "<br />"
	orderText=orderText & rs("ordCity") & ", " & rs("ordState") & "<br />"
	orderText=orderText & rs("ordZip") & "<br />"
	orderText=orderText & rs("ordCountry") & "<br />"
	orderText=orderText & xxEmail & ": " & rs("ordEmail") & "<br />"
	custEmail=rs("ordEmail")
	orderText=orderText & xxPhone & ": " & rs("ordPhone") & "<br />"
	if trim(extraorderfield2)<>"" then orderText=orderText & extraorderfield2 & ": " & rs("ordExtra2") & "<br />"
	if trim(rs("ordShipName")&"")<>"" OR trim(rs("ordShipLastName")&"")<>"" OR trim(rs("ordShipAddress")&"")<>"" then
		orderText=orderText & xxShpDet & ": " & "<br />"
		if trim(extraorderfield1)<>"" AND trim(rs("ordShipExtra1")&"")<>"" then orderText=orderText & extraorderfield1 & ": " & rs("ordShipExtra1") & "<br />"
		orderText=orderText & rs("ordShipName")&IIfVr(trim(rs("ordShipName")&"")<>"" AND trim(rs("ordShipLastName")&"")<>"", " ", "")&rs("ordShipLastName") & "<br />"
		orderText=orderText & rs("ordShipAddress") & "<br />"
		if trim(rs("ordShipAddress2"))<>"" then orderText=orderText & rs("ordShipAddress2") & "<br />"
		orderText=orderText & rs("ordShipCity") & ", " & rs("ordShipState") & "<br />"
		orderText=orderText & rs("ordShipZip") & "<br />"
		orderText=orderText & rs("ordShipCountry") & "<br />"
		if trim(rs("ordShipPhone")&"")<>"" then orderText=orderText & xxPhone & ": " & rs("ordShipPhone") & "<br />"
		if trim(extraorderfield2)<>"" AND trim(rs("ordShipExtra2")&"")<>"" then orderText=orderText & extraorderfield2 & ": " & rs("ordShipExtra2") & "<br />"
	end if
	ordShipType=rs("ordShipType")
	if ordShipType<>"" then
		orderText=orderText & "<br />" & xxShpMet & ": " & ordShipType
		if willpickuptext<>ordShipType then
			if (rs("ordComLoc") AND 2)=2 then orderText=orderText & xxWtIns else if forceinsuranceselection then orderText=orderText & xxNoWtIn
		end if
		orderText=orderText & "<br />"
		if (rs("ordComLoc") AND 1)=1 then orderText=orderText & xxCerCLo & "<br />"
		if (rs("ordComLoc") AND 4)=4 then orderText=orderText & xxSatDeR & "<br />"
	end if
	if trim(extracheckoutfield1)<>"" AND trim(rs("ordCheckoutExtra1")&"")<>"" then orderText=orderText & extracheckoutfield1 & ": " & rs("ordCheckoutExtra1") & "<br />"
	if trim(extracheckoutfield2)<>"" AND trim(rs("ordCheckoutExtra2")&"")<>"" then orderText=orderText & extracheckoutfield2 & ": " & rs("ordCheckoutExtra2") & "<br />"
	ordAddInfo=trim(rs("ordAddInfo"))
	if ordAddInfo<>"" then
		orderText=orderText & "<br />" & xxAddInf & ": " & "<br />"
		orderText=orderText & replace(replace(ordAddInfo,vbLf,"<br />"),vbCr,"") & "<br />"
	end if
	ordTotal=rs("ordTotal")
	ordStateTax=rs("ordStateTax")
	ordDiscount=rs("ordDiscount")
	ordDiscountText=rs("ordDiscountText")
	ordCountryTax=rs("ordCountryTax")
	ordHSTTax=rs("ordHSTTax")
	ordShipping=rs("ordShipping")
	ordHandling=rs("ordHandling")
	affilID=trim(rs("ordAffiliate"))
	ordCity=rs("ordCity")
	ordState=rs("ordState")
	ordCountry=rs("ordCountry")
	ordEmail=rs("ordEmail")
	if FALSE AND sendcustemail AND NOT isresendemail then
		set rsdn=Server.CreateObject("ADODB.RecordSet")
		firsttime=TRUE
		ordGrandTotal=(ordTotal+ordStateTax+ordCountryTax+ordHSTTax+ordShipping+ordHandling)-ordDiscount
		sSQL="SELECT dnID,dnLastUpdated FROM devicenotifications"
		rs2.open sSQL,cnn,0,1
		do while NOT rs2.EOF
			badgecount=0
			sSQL="SELECT COUNT(*) AS tcnt FROM orders WHERE ordDate>"&vsusdatetime(rs2("dnLastUpdated"))
			rsdn.open sSQL,cnn,0,1
				if isnull(badgecount) then badgecount=0 else badgecount=rsdn("tcnt")
			rsdn.close
			if callxmlfunction("https://www.ecommercetemplates.com/applesalealert.asp?id=" & rs2("dnID") & "&badges=" & badgecount & "&store=" & server.urlencode(storeurl) & "&amount=" & server.urlencode(FormatEuroCurrency(ordGrandTotal)),"user=" & server.urlencode(currConvUser) & "&pass=" & server.urlencode(currConvPw),sresult,"","Msxml2.ServerXMLHTTP",errormsg,12) then
				if trim(sresult)<>"" AND firsttime then call notification_err_msg(sresult)
			elseif firsttime then
				call notification_err_msg(errormsg)
			end if
			firsttime=FALSE
			rs2.movenext
		loop
		rs2.close
		set rsdn=nothing
	end if
else
	print "<div style=""padding:40px;text-align:center"">Cannot find details for order id: " & sorderid & "</div>"
	rs.close
	exit sub
end if
rs.close
saveCustomerDetails=orderText
if loyaltypoints<>"" then orderText=orderText & "%loyaltypointplaceholder%"
orderText=orderText & "%digidownloadplaceholder%"
reviewlinks=""
loyaltypointtotal=0
sSQL="SELECT cartProdId,cartOrigProdId,cartProdName,cartProdPrice,cartQuantity,cartID,cartGiftWrap,pDropship,pDisplay,pStaticPage,pStaticURL,pSKU"&IIfVr(digidownloads=TRUE,",pDownload","")&",cartGiftMessage FROM cart LEFT JOIN products ON cart.cartProdId=products.pID WHERE cartOrderID="&replace(sorderid,"'","") & " ORDER BY cartID"
rs.open sSQL,cnn,0,1
if NOT rs.EOF then
	do while not rs.EOF
		if rs("cartProdId")=giftcertificateid then
			sSQL="UPDATE giftcertificate SET gcAuthorized=1,gcOrigAmount="&rs("cartProdPrice")&",gcRemaining="&rs("cartProdPrice")&" WHERE gcCartID=" & rs("cartID")
			ect_query(sSQL)
			if sendcustemail then
				sSQL="SELECT "&getlangid("giftcertsubject",4096)&","&getlangid("giftcertemail",4096)&" FROM emailmessages WHERE emailID=1"
				rs2.Open sSQL,cnn,0,1
					giftcertsubject=trim(rs2(getlangid("giftcertsubject",4096)))
					emailBody=trim(rs2(getlangid("giftcertemail",4096)))
				rs2.Close
				sSQL="SELECT "&getlangid("giftcertsendersubject",4096)&","&getlangid("giftcertsender",4096)&" FROM emailmessages WHERE emailID=1"
				rs2.Open sSQL,cnn,0,1
					senderSubject=trim(rs2(getlangid("giftcertsendersubject",4096)))
					senderBody=trim(rs2(getlangid("giftcertsender",4096)))
				rs2.Close
				sSQL="SELECT gcID,gcTo,gcFrom,gcEmail,gcMessage FROM giftcertificate WHERE gcCartID="&rs("cartID")
				rs2.Open sSQL,cnn,0,1
					emailBody=replace(emailBody, "%toname%", rs2("gcTo"))
					emailBody=replace(emailBody, "%fromname%", rs2("gcFrom"))
					emailBody=replace(emailBody, "%value%", FormatEuroCurrency(rs("cartProdPrice")))
					emailBody=replaceemailtxt(emailBody, "%message%", trim(rs2("gcMessage")&""), replaceone)
					emailBody=replace(emailBody, "%storeurl%", storeurl)
					emailBody=replace(emailBody, "%certificateid%", rs2("gcID"))
					emailBody=replace(emailBody, "<br />", emlNl)
					Call DoSendEmailEO(rs2("gcEmail"), sEmail, "", replace(giftcertsubject, "%fromname%", rs2("gcFrom")), emailBody, emailObject, themailhost, theuser, thepass)
					
					senderBody=replace(senderBody, "%toname%", rs2("gcTo"))
					Call DoSendEmailEO(custEmail,sEmail,"", replace(senderSubject, "%toname%", rs2("gcTo")),senderBody & emlNl & emailBody,emailObject,themailhost,theuser,thepass)
				rs2.Close
			end if
		end if
		reviewprodid=rs("cartProdId")
		reviewisstatic=rs("pStaticPage")
		reviewstaticurl=rs("pStaticURL")
		reviewprodname=rs("cartProdName")
		if trim(rs("cartOrigProdId")&"")<>"" then
			sSQL="SELECT pName,pStaticPage,pStaticURL FROM products WHERE pID='"&escape_string(rs("cartOrigProdId"))&"'"
			rs2.open sSQL,cnn,0,1
			if NOT rs2.EOF then
				reviewprodid=rs("cartOrigProdId")
				reviewisstatic=rs2("pStaticPage")
				reviewstaticurl=rs2("pStaticURL")
				reviewprodname=rs2("pName")
			end if
			rs2.close
		end if
		if isnull(reviewisstatic) then
			thelink=""
		else
			thelink=storeurl & getdetailsurl(reviewprodid,reviewisstatic,reviewprodname,trim(reviewstaticurl&""),"review=true","")
		end if
		if htmlemails=TRUE AND thelink<>"" then thelink="<a href=""" & thelink & """>" & thelink & "</a>" else thelink=replace(thelink,"&amp;","&")
		if thelink<>"" then reviewlinks=reviewlinks & thelink & emlNl
		localhasdownload=FALSE
		if digidownloads=TRUE then
			if trim(rs("pDownload")&"")<>"" then localhasdownload=TRUE
		end if
		saveCartItems="--------------------------" & "<br />"
		saveCartItems=saveCartItems & xxPrId & ": " & rs("cartProdId") & "<br />"
		saveCartItems=saveCartItems & xxPrNm & ": " & rs("cartProdName") & "<br />"
		saveCartItems=saveCartItems & xxQuant & ": " & rs("cartQuantity") & "<br />"
		if rs("cartGiftWrap")<>0 then
			saveCartItems=saveCartItems & xxGWrSel & "<br />"
			cartGiftMessage=trim(rs("cartGiftMessage"))
			if cartGiftMessage<>"" then
				saveCartItems=saveCartItems & xxGifMes & ": " & cartGiftMessage & "<br />"
			end if
		end if
		orderText=orderText & saveCartItems
		theoptions=""
		theoptionspricediff=0
		sSQL="SELECT coOptGroup,coCartOption,coPriceDiff,optRegExp FROM cartoptions LEFT JOIN options ON cartoptions.coOptID=options.optID WHERE coCartID="&rs("cartID") & " ORDER BY coID"
		rs2.Open sSQL,cnn,0,1
		do while NOT rs2.EOF
			theoptionspricediff=theoptionspricediff + rs2("coPriceDiff")
			optionline=IIfVr(htmlemails=true,"&nbsp;&nbsp;&nbsp;&nbsp;>&nbsp;","> > > ") & rs2("coOptGroup") & " : " & replace(rs2("coCartOption")&"", vbCrLf, "<br />")
			theoptions=theoptions & optionline
			saveCartItems=saveCartItems & optionline & "<br />"
			if rs2("coPriceDiff")=0 OR hideoptpricediffs=TRUE OR nopriceanywhere then
				theoptions=theoptions & "<br />"
			else
				theoptions=theoptions & " ("
				if rs2("coPriceDiff") > 0 then theoptions=theoptions & "+"
				theoptions=theoptions & FormatEmailEuroCurrency(rs2("coPriceDiff")) & ")" & "<br />"
			end if
			if rs2("optRegExp")="!!" then localhasdownload=FALSE
			rs2.MoveNext
		loop
		rs2.Close
		if NOT nopriceanywhere then orderText=orderText & xxUnitPr & ": " & IIfVr(hideoptpricediffs=TRUE,FormatEmailEuroCurrency(rs("cartProdPrice")+theoptionspricediff),FormatEmailEuroCurrency(rs("cartProdPrice"))) & "<br />"
		orderText=orderText & theoptions
		if rs("pDropship")<>0 then
			index=0
			do while TRUE
				if index>=UBOUND(dropShippers,2) then Redim Preserve dropShippers(2,index+10)
				if dropShippers(0, index)="" OR dropShippers(0, index)=rs("pDropship") then exit do
				index=index+1
			loop
			dropShippers(0, index)=rs("pDropship")
			dropShippers(1, index)=dropShippers(1, index) & saveCartItems & IIfVs(rs("pSKU")<>"","pSKU:" & rs("pSKU") & "<br />")
		end if
		if localhasdownload=TRUE then hasdownload=TRUE
		rs.MoveNext
	loop
	orderText=orderText & "--------------------------" & "<br />"
	if NOT nopriceanywhere then
		orderText=orderText & xxOrdTot & " : " & FormatEmailEuroCurrency(ordTotal) & "<br />"
		if combineshippinghandling=TRUE then
			orderText=orderText & xxShipHa & " : " & FormatEmailEuroCurrency(ordShipping + ordHandling) & "<br />"
		else
			if shipType<>0 then orderText=orderText & xxShippg & " : " & FormatEmailEuroCurrency(ordShipping) & "<br />"
			if cdbl(ordHandling)<>0.0 then orderText=orderText & xxHndlg & " : " & FormatEmailEuroCurrency(ordHandling) & "<br />"
		end if
		if cdbl(ordDiscount)<>0.0 then orderText=orderText & xxDscnts & " : " & FormatEmailEuroCurrency(ordDiscount) & "<br />"
		if cdbl(ordStateTax)<>0.0 then orderText=orderText & xxStaTax & " : " & FormatEmailEuroCurrency(ordStateTax) & "<br />"
		if cdbl(ordCountryTax)<>0.0 then orderText=orderText & xxCntTax & " : " & FormatEmailEuroCurrency(ordCountryTax) & "<br />"
		if cdbl(ordHSTTax)<>0.0 then orderText=orderText & xxHST & " : " & FormatEmailEuroCurrency(ordHSTTax) & "<br />"
		ordGrandTotal=(ordTotal+ordStateTax+ordCountryTax+ordHSTTax+ordShipping+ordHandling)-ordDiscount
		orderText=orderText & xxGndTot & " : " & FormatEmailEuroCurrency(ordGrandTotal) & "<br />"
	end if
else
	print "&nbsp;<br />&nbsp;<br />&nbsp;<br /><p align=""center"">Cannot find details for cart id: " & sorderid & "</p><br />&nbsp;"
	rs.close
	exit sub
end if
rs.close
if loyaltypoints<>"" AND orderloyaltypoints=0 AND sendmanufemail then
	loyaltypointtotal=int((ordTotal-ordDiscount)*loyaltypoints)
	if loyaltypointtotal>0 then
		if loyaltypointsnowholesale OR loyaltypointsnopercentdiscount then
			sSQL="SELECT clActions FROM customerlogin WHERE clID=" & ordClientID
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if loyaltypointsnowholesale AND (rs("clActions") AND 8)=8 then loyaltypointtotal=0
				if loyaltypointsnopercentdiscount AND (rs("clActions") AND 16)=16 then loyaltypointtotal=0
			end if
			rs.close
		end if
		sSQL="UPDATE orders SET loyaltyPoints=" & loyaltypointtotal & " WHERE ordID="&replace(sorderid,"'","")
		ect_query(sSQL)
		sSQL="UPDATE customerlogin SET loyaltyPoints=loyaltyPoints+" & loyaltypointtotal & " WHERE clID="&ordClientID
		ect_query(sSQL)
	end if
end if
numloyaltypoints=0
if loyaltypoints<>"" then
	sSQL="SELECT loyaltyPoints FROM orders WHERE ordID="&replace(sorderid,"'","")
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then numloyaltypoints=rs("loyaltyPoints")
	rs.close
	if numloyaltypoints>0 then
		orderText=replace(orderText,"%loyaltypointplaceholder%",xxLoyPoi & ": " & numloyaltypoints & "<br />")
	else
		orderText=replace(orderText,"%loyaltypointplaceholder%","")
	end if
end if
if hasdownload=TRUE AND digidownloademail<>"" then
	fingerprint=HMAC(digidownloadsecret, sorderid & ordAuthNumber & ordSessionID)
	fingerprint=Left(fingerprint, 14)
	digidownloademail=replace(digidownloademail,"%orderid%",ordID)
	digidownloademail=replace(digidownloademail,"%password%",fingerprint)
	digidownloademail=replace(digidownloademail,"%nl%",emlNl)
	orderEmailText=replace(orderText,"%digidownloadplaceholder%",digidownloademail)
else
	orderEmailText=replace(orderText,"%digidownloadplaceholder%","")
end if
orderText=replace(orderText,"%digidownloadplaceholder%","")
emailheader=replaceemailtxt(emailheader, "%reviewlinks%", reviewlinks, replaceone)
receiptheader=replaceemailtxt(receiptheader, "%reviewlinks%", reviewlinks, replaceone)
emailrecpt=getemailrecpt(ordID)
recpt=getrecpt(ordID)
if loyaltypoints<>"" then
	if numloyaltypoints>0 then
		recpt=replace(recpt,"<!--%loyaltypointplaceholder%-->","<div class=""receiptcontainer rcontloyaltypoints""><div class=""receiptleft"">"&xxLoyPoi&"</div><div class=""receiptright"">" & numloyaltypoints & "</div></div>")
		emailrecpt=replace(emailrecpt,"<!--%loyaltypointplaceholder%-->","<tr><td valign=""top""><strong>"&xxLoyPoi&":</strong></td><td valign=""top"" colspan=""3"">" & numloyaltypoints & "</td></tr>")
	else
		recpt=replace(recpt,"<!--%loyaltypointplaceholder%-->","")
		emailrecpt=replace(emailrecpt,"<!--%loyaltypointplaceholder%-->","")
	end if
end if
if hasdownload=FALSE then
	recpt=replace(recpt,"<!--%digidownloadplaceholder%-->","")
	recpt=replace(recpt,"<!--%digidownloaditems%-->","")
	emailrecpt=replace(emailrecpt,"<!--%digidownloadplaceholder%-->","")
	emailrecpt=replace(emailrecpt,"<!--%digidownloaditems%-->","")
end if
emlhdrs="<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd""><html xmlns=""http://www.w3.org/1999/xhtml""><head><style type=""text/css"">body{font-size:11px; font-family: Tahoma, Helvetica, Arial, Verdana}hr{height: 0;border-width: 1px 0 0 0;border-style: solid;border-color: #006AC8;}</style></head>"
set toregexp=new regexp
toregexp.pattern="<html(.*?)>"
toregexp.ignorecase=TRUE
hasheaders=toregexp.test(emailheader)
set toregexp=Nothing
if sendstoreemail AND allStoreEmails<>"" then
	allemailsarray=split(allStoreEmails,",")
	for each sstoreemail in allemailsarray
		if htmlemails=TRUE then
			Call DoSendEmailEO(sstoreemail,sEmail,custEmail,replace(replace(xxOrdStr, "%ordername%", ordName), "%orderid%", sorderid), IIfVs(NOT hasheaders,emlhdrs & "<body class=""receiptbody"">") & replace(replace(emailheader, "%messagebody%", replace(emailrecpt,"<!--%digidownloadplaceholder%-->","<tr><td valign=""top""><strong>" & xxDigPro & ":</strong></td><td valign=""top"" colspan=""3"">" & digidownloademail & "</td></tr>")), "<br />", emlNl) & IIfVs(NOT hasheaders,"</body></html>"),emailObject,themailhost,theuser,thepass)
		else
			Call DoSendEmailEO(sstoreemail,sEmail,custEmail,replace(replace(xxOrdStr, "%ordername%", ordName), "%orderid%", sorderid),replace(replace(emailheader, "%messagebody%", orderEmailText), "<br />", emlNl),emailObject,themailhost,theuser,thepass)
		end if
	next
end if
' And one for the customer
if sendcustemail then
	if htmlemails=TRUE then
		Call DoSendEmailEO(custEmail,sEmail,"",replace(replace(emailsubject, "%ordername%", ordName), "%orderid%", sorderid), IIfVs(NOT hasheaders,emlhdrs & "<body class=""receiptbody"">") & replace(replace(emailheader, "%messagebody%", replace(emailrecpt,"<!--%digidownloadplaceholder%-->","<tr><td valign=""top""><strong>" & xxDigPro & ":</strong></td><td valign=""top"" colspan=""3"">" & digidownloademail & "</td></tr>")), "<br />", emlNl) & IIfVs(NOT hasheaders,"</body></html>"),emailObject,themailhost,theuser,thepass)
	else
		Call DoSendEmailEO(custEmail,sEmail,"",replace(replace(emailsubject, "%ordername%", ordName), "%orderid%", sorderid),IIfVr(trim(xxTouSoo)<>"",xxTouSoo & emlNl & emlNl, "") & replace(replace(emailheader, "%messagebody%", orderEmailText), "<br />", emlNl),emailObject,themailhost,theuser,thepass)
	end if
end if
languageid=savelangid
' Drop Shippers / Manufacturers
if sendmanufemail then
	for index=0 to UBOUND(dropShippers,2)
		if dropShippers(0, index)="" then exit for
		sSQL="SELECT dsEmail,dsAction,dsEmailHeader FROM dropshipper WHERE dsID="&dropShippers(0, index)
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			if (rs("dsAction") AND 1)=1 OR sendmanufemail=2 then
				dsEmailHeader=trim(rs("dsEmailHeader")&"")
				Call DoSendEmailEO(trim(rs("dsEmail")),sEmail,"",replace(dropshipsubject, "%orderid%", sorderid),replace(replace(dropshipheader, "%messagebody%", IIfVr(dsEmailHeader<>"", emlNl & replace(dsEmailHeader, "%nl%", emlNl) & emlNl, "") & saveCustomerDetails & dropShippers(1, index)), "<br />", emlNl),emailObject,themailhost,theuser,thepass)
			end if
		end if
		rs.close
	next
end if
if sendaffilemail then
	if affilID<>"" then
		sSQL="SELECT affilEmail,affilInform FROM affiliates WHERE affilID='"&replace(affilID,"'","")&"'"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			if int(rs("affilInform"))=1 then
				affiltext=xxAff1 & " "&FormatEmailEuroCurrency(ordTotal-ordDiscount)&"."&emlNl&emlNl&xxAff2&emlNl&emlNl&xxThnks&emlNl
				Call DoSendEmailEO(trim(rs("affilEmail")),sEmail,"",replace(xxAff3, "%orderid%", sorderid),affiltext,emailObject,themailhost,theuser,thepass)
			end if
		end if
		rs.close
	end if
end if
if doshowhtml then
		if hasimageupload then
			filetypes=".bmp,.gif,.jpg,.jpe,.jpg,.pdf,.png,.tif,.tiff"
			maxfilesize=2*1024*1024
			if uploadfiletypes<>"" then filetypes=uploadfiletypes
			if uploadmaxfilesize<>"" then maxfilesize=uploadmaxfilesize
			print "<div class=""iuopaque"" id=""imageuploadopdiv""" & IIfVs(imageuploadbybutton," style=""display:none""") & ">"
				print "<div class=""iuwrap"" id=""imageuploaddiv"">"
					print "<div style=""padding:20px;text-align:center;font-size:1.25em;font-weight:bold;color:#333"">" & xxImUplo & "</div>"
%>
<div class="imageupload imageuploadhead"><%=xxImgUpl%></div>
<div class="imageupload imageuploadfile"><input class="ectfileinput" type="file" id="fileselectinput" value="<%=xxChoImg%>" accept="<%=filetypes%>" style="width:80%" /></div>
<div class="imageupload imageuploadcomments"><%=xxImgCom%><input type="text" id="comments" style="width:50%" maxlength="512" /></div>
<div class="imageupload imageuploadprogress"><progress id="progressBar" value="0" max="100" style="width:80%"></progress></div>
<div class="imageupload uploadedimages" id="uploadedimages" style="display:none">
	<div class="imageuploadtable" id="uploadedimagestable">
		<div class="imageuploadrow"><div><%=xxImgFin%></div><div><%=xxImgSta%></div></div>
	</div>
</div>
<div class="imageupload imageuploadbuttons">
	<input class="ectbutton" id="startupload" type="button" value="<%=xxUplImg%>" onclick="dosubmitimage()" disabled="disabled" />
	&nbsp; <input class="ectbutton" type="button" value="<%=xxDonUpl%>" onclick="imuploaddone()" />
</div>
<script>
var fileselectinput=document.getElementById("fileselectinput");
var ectuploadfilename='', ectuploadimgsrc='';
var imuploadhasselected=false;
var fuajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
fileselectinput.addEventListener("change", function(){ hasselectedimage(this); });
function doneimageupload(){
	if(fuajaxobj.readyState==4){
		var imgstatus='<div class="receiptheadrow"><div>'+(ectuploadfilename.length>20?ectuploadfilename.substr(0,8)+'...'+ectuploadfilename.substr(-8):ectuploadfilename)+'</div><div>';
		if(fuajaxobj.status==200){
			var restxt=fuajaxobj.responseText;
			if(restxt=='SUCCESS')
				imgstatus+='&#9989;';
			else{
				if(restxt=='MAXIMAGES') restxt='Max number of images has been uploaded';
				imgstatus+='&#10060; '+restxt;
			}
		}else{
			imgstatus+=fuajaxobj.status+":Refused by server";
			document.getElementById("progressBar").value=0;
		}
		document.getElementById("uploadedimages").style.display='';
		document.getElementById("uploadedimagestable").innerHTML+=imgstatus+'</div></div>';
		document.getElementById("fileselectinput").value='';
		document.getElementById("comments").value='';
		document.getElementById('startupload').disabled='disabled';
		imuploadhasselected=false;
	}
}
function imuploaddone(){
	var dodone=true;
	if(imuploadhasselected) dodone=confirm('You have selected a file but not uploaded it. Please click cancel to go back and upload the file or ok to continue.');
	if(dodone){
		document.getElementById('imageuploadopdiv').style.display='none';
		document.getElementById('imageuploadbutton').style.display='';
	}
}
function progressHandler(event){
	document.getElementById("progressBar").value=Math.round((event.loaded / event.total) * 100);
}
function errorHandler(event){
	alert("Upload Failed");
}
function abortHandler(event){
	alert("Upload Aborted");
}
function dosubmitimage(){
	var postdata='orderid=<%=ordID%>';
	postdata+='&filename='+encodeURIComponent(ectuploadfilename) +
		'&session=<%=ordSessionID %>' +
		'&check=<%=hex_sha1("imageupload^" & ordID & "^" & adminSecret & "^fromect" & ordSessionID)%>' +
		'&comments='+encodeURIComponent(document.getElementById('comments').value) +
		'&imgsrc='+encodeURIComponent(ectuploadimgsrc);
	fuajaxobj.onreadystatechange=doneimageupload;
	fuajaxobj.upload.addEventListener("progress", progressHandler, false);
	fuajaxobj.addEventListener("error", errorHandler, false);
	fuajaxobj.addEventListener("abort", abortHandler, false);
	fuajaxobj.open("POST",'vsadmin/ajaxservice.asp?action=imageupload',true);
	fuajaxobj.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
	fuajaxobj.send(postdata);
}
function hasselectedimage(input){
	var reader;
	if(input.files && input.files[0]){
		reader=new FileReader();
		reader.onload=function(e){
			if(input.files[0].size><%=maxfilesize%>){
				alert('The file you have selected is too large to upload. (Max. <%=(maxfilesize/(1024*1024))&"M"%>)');
			}else{
				ectuploadimgsrc=e.target.result;
				ectuploadfilename=input.files[0].name;
				document.getElementById('startupload').disabled='';
				document.getElementById("progressBar").value=0;
				imuploadhasselected=true;
			}
		}
		reader.readAsDataURL(input.files[0]);
	}
}
</script>
<%
				print "</div>"
			print "</div>"
		end if
%>
<script>
<!--
function doprintcontent(){
	var stylesheetlist='';
	if(document.styleSheets){
		for(var dpci=0;dpci<document.styleSheets.length;dpci++){
			if(document.styleSheets[dpci].href){
				stylesheetlist+='<link rel="stylesheet" type="text/css" href="'+document.styleSheets[dpci].href+'" />\n';
			}
		}
	}
	var prnttext='<html><head>'+stylesheetlist+'</head><body onload="window.print()"><div class="printbody">\n';
	prnttext+=document.getElementById('printcontent').innerHTML+"\n";
	prnttext+='</div></body></'+'html>\n';
	var newwin=window.open("","printit",'menubar=no, scrollbars=yes, width=600, height=450, directories=no,location=no,resizable=yes,status=no,toolbar=no');
	newwin.document.open();
	newwin.document.write(prnttext);
	newwin.document.close();
}
//-->
</script>
<%		if digidownloads<>TRUE then %>
			<div class="orderreceipt">
<%			if xxThkYou<>"" then print "<div class=""recptthanks"">" & xxThkYou & "</div>" %>
				<div id="printcontent" class="recptbody"><%=replace(replace(receiptheader, "%messagebody%", recpt),"%nl%","<br />")%></div>
				<div class="receiptbuttons">
<%			if xxRecEml<>"" then print "<div class=""receiptrecemail"">" & xxRecEml & "</div>"
			print "<div class=""receiptcontinueshopping"">" & imageorbutton(imgcontinueshopping,"&nbsp;"&xxCntShp&"&nbsp;","continueshopping",IIfVr(thankspagecontinue<>"",thankspagecontinue,storeurl),IIfVr(thankspagecontinue="javascript:history.go(-1)",TRUE,FALSE)) & "</div>"
			print "<div class=""receiptprintversion"">" & imageorbutton(imgprintversion,"&nbsp;"&xxPrint&"&nbsp;","printversion","doprintcontent();printwindow()",TRUE) & "</div>"
%>				</div>
			</div>
<%		end if
end if
end sub
%>