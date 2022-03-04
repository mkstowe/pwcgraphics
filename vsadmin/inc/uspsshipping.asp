<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
dim fedexnamespace
function getshippingerror(shiperrmsg)
	getshippingerror="<div class=""shippingerror"">There seems to be a problem connecting to the shipping rates server. Please wait a few moments and refresh your browser" & IIfVs(adminAltRates<>0,", or try a different shipping carrier") & ".</div><div class=""shiperrortechdetails"" style=""font-size:10px;color:#000000;margin-top:3px"">" & shiperrmsg & "</div>"
end function
function wantcarrierinsurance()
	wantcarrierinsurance=(abs(addshippinginsurance)=1 OR (abs(addshippinginsurance)=2 AND wantinsurance_)) AND NOT nocarrierinsurancerates
end function
sub sortshippingarray()
	maxshipoptions=UBOUND(intShipping,2)
	maxallocateditem=0
	for ssaindex=0 to maxshipoptions
		if intShipping(3,ssaindex)<>0 then maxallocateditem=ssaindex
	next
	for ssaindex=0 to maxallocateditem+1
		intShipping(2,ssaindex)=cdbl(intShipping(2,ssaindex))
		if intShipping(3,ssaindex)<>0 then
			for ssaindex2=ssaindex+1 to maxallocateditem+1
				if intShipping(3,ssaindex)<>0 AND intShipping(3,ssaindex2)<>0 AND intShipping(0,ssaindex)=intShipping(0,ssaindex2) then
					if intShipping(2,ssaindex)<=intShipping(2,ssaindex2) AND intShipping(4,ssaindex2)=0 then
						if intShipping(4,ssaindex2)=0 then intShipping(3,ssaindex2)=0
					elseif intShipping(2,ssaindex)>=intShipping(2,ssaindex2) then
						if intShipping(4,ssaindex)=0 then intShipping(3,ssaindex)=0
					end if
				end if
			next
		end if
	next
	maxshipoptions=maxallocateditem+1
	for ssaindex2=0 to maxshipoptions
		for ssaindex=1 to maxshipoptions
			if intShipping(3,ssaindex-1)=0 OR ((intShipping(3,ssaindex)<>0 AND (cdbl(intShipping(2,ssaindex))<cdbl(intShipping(2,ssaindex-1))))) then
				for iii=0 to UBOUND(intShipping)
					ttt=intShipping(iii,ssaindex) : intShipping(iii,ssaindex)=intShipping(iii,ssaindex-1) : intShipping(iii,ssaindex-1)=ttt
				next
			end if
		next
	next
end sub
function ParseDHLXMLOutput(sXML, international, byRef errormsg, byRef errorcode, byRef intShipping)
	noError=TRUE
	set xmlDoc=Server.CreateObject("MSXML2.DOMDocument")
	xmlDoc.validateOnParse=FALSE
	xmlDoc.loadXML(sXML)
	Set nodeList=xmlDoc.getElementsByTagName("QtdShp")
	if nodeList.length=0 then
		Set t2=xmlDoc.getElementsByTagName("ConditionData")
		errormsg=t2.Item(0).firstChild.nodeValue
		noError=FALSE
	else
		for i=0 to nodeList.length - 1
			set n=nodeList.item(i)
			shippingcharge=0
			Set t2=n.getElementsByTagName("ShippingCharge")
			if t2.length>0 then shippingcharge=cdbl(t2.Item(0).firstChild.nodeValue)
			
			if shippingcharge>0 then
				Set t2=n.getElementsByTagName("GlobalProductCode")
				Set s2=t2.Item(0)
				serviceid=s2.firstChild.nodeValue
				l=0
				do while (intShipping(5, l)<>serviceid AND intShipping(5, l)<>"")
					l=l + 1
				loop
				intShipping(5, l)=serviceid
				
				Set t2=n.getElementsByTagName("ProductShortName")
				intShipping(0, l)=t2.Item(0).firstChild.nodeValue
				
				if NOT noshipdateestimate then
					Set t2=n.getElementsByTagName("TotalTransitDays")
					intShipping(1, l)=cint(t2.Item(0).firstChild.nodeValue)
					Set t2=n.getElementsByTagName("PickupPostalLocAddDays")
					if is_numeric(t2.Item(0).firstChild.nodeValue) then intShipping(1, l)=intShipping(1, l) + cint(t2.Item(0).firstChild.nodeValue)
					Set t2=n.getElementsByTagName("DeliveryPostalLocAddDays")
					if is_numeric(t2.Item(0).firstChild.nodeValue) then intShipping(1, l)=intShipping(1, l) + cint(t2.Item(0).firstChild.nodeValue)
					intShipping(1, l)=intShipping(1, l) & " " & xxDays
				end if

				Set t2=n.getElementsByTagName("TotalTaxAmount")
				if t2.length>0 then shiptax=cdbl(t2.Item(0).firstChild.nodeValue)
				intShipping(2, l)=shippingcharge-shiptax

				wantthismethod=checkUPSShippingMeth(serviceid,discntsApp,showAs)
				if NOT wantthismethod then
					intShipping(3, l)=0
				else
					intShipping(0, l)=showAs
					intShipping(4, l)=discntsApp
					if discountshippingdhl<>"" then intShipping(2, l)=vsround(intShipping(2, l)*(1+discountshippingdhl/100.0),2)
					intShipping(3, l)=TRUE
				end if
			end if
		next
	end if
	ParseDHLXMLOutput=noError
end function
function ParseUSPSXMLOutput(sXML,international,byRef errormsg,byRef intShipping)
Dim noError, nodeList, packCost, xmlDoc, e, i, j, k, l, n, t, t2, s2
	noError=TRUE
	packCost=0
	errormsg=""
	gotxml=false
	on error resume next
	err.number=0
	set xmlDoc=Server.CreateObject("MSXML2.DOMDocument")
	if err.number=0 then gotxml=true
	if NOT gotxml then
		err.number=0
		set xmlDoc=Server.CreateObject("MSXML.DOMDocument")
		if err.number=0 then gotxml=true
	end if
	on error goto 0
	xmlDoc.validateOnParse=FALSE
	xmlDoc.loadXML(sXML)
	if xmlDoc.documentElement.nodeName="Error" then 'Top-level Error
		noError=FALSE
		Set nodeList=xmlDoc.getElementsByTagName("Error")
		Set n=nodeList.Item(0)
		for i=0 To n.childNodes.length - 1
			Set e=n.childNodes.Item(i)
			Select Case e.nodeName
				Case "Source"
				Case "Number"
				Case "Description"
					errormsg=e.firstChild.nodeValue
				Case "HelpFile"
				Case "HelpContext"
			end Select
		next
	else 'no Top-level Error
		Set nodeList=xmlDoc.getElementsByTagName("Package")
		for i=0 To nodeList.length - 1
			Set n=nodeList.Item(i)
			tmpArr=Split(n.getAttribute("ID"),"xx")
			quantity=Int(tmpArr(2))
			thisService=tmpArr(0)
			for j=0 To n.childNodes.length - 1
				Set e=n.childNodes.Item(j)
				if e.nodeName="Error" then 'Lower-level error
					errnum=0 : errdesc=""
					for k=0 To e.childNodes.length - 1
						Set t=e.childNodes.Item(k)
						Select Case t.nodeName
							Case "Number"
								errnum=t.firstChild.nodeValue
							Case "Description"
								errdesc=t.firstChild.nodeValue
								if dumpshippingxml=TRUE then print "USPS warning: " & t.firstChild.nodeValue & "<br>"
						end Select
					next
					if errnum="-2147219497" OR errnum="-2147219498" OR errnum="-2147219433" then ' Invalid Zip
						noError=FALSE
						if errnum="-2147219497" then errormsg=xxInvZip else errormsg=errdesc
					end if
				else
					select Case e.nodeName
						Case "Postage"
							if international="" then
								specialservices=0
								for ctr1=0 To e.childNodes.length - 1
									Set ee=e.childNodes.Item(ctr1)
									Select Case ee.nodeName
									Case "SpecialServices"
										for ctr2=0 To ee.childNodes.length - 1
											set eee=ee.childNodes.Item(ctr2)
											select Case eee.nodeName
											case "SpecialService"
												thisserviceid=""
												for ctr3=0 To eee.childNodes.length - 1
													set eeee=eee.childNodes.Item(ctr3)
													select Case eeee.nodeName
													case "ServiceID"
														thisserviceid=eeee.firstChild.nodeValue
													case "Price"
														thisprice=eeee.firstChild.nodeValue
													end select
												next
												if thisserviceid="100" OR thisserviceid="101" OR thisserviceid="125" then
													specialservices=cDbl(thisprice)
												end if
											end select
										next
									end select
								next
								
								gotrates=FALSE
								set rateelem=e.getElementsByTagName("CommercialPlusRate")
								if rateelem.length>0 then set e=rateelem.item(0) : gotrates=TRUE
								if NOT gotrates then
									set rateelem=e.getElementsByTagName("CommercialRate")
									if rateelem.length>0 then set e=rateelem.item(0) : gotrates=TRUE
								end if
								if NOT gotrates then
									set e=e.getElementsByTagName("Rate").Item(0)
								end if
								l=0
								do while (intShipping(5, l) <> thisService AND intShipping(5, l)<>"")
									l=l + 1
								loop
								intShipping(5, l)=thisService
								intShipping(8, l)=intShipping(8, l)+specialservices
								if noshipdateestimate then
									intShipping(1, l)=""
								elseif thisService="PARCEL" then
									intShipping(1, l)="2-7 " & xxDays
								elseif instr(lcase(thisService),"express")>0 then
									intShipping(1, l)="Overnight to most areas"
								elseif instr(lcase(thisService),"priority")>0 then
									intShipping(1, l)="2-3 " & xxDays
								elseif thisService="BPM" then
									intShipping(1, l)="2-7 " & xxDays
								elseif thisService="Media" then
									intShipping(1, l)="2-7 " & xxDays
								elseif thisService="FIRST-CLASS" then
									intShipping(1, l)="1-3 " & xxDays
								end if
								intShipping(2, l)=intShipping(2, l) + (e.firstChild.nodeValue * quantity)
								intShipping(3, l)=intShipping(3, l) + 1
								wantthismethod=FALSE
								for index2=0 to UBOUND(uspsmethods,2)
									if replace(thisService,"-"," ")=replace(uspsmethods(0,index2),"-"," ") then intShipping(0, l)=uspsmethods(2,index2) : wantthismethod=TRUE : exit for
								next
								if NOT wantthismethod then intShipping(3, l)=0
							end if
						Case "Service"
							if international<>"" then
								specialservices=0
								for ctr1=0 To e.childNodes.length - 1
									Set ee=e.childNodes.Item(ctr1)
									Select Case ee.nodeName
									Case "ExtraServices"
										for ctr2=0 To ee.childNodes.length - 1
											set eee=ee.childNodes.Item(ctr2)
											select Case eee.nodeName
											case "ExtraService"
												thisserviceid=""
												for ctr3=0 To eee.childNodes.length - 1
													set eeee=eee.childNodes.Item(ctr3)
													select Case eeee.nodeName
													case "ServiceID"
														thisserviceid=eeee.firstChild.nodeValue
													case "Price"
														thisprice=eeee.firstChild.nodeValue
													end select
												next
												if thisserviceid="106" OR thisserviceid="107" OR thisserviceid="108" then
													specialservices=cDbl(thisprice)
												end if
											end select
										next
									end select
								next
								
								serviceerror=FALSE : wantthismethod=FALSE
								serviceid=e.getAttribute("ID")
								Set t2=e.getElementsByTagName("SvcDescription")
								Set s2=t2.Item(0)
								l=0
								do while (intShipping(5, l)<>serviceid AND intShipping(5, l)<>"")
									l=l + 1
								loop
								intShipping(5, l)=serviceid
								intShipping(8, l)=intShipping(8, l)+specialservices
								if NOT noshipdateestimate then
									Set t2=e.getElementsByTagName("SvcCommitments")
									if t2.length>0 then
										Set s2=t2.Item(0)
										intShipping(1, l)=replace(s2.firstChild.nodeValue&""," to many major markets","")
									end if
								end if
								Set t2=e.getElementsByTagName("ServiceErrors")
								if t2.length>0 then serviceerror=TRUE
								if NOT serviceerror then
									Set t2=e.getElementsByTagName("Postage")
									Set s2=t2.Item(0)
									intShipping(2, l)=intShipping(2, l) + (s2.firstChild.nodeValue * quantity)
									intShipping(3, l)=intShipping(3, l) + 1
									for index2=0 to UBOUND(uspsmethods,2)
										if serviceid=uspsmethods(0,index2) then intShipping(0, l)=uspsmethods(2,index2) : wantthismethod=TRUE : exit for
									next
								end if
								if NOT wantthismethod then intShipping(3, l)=0
							else
								thisService=e.firstChild.nodeValue
							end if
					end Select
				end if
			next
			packCost=0
		next
	end if
	set xmlDoc=nothing
	if discountshippingusps<>"" then
		for uspsind=0 to UBOUND(uspsmethods,2)
			if intShipping(3, uspsind)>0 then intShipping(2, uspsind)=vsround(intShipping(2, uspsind)*(1+discountshippingusps/100.0),2)
		next
	end if
	ParseUSPSXMLOutput=noError
end function
function checkUPSShippingMeth(method, byRef discountsApply, byRef showAs)
	retval=FALSE
	discountsApply=0
	if isarray(uspsmethods) then
		for xx=0 to UBOUND(uspsmethods,2)
			if method=uspsmethods(0,xx) then
				retval=true
				discountsApply=uspsmethods(1,xx)
				showAs=uspsmethods(2,xx)
				exit for
			end if
		next
	end if
	checkUPSShippingMeth=retval
end function
function ParseUPSXMLOutput(xmlDoc, international, byRef errormsg, byRef errorcode, byRef intShipping)
Dim noError, nodeList, e, i, j, k, l, n, t, t2, indexus
	noError=TRUE
	indexus=0
	l=0
	errormsg=""
	if len(xmldoc.xml)<40 then
		noError=FALSE
		errormsg="Invalid Response From UPS Server"
	else
		Set t2=xmlDoc.getElementsByTagName("RatingServiceSelectionResponse").Item(0)
		for j=0 to t2.childNodes.length - 1
			Set n=t2.childNodes.Item(j)
			if n.nodename="Response" then
				for i=0 To n.childNodes.length - 1
					Set e=n.childNodes.Item(i)
					if e.nodeName="ResponseStatusCode" then
						noError=Int(e.firstChild.nodeValue)=1
					end if
					if e.nodeName="Error" then
						for k=0 To e.childNodes.length - 1
							Set t=e.childNodes.Item(k)
							Select Case t.nodeName
								Case "ErrorCode"
									errorcode=t.firstChild.nodeValue
								Case "ErrorSeverity"
									if t.firstChild.nodeValue="Transient" then errormsg=errormsg & "<div>This is a temporary error. Please wait a few moments then refresh this page.<br />" & errormsg & "</div>"
								Case "ErrorDescription"
									errormsg=errormsg & "<div>" & t.firstChild.nodeValue & "</div>"
							end Select
						next
					end if
				next
			elseif n.nodename="RatedShipment" then
				wantthismethod=true
				negotiatedrate=""
				for i=0 To n.childNodes.length - 1
					Set e=n.childNodes.Item(i)
					Select Case e.nodeName
						Case "Service"
							for k=0 To e.childNodes.length - 1
								Set t=e.childNodes.Item(k)
								if t.nodeName="Code" then
									Select Case cStr(t.firstChild.nodeValue)
										Case "01"
											intShipping(0, l)="UPS Next Day Air&reg;"
										Case "02"
											intShipping(0, l)="UPS 2nd Day Air&reg;"
										Case "03"
											intShipping(0, l)="UPS Ground"
										Case "07"
											intShipping(0, l)="UPS Worldwide Express&reg;"
										Case "08"
											intShipping(0, l)="UPS Worldwide Expedited&reg;"
										Case "11"
											intShipping(0, l)="UPS Standard"
										Case "12"
											intShipping(0, l)="UPS 3 Day Select&reg;"
										Case "13"
											intShipping(0, l)="UPS Next Day Air Saver&reg;"
										Case "14"
											intShipping(0, l)="UPS Next Day Air&reg; Early A.M.&reg;"
										Case "54"
											intShipping(0, l)="UPS Worldwide Express Plus&reg;"
										Case "59"
											intShipping(0, l)="UPS 2nd Day Air A.M.&reg;"
										Case "65"
											if origCountryCode="US" AND shipCountryCode<>"US" then
												intShipping(0, l)="UPS Worldwide Saver&reg;"
											else
												intShipping(0, l)="UPS Express Saver&reg;"
											end if
									end Select
									wantthismethod=checkUPSShippingMeth(t.firstChild.nodeValue, discntsApp, notUsed)
									intShipping(4, l)=discntsApp
								end if
							next
						Case "TotalCharges"
							for k=0 To e.childNodes.length - 1
								Set t=e.childNodes.Item(k)
								if t.nodeName="MonetaryValue" then intShipping(2, l)=cdbl(t.firstChild.nodeValue)
							next
						Case "ServiceOptionsCharges"
							for k=0 To e.childNodes.length - 1
								Set t=e.childNodes.Item(k)
								if t.nodeName="MonetaryValue" then intShipping(8, l)=cdbl(t.firstChild.nodeValue)
							next
						Case "GuaranteedDaysToDelivery"
							if e.childNodes.length>0 AND NOT noshipdateestimate then
								if e.firstChild.nodeValue="1" then
									intShipping(1, l)="1 " & xxDay & intShipping(1, l)
								else
									intShipping(1, l)=e.firstChild.nodeValue & " " & xxDays & intShipping(1, l)
								end if
							end if
						Case "ScheduledDeliveryTime"
							if e.childNodes.length>0 AND NOT noshipdateestimate then intShipping(1, l)=intShipping(1, l) & " by " & e.firstChild.nodeValue
						Case "NegotiatedRates"
							if e.childNodes.length>0 then
								set obj3=e.childNodes.Item(0).getElementsByTagName("MonetaryValue")
								if obj3.length>0 then
									negotiatedrate=obj3.item(0).firstChild.nodeValue
								end if
							end if
						Case "RatedShipmentWarning"
							if e.childNodes.length>0 then
								if instr(e.firstChild.nodeValue, "Commercial to Residential") then
									commercialloc_=FALSE
									if (ordComLoc AND 1)=1 then ordComLoc=ordComLoc-1
								end if
							end if
					end select
				next
				if negotiatedrate<>"" AND upsnegdrates=TRUE then intShipping(2, l)=cdbl(negotiatedrate)
				if wantthismethod=true then 
					if discountshippingups<>"" then intShipping(2, l)=vsround(intShipping(2, l)*(1+discountshippingups/100.0),2)
					intShipping(3, l)=TRUE
					l=l + 1
				else
					intShipping(1, l)=""
				end if
				wantthismethod=true
			end if
		next
	end if
	ParseUPSXMLOutput=noError
end function
function ParseCanadaPostXMLOutput(sXML,international, byRef errormsg,byRef errorcode,byRef intShipping)
Dim noError, nodeList, e, i, j, k, l, n, t, t2, indexus
	noError=TRUE
	indexus=0
	l=0
	cphandlingcharge=0
	errormsg=""
	set xmlDoc=Server.CreateObject("MSXML2.DOMDocument")
	xmlDoc.validateOnParse=FALSE
	xmlDoc.loadXML(sXML)
	Set t2=xmlDoc.getElementsByTagName("faultstring")
	if t2.length>0 then
		if t2.item(0).hasChildNodes then
			errormsg=t2.Item(0).firstChild.nodeValue
			if instr(errormsg,"}PostalCodeType")>0 OR instr(errormsg,"}ZipCodeType")>0 then
				if storelang="fr" then errormsg="Code Postal Invalide: " & destZip else errormsg="Invalid Postal Code: " & destZip
			end if
			noError=FALSE
		end if
	else
		Set t2=xmlDoc.getElementsByTagName("price-quote")
		for j=0 to t2.length - 1
			Set n=t2.Item(j)
			for i=0 To n.childNodes.length - 1
				Set e=n.childNodes.Item(i)
				if e.nodeName="service-code" then
					wantthismethod=checkUPSShippingMeth(e.firstChild.nodeValue, discntsApp, notUsed)
					intShipping(4, l)=discntsApp
				elseif e.nodeName="service-name" then
					intShipping(0, l)=e.firstChild.nodeValue
				elseif e.nodeName="price-details" then
					for k=0 To e.childNodes.length - 1
						Set ee=e.childNodes.Item(k)
						if ee.nodeName="due" then
							intShipping(2, l)=cdbl(ee.firstChild.nodeValue)
						elseif ee.nodeName="options" then
							for ctr2=0 To ee.childNodes.length - 1
								set eee=ee.childNodes.Item(ctr2)
								select Case eee.nodeName
								case "option"
									thisserviceid=""
									for ctr3=0 To eee.childNodes.length - 1
										set eeee=eee.childNodes.Item(ctr3)
										select Case eeee.nodeName
										case "option-code"
											thisserviceid=eeee.firstChild.nodeValue
										case "option-price"
											thisprice=eeee.firstChild.nodeValue
										end select
									next
									if thisserviceid="COV" then intShipping(8, l)=cdbl(thisprice)
								end select
							next
						end if
					next
				elseif e.nodeName="service-standard" AND NOT noshipdateestimate then
					set ep=e.getElementsByTagName("expected-delivery-date")
					if ep.length>0 then
						if IsDate(ep.item(0).firstChild.nodeValue) then
							numdays=DateValue(ep.item(0).firstChild.nodeValue) - Date()
							intShipping(1, l)=numdays & " " & IIfVr(numdays<2,xxDay,xxDays) & intShipping(1, l)
						else
							intShipping(1, l)=ep.item(0).firstChild.nodeValue & intShipping(1, l)
						end if
					end if
				end if
			next
			if wantthismethod=true then
				if discountshippingcanadapost<>"" then intShipping(2, l)=vsround(intShipping(2, l)*(1+discountshippingcanadapost/100.0),2)
				intShipping(3, l)=TRUE
				l=l + 1
			else
				intShipping(1, l)=""
			end if
			wantthismethod=true
		next
	end if
	ParseCanadaPostXMLOutput=noError
end function
function getuspscontainer(gpcweight,theservice)
	ispriority=theservice="PRIORITY"
	isexpress=theservice="EXPRESS"
	getuspscontainer=""
	if ispriority AND gpcweight<=70 AND (packdims(0)<=12.25 AND packdims(1)<=12.25 AND packdims(2)<=6) then getuspscontainer="lg flat rate box"
	if ispriority AND gpcweight<=70 AND ((packdims(0)<=11 AND packdims(1)<=8.5 AND packdims(2)<=5.5) OR (packdims(0)<=13.625 AND packdims(1)<=11.875 AND packdims(2)<=3.375)) then getuspscontainer="md flat rate box"
	if ispriority AND gpcweight<=70 AND (packdims(0)<=8.625 AND packdims(1)<=5.375 AND packdims(2)<=1.625) then getuspscontainer="sm flat rate box"
	if gpcweight<=70 AND (packdims(0)<=12.5 AND packdims(1)<=9.5 AND packdims(2)<=1) then getuspscontainer="flat rate envelope"
	if packdims(0)<=0 OR packdims(1)<=0 OR packdims(2)<=0 then getuspscontainer=""
end function
function addUSPSDomestic(id,service,orig,dest,iWeight,quantity,size,machinable,packcost)
	Dim sXML
	iTotItems=iTotItems + 1
	sXML=""
	pounds=int(iWeight)
	ounces=vsceil((iWeight-pounds)*16)
	if pounds=0 AND ounces=0 then ounces=1
	if IsArray(uspsmethods) then
		if (adminUnits AND 12)<>0 then
			totaldims=packdims(0) + (2 * (packdims(1) + packdims(2)))
			if totaldims>84 then size="LARGE"
			if totaldims>108 then size="OVERSIZE"
		end if
		for indexus=0 TO UBOUND(uspsmethods,2)
			packsize=size
			if uspsmethods(0,indexus)<>"" then
				sXML=sXML & "<Package ID="""&replace(uspsmethods(0,indexus)," ","-")&"xx"&id&"xx"&quantity&""">"
				sXML=sXML & "<Service>"&uspsmethods(0,indexus)&"</Service>"
				if uspsmethods(0,indexus)="FIRST CLASS" OR uspsmethods(0,indexus)="First Class Commercial" then sXML=sXML & "<FirstClassMailType>"&IIfVr(firstclassmailtype<>"",firstclassmailtype,"PARCEL")&"</FirstClassMailType>"
				sXML=sXML & "<ZipOrigination>"&orig&"</ZipOrigination><ZipDestination>"&left(dest, 5)&"</ZipDestination>"
				sXML=sXML & "<Pounds>"&pounds&"</Pounds><Ounces>"&ounces&"</Ounces>"
				thecontainer="VARIABLE"
				if uspsprioritycontainer="flat rate box" then uspsprioritycontainer="md flat rate box"
				if instr(uspsexpresscontainer,"flat rate box")>0 then uspsexpresscontainer=""
				tempcontainer=uspsprioritycontainer
				if uspsmethods(0,indexus)="PRIORITY" then
					if (adminUnits AND 12)<>0 then
						if (packdims(0) * packdims(1) * packdims(2))>1728 AND packsize="REGULAR" then packsize="LARGE"
						if tempcontainer<>"" then
							tempcontainer=getuspscontainer(iWeight,"PRIORITY")
							if uspsprioritycontainer<>"auto" then
								if tempcontainer="" OR (uspsprioritycontainer="md flat rate box" AND tempcontainer="lg flat rate box") OR (uspsprioritycontainer="sm flat rate box" AND (tempcontainer="lg flat rate box" OR tempcontainer="md flat rate box")) OR (uspsprioritycontainer="flat rate envelope" AND (tempcontainer="lg flat rate box" OR tempcontainer="md flat rate box" OR tempcontainer="sm flat rate box")) then
									uspsmethods(0,indexus)=""
								else
									tempcontainer=uspsprioritycontainer
								end if
							end if
						end if
					end if
					if tempcontainer="" OR tempcontainer="auto" then thecontainer=IIfVr(packsize="LARGE","rectangular","") else thecontainer=tempcontainer
				end if
				tempcontainer=uspsexpresscontainer
				if uspsmethods(0,indexus)="EXPRESS" then
					if (adminUnits AND 12)<>0 AND tempcontainer<>"" then tempcontainer=getuspscontainer(iWeight,"EXPRESS")
					if uspsexpresscontainer<>"auto" then
						if uspsexpresscontainer="flat rate envelope" AND tempcontainer="flat rate box" then uspsmethods(0,indexus)="" else tempcontainer=uspsexpresscontainer
					end if
					if tempcontainer="" OR tempcontainer="auto" then thecontainer="" else thecontainer=tempcontainer
				end if
				sXML=sXML & "<Container>"&thecontainer&"</Container><Size>"&packsize&"</Size>"
				if (adminUnits AND 12)<>0 AND packdims(0)>0 AND packdims(1)>0 AND packdims(2)>0 then sXML=sXML & "<Width>" & vsround(packdims(1),1) & "</Width><Length>" & vsround(packdims(0),1) & "</Length><Height>" & vsround(packdims(2),1) & "</Height>"
				sXML=sXML & "<Value>" & -int(-packcost) & "</Value>"
				sXML=sXML & "<Machinable>"&machinable&"</Machinable></Package>"
			end if
		next
	end if
	addUSPSDomestic=sXML
end function
function doesfitinbox(blen,bwid,bhei)
	doesfitinbox=TRUE
	if packdims(0)>blen OR packdims(1)>bwid OR packdims(2)>bhei then
		doesfitinbox=FALSE
	end if
	if NOT doesfitinbox AND packdims(7)>=3 then
		if packdims(4)<=blen AND packdims(5)<=bwid AND packdims(6)<=bhei AND packdims(4)<=(blen*bwid*bhei) then
			doesfitinbox=TRUE
		end if
	end if
end function
function addUSPSInternational(id,iWeight,quantity,mailtype,country,packcost)
	Dim sXML
	iTotItems=iTotItems + 1
	if (adminUnits AND 12)<>0 then
		if isarray(uspsmethods) then
			lenplusgirth=packdims(0) + (2 * (packdims(1) + packdims(2)))
			for xx=0 to UBOUND(uspsmethods,2)
				if shipCountryCode="AD" OR shipCountryCode="AT" OR shipCountryCode="BE" OR shipCountryCode="CH" OR shipCountryCode="CN" OR shipCountryCode="CZ" OR shipCountryCode="DE" OR shipCountryCode="DK" OR shipCountryCode="ES" OR shipCountryCode="FI" OR shipCountryCode="FR" OR shipCountryCode="GR" OR shipCountryCode="HK" OR shipCountryCode="IE" OR shipCountryCode="IT" OR shipCountryCode="JP" OR shipCountryCode="LI" OR shipCountryCode="LU" OR shipCountryCode="MC" OR shipCountryCode="MT" OR shipCountryCode="NL" OR shipCountryCode="NO" OR shipCountryCode="PT" OR shipCountryCode="SE" OR shipCountryCode="VA" then
					if packdims(0)>60 OR lenplusgirth>108 then ' Express Mail
						if uspsmethods(0,xx)="1" then uspsmethods(0,xx)="xxx"
					end if
				elseif shipCountryCode="CA" then
					if packdims(0)>42 OR lenplusgirth>79 then
						if uspsmethods(0,xx)="1" then uspsmethods(0,xx)="xxx"
					end if
				else
					if packdims(0)>36 OR lenplusgirth>79 then
						if uspsmethods(0,xx)="1" then uspsmethods(0,xx)="xxx"
					end if
				end if
				if shipCountryCode="CA" OR shipCountryCode="HK" then ' Priority Mail
					if lenplusgirth>108 then
						if uspsmethods(0,xx)="2" then uspsmethods(0,xx)="xxx"
					end if
				elseif shipCountryCode="AD" OR shipCountryCode="AT" OR shipCountryCode="BE" OR shipCountryCode="CH" OR shipCountryCode="CZ" OR shipCountryCode="DE" OR shipCountryCode="DK" OR shipCountryCode="ES" OR shipCountryCode="FI" OR shipCountryCode="FR" OR shipCountryCode="GI" OR shipCountryCode="GB" OR shipCountryCode="GR" OR shipCountryCode="IE" OR shipCountryCode="IT" OR shipCountryCode="JP" OR shipCountryCode="LI" OR shipCountryCode="LU" OR shipCountryCode="MC" OR shipCountryCode="MT" OR shipCountryCode="NL" OR shipCountryCode="NO" OR shipCountryCode="NZ" OR shipCountryCode="PL" OR shipCountryCode="PT" OR shipCountryCode="SE" OR shipCountryCode="VA" then
					if packdims(0)>60 OR lenplusgirth>108 then
						if uspsmethods(0,xx)="2" then uspsmethods(0,xx)="xxx"
					end if
				else
					if packdims(0)>42 OR lenplusgirth>79 then
						if uspsmethods(0,xx)="2" then uspsmethods(0,xx)="xxx"
					end if
				end if
				if iWeight>70 OR packdims(0)>46 OR packdims(1)>46 OR packdims(2)>35 OR lenplusgirth>108 then
					if uspsmethods(0,xx)="4" OR uspsmethods(0,xx)="6" OR uspsmethods(0,xx)="7" then uspsmethods(0,xx)="xxx" ' GXG
				end if
				if uspsmethods(0,xx)="24" then if iWeight>4 OR NOT doesfitinbox(7.5625,5.4375,0.625) then uspsmethods(0,xx)="xxx" ' DVD FRB
				if uspsmethods(0,xx)="16" then if iWeight>4 OR NOT doesfitinbox(8.625, 5.375, 1.625) then uspsmethods(0,xx)="xxx" ' Small FRB
				if uspsmethods(0,xx)="20" then if iWeight>4 OR NOT doesfitinbox(10,6,0.75) then uspsmethods(0,xx)="xxx" ' Small FRE
				if uspsmethods(0,xx)= "9" OR uspsmethods(0,xx)= "26" then if iWeight>20 OR (NOT doesfitinbox(11,8.5,5.5) AND NOT doesfitinbox(13.625,11.875,3.375)) then uspsmethods(0,xx)="xxx" ' FRB
				if uspsmethods(0,xx)="13" then if iWeight>4 OR NOT doesfitinbox(11.5,6.125,0.25) then uspsmethods(0,xx)="xxx" ' FirstClass Letter
				if uspsmethods(0,xx)="11" then if iWeight>20 OR NOT doesfitinbox(12,12,5.5) then uspsmethods(0,xx)="xxx" ' LFRB
				if uspsmethods(0,xx)="8" OR uspsmethods(0,xx)="10" then if iWeight>4 OR NOT doesfitinbox(12.5,9.5,1) then uspsmethods(0,xx)="xxx" ' FRE
				if uspsmethods(0,xx)="17" then if iWeight>4 OR NOT doesfitinbox(15,9.5,0.75) then uspsmethods(0,xx)="xxx" ' Legal FRE
				if uspsmethods(0,xx)="14" then if iWeight>4 OR NOT doesfitinbox(15,12,0.75) then uspsmethods(0,xx)="xxx" ' FirstClass L-E
				if packdims(0)>24 OR (packdims(0)+packdims(1)+packdims(2))>36 then
					if uspsmethods(0,xx)="15" then uspsmethods(0,xx)="xxx" ' FirstClass Package
				end if
			next
		end if
	end if
	pounds=int(iWeight)
	ounces=vsceil((iWeight-pounds)*16)
	if pounds=0 AND ounces=0 then ounces=1
	sXML="<Package ID=""xx"&id&"xx"&quantity&"""><Pounds>"&pounds&"</Pounds><Ounces>"&ounces&"</Ounces><MailType>ALL</MailType>"
	sXML=sXML & "<GXG><POBoxFlag>N</POBoxFlag><GiftFlag>N</GiftFlag></GXG>"
	sXML=sXML & "<ValueOfContents>" & -int(-packcost) & "</ValueOfContents>"
	sXML=sXML & "<Country>"&country&"</Country><Container>RECTANGULAR</Container><Size>REGULAR</Size>"
	if (adminUnits AND 12)<>0 AND -int(-packdims(0))>0 AND -int(-packdims(1))>0 AND -int(-packdims(2))>0 then sXML=sXML & "<Width>" & vsround(vrmax(packdims(2),6),2) & "</Width><Length>" & vsround(vrmax(packdims(0),4),2) & "</Length><Height>" & vsround(packdims(1),2) & "</Height><Girth>" & -int(-((packdims(1)*2)+(packdims(2)*2))) & "</Girth>" else sXML=sXML & "<Width>6</Width><Length>4</Length><Height>0.1</Height><Girth></Girth>"
	sXML=sXML & "<OriginZip>"&origZip&"</OriginZip>"
	addUSPSInternational=sXML & "<CommercialFlag>N</CommercialFlag></Package>"
end function
function addUPSInternational(iWeight,adminUnits,packTypeCode,country,packcost,dimens)
	Dim sXML
	if iWeight < 0.1 then iWeight=0.1
	sXML="<Package><PackagingType><Code>"&packTypeCode&"</Code><Description>Package</Description></PackagingType>"
	if dimens(0)>0 AND dimens(1)>0 AND dimens(2)>0 then sXML=sXML & "<Dimensions><Length>" & vsround(dimens(0),0) & "</Length><Width>" & vsround(dimens(1),0) & "</Width><Height>" & vsround(dimens(2),0) & "</Height><UnitOfMeasurement><Code>"&IIfVr((adminUnits AND 12)=4,"IN","CM")&"</Code></UnitOfMeasurement></Dimensions>"
	sXML=sXML & "<Description>Rate Shopping</Description><PackageWeight><UnitOfMeasurement><Code>"&IIfVr((adminUnits AND 1)=1,"LBS","KGS")&"</Code></UnitOfMeasurement><Weight>"&iWeight&"</Weight></PackageWeight><PackageServiceOptions>"
	if wantcarrierinsurance() then
		if packcost>50000 then packcost=50000
		sXML=sXML & "<InsuredValue><CurrencyCode>" & countryCurrency & "</CurrencyCode><MonetaryValue>" & FormatNumber(packcost,2,-1,0,0) & "</MonetaryValue></InsuredValue>"
	end if
	if ordPayProvider<>"" then
		if int(ordPayProvider)=codpaymentprovider then sXML=sXML & "<COD><CODFundsCode>0</CODFundsCode><CODCode>3</CODCode><CODAmount><CurrencyCode>"&countryCurrency&"</CurrencyCode><MonetaryValue>" & FormatNumber(packcost,2,-1,0,0) & "</MonetaryValue></CODAmount></COD>"
	end if
	if signatureoption="indirect" then
		sXML=sXML & "<DeliveryConfirmation><DCISType>1</DCISType></DeliveryConfirmation>"
	elseif signatureoption="direct" then
		sXML=sXML & "<DeliveryConfirmation><DCISType>2</DCISType></DeliveryConfirmation>"
	elseif signatureoption="adult" then
		sXML=sXML & "<DeliveryConfirmation><DCISType>3</DCISType></DeliveryConfirmation>"
	end if
	addUPSInternational=sXML & "</PackageServiceOptions></Package>"
end function
function addDHLPackage(iWeight,adminUnits,packTypeCode,country,packcost,dimens)
	addDHLPackage="<Piece><PieceID>" & packnumber & "</PieceID>"
	if dimens(0)>0 AND dimens(1)>0 AND dimens(2)>0 then addDHLPackage=addDHLPackage & "<Height>" & vsround(dimens(0),0) & "</Height><Depth>" & vsround(dimens(1),0) & "</Depth><Width>" & vsround(dimens(2),0) & "</Width>"
	addDHLPackage=addDHLPackage & "<Weight>" & vsround(iWeight,2) & "</Weight></Piece>"
	packnumber=packnumber+1
end function
function dhlcalculate(sXML,international,byRef errormsg,byRef intShipping)
	if destZip="" AND NOT zipisoptional(shipCountryID) then
		errormsg=xxPlsZip
		dhlcalculate=FALSE
	elseif callxmlfunction("https://xmlpi" & IIfVs(upstestmode,"test") & "-ea.dhl.com/XMLShippingServlet", sXML, xmlres, "", "Msxml2.ServerXMLHTTP", errormsg, FALSE) then
		if dumpshippingxml then call dumpxmlout(sXML,xmlres)
		saveLCID=Session.LCID
		Session.LCID=1033
		dhlcalculate=ParseDHLXMLOutput(xmlres,international,errormsg,errorcode,intShipping)
		sortshippingarray()
		Session.LCID=saveLCID
		set xmlres=nothing
	else
		errormsg=getshippingerror(errormsg)
		dhlcalculate=FALSE
	end if
end function
function addCanadaPostPackage(iWeight,adminUnits,packTypeCode,country,packcost,dimens)
	if iWeight < 0.1 then iWeight=0.1
	tmpXML=""
	if wantcarrierinsurance() then
		tmpXML=tmpXML & "<options><option><option-code>COV</option-code><option-amount>" & -int(-packcost) & "</option-amount></option></options>"
	end if
	tmpXML=tmpXML & "<parcel-characteristics><weight>" & vsround(IIfVr((adminUnits AND 1)=1, iWeight * 0.453592, iWeight),3) & "</weight>"
	if (adminUnits AND 12)=4 then dimens(0)=dimens(0)*2.54 : dimens(1)=dimens(1)*2.54 : dimens(2)=dimens(2)*2.54
	if dimens(0)>0 AND dimens(1)>0 AND dimens(2)>0 then tmpXML=tmpXML & "<dimensions><length>" & vsround(dimens(0),1) & "</length><width>" & vsround(dimens(1),1) & "</width><height>" & vsround(dimens(2),1) & "</height></dimensions>"
	tmpXML=tmpXML & "</parcel-characteristics>"
	addCanadaPostPackage=tmpXML
end function
function addFedexPackage(iWeight,packcost,dimens)
	Session.LCID=1033
	tmpXML="<v9:RequestedPackageLineItems>"
	if iWeight<0.1 then iWeight=0.1
	if shipType=8 AND iWeight<1 then iWeight=0.1
	if shipType<>8 AND wantcarrierinsurance() then
		tmpXML=tmpXML & "<v9:InsuredValue><v9:Currency>" & countryCurrency & "</v9:Currency><v9:Amount>" & FormatNumber(packcost,2,-1,0,0) & "</v9:Amount></v9:InsuredValue>"
	end if
	tmpXML=tmpXML & "<v9:Weight><v9:Units>"&IIfVr((adminUnits AND 1)=1,"LB","KG")&"</v9:Units><v9:Value>"&formatnumber(iWeight,1,-1,0,0)&"</v9:Value></v9:Weight>"
	if dimens(0)>0 AND dimens(1)>0 AND dimens(2)>0 then
		if shipType=8 then
			if (adminUnits AND 12)=4 then dimens(0)=vrmax(dimens(0),6) : dimens(1)=vrmax(dimens(1),4) : dimens(2)=vrmax(dimens(2),1) else dimens(0)=vrmax(dimens(0),15) : dimens(1)=vrmax(dimens(1),10) : dimens(2)=vrmax(dimens(2),3)
		end if
		tmpXML=tmpXML & "<v9:Dimensions><v9:Length>" & vsround(dimens(0),0) & "</v9:Length><v9:Width>" & vsround(dimens(1),0) & "</v9:Width><v9:Height>" & vsround(dimens(2),0) & "</v9:Height><v9:Units>"&IIfVr((adminUnits AND 12)=4,"IN","CM")&"</v9:Units></v9:Dimensions>"
	end if
	if packaging<>"" AND shipType=8 then tmpXML=tmpXML & "<v9:PhysicalPackaging>"&UCASE(packaging)&"</v9:PhysicalPackaging>"
	tmpXML=tmpXML & "<v9:SpecialServicesRequested>"
	if signaturerelease_ AND allowsignaturerelease=TRUE then
	elseif signatureoption="indirect" then
		tmpXML=tmpXML & "<v9:SpecialServiceTypes>SIGNATURE_OPTION</v9:SpecialServiceTypes>"
	elseif signatureoption="direct" then
		tmpXML=tmpXML & "<v9:SpecialServiceTypes>SIGNATURE_OPTION</v9:SpecialServiceTypes>"
	elseif signatureoption="adult" then
		tmpXML=tmpXML & "<v9:SpecialServiceTypes>SIGNATURE_OPTION</v9:SpecialServiceTypes>"
	elseif signatureoption="none" then
		tmpXML=tmpXML & "<v9:SpecialServiceTypes>SIGNATURE_OPTION</v9:SpecialServiceTypes>"
	end if
	if nonstandardcontainer=TRUE then tmpXML=tmpXML & "<v9:SpecialServiceTypes>NON_STANDARD_CONTAINER</v9:SpecialServiceTypes>"
	if dryice=TRUE then tmpXML=tmpXML & "<v9:SpecialServiceTypes>DRY_ICE</v9:SpecialServiceTypes><v9:DryIceWeight><v9:Units>KG</v9:Units><v9:Value>5</v9:Value></v9:DryIceWeight>"
	if dangerousgoods=TRUE then tmpXML=tmpXML & "<v9:SpecialServiceTypes>DANGEROUS_GOODS</v9:SpecialServiceTypes><v9:DangerousGoodsDetail><v9:Accessibility>ACCESSIBLE</v9:Accessibility><v9:CargoAircraftOnly>1</v9:CargoAircraftOnly></v9:DangerousGoodsDetail>"
	if ordPayProvider<>"" then
		if int(ordPayProvider)=codpaymentprovider then tmpXML=tmpXML & "<v9:SpecialServiceTypes>COD</v9:SpecialServiceTypes><v9:CodDetail><v9:CodCollectionAmount><v9:Currency>CAD</v9:Currency><v9:Amount>XXXFEDEXGRANDTOTXXX</v9:Amount></v9:CodCollectionAmount><v9:CollectionType>ANY</v9:CollectionType></v9:CodDetail>"
	end if
	if signaturerelease_ AND allowsignaturerelease=TRUE then
	elseif signatureoption="indirect" then
		tmpXML=tmpXML & "<v9:SignatureOptionDetail><v9:OptionType>INDIRECT</v9:OptionType></v9:SignatureOptionDetail>"
	elseif signatureoption="direct" then
		tmpXML=tmpXML & "<v9:SignatureOptionDetail><v9:OptionType>DIRECT</v9:OptionType></v9:SignatureOptionDetail>"
	elseif signatureoption="adult" then
		tmpXML=tmpXML & "<v9:SignatureOptionDetail><v9:OptionType>ADULT</v9:OptionType></v9:SignatureOptionDetail>"
	elseif signatureoption="none" then
		tmpXML=tmpXML & "<v9:SignatureOptionDetail><v9:OptionType>NO_SIGNATURE_REQUIRED</v9:OptionType></v9:SignatureOptionDetail>"
	end if
	tmpXML=tmpXML & "</v9:SpecialServicesRequested>"
	addFedexPackage=tmpXML & "</v9:RequestedPackageLineItems>"
	Session.LCID=saveLCID
end function
function USPSCalculate(sXML,international,byRef errormsg,byRef intShipping)
	if destZip="" AND NOT zipisoptional(shipCountryID) then
		errormsg=xxPlsZip
		USPSCalculate=FALSE
	elseif callxmlfunction("https://" & IIfVs(debugmode,"stg-") & "production.shippingapis.com/ShippingAPI.dll", "API="&international&"Rate"&IIfVr(international="","V4","V2")&"&XML=" & urlencode(sXML), xmlres, "", "Msxml2.ServerXMLHTTP", errormsg, FALSE) then
		if dumpshippingxml then call dumpxmlout(sXML,xmlres)
		Session.LCID=1033
		USPSCalculate=ParseUSPSXMLOutput(xmlres,international,errormsg,intShipping)
		for ind1=0 to UBOUND(intShipping,2)
			for ind2=ind1+1 to UBOUND(intShipping,2)
				if intShipping(3,ind1)>0 AND intShipping(3,ind2)>0 AND intShipping(5,ind1)=intShipping(5,ind2) AND intShipping(5,ind1)<>"" then
					if cdbl(intShipping(2,ind1))<cdbl(intShipping(2,ind2)) then intShipping(3,ind2)=0 else intShipping(3,ind1)=0
				end if
			next
		next
		sortshippingarray()
		Session.LCID=saveLCID
	else
		errormsg=getshippingerror(errormsg)
		USPSCalculate=FALSE
	end if
end function
function UPSCalculate(sXML,international,byRef errormsg,byRef intShipping)
	xmlres="xml"
	if upstestmode=TRUE then print "UPS Test Mode<br />" : upsurl="wwwcie.ups.com" else upsurl="onlinetools.ups.com"
	if destZip="" AND NOT zipisoptional(shipCountryID) then
		errormsg=xxPlsZip
		UPSCalculate=FALSE
	elseif callxmlfunction("https://"&upsurl&"/ups.app/xml/Rate", sXML, xmlres, "", "Msxml2.ServerXMLHTTP", errormsg, FALSE) then
		if dumpshippingxml then call dumpxmlout(sXML,xmlres.xml)
		saveLCID=Session.LCID
		Session.LCID=1033
		UPSCalculate=ParseUPSXMLOutput(xmlres,international,errormsg,errorcode,intShipping)
		sortshippingarray()
		if errorcode=110971 then errormsg="" ' May differ from published rates.
		if errorcode=111210 then errormsg=xxInvZip
		if errorcode=119070 then errormsg="" ' Large package surcharge.
		'if errorcode=120900 then errormsg=""
		'if errorcode=120901 then errormsg=""
		Session.LCID=saveLCID
		set xmlres=nothing
	else
		errormsg=getshippingerror(errormsg)
		UPSCalculate=FALSE
	end if
end function
Dim iscanadapost
function CanadaPostCalculate(sXML,international,byRef errormsg,byRef intShipping)
	if destZip="" AND NOT zipisoptional(shipCountryID) then
		errormsg=xxPlsZip
		CanadaPostCalculate=FALSE
	else
		iscanadapost=TRUE
		CanadaPostCalculate=callxmlfunction("https://" & IIfVr(canadaposttestmode,"ct.","") & "soa-gw.canadapost.ca/rs/soap/rating/v2",sXML,xmlres,"","WinHTTP.WinHTTPRequest.5.1",errormsg,FALSE)
		iscanadapost=FALSE
		if CanadaPostCalculate then
			saveLCID=Session.LCID
			Session.LCID=1033
			if dumpshippingxml then call dumpxmlout(sXML,xmlres)
			CanadaPostCalculate=ParseCanadaPostXMLOutput(xmlres,international,errormsg,errorcode,intShipping)
			sortshippingarray()
			Session.LCID=saveLCID
		else
			errormsg=getshippingerror(errormsg)
			CanadaPostCalculate=FALSE
		end if
		set xmlres=nothing
	end if
end function
function parsefedexXMLoutput(xmlDoc, international, byRef errormsg, byRef errorcode, byRef intShipping)
	noError=TRUE
	errormsg=""
	l=0
	fns=fedexnamespace
	Set t2=xmlDoc.getElementsByTagName(fns&"RateReply")
	if t2.length=0 then
		noError=FALSE
		set obj3=xmlDoc.getElementsByTagName(fns&"desc")
		if obj3.length>0 then errormsg=obj3.item(0).firstChild.nodeValue else errormsg="Unrecognized response"
	else
		for j=0 to t2.item(0).childNodes.length - 1
			Set n=t2.item(0).childNodes.Item(j)
			if n.nodename=fns&"HighestSeverity" then
				if n.firstChild.nodeValue="ERROR" then
					noError=FALSE
					for ind1=0 To n.childNodes.length - 1
						Set e=n.childNodes.Item(ind1)
						if e.nodeName=fns&"Message" then
							errormsg=errormsg & e.firstChild.nodeValue
						elseif e.nodeName=fns&"Code" then
							errorcode=e.firstChild.nodeValue
						end if
					next
				end if
			elseif n.nodename=fns&"Notifications" then
				iserror=FALSE
				themessage=""
				thecode=""
				for ind1=0 To n.childNodes.length - 1
					Set e=n.childNodes.Item(ind1)
					if e.nodeName=fns&"Message" then
						themessage=e.firstChild.nodeValue
					elseif e.nodeName=fns&"Code" then
						thecode=e.firstChild.nodeValue
					elseif e.nodeName=fns&"Severity" then
						if e.firstChild.nodeValue="ERROR" then iserror=TRUE
					end if
				next
				if iserror then
					errormsg=themessage
					errorcode=thecode
				end if
			elseif n.nodename=fns&"RateReplyDetails" then
				thisratetype=""
				wantthismethod=FALSE
				entryweight=0
				set objweight=n.getElementsByTagName("BilledWeight")
				if objweight.length>0 then
					entryweight=objweight.item(0).firstChild.nodeValue
				end if
				for ind1=0 To n.childNodes.length - 1
					Set e=n.childNodes.Item(ind1)
					if e.nodeName=fns&"ServiceType" then
						thisratetype=""
						theservicename=replace(e.firstChild.nodeValue,"_","")
						wantthismethod=checkUPSShippingMeth(theservicename, discntsApp, showAs)
						' if theservicename="FEDEXGROUND" AND shipCountryCode<>"CA" AND shipCountryCode<>"PR" AND NOT commercialloc_ AND entryweight<=70.0 then wantthismethod=FALSE
						if origCountryCode<>shipCountryCode then
							'if instr(showAs,"FedEx Ground")>0 AND nofedexinternationalground=TRUE then wantthismethod=FALSE
							showAs=replace(showAs, "FedEx Ground", "FedEx International Ground")
						end if
						if wantthismethod then
							intShipping(0, l)=showAs
							intShipping(4, l)=discntsApp
						end if
					elseif e.nodeName=fns&"RatedShipmentDetails" AND thisratetype<>"PAYOR_ACCOUNT_SHIPMENT" then
						for k9=0 To e.childNodes.length - 1
							Set f9=e.childNodes.Item(k9)
							if f9.nodeName=fns&"ShipmentRateDetail" then
								intShipping(2, l)=0
								for m=0 To f9.childNodes.length - 1
									Set g9=f9.childNodes.Item(m)
									if g9.nodeName=fns&"RateType" then
										thisratetype=g9.firstChild.nodeValue
									elseif g9.nodeName=fns&"TotalNetCharge" then
										intShipping(2, l)=intShipping(2, l) + cdbl(g9.getElementsByTagName(fns&"Amount").item(0).firstChild.nodeValue)
									elseif g9.nodeName=fns&"TotalSurcharges" then
										intShipping(8, l)=intShipping(8, l) + cdbl(g9.getElementsByTagName(fns&"Amount").item(0).firstChild.nodeValue)
									elseif g9.nodeName=fns&"TotalFreightDiscounts" then
										if uselistshippingrates=TRUE then intShipping(2, l)=intShipping(2, l) + cdbl(g9.getElementsByTagName(fns&"Amount").item(0).firstChild.nodeValue)
									end if
								next
							end if
						next
					elseif e.nodeName=fns&"DeliveryTimestamp" AND NOT noshipdateestimate then
						numdays=DateValue(replace(e.firstChild.nodeValue,"T"," ")) - Date()
						for fedwdayind=0 to numdays-1
							if weekday(Date()+fedwdayind,7)=1 OR weekday(Date()+fedwdayind,7)=2 then numdays=numdays-1
						next
						if numdays < 1 then numdays=1
						intShipping(1, l)=numdays & " " & IIfVr(numdays<2,xxDay,xxDays)
					elseif e.nodeName=fns&"TransitTime" AND NOT noshipdateestimate then
						if e.firstChild.nodeValue="ONE_DAY" then intShipping(1, l)="1 " & xxDay
						if e.firstChild.nodeValue="TWO_DAYS" then intShipping(1, l)="2 " & xxDays
						if e.firstChild.nodeValue="THREE_DAYS" then intShipping(1, l)="3 " & xxDays
						if e.firstChild.nodeValue="FOUR_DAYS" then intShipping(1, l)="4 " & xxDays
						if e.firstChild.nodeValue="FIVE_DAYS" then intShipping(1, l)="5 " & xxDays
						if e.firstChild.nodeValue="SIX_DAYS" then intShipping(1, l)="6 " & xxDays
						if e.firstChild.nodeValue="SEVEN_DAYS" then intShipping(1, l)="7 " & xxDays
					end if
				next
				if wantthismethod then
					if discountshippingfedex<>"" then intShipping(2, l)=vsround(intShipping(2, l)*(1+discountshippingfedex/100.0),2)
					intShipping(3, l)=TRUE
					l=l + 1
				end if
			end if
		next
	end if
	parsefedexXMLoutput=noError
end function
function fedexcalculate(sXML,international, byRef errormsg, byRef intShipping)
	if destZip="" AND NOT zipisoptional(shipCountryID) then
		errormsg=xxPlsZip
		fedexcalculate=FALSE
	else
		Session.LCID=1033
		xmlres="xml"
		success=callxmlfunction(fedexurl, sXML, xmlres, "", "Msxml2.ServerXMLHTTP", errormsg, FALSE)
		if dumpshippingxml then
			if success then call dumpxmlout(sXML,xmlres.xml) else call dumpxmlout(sXML,"")
		end if
		if success then
			set toregexp=new RegExp
			toregexp.pattern="<(.{1,3}):RateReply"
			set matches=toregexp.execute(xmlres.xml)
			set toregexp=nothing
			if matches.count>0 then fedexnamespace=matches(0).submatches(0)&":" else fedexnamespace=""
			success=parsefedexXMLoutput(xmlres, international, errormsg, errorcode, intShipping)
		end if
		if success then sortshippingarray() else errormsg=getshippingerror(errormsg)
		fedexcalculate=success
		Session.LCID=saveLCID
	end if
end function
function parseauspostXMLoutput(xmlDoc, international, byRef errormsg, byRef errorcode, byRef intShipping)
	noError=TRUE
	errormsg=""
	l=0
	Set t2=xmlDoc.getElementsByTagName("services")
	if t2.length=0 then
		noError=FALSE
		errormsg="Unrecognized response"
	else
		for j=0 to t2.item(0).childNodes.length - 1
			Set n=t2.item(0).childNodes.Item(j)
			if n.nodename="service" then
				wantthismethod=FALSE
				entryweight=0
				set objweight=n.getElementsByTagName("BilledWeight")
				if objweight.length>0 then
					entryweight=objweight.item(0).firstChild.nodeValue
				end if
				for ind1=0 To n.childNodes.length - 1
					Set e=n.childNodes.Item(ind1)
					'print e.nodename & " : " & e.firstChild.nodeValue & "<br>"
					if e.nodeName="code" then
						theservicename=e.firstChild.nodeValue
						wantthismethod=checkUPSShippingMeth(theservicename,discntsApp,showAs)
						if wantthismethod then
							intShipping(0, l)=showAs
							intShipping(4, l)=discntsApp
						end if
					elseif e.nodeName="price" then
						intShipping(2, l)=cdbl(e.firstChild.nodeValue)
					end if
				next
				if wantthismethod then
					if discountshippingauspost<>"" then intShipping(2, l)=vsround(intShipping(2, l)*(1+discountshippingauspost/100.0),2)
					intShipping(3, l)=TRUE
					l=l + 1
				end if
			end if
		next
	end if
	parseauspostXMLoutput=noError
end function
function auspostcalculate(appackweight,international, byRef errormsg, byRef intShipping)
	if international<>"" then
		sXML="international/service.xml?country_code="&shipCountryCode&"&weight="&appackweight
	else
		sXML="domestic/service.xml?from_postcode="&origZip&"&to_postcode="&destZip&"&weight="&appackweight&"&length=" & vrmax(1,vsround(packdims(0),1)) & "&width=" & vrmax(1,vsround(packdims(1),1)) & "&height=" & vrmax(1,vsround(packdims(2),1))
	end if
	xmlfnheaders=array(array("AUTH-KEY",AusPostAPI))
	theurl="https://digitalapi.auspost.com.au/postage/parcel/"&sXML
	xmlres="xml"
	if AusPostAPI="" then
		success=FALSE
		errormsg="You must set your Australia Post API Key"
	else
		success=callxmlfunction(theurl,"",xmlres,"","Msxml2.ServerXMLHTTP",errormsg,FALSE)
		if instr(errormsg," (404)")>0 then errormsg="Invalid Postal Code: " & destZip
	end if
	if success then
		success=parseauspostXMLoutput(xmlres,international,errormsg,errorcode,intShipping)
		if success then sortshippingarray()
	else
		errormsg=getshippingerror(errormsg)
	end if
	if dumpshippingxml then
		if success then call dumpxmlout(sXML,xmlres.xml) else call dumpxmlout(sXML,"")
	end if
	auspostcalculate=success
end function
%>