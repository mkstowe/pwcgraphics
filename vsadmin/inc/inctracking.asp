<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if sDSN="" then response.write "Database connection not set" : response.end
Dim iscanadapost
if request.totalbytes > 10000 then response.end
' ActivityList(0)=Address
' ActivityList(1)=SignedForByName
' ActivityList(2)=Not Used
' ActivityList(3)=Activity -> Status -> StatusType -> Code
' ActivityList(4)=Activity -> Status -> StatusType -> Description
' ActivityList(5)=Activity -> Status -> StatusCode -> Code
' ActivityList(6)=Activity -> Date
' ActivityList(7)=Activity -> Time
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs2=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
incupscopyright=FALSE
incfedexcopyright=FALSE
alternateratesusps=FALSE
alternateratesups=FALSE
alternateratesfedex=FALSE
alternateratescanadapost=FALSE
alternateratesdhl=FALSE
dim fedexnamespace
if adminAltRates>0 then
	sSQL="SELECT altrateid FROM alternaterates WHERE usealtmethod<>0 OR usealtmethodintl<>0"
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		if rs("altrateid")=3 then alternateratesusps=TRUE
		if rs("altrateid")=4 then alternateratesups=TRUE
		if rs("altrateid")=6 then alternateratescanadapost=TRUE
		if rs("altrateid")=7 OR rs("altrateid")=8 then alternateratesfedex=TRUE
		if rs("altrateid")=9 then alternateratesdhl=TRUE
		rs.movenext
	loop
	rs.close
end if
theshiptype=""
canadaposttrackurl="https://" & IIfVr(canadaposttestmode,"ct.","") & "soa-gw.canadapost.ca/vis/soap/track"
if getpost("trackno")<>"" then
	sSQL="SELECT adminPacking,adminCanPostUser,adminCanPostLogin,adminCanPostPass,adminUSPSUser,smartPostHub,adminUPSUser,adminUPSpw,adminUPSAccess,adminUPSAccount,adminUPSNegotiated,FedexAccountNo,FedexMeter,FedexUserKey,FedexUserPwd,DHLSiteID,DHLSitePW,DHLAccountNo FROM adminshipping WHERE adminShipID=1"
	rs.open sSQL,cnn,0,1
	adminCanPostLogin=trim(rs("adminCanPostLogin"))
	adminCanPostPass=trim(rs("adminCanPostPass"))
	adminCanPostUser=trim(rs("adminCanPostUser"))
	packtogether=int(rs("adminPacking"))=1
	uspsUser=rs("adminUSPSUser")
	smartPostHub=rs("smartPostHub")
	upsUser=upsdecode(rs("adminUPSUser"), "")
	upsPw=upsdecode(rs("adminUPSpw"), "")
	upsAccess=rs("adminUPSAccess")
	upsAccount=rs("adminUPSAccount")
	upsnegdrates=(cint(rs("adminUPSNegotiated"))<>0)
	fedexaccount=rs("FedexAccountNo")
	fedexmeter=rs("FedexMeter")
	fedexuserkey=rs("FedexUserKey")
	fedexuserpwd=rs("FedexUserPwd")
	DHLSiteID=rs("DHLSiteID")
	DHLSitePW=rs("DHLSitePW")
	DHLAccountNo=rs("DHLAccountNo")
	rs.close
end if
if request("carrier")<>"" then
	theshiptype=request("carrier")
else
	if trim(request("trackno"))<>"" then theshiptype=getcarrierfromtrack(request("trackno"),thelink)
	if theshiptype="" then
		possshiptypes=0
		if defaulttrackingcarrier<>"" then theshiptype=defaulttrackingcarrier else theshiptype="ups"
		if shipType=3 OR alternateratesusps OR instr(1, trackingcarriers, "usps", 1)>0 then
			theshiptype="usps"
			possshiptypes=possshiptypes+1
		end if
		if shipType=4 OR alternateratesups OR instr(1, trackingcarriers, "ups", 1)>0 then
			theshiptype="ups"
			incupscopyright=true
			possshiptypes=possshiptypes+1
		end if
		if shipType=6 OR alternateratescanadapost OR instr(1, trackingcarriers, "canadapost", 1)>0 then
			theshiptype="canadapost"
			possshiptypes=possshiptypes+1
		end if
		if shipType=7 OR shipType=8 OR alternateratesfedex OR instr(1, trackingcarriers, "fedex", 1)>0 then
			theshiptype="fedex"
			incfedexcopyright=true
			possshiptypes=possshiptypes+1
		end if
		if shipType=9 OR alternateratesdhl OR instr(1, trackingcarriers, "dhl", 1)>0 then
			theshiptype="dhl"
			possshiptypes=possshiptypes+1
		end if
		if possshiptypes>1 then theshiptype="undecided"
	end if
end if
%>
<script>
<!--
function viewlicense()
{
	var prnttext='<html><head><STYLE TYPE="text/css">A:link {COLOR: #333333; TEXT-DECORATION: none}A:visited {COLOR: #333333; TEXT-DECORATION: none}A:active {COLOR: #333333; TEXT-DECORATION: none}A:hover {COLOR: #f39000; TEXT-DECORATION: none}TD {FONT-FAMILY: Verdana;}P {FONT-FAMILY: Verdana;}HR {color: #637BAD;height: 1px;}</STYLE></head><body><table width="100%" border="0" cellspacing="1" cellpadding="3">\n';
	prnttext += '<tr><td colspan="2" align="center"><a href="javascript:window.close()"><strong>Close Window</strong></a></td></tr>';
	prnttext += '<tr><td width="40"><img src="images/upslogo.png"  alt="UPS" /></td><td><p><span style="font-size:16px;font-family:Verdana;font-weight:bold">UPS Tracking Terms and Conditions</span></p></td></tr>';
	prnttext += '<tr><td colspan="2"><p><span style="font-size:12px;font-family:Verdana">The UPS package tracking systems accessed via this Web Site (the &quot;Tracking Systems&quot;) and tracking information obtained through this Web Site (the &quot;Information&quot;) are the private property of UPS. UPS authorizes you to use the Tracking Systems solely to track shipments tendered by or for you to UPS for delivery and for no other purpose. Without limitation, you are not authorized to make the Information available on any web site or otherwise reproduce, distribute, copy, store, use or sell the Information for commercial gain without the express written consent of UPS. This is a personal service, thus your right to use the Tracking Systems or Information is non-assignable. Any access or use that is inconsistent with these terms is unauthorized and strictly prohibited.</span></p></td></tr>';
	prnttext += '<tr><td colspan="2" align="center"><hr /><span style="font-size:10px;font-family:Verdana"><%=replace(xxUPStm,"'","\'")%></span></td></tr>';
	prnttext += '<tr><td colspan="2" align="center">&nbsp;<br /><a href="javascript:window.close()"><strong>Close Window</strong></a></td></tr>';
	prnttext += '</table></body></'+'html>';
	var newwin=window.open("","viewlicense",'menubar=no, scrollbars=yes, width=500, height=420, directories=no,location=no,resizable=yes,status=no,toolbar=no');
	newwin.document.open();
	newwin.document.write(prnttext);
	newwin.document.close();
}
function checkaccept()
{
  if (document.trackform.agreeconds.checked==false)
  {
    alert("Please note: To track your package(s), you must accept the UPS Tracking Terms and Conditions by selecting the checkbox below.");
    return (false);
  }else{
	document.trackform.submit();
  }
  return (true);
}
//-->
</script>
<%
if theshiptype="laposte" then %>
	<form method="get" name="trackform" action="https://www.laposte.fr/outils/suivre-vos-envois" target="_blank">
		<div class="ectdiv ecttracking">
			<div class="ectdivcontainer">
				<div class="ectdivleft">Please enter your La Poste Tracking Number</div>
				<div class="ectdivright"><input type="text" size="30" name="trackno" value="<% print htmlspecials(request("trackno"))%>" /></div>
			</div>
			<div class="ectdivcontainer">
				<div class="ectdivleft">Show Activity</div>
				<div class="ectdivright"><select name="activity" size="1"><option value="LAST">Show Last Activity Only</option><option value="ALL"<% if getpost("activity")="ALL" then print " selected=""selected"""%>>Show All Activity</option></select></div>
			</div>
			<div class="ectdiv2column"><% print imageorsubmit(imgtrackpackage,"Track Package","trackpackage")%></div>
		</div>
	</form>
<%
elseif theshiptype="canadapost" then %>
	<form method="post" name="trackform" action="tracking<%=extension%>">
	<input type="hidden" name="carrier" value="canadapost" />
      <div class="ectdiv ecttracking">
		<div class="ectdivhead">
			<div class="trackinglogo"><img src="images/canadapost.gif" alt="CanadaPost" /></div>
			<div class="trackingtext">Canada Post<small>&reg;</small> Tracking Tool</div>
		</div>
<%		
function ParseCanadaPostTrackingOutput(sXML, byRef totActivity, byRef deliverydate, byRef serviceDesc, byRef packagecount, byRef shiptoaddress, byRef scheddeldate, byRef signedforby, byRef errormsg, byRef activityList)
	noError=TRUE
	totalCost=0
	packCost=0
	index=0
	errormsg=""
	gotxml=FALSE
	theaddress=""
	' 1111111332936901 1371134583769923
	set xmlDoc=Server.CreateObject("MSXML2.DOMDocument")
	xmlDoc.validateOnParse=False
	xmlDoc.loadXML(sXML)	
	set toregexp=new RegExp
	toregexp.pattern="<(.{3,7}):Body"
	set matches=toregexp.execute(sXML)
	set toregexp=nothing
	cpnamespace=matches(0).submatches(0)
	Set t2=xmlDoc.getElementsByTagName(cpnamespace & ":Body").Item(0)
	for j=0 to t2.childNodes.length - 1
		set n=t2.childNodes.Item(j)
		if n.nodename="soap:Fault" then
			noError=FALSE
			for i=0 To n.childNodes.length - 1
				set e=n.childNodes.Item(i)
				if e.nodeName="faultstring" then
					errormsg=e.firstChild.nodeValue
					if instr(errormsg,"element pin value")>0 then
						if storelang="fr" then errormsg="Num&eacute;ro de Rep&eacute;rage Invalide" else errormsg="Invalid Tracking Number"
					end if
				end if
			next
		elseif n.nodename="tns:get-tracking-detail-response" then
			for i=0 to n.childNodes.length-1
				set e=n.childNodes.Item(i)
				if e.nodeName="messages" then
					for k=0 to e.childNodes.length-1
						set t=e.childNodes.Item(k)
						if t.nodeName="message" then
							set obj2=t.getElementsByTagName("description")
							if obj2.length > 0 then
								if obj2.item(0).hasChildNodes then errormsg=obj2.item(0).firstChild.nodeValue
								noError=FALSE
							end if
						end if
					next
				elseif e.nodeName="tracking-detail" then
					for k=0 to e.childNodes.length - 1
						set t=e.childNodes.item(k)
						if t.nodeName="expected-delivery-date" then
						elseif t.nodeName="significant-events" then
							activityList(totActivity,0)=""
							for pj=0 to t.childNodes.length - 1
								set vj=t.childNodes.item(pj)
								if vj.nodeName="occurrence" then
									for vk=0 to vj.childNodes.length - 1
										set v=vj.childNodes.item(vk)
										if v.nodeName="event-date" then
											activityList(totActivity,6)=v.firstChild.nodeValue
										elseif v.nodeName="event-time" then
											if v.firstChild.nodeValue<>"00:00:00" then activityList(totActivity,7)=v.firstChild.nodeValue
										elseif v.nodeName="event-description" then
											activityList(totActivity,4)=v.firstChild.nodeValue
										elseif v.nodeName="event-site" then
											if v.hasChildNodes then activityList(totActivity,0)=v.firstChild.nodeValue
										elseif v.nodeName="event-province" then
											if v.hasChildNodes then activityList(totActivity,0)=activityList(totActivity,0) & ", " & v.firstChild.nodeValue
										end if
									next
									totActivity=totActivity + 1
								end if
							next
						end if
					next
				end if
			next
		end if
	next
	ParseCanadaPostTrackingOutput=noError
end function
function CanadaPostTrack(trackNo)
	Dim activityList(100,10)
	lastloc="xxxxxx"
	success=TRUE
	' (getpost("activity")="LAST" ? "false" : "true") . "</v4:IncludeDetailedScans>"
	sXML="<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:trac=""http://www.canadapost.ca/ws/soap/track"">" & _
		"<soapenv:Header><wsse:Security soapenv:mustUnderstand=""1"" xmlns:wsse=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"" xmlns:wsu=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd""><wsse:UsernameToken><wsse:Username>" & adminCanPostLogin & "</wsse:Username><wsse:Password>" & adminCanPostPass & "</wsse:Password></wsse:UsernameToken></wsse:Security></soapenv:Header>" & _
		"<soapenv:Body>" & _
		"<trac:get-tracking-detail-request><platform-id>0008107483</platform-id>" & IIfVs(storelang="fr","<locale>FR</locale>") & "<pin>" & replace(trackNo," ","") & "</pin>" & _
		"</trac:get-tracking-detail-request>" & _
		"</soapenv:Body></soapenv:Envelope>"
	iscanadapost=TRUE
	success=callxmlfunction(canadaposttrackurl, sXML, xmlres, "", "Msxml2.ServerXMLHTTP", errormsg, FALSE)
	iscanadapost=FALSE
	if success then
		Session.LCID=1033
		totActivity=0
		success=ParseCanadaPostTrackingOutput(xmlres, totActivity, deliverydate, serviceDesc, packagecount, shiptoaddress, scheduleddeliverydate, signedforby, errormsg, activityList)
		Session.LCID=saveLCID
		if success then
			for index2=0 to totActivity-2
				for index=0 to totActivity-2
					if DateValue(activityList(index,6)&" "&activityList(index,7))>DateValue(activityList(index+1,6)&" "&activityList(index+1,7)) then
						for index3=0 to UBOUND(activityList,2)
							tempArr=activityList(index,index3)
							activityList(index,index3)=activityList(index+1,index3)
							activityList(index+1,index3)=tempArr
						next
					end if
				next
			next
			if trim(serviceDesc)<>"" then %>
	  <div class="ectdivcontainer">
		<div class="ectdivleft">Service Description</div>
		<div class="ectdivright"><% print serviceDesc%></div>
	  </div>
<%			end if
			if trim(packagecount)<>"" then %>
	  <div class="ectdivcontainer">
		<div class="ectdivleft">Package Count</div>
		<div class="ectdivright"><% print packagecount%></div>
	  </div>
<%			end if
			if trim(shiptoaddress)<>"" then %>
	  <div class="ectdivcontainer">
		<div class="ectdivleft">Ship-To Address</div>
		<div class="ectdivright"><% print shiptoaddress%></div>
	  </div>
<%			end if
			if trim(signedforby)<>"" then %>
	  <div class="ectdivcontainer">
		<div class="ectdivleft">Signed For By</div>
		<div class="ectdivright"><% print signedforby%></div>
	  </div>
<%			end if
			if trim(deliverydate)<>"" then %>
	  <div class="ectdivcontainer">
		<div class="ectdivleft">Delivery Date</div>
		<div class="ectdivright"><% print deliverydate%></div>
	  </div>
<%			end if %>
			<div class="ecttrackingresults" style="display:table">
			  <div style="display:table-row" class="tracktablehead">
				<div style="display:table-cell">Location</div>
				<div style="display:table-cell">Description</div>
				<div style="display:table-cell">Date&nbsp;/&nbsp;Time</div>
			  </div>
<%			for index=0 to totActivity-1
				cellbg="class=""ect"&IIfVr(index MOD 2=0,"low","high")&"light"""
%>			  <div style="display:table-row">
			    <div style="display:table-cell" <%=cellbg%>><%
									if lastloc=activityList(index,0) then
										print "<div style=""text-align:center"">&quot;</div>"
									else
										print activityList(index,0)
										lastloc=activityList(index,0)
									end if %></div>
				<div style="display:table-cell" <%=cellbg%>><%	print activityList(index,4)
									if activityList(index,1)<>"" then print "<br />Signed By : " & activityList(index,1) %></div>
				<div style="display:table-cell" <%=cellbg%>><%=DateSerial(Left(activityList(index,6),4),Mid(activityList(index,6),5,2),Mid(activityList(index,6),7,2))%><br />
				<%=activityList(index,7)%></div>
			  </div>
<%			next %>
			</div>
<%		else %>
			<div class="ectdiv2column ectwarning">The tracking system returned the following error : <% print errormsg%></div>
<%		end if
	end if
	CanadaPostTrack=success
end function
if getpost("trackno")<>"" then CanadaPostTrack(getpost("trackno"))
%>
			  <div class="ectdivcontainer">
				<div class="ectdivleft">Please enter your Canada Post Tracking Number</div>
				<div class="ectdivright"><input type="text" size="30" name="trackno" value="<% print htmlspecials(request("trackno"))%>" /></div>
			  </div>
			  <div class="ectdivcontainer">
				<div class="ectdivleft">Show Activity</div>
				<div class="ectdivright"><select name="activity" size="1"><option value="LAST">Show Last Activity Only</option><option value="ALL"<% if getpost("activity")="ALL" then print " selected=""selected"""%>>Show All Activity</option></select></div>
			  </div>
			  <div class="ectdiv2column"><% print imageorsubmit(imgtrackpackage,"Track Package","trackpackage")%></div>
			</div>
	</form>
<%
elseif theshiptype="ups" then
%>
	<form method="post" name="trackform" action="tracking<%=extension%>">
	<input type="hidden" name="carrier" value="ups" />
      <div class="ectdiv ecttracking">
		<div class="ectdivhead">
			<div class="trackinglogo"><img src="images/upslogo.png" alt="UPS" /></div>
			<div class="trackingtext">UPS OnLine Tools&reg; Tracking</div>
		</div>
<%
Function getAddress(t, byRef theAddress)
	signedby=""
	For l=0 To t.childNodes.length - 1
		Set u=t.childNodes.Item(l)
		if u.nodeName="AddressLine1" then
			addressline1=u.firstChild.nodeValue
		elseif u.nodeName="AddressLine2" then
			addressline2=u.firstChild.nodeValue
		elseif u.nodeName="AddressLine3" then
			addressline3=u.firstChild.nodeValue
		elseif u.nodeName="City" then
			city=u.firstChild.nodeValue
		elseif u.nodeName="StateProvinceCode" then
			statecode=u.firstChild.nodeValue
		elseif u.nodeName="PostalCode" then
			postcode=u.firstChild.nodeValue
		elseif u.nodeName="CountryCode" then
			sSQL="SELECT countryName FROM countries WHERE countryCode='" & u.firstChild.nodeValue & "'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				countrycode=rs("countryName")
			else
				countrycode=u.firstChild.nodeValue
			end if
			rs.close
		end if
	next
	theAddress=""
	if addressline1<>"" then theAddress=theAddress & addressline1 & "<br />"
	if addressline2<>"" then theAddress=theAddress & addressline2 & "<br />"
	if addressline3<>"" then theAddress=theAddress & addressline3 & "<br />"
	if city<>"" then theAddress=theAddress & city & "<br />"
	if statecode<>"" AND postcode<>"" then
		theAddress=theAddress & statecode & ", " & postcode & "<br />"
	else
		if statecode<>"" then theAddress=theAddress & statecode & "<br />"
		if postcode<>"" then theAddress=theAddress & postcode & "<br />"
	end if
	if countrycode<>"" then theAddress=theAddress & countrycode & "<br />"
End Function
Function ParseUPSTrackingOutput(sXML, byRef totActivity, byRef shipperNo, byRef serviceDesc, byRef shipperaddress, byRef shiptoaddress, byRef scheddeldate, byRef rescheddeldate, byRef errormsg, byRef activityList)
Dim noError, nodeList, packCost, xmlDoc, e, i, j, k, n, t, t2, index
	noError=True
	totalCost=0
	packCost=0
	index=0
	errormsg=""
	gotxml=false
	theaddress=""
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
	xmlDoc.validateOnParse=False
	xmlDoc.loadXML (sXML)
	Set t2=xmlDoc.getElementsByTagName("TrackResponse").Item(0)
	for j=0 to t2.childNodes.length - 1
		Set n=t2.childNodes.Item(j)
		if n.nodename="Response" then
			For i=0 To n.childNodes.length - 1
				Set e=n.childNodes.Item(i)
				if e.nodeName="ResponseStatusCode" then
					noError=Int(e.firstChild.nodeValue)=1
				end if
				if e.nodeName="Error" then
					errormsg=""
					For k=0 To e.childNodes.length - 1
						Set t=e.childNodes.Item(k)
						Select Case t.nodeName
							Case "ErrorSeverity"
								if t.firstChild.nodeValue="Transient" then errormsg="This is a temporary error. Please wait a few moments then refresh this page.<br />" & errormsg
							Case "ErrorDescription"
								errormsg=errormsg & t.firstChild.nodeValue
						End Select
					Next
				end if
				' print "The Nodename is : " & e.nodeName & ":" & e.firstChild.nodeValue & "<br />"
			Next
		elseif n.nodename="Shipment" then
			For i=0 To n.childNodes.length - 1
				Set e=n.childNodes.Item(i)
				' print "Nodename is : " & e.nodeName & "<br />"
				Select Case e.nodeName
					Case "Shipper"
						For k=0 To e.childNodes.length - 1
							Set t=e.childNodes.Item(k)
							if t.nodeName="ShipperNumber" then
								shipperNo=t.firstChild.nodeValue
							elseif t.nodeName="Address" then
								call getAddress(t, shipperaddress)
							end if
						Next
					Case "ShipTo"
						For k=0 To e.childNodes.length - 1
							Set t=e.childNodes.Item(k)
							if t.nodeName="Address" then
								call getAddress(t, shiptoaddress)
							end if
						Next
					Case "ScheduledDeliveryDate"
						scheddeldate=e.firstChild.nodeValue
					Case "Service"
						For k=0 To e.childNodes.length - 1
							Set t=e.childNodes.Item(k)
							if t.nodeName="X_Code_X" then
								Select Case Int(t.firstChild.nodeValue)
									Case 1
										serviceDesc="Next Day Air"
									Case 2
										serviceDesc="2nd Day Air"
									Case 3
										serviceDesc="Ground Service"
									Case 7
										serviceDesc="Worldwide Express"
									Case 8
										serviceDesc="Worldwide Expedited"
									Case 11
										serviceDesc="Standard service"
									Case 12
										serviceDesc="3-Day Select"
									Case 13
										serviceDesc="Next Day Air Saver"
									Case 14
										serviceDesc="Next Day Air Early AM"
									Case 54
										serviceDesc="Worldwide Express Plus"
									Case 59
										serviceDesc="2nd Day Air AM"
									Case 64
										serviceDesc="UPS Express NA1"
									Case 65
										serviceDesc="Express Saver"
								End Select
								' print "The service code is : " & t.nodeName & ":" & t.firstChild.nodeValue & "<br />"
							elseif t.nodeName="Description" then
								serviceDesc=t.firstChild.nodeValue
							end if
						Next
					Case "Package"
						For k=0 To e.childNodes.length - 1
							Set t=e.childNodes.Item(k)
							if t.nodeName="RescheduledDeliveryDate" then
								rescheddeldate=t.firstChild.nodeValue
							elseif t.nodeName="Activity" then
								For l=0 To t.childNodes.length - 1
									Set u=t.childNodes.Item(l)
									if u.nodeName="ActivityLocation" then
										For m=0 To u.childNodes.length - 1
											Set v=u.childNodes.Item(m)
											if v.nodeName="Address" then
												call getAddress(v, activityList(totActivity,0))
											elseif v.nodeName="Description" then
												description=v.firstChild.nodeValue
											elseif v.nodeName="SignedForByName" then
												activityList(totActivity,1)=v.firstChild.nodeValue
											end if
										Next
									elseif u.nodeName="Status" then
										For m=0 To u.childNodes.length - 1
											Set v=u.childNodes.Item(m)
											if v.nodeName="StatusType" then
												For nn=0 To v.childNodes.length - 1
													Set w=v.childNodes.Item(nn)
													if w.nodeName="Code" then
														activityList(totActivity,3)=w.firstChild.nodeValue
													elseif w.nodeName="Description" then
														activityList(totActivity,4)=w.firstChild.nodeValue
													end if
												next
											elseif v.nodeName="StatusCode" then
												For nn=0 To v.childNodes.length - 1
													Set w=v.childNodes.Item(nn)
													if w.nodeName="Code" then
														activityList(totActivity,5)=w.firstChild.nodeValue
													end if
												next
											end if
										Next
									else
										if u.nodeName="Date" then
											activityList(totActivity,6)=u.firstChild.nodeValue
										elseif u.nodeName="Time" then
											activityList(totActivity,7)=u.firstChild.nodeValue
										end if
									end if
								Next
								totActivity=totActivity + 1
								' print "<HR>"
							end if
						Next
				End select
			Next
		end if
	Next
	ParseUPSTrackingOutput=noError
end Function
function UPSTrack(trackNo)
	Dim i, activityList(100,10),success,lastloc
	success=TRUE
	lastloc="xxxxxx"
	sXML="<?xml version=""1.0""?><AccessRequest xml:lang=""en-US""><AccessLicenseNumber>"&upsAccess&"</AccessLicenseNumber><UserId>"&upsUser&"</UserId><Password>"&upsPw&"</Password></AccessRequest>"
	sXML=sXML & "<?xml version=""1.0""?><TrackRequest xml:lang=""en-US""><Request><TransactionReference><CustomerContext>Example 3</CustomerContext><XpciVersion>1.0001</XpciVersion></TransactionReference><RequestAction>Track</RequestAction><RequestOption>"
	if getpost("activity")="LAST" then sXML=sXML & "none" else sXML=sXML & "activity"
	sXML=sXML & "</RequestOption></Request>"
	if false then
		sXML=sXML & "<ReferenceNumber><Value>"&trackNo&"</Value></ReferenceNumber>"
		sXML=sXML & "<ShipperNumber>116593</ShipperNumber></TrackRequest>"
	else
		sXML=sXML & "<TrackingNumber>"&trackNo&"</TrackingNumber></TrackRequest>"
	end if
	xmlres="xml"
	if upstestmode=TRUE then print "UPS Test Mode<br />" : upsurl="wwwcie.ups.com" else upsurl="onlinetools.ups.com"
	' print replace(replace(sXML,"</","&lt;/"),"<","<br />&lt;")&"<hr />"
	if callxmlfunction("https://"&upsurl&"/ups.app/xml/Track", sXML, xmlres, "", "Msxml2.ServerXMLHTTP", errormsg, FALSE) then
		saveLCID=Session.LCID
		Session.LCID=1033
		totActivity=0
		' print Replace(Replace(xmlres.xml,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
		success=ParseUPSTrackingOutput(xmlres.xml, totActivity, shipperNo, serviceDesc, shipperaddress, shiptoaddress, scheduleddeliverydate, rescheddeliverydate, errormsg, activityList)
		Session.LCID=saveLCID
		if success then
			for index2=0 to totActivity-2
				for index=0 to totActivity-2
					if Int(activityList(index,6)&activityList(index,7))>Int(activityList(index+1,6)&activityList(index+1,7)) then
						for index3=0 to UBOUND(activityList,2)
							tempArr=activityList(index,index3)
							activityList(index,index3)=activityList(index+1,index3)
							activityList(index+1,index3)=tempArr
						next
					end if
				next
			next
			if trim(shipperNo)<>"" then %>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Shipper Number</div>
			<div class="ectdivright"><%=shipperNo%></div>
		  </div>
		<%	end if
			if trim(serviceDesc)<>"" then %>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Service Description</div>
			<div class="ectdivright"><%=serviceDesc%></div>
		  </div>
		<%	end if
			if trim(shipperaddress)<>"" then %>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Shipper Address</div>
			<div class="ectdivright"><%=shipperaddress%></div>
		  </div>
		<%	end if
			if trim(shiptoaddress)<>"" then %>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Ship-To Address</div>
			<div class="ectdivright"><%=shiptoaddress%></div>
		  </div>
		<%	end if
			if trim(scheduleddeliverydate)<>"" then %>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Sched. Delivery Date</div>
			<div class="ectdivright"><%=DateSerial(Left(scheduleddeliverydate,4),Mid(scheduleddeliverydate,5,2),Mid(scheduleddeliverydate,7,2)) %></div>
		  </div>
		<%	end if
			if trim(rescheddeliverydate)<>"" then %>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">ReSched. Delivery Date</div>
			<div class="ectdivright"><%=DateSerial(Left(rescheddeliverydate,4),Mid(rescheddeliverydate,5,2),Mid(rescheddeliverydate,7,2)) %></div>
		  </div>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Note</div>
			<div class="ectdivright">Your package is in the UPS system and has a rescheduled delivery date of <%=DateSerial(Left(rescheddeliverydate,4),Mid(rescheddeliverydate,5,2),Mid(rescheddeliverydate,7,2)) %></div>
		  </div>
		<%	end if %>
		<div class="ecttrackingresults" style="display:table">
		  <div style="display:table-row" class="tracktablehead">
			<div style="display:table-cell">Location</div>
			<div style="display:table-cell">Description</div>
			<div style="display:table-cell">Date&nbsp;/&nbsp;Time</div>
		  </div>
<%			for index=0 to totActivity-1 
				cellbg="class=""ect"&IIfVr(index MOD 2=0,"low","high")&"light""" %>
		  <div style="display:table-row">
			<div style="display:table-cell" <%=cellbg%>><%	if lastloc=activityList(index,0) then
									print "<div style=""text-align:center"">&quot;</div>"
								else
									print activityList(index,0)
									lastloc=activityList(index,0)
								end if %></div>
			<div style="display:table-cell" <%=cellbg%>><%	print activityList(index,4)
								if activityList(index,1)<>"" then print "<br />Signed By : " & activityList(index,1) %></div>
			<div style="display:table-cell" <%=cellbg%>><%=DateSerial(Left(activityList(index,6),4),Mid(activityList(index,6),5,2),Mid(activityList(index,6),7,2))%><br />
			<%=TimeSerial(Left(activityList(index,7),2),Mid(activityList(index,7),3,2),Mid(activityList(index,7),5,2))%></div>
		  </div>
<%			next %>
		</div>
<%		end if
	else
		success=FALSE
	end If
	if NOT success then %>
		<div class="ectdiv2column ectwarning">The tracking system returned the following error : <%=errormsg%></div>
<%	end if
	UPSTrack=success
end function
if getpost("trackno")<>"" then
	UPSTrack(getpost("trackno"))
end if
%>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Please enter your UPS Tracking Number</div>
			<div class="ectdivright"><input type="text" size="30" name="trackno" value="<%=htmlspecials(trim(request("trackno")))%>" /></div>
		  </div>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Show Activity</div>
			<div class="ectdivright"><select name="activity" size="1"><option value="LAST">Show Last Activity Only</option><option value="ALL"<% if getpost("activity")="ALL" then print " selected=""selected"""%>>Show All Activity</option></select></div>
		  </div>
		  <div class="ectdiv2column"><%=imageorbutton(imgviewlicense,"View License","viewlicense","viewlicense()",TRUE)&" "&imageorbutton(imgtrackpackage,"Track Package","trackpackage","checkaccept()",TRUE)%></div>
		  <div class="ectdiv2column"><input type="checkbox" name="agreeconds" value="ON" <%if getpost("agreeconds")="ON" then print "checked"%> /> By selecting this box and the "Track Package" button, I agree to these <a class="ectlink" href="javascript:viewlicense();">Terms and Conditions</a>.</div>
		  <div class="trackingcopyright"><%=xxUPStm%></div>
	  </div>
	</form>
<%
elseif theshiptype="usps" then
%>
	<form method="post" name="trackform" action="tracking<%=extension%>">
	<input type="hidden" name="carrier" value="usps" />
	  <div class="ectdiv ecttracking">
		<div class="ectdivhead">
			<div class="trackinglogo"><img src="images/usps_logo.png" alt="USPS" /></div>
			<div class="trackingtext">USPS Tracking Tool</div>
		</div>
<%
Function ParseUSPSTrackingOutput(sXML, byRef totActivity, onlylast, byRef serviceDesc, byRef shipperaddress, byRef shiptoaddress, byRef scheddeldate, byRef rescheddeldate, byRef errormsg, byRef activityList)
Dim noError, nodeList, packCost, xmlDoc, e, i, j, k, n, t, t2, index
	noError=True
	totalCost=0
	packCost=0
	index=0
	errormsg=""
	gotxml=false
	theaddress=""
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
	xmlDoc.validateOnParse=False
	xmlDoc.loadXML (sXML)
	If xmlDoc.documentElement.nodeName="Error" then 'Top-level Error
		Set t2=xmlDoc.getElementsByTagName("Error").Item(0)
		noError=FALSE
		for j=0 to t2.childNodes.length - 1
			Set n=t2.childNodes.Item(j)
			if n.nodename="Description" then
				errormsg=n.firstChild.nodeValue
			end if
		next
	else
		Set t2=xmlDoc.getElementsByTagName("TrackInfo").Item(0)
		for j=0 to t2.childNodes.length - 1
			companyname= ""
			city=""
			statecode=""
			postcode=""
			countrycode=""
			Set n=t2.childNodes.Item(j)
			if n.nodename="Error" then
				For i=0 To n.childNodes.length - 1
					Set e=n.childNodes.Item(i)
					if e.nodeName="Description" then
						errormsg=e.firstChild.nodeValue
						noError=FALSE
					end if
					' print "The Nodename is : " & e.nodeName & ":" & e.firstChild.nodeValue & "<br />"
				Next
			elseif n.nodename="TrackDetail" then
				if NOT onlylast then
					For i=0 To n.childNodes.length - 1
						Set e=n.childNodes.Item(i)
						' print "Nodename is : " & e.nodeName & "<br />"
						Select Case e.nodeName
							Case "EventDate"
								if e.hasChildNodes then activityList(totActivity,6)=e.firstChild.nodeValue else activityList(totActivity,6)=""
							Case "EventTime"
								if e.hasChildNodes then activityList(totActivity,7)=e.firstChild.nodeValue else activityList(totActivity,7)=""
							Case "Event"
								activityList(totActivity,4)=e.firstChild.nodeValue
							Case "EventCity"
								if e.hasChildNodes then city=e.firstChild.nodeValue
							Case "EventState"
								if e.hasChildNodes then statecode=e.firstChild.nodeValue
							Case "EventZIPCode"
								if e.hasChildNodes then postcode=e.firstChild.nodeValue
							Case "EventCountry"
								if e.hasChildNodes then countrycode=e.firstChild.nodeValue
							Case "FirmName"
								if e.hasChildNodes then companyname=e.firstChild.nodeValue
						End select
					Next
					theAddress=""
					if companyname<>"" then theAddress=theAddress & companyname & "<br />"
					if city<>"" then theAddress=theAddress & city & "<br />"
					if statecode<>"" AND postcode<>"" then
						theAddress=theAddress & statecode & ", " & postcode & "<br />"
					else
						if statecode<>"" then theAddress=theAddress & statecode & "<br />"
						if postcode<>"" then theAddress=theAddress & postcode & "<br />"
					end if
					if countrycode<>"" then theAddress=theAddress & countrycode & "<br />"
					activityList(totActivity,0)=theAddress
					totActivity=totActivity + 1
				end if
			elseif n.nodename="TrackSummary" then
				For i=0 To n.childNodes.length - 1
					Set e=n.childNodes.Item(i)
					' print "Nodename is : " & e.nodeName & "<br />"
					Select Case e.nodeName
						Case "EventDate"
							if e.hasChildNodes then scheddeldate=e.firstChild.nodeValue&scheddeldate
						Case "EventTime"
							if e.hasChildNodes then scheddeldate=scheddeldate&" "&e.firstChild.nodeValue
						Case "Event"
							if e.hasChildNodes then serviceDesc=e.firstChild.nodeValue
						Case "EventCity"
							if e.hasChildNodes then city=e.firstChild.nodeValue
						Case "EventState"
							if e.hasChildNodes then statecode=e.firstChild.nodeValue
						Case "EventZIPCode"
							if e.hasChildNodes then postcode=e.firstChild.nodeValue
						Case "EventCountry"
							if e.hasChildNodes then countrycode=e.firstChild.nodeValue
						Case "FirmName"
							if e.hasChildNodes then companyname=e.firstChild.nodeValue
					End select
				Next
				theAddress=""
				if companyname<>"" then theAddress=theAddress & companyname & "<br />"
				if city<>"" then theAddress=theAddress & city & "<br />"
				if statecode<>"" AND postcode<>"" then
					theAddress=theAddress & statecode & ", " & postcode & "<br />"
				else
					if statecode<>"" then theAddress=theAddress & statecode & "<br />"
					if postcode<>"" then theAddress=theAddress & postcode & "<br />"
				end if
				if countrycode<>"" then theAddress=theAddress & countrycode & "<br />"
				shiptoaddress=theAddress
			end if
		Next
	end if
	ParseUSPSTrackingOutput=noError
end Function
function USPSTrack(trackNo)
	Dim objHttp, i, activityList(100,10),success,lastloc
	lastloc="xxxxxx"
	sXML="<TrackFieldRequest USERID="""&uspsUser&"""><TrackID ID="""&replace(trackNo," ","")&"""></TrackID></TrackFieldRequest>"
	if proxyserver<>"" then
		set objHttp=Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
		objHttp.setProxy 2, proxyserver
	else
		set objHttp=Server.CreateObject("Msxml2.ServerXMLHTTP")
	end if
	objHttp.open "POST", "https://production.shippingapis.com/ShippingAPI.dll", false
	on error resume next
	err.number=0
	' print Replace(Replace(sXML,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
	objHttp.Send "API=TrackV2&XML=" & urlencode(sXML)
	on error goto 0
	if err.number <> 0 OR objHttp.status <> 200 then
		errormsg="Error, couldn't connect to USPS server"
		success=false
	else
		saveLCID=Session.LCID
		Session.LCID=1033
		totActivity=0
		' print Replace(Replace(objHttp.responseText,"</","&lt;/"),"<","<br />&lt;")&"<HR>"
		success=ParseUSPSTrackingOutput(objHttp.responseText, totActivity, getpost("activity")="LAST", serviceDesc, shipperaddress, shiptoaddress, scheduleddeliverydate, rescheddeliverydate, errormsg, activityList)
		Session.LCID=saveLCID
		if success then
			for index2=0 to totActivity-2
				for index=0 to totActivity-2
					swapdate=FALSE
					if activityList(index,6)="" then
						swapdate=TRUE
					elseif activityList(index+1,6)="" then
						
					elseif DateValue(activityList(index,6)&" "&activityList(index,7))>DateValue(activityList(index+1,6)&" "&activityList(index+1,7)) then
						swapdate=TRUE
					end if
					if swapdate then
						for index3=0 to UBOUND(activityList,2)
							tempArr=activityList(index,index3)
							activityList(index,index3)=activityList(index+1,index3)
							activityList(index+1,index3)=tempArr
						next
					end if
				next
			next
			if trim(serviceDesc)<>"" then %>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Event</div>
			<div class="ectdivright"><%=serviceDesc%></div>
		  </div>
		<%	end if
			if trim(shiptoaddress)<>"" then %>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Address</div>
			<div class="ectdivright"><%=shiptoaddress%></div>
		  </div>
		<%	end if
			if trim(scheduleddeliverydate)<>"" then %>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Event Date</div>
			<div class="ectdivright"><%=scheduleddeliverydate %></div>
		  </div>
		<%	end if
			if totActivity>0 then %>
		<div class="ecttrackingresults" style="display:table">
		  <div style="display:table-row" class="tracktablehead">
			<div style="display:table-cell">Location</div>
			<div style="display:table-cell">Description</div>
			<div style="display:table-cell">Date&nbsp;/&nbsp;Time</div>
		  </div>
<%				for index=0 to totActivity-1 
					cellbg="class=""ect"&IIfVr(index MOD 2=0,"low","high")&"light"""
%>		  <div style="display:table-row">
			<div style="display:table-cell" <%=cellbg%>><%
									if lastloc=activityList(index,0) then
										print "<div style=""text-align:center"">&quot;</div>"
									else
										print activityList(index,0)
										lastloc=activityList(index,0)
									end if %></div>
			<div style="display:table-cell" <%=cellbg%>><%	print activityList(index,4)
									if activityList(index,1)<>"" then print "<br />Signed By : " & activityList(index,1) %></div>
			<div style="display:table-cell" <%=cellbg%>><%=activityList(index,6)%>
				<br /><%=activityList(index,7)%></div>
		  </div>
<%				next %>
		</div>
<%			end if
		else %>
		<div class="ectdiv2column ectwarning">The tracking system returned the following error : <%=errormsg%></div>
<%		end if
	end If
	USPSTrack=success
	set objHttp=nothing
end function
if getpost("trackno")<>"" then
	USPSTrack(getpost("trackno"))
end if
%>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Please enter your USPS Tracking Number</div>
			<div class="ectdivright"><input type="text" size="30" name="trackno" value="<%=htmlspecials(trim(request("trackno")))%>" /></div>
		  </div>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Show Activity</div>
			<div class="ectdivright"><select name="activity" size="1"><option value="LAST">Show Last Activity Only</option><option value="ALL"<% if getpost("activity")="ALL" OR getpost("activity")="" then print " selected=""selected"""%>>Show All Activity</option></select></div>
		  </div>
		  <div class="ectdiv2column"><%=imageorsubmit(imgtrackpackage,"Track Package","trackpackage")%></div>
	  </div>
	</form>
<%
elseif theshiptype="fedex" then
%>
	<form method="post" name="trackform" action="tracking<%=extension%>">
	<input type="hidden" name="carrier" value="fedex" />
      <div class="ectdiv ecttracking">
		<div class="ectdivhead">
			<div class="trackinglogo"><img src="images/fedexlogo.png" alt="FedEx" /></div>
			<div class="trackingtext">FedEx<small>&reg;</small> Tracking Tool</div>
		</div>
<%
function getFedExAddress(t, byRef theAddress)
	signedby=""
	For l=0 To t.childNodes.length - 1
		Set u=t.childNodes.Item(l)
		if u.nodeName="AddressLine1" then
			addressline1=u.firstChild.nodeValue
		elseif u.nodeName="AddressLine2" then
			addressline2=u.firstChild.nodeValue
		elseif u.nodeName="AddressLine3" then
			addressline3=u.firstChild.nodeValue
		elseif u.nodeName=fns&"City" then
			city=u.firstChild.nodeValue
		elseif u.nodeName=fns&"StateOrProvinceCode" then
			statecode=u.firstChild.nodeValue
		elseif u.nodeName=fns&"PostalCode" then
			postcode=u.firstChild.nodeValue
		elseif u.nodeName=fns&"CountryCode" then
			sSQL="SELECT countryName FROM countries WHERE countryCode='" & u.firstChild.nodeValue & "'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				countrycode=rs("countryName")
			else
				countrycode=u.firstChild.nodeValue
			end if
			rs.close
		end if
	next
	theAddress=""
	if addressline1<>"" then theAddress=theAddress & addressline1 & "<br />"
	if addressline2<>"" then theAddress=theAddress & addressline2 & "<br />"
	if addressline3<>"" then theAddress=theAddress & addressline3 & "<br />"
	if city<>"" then theAddress=theAddress & city & "<br />"
	if statecode<>"" AND postcode<>"" then
		theAddress=theAddress & statecode & ", " & postcode & "<br />"
	else
		if statecode<>"" then theAddress=theAddress & statecode & "<br />"
		if postcode<>"" then theAddress=theAddress & postcode & "<br />"
	end if
	if countrycode<>"" then theAddress=theAddress & countrycode & "<br />"
end function
function ParseFedexTrackingOutput(xmlDoc, byRef totActivity, byRef deliverydate, byRef serviceDesc, byRef packagecount, byRef shiptoaddress, byRef scheddeldate, byRef signedforby, byRef errormsg, byRef activityList)
	Dim noError, nodeList, packCost, e, i, j, k, n, t, t2, index
	noError=True
	totalCost=0
	packCost=0
	index=0
	errormsg=""
	theaddress=""
	fns=fedexnamespace
	if fns<>"" then fns=fns&":"
	Set t2=xmlDoc.getElementsByTagName(fns&"TrackReply").Item(0)
	for j=0 to t2.childNodes.length - 1
		Set e=t2.childNodes.Item(j)
		' print "e.nodeName: " & e.nodeName & "<br>"
		Select Case e.nodeName
			Case fns&"HighestSeverity"
				noError=(e.firstChild.nodeValue<>"ERROR" AND e.firstChild.nodeValue<>"FAILURE")
			Case fns&"Notifications"
				For k=0 To e.childNodes.length - 1
					Set t=e.childNodes.Item(k)
					if t.nodeName=fns&"Message" then errormsg=t.firstChild.nodeValue
				Next
			Case fns&"TrackDetails"
				For k=0 To e.childNodes.length - 1
					Set fxw=e.childNodes.Item(k)
					' print "fxw.nodeName: " & fxw.nodeName & "<br>"
					Select Case fxw.nodeName
						Case fns&"DeliverySignatureName"
							signedforby=fxw.firstChild.nodeValue
						Case fns&"DestinationAddress"
							call getFedExAddress(fxw, shiptoaddress)
						Case "DeliveredDate"
							deliverydate=fxw.firstChild.nodeValue & deliverydate
						Case "DeliveredTime"
							deliverydate=deliverydate & " " & fxw.firstChild.nodeValue
						Case fns&"ServiceType"
							serviceDesc=fxw.firstChild.nodeValue
						Case fns&"PackageCount"
							packagecount=fxw.firstChild.nodeValue
						Case fns&"Events"
							For kfx=0 To fxw.childNodes.length - 1
								Set t=fxw.childNodes.Item(kfx)
								if t.nodeName=fns&"Timestamp" then
									activityList(totActivity,6)=t.firstChild.nodeValue
								elseif t.nodeName="Time" then
									activityList(totActivity,7)=t.firstChild.nodeValue
								elseif t.nodeName="StatusExceptionCode" then
									activityList(totActivity,3)=t.firstChild.nodeValue
								elseif t.nodeName=fns&"EventDescription" OR t.nodeName="StatusExceptionDescription" then
									if t.firstChild.nodeValue <> "Package status" then activityList(totActivity,4)=t.firstChild.nodeValue
								elseif t.nodeName=fns&"Address" then
									call getFedExAddress(t, activityList(totActivity,0))
								end if
							Next
							if activityList(totActivity,4)<>"" then totActivity=totActivity + 1
					End select
				Next
		End select
	Next
	ParseFedexTrackingOutput=noError
end function
function FedexTrack(trackNo)
	Dim objHttp, i, activityList(100,10),success,lastloc
	lastloc="xxxxxx"
sXML ="<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:v4=""http://fedex.com/ws/track/v4"">" & _
"   <soapenv:Header/>" & _
"   <soapenv:Body>" & _
"      <v4:TrackRequest>" & _
"         <v4:WebAuthenticationDetail>" & _
"            <v4:CspCredential>" & _
"               <v4:Key>mKOUqSP4CS0vxaku</v4:Key>" & _
"               <v4:Password>IAA5db3Pmhg3lyWW6naMh4Ss2</v4:Password>" & _
"            </v4:CspCredential>" & _
"            <v4:UserCredential>" & _
"               <v4:Key>" & fedexuserkey & "</v4:Key>" & _
"               <v4:Password>" & fedexuserpwd & "</v4:Password>" & _
"            </v4:UserCredential>" & _
"         </v4:WebAuthenticationDetail>" & _
"         <v4:ClientDetail>" & _
"            <v4:AccountNumber>" & fedexaccount & "</v4:AccountNumber>" & _
"            <v4:MeterNumber>" & fedexmeter & "</v4:MeterNumber>" & _
"            <v4:ClientProductId>IBTB</v4:ClientProductId>" & _
"            <v4:ClientProductVersion>3272</v4:ClientProductVersion>" & _
"         </v4:ClientDetail>" & _
"         <v4:TransactionDetail>" & _
"            <v4:CustomerTransactionId>track Request</v4:CustomerTransactionId>" & _
"         </v4:TransactionDetail>" & _
"         <v4:Version>" & _
"            <v4:ServiceId>trck</v4:ServiceId>" & _
"            <v4:Major>4</v4:Major>" & _
"            <v4:Intermediate>1</v4:Intermediate>" & _
"            <v4:Minor>0</v4:Minor>" & _
"         </v4:Version>" & _
"         <v4:PackageIdentifier>" & _
"            <v4:Value>"&trackNo&"</v4:Value>" & _
"            <v4:Type>TRACKING_NUMBER_OR_DOORTAG</v4:Type>" & _
"         </v4:PackageIdentifier>" & _
"         <v4:IncludeDetailedScans>" & IIfVr(getpost("activity")="LAST","false","true") & "</v4:IncludeDetailedScans>" & _
"      </v4:TrackRequest>" & _
"   </soapenv:Body>" & _
"</soapenv:Envelope>"

	xmlres="xml"
	success=callxmlfunction(fedexurl, sXML, xmlres, "", "Msxml2.ServerXMLHTTP", errormsg, FALSE)
	if success then
		Session.LCID=1033
		totActivity=0
		set toregexp=new RegExp
		toregexp.pattern="<(.{1,3}):TrackReply"
		set matches=toregexp.execute(xmlres.xml)
		set toregexp=nothing
		if matches.count > 0 then fedexnamespace=matches(0).submatches(0) else fedexnamespace=""
		success=ParseFedexTrackingOutput(xmlres, totActivity, deliverydate, serviceDesc, packagecount, shiptoaddress, scheduleddeliverydate, signedforby, errormsg, activityList)
		Session.LCID=saveLCID
		if success then
			for index2=0 to totActivity-2
				for index=0 to totActivity-2
					if (activityList(index,6)&activityList(index,7))>(activityList(index+1,6)&activityList(index+1,7)) then
						for index3=0 to UBOUND(activityList,2)
							tempArr=activityList(index,index3)
							activityList(index,index3)=activityList(index+1,index3)
							activityList(index+1,index3)=tempArr
						next
					end if
				next
			next
			if trim(serviceDesc)<>"" then %>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Service Description</div>
			<div class="ectdivright"><%=replace(serviceDesc,"_"," ")%></div>
		  </div>
		<%	end if
			if trim(packagecount)<>"" then %>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Package Count</div>
			<div class="ectdivright"><%=packagecount%></div>
		  </div>
		<%	end if
			if trim(shiptoaddress)<>"" then %>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Ship-To Address</div>
			<div class="ectdivright"><%=shiptoaddress%></div>
		  </div>
		<%	end if
			if trim(signedforby)<>"" then %>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Signed For By</div>
			<div class="ectdivright"><%=signedforby %></div>
		  </div>
		<%	end if
			if trim(deliverydate)<>"" then %>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Delivery Date</div>
			<div class="ectdivright"><%=deliverydate%></div>
		  </div>
		<%	end if %>
		<div class="ecttrackingresults" style="display:table">
		  <div style="display:table-row" class="tracktablehead">
			<div style="display:table-cell">Location</div>
			<div style="display:table-cell">Description</div>
			<div style="display:table-cell">Date&nbsp;/&nbsp;Time</div>
		  </div>
<%			for index=0 to totActivity-1 
				cellbg="class=""ect"&IIfVr(index MOD 2=0,"low","high")&"light""" %>
		  <div style="display:table-row">
			<div style="display:table-cell" <%=cellbg%>><%	if lastloc=activityList(index,0) then
									print "<div style=""text-align:center"">&quot;</div>"
								else
									print activityList(index,0)
									lastloc=activityList(index,0)
								end if %></div>
			<div style="display:table-cell" <%=cellbg%>><%	print activityList(index,4)
								if activityList(index,1)<>"" then print "<br />Signed By : " & activityList(index,1) %></div>
			<div style="display:table-cell" <%=cellbg%>><%
				fxdate=activityList(index,6)
				tpos=instr(fxdate,"T")
				if instr(tpos,fxdate,"-")=0 then offsetpos=instr(tpos,fxdate,"+") else offsetpos=instr(tpos,fxdate,"-")
				if offsetpos>0 then fxdate=left(fxdate,offsetpos-1)
				fxdate=replace(fxdate,"T"," ")
				print DateValue(fxdate)
				print "<br />" & right(fxdate,len(fxdate)-instr(fxdate," "))
				%><br /><%=activityList(index,7)%></div>
		  </div>
<%			next %>
		</div>
<%		else %>
		<div class="ectdiv2column ectwarning">The tracking system returned the following error : <%=errormsg%></div>
<%		end if
	end If
	FedexTrack=success
	set objHttp=nothing
end function
if getpost("trackno")<>"" then
	FedexTrack(getpost("trackno"))
end if
%>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Please enter your FedEx Tracking Number</div>
			<div class="ectdivright"><input type="text" size="30" name="trackno" value="<%=htmlspecials(trim(request("trackno")))%>" /></div>
		  </div>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Show Activity</div>
			<div class="ectdivright"><select name="activity" size="1"><option value="LAST">Show Last Activity Only</option><option value="ALL"<% if getpost("activity")="ALL" then print " selected=""selected"""%>>Show All Activity</option></select></div>
		  </div>
		  <div class="ectdiv2column"><%=imageorsubmit(imgtrackpackage,"Track Package","trackpackage")%></div>
		  <div class="trackingcopyright"><%=fedexcopyright%></div>
	  </div>
	</form>
<%
elseif theshiptype="dhl" then
%>
	<form method="post" name="trackform" action="tracking<%=extension%>">
	<input type="hidden" name="carrier" value="dhl" />
      <div class="ectdiv ecttracking">
		<div class="ectdivhead">
			<div class="trackinglogo"><img src="images/dhllogo.gif" alt="DHL" style="margin-left:10px" /></div>
			<div class="trackingtext">DHL<small>&reg;</small> Tracking Tool</div>
		</div>
<%
function getDHLDescription(t, byRef theAddress)
	signedby=""
	For l=0 To t.childNodes.length - 1
		Set u=t.childNodes.Item(l)
		if u.nodeName="Description" then
			addressline1=u.firstChild.nodeValue
		end if
	next
	theAddress=addressline1
end function
function ParseDHLTrackingOutput(xmlDoc, byRef totActivity, byRef deliverydate, byRef origservicearea, byRef shiptoaddress, byRef scheddeldate, byRef signedforby, byRef errormsg, byRef activityList)
	Dim noError, nodeList, packCost, e, i, j, k, n, t, t2, index
	noError=True
	totalCost=0
	packCost=0
	index=0
	errormsg=""
	theaddress=""
	Set t2=xmlDoc.getElementsByTagName("AWBInfo").Item(0)
	for j=0 to t2.childNodes.length - 1
		Set e=t2.childNodes.Item(j)
		Select Case e.nodeName
			Case "HighestSeverity"
				noError=(e.firstChild.nodeValue<>"ERROR" AND e.firstChild.nodeValue<>"FAILURE")
			Case "Status"
				For k=0 To e.childNodes.length - 1
					Set t=e.childNodes.Item(k)
					if t.nodeName="ActionStatus" then
						noError=(t.firstChild.nodeValue="success")
						errormsg=t.firstChild.nodeValue
					end if
				Next
			Case "ShipmentInfo"
				For k=0 To e.childNodes.length - 1
					Set fxw=e.childNodes.Item(k)
					Select Case fxw.nodeName
						Case "OriginServiceArea"
							call getDHLDescription(fxw, origservicearea)
						Case "DestinationServiceArea"
							call getDHLDescription(fxw, shiptoaddress)
						Case "DeliveredDate"
							deliverydate=fxw.firstChild.nodeValue & deliverydate
						Case "DeliveredTime"
							deliverydate=deliverydate & " " & fxw.firstChild.nodeValue
						Case "ServiceType"
							serviceDesc=fxw.firstChild.nodeValue
						Case "PackageCount"
							packagecount=fxw.firstChild.nodeValue
						Case "ShipmentEvent"
							For kfx=0 To fxw.childNodes.length - 1
								Set t=fxw.childNodes.Item(kfx)
								if t.nodeName="Date" then
									activityList(totActivity,6)=t.firstChild.nodeValue
								elseif t.nodeName="Time" then
									activityList(totActivity,7)=t.firstChild.nodeValue
								elseif t.nodeName="Signatory" then
									if t.haschildnodes then signedforby=t.firstChild.nodeValue
								elseif t.nodeName="ServiceEvent" then
									set obj2=t.getElementsByTagName("Description")
									if obj2.length > 0 then
										if obj2.item(0).hasChildNodes then activityList(totActivity,4)=obj2.item(0).firstChild.nodeValue
									end if
								elseif t.nodeName="ServiceArea" then
									set obj2=t.getElementsByTagName("Description")
									if obj2.length > 0 then
										if obj2.item(0).hasChildNodes then activityList(totActivity,0)=obj2.item(0).firstChild.nodeValue
									end if
								end if
							Next
							if activityList(totActivity,4)<>"" then totActivity=totActivity + 1
					End select
				Next
		End select
	Next
	ParseDHLTrackingOutput=noError
end function
function DHLTrack(trackNo)
	Dim objHttp, i, activityList(100,10),success,lastloc
	lastloc="xxxxxx"
	sXML="<?xml version=""1.0"" encoding=""utf-8"" ?><req:KnownTrackingRequest xmlns:req=""http://www.dhl.com"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.dhl.com TrackingRequestKnown.xsd"">" & _
	"<Request><ServiceHeader><SiteID>" & DHLSiteID & "</SiteID><Password>" & DHLSitePW & "</Password></ServiceHeader></Request>" & _
	"<LanguageCode>en</LanguageCode><AWBNumber>"&trackNo&"</AWBNumber><LevelOfDetails>" & IIfVr(getpost("activity")="LAST","LAST_CHECK_POINT_ONLY","ALL_CHECK_POINTS") & "</LevelOfDetails><PiecesEnabled>S</PiecesEnabled></req:KnownTrackingRequest>"
	xmlres="xml"
	success=callxmlfunction("https://xmlpi" & IIfVs(upstestmode,"test") & "-ea.dhl.com/XMLShippingServlet", sXML, xmlres, "", "Msxml2.ServerXMLHTTP", errormsg, FALSE)
	if success then
		Session.LCID=1033
		totActivity=0
		success=ParseDHLTrackingOutput(xmlres, totActivity, deliverydate, origservicearea,shiptoservicearea, scheduleddeliverydate, signedforby, errormsg, activityList)
		Session.LCID=saveLCID
		if success then
			for index2=0 to totActivity-2
				for index=0 to totActivity-2
					if (activityList(index,6)&activityList(index,7))>(activityList(index+1,6)&activityList(index+1,7)) then
						for index3=0 to UBOUND(activityList,2)
							tempArr=activityList(index,index3)
							activityList(index,index3)=activityList(index+1,index3)
							activityList(index+1,index3)=tempArr
						next
					end if
				next
			next
			if trim(origservicearea)<>"" then %>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Origin Service Area</div>
			<div class="ectdivright"><%=origservicearea%></div>
		  </div>
		<%	end if
			if trim(shiptoservicearea)<>"" then %>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Destination Service Area</div>
			<div class="ectdivright"><%=shiptoservicearea%></div>
		  </div>
		<%	end if
			if trim(signedforby)<>"" then %>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Signed For By</div>
			<div class="ectdivright"><%=signedforby %></div>
		  </div>
		<%	end if
			if trim(deliverydate)<>"" then %>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Delivery Date</div>
			<div class="ectdivright"><%=deliverydate%></div>
		  </div>
		<%	end if %>
		<div class="ecttrackingresults" style="display:table">
		  <div style="display:table-row" class="tracktablehead">
			<div style="display:table-cell">Location</div>
			<div style="display:table-cell">Description</div>
			<div style="display:table-cell">Date&nbsp;/&nbsp;Time</div>
		  </div>
<%			for index=0 to totActivity-1 
				cellbg="class=""ect"&IIfVr(index MOD 2=0,"low","high")&"light"""
%>		  <div style="display:table-row">
			<div style="display:table-cell" <%=cellbg%>><%	if lastloc=activityList(index,0) then
									print "<div style=""text-align:center"">&quot;</div>"
								else
									print activityList(index,0)
									lastloc=activityList(index,0)
								end if %></div>
			<div style="display:table-cell" <%=cellbg%>><%	print activityList(index,4)
								if activityList(index,1)<>"" then print "<br />Signed By : " & activityList(index,1) %></div>
			<div style="display:table-cell" <%=cellbg%>><%
				fxdate=activityList(index,6)
				print DateValue(fxdate)
			%><br /><%=activityList(index,7)%></div>
		  </div>
<%			next %>
		</div>
<%		else %>
		<div class="ectdiv2column ectwarning">The tracking system returned the following error : <%=errormsg%></div>
<%		end if
	end If
	DHLTrack=success
	set objHttp=nothing
end function
if getpost("trackno")<>"" then
	DHLTrack(getpost("trackno"))
end if
%>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Please enter your DHL Tracking Number</div>
			<div class="ectdivright"><input type="text" size="30" name="trackno" value="<%=htmlspecials(trim(request("trackno")))%>" /></div>
		  </div>
		  <div class="ectdivcontainer">
			<div class="ectdivleft">Show Activity</div>
			<div class="ectdivright"><select name="activity" size="1"><option value="LAST">Show Last Activity Only</option><option value="ALL"<% if getpost("activity")="ALL" then print " selected=""selected"""%>>Show All Activity</option></select></div>
		  </div>
		  <div class="ectdiv2column"><%=imageorsubmit(imgtrackpackage,"Track Package","trackpackage")%></div>
	  </div>
	</form>
<%
elseif theshiptype="royalmail" then %>
<script>
function postrmtrackno(){
	document.getElementById('rmtrackform').action='https://www.royalmail.com/portal/rm/track?trackNumber=';
	document.getElementById('rmtrackform').action+=document.getElementById('rmtrackno').value;
	document.getElementById('rmtrackform').submit();
}
</script>
	<form method="post" id="rmtrackform" name="trackform" action="" target="_blank">
		<input type="hidden" name="carrier" value="royalmail" />
		<div class="ectdiv ecttracking">
			<div class="ectdivhead">
				<div class="trackinglogo"><img src="images/royalmaillogo.png" alt="Royal Mail" style="margin-left:10px" /></div>
				<div class="trackingtext">Royal Mail<small>&reg;</small> Tracking Tool</div>
			</div>
			<div class="ectdivcontainer">
				<div class="ectdivleft">Please enter your Royal Mail Tracking Number</div>
				<div class="ectdivright"><input type="text" size="30" id="rmtrackno" name="trackno" value="<% print htmlspecials(getrequest("trackno"))%>" /></div>
			</div>
			<div class="ectdiv2column"><%=imageorbutton(imgtrackpackage,"Track Package","trackpackage","postrmtrackno()",TRUE)%></div>
		</div>
	</form>
<%
else ' undecided
%>
	<form method="post" action="tracking<%=extension%>">
	<input type="hidden" name="carrier" id="carrier" value="xxxxxx" />
	  <div class="ectdiv ecttracking">
		<div class="ectdivhead trackingpleaseselect">Please select your shipping carrier.</div>
<%		if shipType=4 OR alternateratesups OR instr(1, trackingcarriers, "ups", 1)>0 then %>
		<div class="ectdivcontainer">
			<div class="trackingselectlogo"><img src="images/upslogo.png" alt="UPS" /></div>
			<div class="ectdivleft">Products shipped via UPS</div>
			<div class="ectdivright"><%=imageorsubmit(imgtrackinggo,xxGo&""" onclick=""document.getElementById('carrier').value='ups'","ectbutton trackinggo")%></div>
		</div>
<%		end if
		if shipType=3 OR alternateratesusps OR instr(1, trackingcarriers, "usps", 1)>0 then %>
		<div class="ectdivcontainer">
			<div class="trackingselectlogo"><img src="images/usps_logo.png" alt="USPS" /></div>
			<div class="ectdivleft">Products shipped via USPS</div>
			<div class="ectdivright"><%=imageorsubmit(imgtrackinggo,xxGo&""" onclick=""document.getElementById('carrier').value='usps'","ectbutton trackinggo")%></div>
		</div>
<%		end if
		if shipType=7 OR shipType=8 OR alternateratesfedex OR instr(1, trackingcarriers, "fedex", 1)>0 then %>
		<div class="ectdivcontainer">
			<div class="trackingselectlogo"><img src="images/fedexlogo.png" alt="FedEx" /></div>
			<div class="ectdivleft">Products shipped via FedEx</div>
			<div class="ectdivright"><%=imageorsubmit(imgtrackinggo,xxGo&""" onclick=""document.getElementById('carrier').value='fedex'","ectbutton trackinggo")%></div>
		</div>
<%		end if
		if shipType=9 OR alternateratesdhl OR instr(1, trackingcarriers, "dhl", 1)>0 then %>
		<div class="ectdivcontainer">
			<div class="trackingselectlogo"><img src="images/dhllogo.gif" alt="UPS" /></div>
			<div class="ectdivleft">Products shipped via DHL</div>
			<div class="ectdivright"><%=imageorsubmit(imgtrackinggo,xxGo&""" onclick=""document.getElementById('carrier').value='dhl'","ectbutton trackinggo")%></div>
		</div>
<%		end if
		if shipType=6 OR alternateratescanadapost OR instr(1, trackingcarriers, "canadapost", 1)>0 then %>
		<div class="ectdivcontainer">
			<div class="trackingselectlogo"><img src="images/canadapost.gif" alt="Canada Post" /></div>
			<div class="ectdivleft">Products shipped via Canada Post</div>
			<div class="ectdivright"><%=imageorsubmit(imgtrackinggo,xxGo&""" onclick=""document.getElementById('carrier').value='canadapost'","ectbutton trackinggo")%></div>
		</div>
<%		end if
		if instr(1, trackingcarriers, "royalmail", 1)>0 then %>
		<div class="ectdivcontainer">
			<div class="trackingselectlogo"><img src="images/royalmaillogo.png" alt="Royal Mail" /></div>
			<div class="ectdivleft">Products shipped via Royal Mail</div>
			<div class="ectdivright"><%=imageorsubmit(imgtrackinggo,xxGo&""" onclick=""document.getElementById('carrier').value='royalmail'","ectbutton trackinggo")%></div>
		</div>
<%		end if %>
	  </div>
	</form>
	  <div class="ectdiv ecttracking">
<%	if incupscopyright then %>
        <div class="trackingcopyright"><%=xxUPStm%></div>
<%	end if
	if incfedexcopyright then %>
        <div class="trackingcopyright"><%=fedexcopyright%></div>
<%	end if %>
	  </div>
<%
end if
cnn.Close
set rs=nothing
set cnn=nothing
%>
