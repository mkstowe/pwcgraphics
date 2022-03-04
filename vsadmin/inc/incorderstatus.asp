<!--#include file="incemail.asp"-->
<!--#include file="md5.asp"-->
<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if sDSN="" then response.write "Database connection not set" : response.end
Dim netnav, success
ordGrandTotal=0 : ordTotal=0 : ordStateTax=0 : ordHSTTax=0 : ordCountryTax=0 : ordShipping=0 : ordHandling=0 : ordDiscount=0
affilID="" : ordCity="" : ordState="" : ordCountry="" : ordDiscountText="" : ordEmail=""
if request.totalbytes > 10000 then response.end
success = true
digidownloads=false
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
if getpost("posted")="1" then
	email = replace(getpost("email"),"'","")
	ordid = replace(getpost("ordid"),"'","")
	if NOT is_numeric(ordid) then
		success = false
		errormsg = xxStaEr1
	elseif email<>"" AND ordid<>"" then
		sSQL = "SELECT ordStatus,ordStatusDate,"&getlangid("statPublic",64)&",ordTrackNum,ordAuthNumber,ordStatusInfo FROM orders INNER JOIN orderstatus ON orders.ordStatus=orderstatus.statID WHERE ordID=" & ordid & " AND ordEmail='" & email & "'"
		rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				ordStatus = rs("ordStatus")
				ordStatusDate = rs("ordStatusDate")
				statPublic = rs(getlangid("statPublic",64))
				ordAuthNumber = trim(rs("ordAuthNumber")&"")
				ordStatusInfo = trim(rs("ordStatusInfo")&"")
				ordTrackNum = trim(rs("ordTrackNum")&"")
				if trackingnumtext = "" then trackingnumtext=xxTrackT
				if ordTrackNum<>"" then trackingnum=replace(trackingnumtext, "%s", ordTrackNum) else trackingnum=""
				trackingnum = replace(trackingnum, "%nl%", "<br>")
				' if dateadjust<>"" then ordStatusDate = DateAdd("h",dateadjust,ordStatusDate)
			else
				success = false
				errormsg = xxStaEr2
			end if
		rs.close
	else
		success = false
		errormsg = xxStaEnt
	end if
end if
%>		<form method="post" name="statusform" action="orderstatus<%=extension%>">
		  <input type="hidden" name="posted" value="1" />
			<div class="ectdiv ectorderstatus">
<%		if getpost("posted")="1" AND success then %>
			  <div class="ectdivhead"><%=xxStaVw%></div>
			  <div class="ectdiv2column"><%=xxStaCur & " " & ordid%></div>
			  <div class="ectdivcontainer">
			    <div class="ectdivleft"><%=xxStatus%></div>
				<div class="ectdivright"><%=statPublic%></div>
			  </div>
			  <div class="ectdivcontainer">
			    <div class="ectdivleft"><%=xxDate%></div>
				<div class="ectdivright"><%=FormatDateTime(ordStatusDate, 1)%></div>
			  </div>
			  <div class="ectdivcontainer">
			    <div class="ectdivleft"><%=xxTime%></div>
				<div class="ectdivright"><%=FormatDateTime(ordStatusDate, 4)%></div>
			  </div>
<%			if trackingnum<>"" then %>
			  <div class="ectdivcontainer">
			    <div class="ectdivleft"><%=xxTraNum%></div>
				<div class="ectdivright"><%
				tracknumarr=split(trackingnum,",")
				for uoindex=0 to UBOUND(tracknumarr)
					thecarrier=getcarrierfromtrack(tracknumarr(uoindex),thelink)
					print "<p class=""tracknumline"">"
					if thelink<>"" then print "<a href=""" & thelink & tracknumarr(uoindex) & """ target=""_blank"">" & tracknumarr(uoindex) & "</a>" else print tracknumarr(uoindex)
					print "</p>"
				next
				%></div>
			  </div>
<%			end if
			if ordStatusInfo<>"" then %>
			  <div class="ectdivcontainer">
			    <div class="ectdivleft"><%=xxAddInf%></div>
				<div class="ectdivright"><%=ordStatusInfo%></div>
			  </div>
<%			end if 
			if ordAuthNumber>"" then %>
			  <div class="ectdiv2column"><%
					xxThkYou=""
					xxRecEml=""
					call do_order_success(ordid,"",FALSE,TRUE,FALSE,FALSE,FALSE) %>
			  </div>
<%			end if
		else %>
			  <div class="ectdivhead"><%=xxStaVw%></div>
<%		end if %>
			  <div class="ectdiv2column"><%=xxStaEnt%></div>
<%		if NOT success then %>
			  <div class="ectdivcontainer">
			    <div class="ectdivleft"><%=xxStaErr%></div>
				<div class="ectdivright ectwarning"><%=errormsg%></div>
			  </div>
<%		end if %>
			  <div class="ectdivcontainer">
			    <div class="ectdivleft"><%=xxOrdId%></div>
				<div class="ectdivright"><input type="text" size="20" name="ordid" value="<%=htmlspecials(getpost("ordid"))%>" /></div>
			  </div>
			  <div class="ectdivcontainer">
			    <div class="ectdivleft"><%=xxEmail%></div>
				<div class="ectdivright"><input type="text" size="30" name="email" value="<%=htmlspecials(getpost("email"))%>" /></div>
			  </div>
			  <div class="ectdiv2column"><%=imageorsubmit(imgvieworderstatus,xxStaVw,"vieworderstatus")%></div>
			</div>
		  </form>
<%
set rs = nothing
set cnn = nothing
%>