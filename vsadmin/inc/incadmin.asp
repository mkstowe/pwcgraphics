<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
Dim sSQL,rs,alldata,storeVersion,success,cnn,errtext
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
const ecthelpbaseurl="https://www.ecommercetemplates.com/help/ecommplus/"
success=0
set rs =Server.CreateObject("ADODB.RecordSet")
set rs2=Server.CreateObject("ADODB.RecordSet")
set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
on error resume next
sSQL = "UPDATE admin SET adminID=1 WHERE adminID=1"
ect_query(sSQL)
if err.number<>0 then
	errtext = errtext & "Your database is not writeable. This probably means you just need to set the permissions on the directory the database is in to allow writing.<br />Your host can help you with this.<br />" & vbCrLf
	success = -1
end if
on error goto 0
if debugmode=TRUE then
	errtext = errtext & yyDebug & "<br />" & vbCrLf
	success = -1
end if
if SESSION("loginid")=0 AND getget("act")="eventsclear" then
	sSQL="DELETE FROM auditlog"
	cnn.execute(sSQL) %>
	<div style="margin:100px 0;text-align:center">
		Your Event Log has been cleared.
	</div>
<%
	print "<meta http-equiv=""refresh"" content=""2; URL=admin.asp"" />"
elseif SESSION("loginid")=0 AND getget("act")="events" then
	call logevent(SESSION("loginuser"),"EVENTLOG",TRUE,"admin.asp","VIEW LOG")
	sSQL = "SELECT userID,eventType,eventDate,eventSuccess,eventOrigin,areaAffected FROM auditlog ORDER BY logID DESC"
%>
<div class="heading">
	<form method="post" action="dumporders.asp">
	<input type="hidden" name="act" value="dumpevents" />
	<input type="submit" value="Dump Event Log" /> Event Log
	
	<input type="button" value="Clear Event Log" onclick="if(confirm('Are you sure you want to clear the event log?')) document.location='admin.asp?act=eventsclear'" style="float:right" />
	</form>
</div>
<table width="100%" class="stackable admin-table-a">
  <thead>
	<tr>
	  <th scope="col">User ID</th>
	  <th scope="col">Event Type</th>
	  <th scope="col">Success</th>
	  <th scope="col">Origin</th>
	  <th scope="col">Area Affected</th>
	  <th scope="col">Date</th>
	</tr>
  </thead>
<%	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		do while NOT rs.EOF
			if rs("eventSuccess")<>0 then startfont="" : endfont="" else startfont="<span style=""color:#FF0000"">" : endfont="</span>" %>
  <tr>
	<td><%=startfont & htmlspecials(IIfVr(trim(rs("userID")&"")<>"",rs("userID"),"-")) & endfont%></td>
	<td><%=startfont & htmlspecials(IIfVr(trim(rs("eventType")&"")<>"",rs("eventType"),"-")) & endfont%></td>
	<td><%=startfont & htmlspecials(IIfVr(rs("eventSuccess")<>0,"TRUE","FALSE")) & endfont%></td>
	<td><%=startfont & htmlspecials(IIfVr(trim(rs("eventOrigin")&"")<>"",rs("eventOrigin"),"-")) & endfont%></td>
	<td><%=startfont & htmlspecials(IIfVr(trim(rs("areaAffected")&"")<>"",rs("areaAffected"),"-")) & endfont%></td>
	<td><%=startfont & htmlspecials(IIfVr(trim(rs("eventDate")&"")<>"",rs("eventDate"),"-")) & endfont%></td>
  </tr>
<%			rs.movenext
		loop
	else %>
  <tr>
    <td class="new" colspan="6" align="center">No events in log.</td>
  </tr>
<%	end if
	rs.close
%>
</table>
<%
else
	if dateadjust="" then dateadjust=0
	sSQL = "SELECT adminVersion,adminUser,adminPassword FROM admin WHERE adminID=1"
	rs.open sSQL,cnn,0,1
	storeVersion = rs("adminVersion")
	adminUser = rs("adminUser")
	adminPassword = rs("adminPassword")
	rs.close
	alreadygotadmin = getadminsettings()
	neworders = 0
	if getget("writeck")="no" then
		response.cookies("WRITECKL")=""
		response.cookies("WRITECKL").Expires = Date()-30
		response.cookies("WRITECKP")=""
		response.cookies("WRITECKP").Expires = Date()-30
		print "<meta http-equiv=""Refresh"" content=""3; URL=admin.asp"" />"
		success=1
	end if
	if Mid(SESSION("loggedonpermissions"),2,1)="X" then
%>
<div class="row">
	<div class="one_fourth home_boxes">
		<div class="full_width round_all box">
		<div class="box_title round-top" id="newordersdiv"><a href="<%=ecthelpbaseurl%>help.asp#orders" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a>
		<h3><a href="adminorders.asp"><%=yyVwOrd%></a></h3>
		</div>
		<div class="box_new" id="neworders" onclick="document.location='adminorders.asp'">-</div>
		</div>
		<div class="full_width round_all box">
		<div class="box_title round-top" id="newgiftcertdiv"><a href="<%=ecthelpbaseurl%>help.asp#giftcert" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a>
		<h3><a href="admingiftcert.asp">Gift Certificates</a></h3>
		</div>
		<div class="box_new" id="newgiftcert" onclick="document.location='admingiftcert.asp'">-</div>
		</div>
	</div>
	<div class="three_fourths last">
		<table id="latestorders" width="100%" class="quickstats neworders" style="margin-top:0;">
		<tr><th>Customer</th><th>Date</th><th>Status</th><th>Total</th></tr>
		</table>
	</div>
</div>
<div id="equalize" class="row home_boxes">
	<div class="one_sixth round_all box">
		<div class="box_title round-top" id="newaffiliatediv"><a href="<%=ecthelpbaseurl%>help.asp#affiliate" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a>
		<h3><a href="adminaffil.asp"><%=yyVwAff%></a></h3>
		</div>
		<div class="box_new" id="newaffiliate" onclick="document.location='adminaffil.asp'">-</div>
	</div>
	<div class="one_sixth round_all box">
		<div class="box_title round-top" id="newratingsdiv"><a href="<%=ecthelpbaseurl%>help.asp#ratings" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a>
		<h3><a href="adminratings.asp"><%=yyVwRat%></a></h3>
		</div>
		<div class="box_new" id="newratings" onclick="document.location='adminratings.asp'">-</div>
	</div>
	<div class="one_sixth last_third round_all box">
		<div class="box_title round-top" id="newaccountsdiv"><a href="<%=ecthelpbaseurl%>help.asp#clientlogin" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a>
		<h3><a href="adminclientlog.asp"><%=yyCliLog%></a></h3>
		</div>
		<div class="box_new" id="newaccounts" onclick="document.location='adminclientlog.asp'">-</div>
	</div>
	<%		if notifybackinstock then %>
	<div class="one_sixth round_all box">
		<div class="box_title round-top" id="newstocknotifydiv"><a href="<%=ecthelpbaseurl%>help.asp#notifystock" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a>
		<h3><a href="adminprods.asp?act=stknot"><%=yyStkNot%></a></h3>
		</div>
		<div class="box_new" id="newstocknotify" onclick="document.location='adminprods.asp?act=stknot'">-</div>
	</div>
	<%		end if
			if SESSION("loginid")=0 then %>
	<div class="one_sixth round_all box">
		<div class="box_title round-top" id="newlogeventsdiv"><a href="<%=ecthelpbaseurl%>help.asp#actlog" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a>
		<h3><a href="admin.asp?act=events">Activity Log</a></h3>
		</div>
		<div class="box_new" id="newlogevents" onclick="document.location='admin.asp?act=events'">-</div>
	</div>
	<%		end if %>
	<div class="one_sixth last round_all box">
		<div class="box_title round-top" id="newmaillistdiv"><a href="<%=ecthelpbaseurl%>help.asp#maillist" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a>
		<h3><a href="adminmailinglist.asp"><%=yyMaLiMa%></a></h3>
		</div>
		<div class="box_new" id="newmaillist" onclick="document.location='adminmailinglist.asp'">-</div>
	</div>
</div>
<%	end if
	if Mid(SESSION("loggedonpermissions"),15,1)="X" then %>
<div class="row">
<%
' this month, last month and this month last year order totals	
dim thismonthorders(2),lastmonthorders(2),yearorders(2),last12(2),thismonthtotal(2),lastmonthtotal(2),yeartotal(2),last12total(2)
if homeordersstatus<>"" then ordersstatus=homeordersstatus else ordersstatus="3,4,5,6,7,8,9,10,11,12,13,14,15,16,17"
for index=0 to 1
	thismonthorders(index)=0 : thismonthtotal(index)=0
	lastmonthorders(index)=0 : lastmonthtotal(index)=0
	yearorders(index)=0 : yeartotal(index)=0
	last12(index)=0 : last12total(index)=0
	alltime=0 : alltimetotal=0
next
sSQL = " SELECT COUNT(*) AS totalorders, SUM(ordTotal-ordDiscount) AS totalvalue FROM orders WHERE ordStatus IN (" & ordersstatus & ") AND ordDate BETWEEN " & vsusdate(dateserial(year(date()),month(date()),1))&" AND " & vsusdate(date()+1)
rs.open sSQL,cnn,0,1
if NOT rs.EOF then
	if NOT isnull(rs("totalvalue")) then thismonthorders(0)=rs("totalorders") : thismonthtotal(0)=rs("totalvalue")
end if
rs.close
sSQL = " SELECT COUNT(*) AS totalorders, SUM(ordTotal-ordDiscount) AS totalvalue FROM orders WHERE ordStatus IN (" & ordersstatus & ") AND ordDate BETWEEN " & vsusdate(dateserial(year(date())-1,month(date()),1))&" AND " & vsusdate(dateserial(year(date())-1,month(date())+1,1))
rs.open sSQL,cnn,0,1
if NOT rs.EOF then
	if NOT isnull(rs("totalvalue")) then thismonthorders(1)=rs("totalorders") : thismonthtotal(1)=rs("totalvalue")
end if
rs.close

sSQL = " SELECT COUNT(*) AS totalorders, SUM(ordTotal-ordDiscount) AS totalvalue FROM orders WHERE ordStatus IN (" & ordersstatus & ") AND ordDate BETWEEN " & vsusdate(dateserial(year(date()),month(date())-1,1))&" AND " & vsusdate(dateserial(year(date()),month(date()),1))
rs.open sSQL,cnn,0,1
if NOT rs.EOF then
	if NOT isnull(rs("totalvalue")) then lastmonthorders(0)=rs("totalorders") : lastmonthtotal(0)=rs("totalvalue")
end if
rs.close
sSQL = " SELECT COUNT(*) AS totalorders, SUM(ordTotal-ordDiscount) AS totalvalue FROM orders WHERE ordStatus IN (" & ordersstatus & ") AND ordDate BETWEEN " & vsusdate(dateserial(year(date())-1,month(date())-1,1))&" AND " & vsusdate(dateserial(year(date())-1,month(date()),1))
rs.open sSQL,cnn,0,1
if NOT rs.EOF then
	if NOT isnull(rs("totalvalue")) then lastmonthorders(1)=rs("totalorders") : lastmonthtotal(1)=rs("totalvalue")
end if
rs.close

sSQL = " SELECT COUNT(*) AS totalorders, SUM(ordTotal-ordDiscount) AS totalvalue FROM orders WHERE ordStatus IN (" & ordersstatus & ") AND ordDate BETWEEN " & vsusdate(dateserial(year(date()),1,1))&" AND " & vsusdate(date()+1)
rs.open sSQL,cnn,0,1
if NOT rs.EOF then
	if NOT isnull(rs("totalvalue")) then yearorders(0)=rs("totalorders") : yeartotal(0)=rs("totalvalue")
end if
rs.close
sSQL = " SELECT COUNT(*) AS totalorders, SUM(ordTotal-ordDiscount) AS totalvalue FROM orders WHERE ordStatus IN (" & ordersstatus & ") AND ordDate BETWEEN " & vsusdate(dateserial(year(date())-1,1,1))&" AND " & vsusdate(dateserial(year(date())-1,month(date()),day(date())))
rs.open sSQL,cnn,0,1
if NOT rs.EOF then
	if NOT isnull(rs("totalvalue")) then yearorders(1)=rs("totalorders") : yeartotal(1)=rs("totalvalue")
end if
rs.close

sSQL = " SELECT COUNT(*) AS totalorders, SUM(ordTotal-ordDiscount) AS totalvalue FROM orders WHERE ordStatus IN (" & ordersstatus & ") AND ordDate BETWEEN " & vsusdate(dateserial(year(date())-1,month(date()),day(date())))&" AND " & vsusdate(date()+1)
rs.open sSQL,cnn,0,1
if NOT rs.EOF then
	if NOT isnull(rs("totalvalue")) then last12(0)=rs("totalorders") : last12total(0)=rs("totalvalue")
end if
rs.close
sSQL = " SELECT COUNT(*) AS totalorders, SUM(ordTotal-ordDiscount) AS totalvalue FROM orders WHERE ordStatus IN (" & ordersstatus & ") AND ordDate BETWEEN " & vsusdate(dateserial(year(date())-2,month(date()),day(date())))&" AND " & vsusdate(dateserial(year(date())-1,month(date()),day(date())))
rs.open sSQL,cnn,0,1
if NOT rs.EOF then
	if NOT isnull(rs("totalvalue")) then last12(1)=rs("totalorders") : last12total(1)=rs("totalvalue")
end if
rs.close

sSQL = " SELECT COUNT(*) AS totalorders, SUM(ordTotal-ordDiscount) AS totalvalue FROM orders WHERE ordStatus IN (" & ordersstatus & ")"
rs.open sSQL,cnn,0,1
if NOT rs.EOF then
	if NOT isnull(rs("totalvalue")) then alltime=rs("totalorders") : alltimetotal=rs("totalvalue")
end if
rs.close

' END this month, last month and this month last year order totals

lmdate=date()-day(date())
if is_numeric(getget("tsdays")) then
	tsdays=cint(getget("tsdays"))
	call setacookie("tsdays",tsdays,365)
else
	tsdays=request.cookies("tsdays")
	if NOT is_numeric(tsdays) then tsdays=30
end if
dbdate=date()-tsdays
sincedate=date()-tsdays
dbweekdate=date()-7
thismonth=monthname(month(date()),TRUE)
lastmonth=monthname(month(lmdate),TRUE)
lastyear=monthname(month(lmdate),TRUE)&" "&(year(date())-1)
%>
<div class="one_third">
<h3 class="round_top">Order Stats<br><span style="font-weight:normal;font-size:10px">This month, last month, last year</span></h3>
<table class="quickstats">
<tr><th>&nbsp;</th><th>&nbsp;</th><th style="text-align:right;">No</th><th style="text-align:right;">Value</th></tr>
<tr>
	<td style="text-align:right;vertical-align:middle" rowspan="2"><%=thismonth%></td>
	<td style="text-align:right;font-size:0.9em">This Year</td>
	<td style="text-align:right;font-size:0.9em" title="<%=FormatEuroCurrency(thismonthtotal(0)/IIfVr(thismonthorders(0)=0,1,thismonthorders(0)))%>"><%=thismonthorders(0) & "</td><td style=""text-align:right;font-size:0.9em"">" & FormatEuroCurrency(thismonthtotal(0))%></td>
</tr>
<tr>
	<td style="text-align:right;font-size:0.9em;color:#FF6060">Last Year</td>
	<td style="text-align:right;font-size:0.9em;color:#FF6060" title="<%=FormatEuroCurrency(thismonthtotal(1)/IIfVr(thismonthorders(1)=0,1,thismonthorders(1)))%>"><%=thismonthorders(1) & "</td><td style=""text-align:right;font-size:0.9em;color:#FF6060"">" & FormatEuroCurrency(thismonthtotal(1))%></td>
</tr>
<tr>
	<td style="text-align:right;vertical-align:middle" rowspan="2"><%=lastmonth%></td>
	<td style="text-align:right;font-size:0.9em">This Year</td>
	<td style="text-align:right;font-size:0.9em" title="<%=FormatEuroCurrency(lastmonthtotal(0)/IIfVr(lastmonthorders(0)=0,1,lastmonthorders(0)))%>"><%=lastmonthorders(0) & "</td><td style=""text-align:right;font-size:0.9em"">" & FormatEuroCurrency(lastmonthtotal(0))%></td>
</tr>
<tr>
	<td style="text-align:right;font-size:0.9em;color:#FF6060">Last Year</td>
	<td style="text-align:right;font-size:0.9em;color:#FF6060" title="<%=FormatEuroCurrency(lastmonthtotal(1)/IIfVr(lastmonthorders(1)=0,1,lastmonthorders(1)))%>"><%=lastmonthorders(1) & "</td><td style=""text-align:right;font-size:0.9em;color:#FF6060"">" & FormatEuroCurrency(lastmonthtotal(1))%></td>
</tr>
<tr>
	<td style="text-align:right;vertical-align:middle" rowspan="2"><div title="January 1 - Now"><%="Jan 1 &raquo;"%></div></td>
	<td style="text-align:right;font-size:0.9em">This Year</td>
	<td style="text-align:right;font-size:0.9em" title="<%=FormatEuroCurrency(yeartotal(0)/IIfVr(yearorders(0)=0,1,yearorders(0)))%>"><%=yearorders(0) & "</td><td style=""text-align:right;font-size:0.9em"">" & FormatEuroCurrency(yeartotal(0))%></td>
</tr>
<tr>
	<td style="text-align:right;font-size:0.9em;color:#FF6060">Last Year</td>
	<td style="text-align:right;font-size:0.9em;color:#FF6060" title="<%=FormatEuroCurrency(yeartotal(1)/IIfVr(yearorders(1)=0,1,yearorders(1)))%>"><%=yearorders(1) & "</td><td style=""text-align:right;font-size:0.9em;color:#FF6060"">" & FormatEuroCurrency(yeartotal(1))%></td>
</tr>
<tr>
	<td style="text-align:right;vertical-align:middle" rowspan="2"><div title="Last 12 Months"><%="12 Mo."%></div></td>
	<td style="text-align:right;font-size:0.9em">This Year</td>
	<td style="text-align:right;font-size:0.9em" title="<%=FormatEuroCurrency(last12total(0)/IIfVr(last12(0)=0,1,last12(0)))%>"><%=last12(0) & "</td><td style=""text-align:right;font-size:0.9em"">" & FormatEuroCurrency(last12total(0))%></td>
</tr>
<tr>
	<td style="text-align:right;font-size:0.9em;color:#FF6060">Last Year</td>
	<td style="text-align:right;font-size:0.9em;color:#FF6060" title="<%=FormatEuroCurrency(last12total(1)/IIfVr(last12(1)=0,1,last12(1)))%>"><%=last12(1) & "</td><td style=""text-align:right;font-size:0.9em;color:#FF6060"">" & FormatEuroCurrency(last12total(1))%></td>
</tr>
<tr>
	<td style="text-align:right;white-space:nowrap"><div title="All Time Sales"><%="All Time"%></div></td>
	<td style="font-size:0.9em">&nbsp;</td>
	<td style="text-align:right;font-size:0.9em" title="<%=FormatEuroCurrency(alltimetotal/IIfVr(alltime=0,1,alltime))%>"><%=alltime & "</td><td style=""text-align:right;font-size:0.9em"">" & FormatCurrencyZeroDP(alltimetotal)%></td>
</tr>
</table>
</div>

<div class="one_third">
<h3 class="round_top">Top Sellers: Last <select size="1" name="tsdays" class="adminhomestats" onchange="document.location='admin.asp?tsdays='+this[this.selectedIndex].value">
		<option value="1"<% if tsdays=1 then print " selected=""selected"""%>>1</option>
		<option value="7"<% if tsdays=7 then print " selected=""selected"""%>>7</option>
		<option value="30"<% if tsdays=30 then print " selected=""selected"""%>>30</option>
		<option value="60"<% if tsdays=60 then print " selected=""selected"""%>>60</option>
		<option value="90"<% if tsdays=90 then print " selected=""selected"""%>>90</option>
		<option value="180"<% if tsdays=180 then print " selected=""selected"""%>>180</option>
		<option value="365"<% if tsdays=365 then print " selected=""selected"""%>>365</option>
		<option value="0"<% if tsdays=0 then print " selected=""selected"""%>>All</option>
		</select> days<br><span style="font-weight:normal;font-size:10px">Since <%=sincedate%></span></h3>
<table class="quickstats">
<tr><th style="width:10%"></th><th style="text-align:left;">Prod</th><th>Sold</th><th style="text-align:right;">Value</th></tr>
<%
count=1
prevbought=0
sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 10 ")&"cartProdId,cartProdName, SUM(cartQuantity) AS numbought, SUM(cartProdPrice*cartQuantity) AS totalvalue FROM cart INNER JOIN orders ON cart.cartOrderID=orders.ordID WHERE "
if tsdays>0 then sSQL=sSQL&"cartDateAdded>=" & vsusdate(dbdate)&" AND "
sSQL=sSQL&"cartCompleted=1 AND ordStatus>=3 GROUP BY cartProdId,cartProdName ORDER BY SUM(cartQuantity) DESC, SUM(cartProdPrice*cartQuantity) DESC"&IIfVs(mysqlserver=TRUE," LIMIT 0,10")
rs.open sSQL,cnn,0,1
do while NOT rs.EOF
	thelink=""
	sSQL="SELECT pID,"&getlangid("pName",1)&",pStaticPage,pStaticURL,pDisplay FROM products WHERE pDisplay<>0 AND pID='"&escape_string(rs("cartProdId"))&"'"
	rs2.open sSQL,cnn,0,1
	if NOT rs2.EOF then thelink=storeurl & getdetailsurl(rs2("pID"),rs2("pStaticPage"),rs2(getlangid("pName",1)),trim(rs2("pStaticURL")&""),"","")
	rs2.close
	print "<tr><td style=""width:10%""><strong>"&count&"</strong></td><td style=""text-align:left;""><a href=""" & thelink & """ target=""_blank"">"&rs("cartProdName")&"</a></td><td>"&rs("numbought")&"</td><td style=""text-align:right;"">"&FormatEuroCurrency(rs("totalvalue"))&"</td></tr>"
	count=count+1
	rs.movenext
loop
rs.close
%>
</table>
</div>
<div class="one_third last">
<h3 class="round_top">Top Customers: Last <%=tsdays%> days<br><span style="font-weight:normal;font-size:10px">Since <%=sincedate%></span></h3>
<table class="quickstats">
<tr><th style="width:10%"></th><th style="text-align:left;">Customer</th><th style="text-align:right;">Spent</th></tr>
<%
count=1
sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 10 ")&"ordName,ordLastName,SUM(ordTotal-ordDiscount) AS spent FROM orders WHERE ordStatus IN (" & ordersstatus & ")"
if tsdays>0 then sSQL=sSQL&"AND ordDate>=" & vsusdate(dbdate)&" "
sSQL=sSQL&"GROUP BY ordName,ordLastName ORDER BY SUM(ordTotal-ordDiscount) DESC"&IIfVs(mysqlserver=TRUE," LIMIT 0,10")
rs.open sSQL,cnn,0,1
do while NOT rs.EOF
	print "<tr><td style=""width:10%""><strong>"&count&"</strong></td><td style=""text-align:left;"">"&trim(rs("ordName")&" "&rs("ordLastName"))&"</td><td style=""text-align:right;"">"&FormatEuroCurrency(rs("spent"))&"</td></tr>"
	count=count+1
	rs.movenext
loop
rs.close
%>
</table>
</div>
</div>
<%	end if
	if Mid(SESSION("loggedonpermissions"),2,1)="X" then
%>
<script>
/* <![CDATA[ */
var dashajob;
function updatedashboardcb(){
	if(dashajob.readyState==4){
		var dbarray=dashajob.responseText.split("&"),newrow,newcell;
		document.getElementById('neworders').innerHTML=dbarray[0];
		if(dbarray[0]>0&&document.getElementById('newordersdiv').className.indexOf('new_alert')<0)document.getElementById('newordersdiv').className+=' new_alert';
		document.getElementById('newratings').innerHTML=dbarray[1];
		if(dbarray[1]>0&&document.getElementById('newratingsdiv').className.indexOf('new_alert')<0)document.getElementById('newratingsdiv').className+=' new_alert';
		document.getElementById('newaccounts').innerHTML=dbarray[2];
		if(dbarray[2]>0&&document.getElementById('newaccountsdiv').className.indexOf('new_alert')<0)document.getElementById('newaccountsdiv').className+=' new_alert';
		document.getElementById('newmaillist').innerHTML=dbarray[3];
		if(dbarray[3]>0&&document.getElementById('newmaillistdiv').className.indexOf('new_alert')<0)document.getElementById('newmaillistdiv').className+=' new_alert';
		document.getElementById('newaffiliate').innerHTML=dbarray[4];
		if(dbarray[4]>0&&document.getElementById('newaffiliatediv').className.indexOf('new_alert')<0)document.getElementById('newaffiliatediv').className+=' new_alert';
		document.getElementById('newgiftcert').innerHTML=dbarray[5];
		if(dbarray[5]>0&&document.getElementById('newgiftcertdiv').className.indexOf('new_alert')<0)document.getElementById('newgiftcertdiv').className+=' new_alert';
<%		if notifybackinstock then %>
		document.getElementById('newstocknotify').innerHTML=dbarray[6];
		if(dbarray[6]>0&&document.getElementById('newstocknotifydiv').className.indexOf('new_alert')<0)document.getElementById('newstocknotifydiv').className+=' new_alert';
<%		end if
		if SESSION("loginid")=0 then %>
		document.getElementById('newlogevents').innerHTML=dbarray[7];
		if(dbarray[7]>0&&document.getElementById('newlogeventsdiv').className.indexOf('new_alert')<0)document.getElementById('newlogeventsdiv').className+=' new_alert';
<%		end if %>
		var ordtable=document.getElementById('latestorders');
		for(var dbind=0;dbind<dbarray.length-8;dbind++){
			var orddetails=dbarray[8+dbind].split('|');
			if(ordtable.rows.length<dbind+2){
				newrow=ordtable.insertRow(-1);
				for(var ncind=0;ncind<4;ncind++)newrow.insertCell(-1);
			}else
				newrow=ordtable.rows[dbind+1];
			var ordid=orddetails[0];
			newrow.setAttribute('onclick', 'document.location="adminorders.asp?id='+ordid+'"');
			newrow.cells[0].innerHTML='<a href="adminorders.asp?id='+ordid+'">'+decodeURIComponent(orddetails[1])+'</a>';
			newrow.cells[1].innerHTML=decodeURIComponent(orddetails[3]);
			newrow.cells[2].innerHTML=decodeURIComponent(orddetails[4]);
			newrow.cells[3].innerHTML=decodeURIComponent(orddetails[5]);
		}
		setTimeout(updatedashboard,90000);
	}
}
function updatedashboard(){
	dashajob=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
	dashajob.onreadystatechange=updatedashboardcb;
	dashajob.open("GET", "ajaxservice.asp?action=dashboard", true);
	dashajob.send(null);
}
updatedashboard();
/* ]]> */
</script>
<%	end if %>
<div class="row">
<h3 class="round_top"><%=yyStoAdm%></h3>
<table width="100%" class="admin-table-b">
  <thead>
	<tr>
	  <th scope="col"><%=yyAdmLnk%></th>
	  <th scope="col"><%=yyDesc%></th>
	  <th scope="col"><%=yyHlpFil%></th>
	</tr>
  </thead>
  <tr>
    <td><a href="adminmain.asp"><%=yyEdAdm%></a></td>
    <td><%=yyDBGlob%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#admin" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
 <tr>
    <td><a href="adminlogin.asp"><%=yyCngPw%></a></td>
    <td><%=yyDBLogA%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#uname" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
     <tr>
    <td><a href="adminpayprov.asp"><%=yyEdPPro%></a></td>
    <td><%=yyDBConP%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#payprov" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
    <tr>
    <td><a href="adminordstatus.asp"><%=yyEdOSta%></a></td>
    <td><%=yyDBConO%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#ordstat" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
  <tr>
    <td><a href="adminemailmsgs.asp"><%=yyEmlAdm%></a></td>
    <td><%=yyDBConE%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#emailadmin" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
  <tr>
    <td><a href="admincontent.asp"><%=yyContReg%></a></td>
    <td><%=yyContExp%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#contreg" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
  <tr>
    <td><a href="adminipblock.asp"><%=yyIPBlock%></a></td>
    <td><%=yyDBBkIP%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#ipblock" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
</table>
	
<h3 class="round_top double_top"><%=yyPrdAdm%></h3>
<table width="100%" class="admin-table-b">
  <thead>
	<tr>
	  <th scope="col"><%=yyAdmLnk%></th>
	  <th scope="col"><%=yyDesc%></th>
	  <th scope="col"><%=yyHlpFil%></th>
	</tr>
  </thead>
  <tr>
    <td><a href="adminprods.asp"><%=yyEdPrd%></a></td>
    <td><%=yyDBMaPI%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#prods" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
    <tr>
    <td><a href="adminprodopts.asp"><%=yyEdOpt%></a></td>
    <td><%=yyDBPrAt%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#prodopt" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
  <tr>
    <td><a href="admincats.asp"><%=yyEdCat%></a></td>
    <td><%=yyDBCats%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#cats" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
  <tr>
    <td><a href="admindiscounts.asp"><%=yyDisCou%></a></td>
    <td><%=yyDBSOFS%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#discounts" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
  <tr>
    <td><a href="adminpricebreak.asp"><%=yyEdPrBk%></a></td>
    <td><%=yyDBBuPr%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#pricebreak" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
  <tr>
    <td><a href="admingiftcert.asp"><%=yyGCMan%></a></td>
    <td><%=yyDBGifC%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#giftcert" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
  <tr>
    <td><a href="adminmanufacturer.asp"><%=yyEdManu%></a></td>
    <td><%=yyDBManD%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#manuf" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
  <tr>
    <td><a href="adminsearchcriteria.asp"><%=yyEdSeCr%></a></td>
    <td><%=yyCrSeCr%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#searcr" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
  <tr>
    <td><a href="admincsv.asp"><%=yyCSVUpl%></a></td>
    <td><%=yyDBBUpI%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#csv" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
</table>	
	
<h3 class="round_top double_top"><%=yyShpAdm%></h3>
<table width="100%" class="admin-table-b">
  <thead>
	<tr>
	  <th scope="col"><%=yyAdmLnk%></th>
	  <th scope="col"><%=yyDesc%></th>
	  <th scope="col"><%=yyHlpFil%></th>
	</tr>
  </thead>
  <tr>
    <td><a href="adminstate.asp"><%=yyEdSta%></a></td>
    <td><%=yyDBStat%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#state" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
  <tr>
    <td><a href="admincountry.asp"><%=yyEdCnt%></a></td>
    <td><%=yyDBCoun%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#country" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
  <tr>
    <td><a href="adminzones.asp"><%=yyEdPzon%></a></td>
    <td><%=yyDBSZon%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#pzone" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
  <tr>
    <td><a href="adminuspsmeths.asp"><%=yyShmReg%></a></td>
    <td><%=yyDBMSHO%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#shipmeth" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
  <tr>
     <td><a href="admindropship.asp"><%=yyEdDrSp%></a></td>
    <td><%=yyDBDSDe%></td>
    <td><a href="<%=ecthelpbaseurl%>help.asp#droshp" target="ttshelp" class="online_help" title="<%=yyOnlHlp%>">?</a></td>
  </tr>
</table>

<h3 class="round_top double_top">Debug Info</h3>
<table width="100%" class="admin-table-b">
  <tr>
    <td>Server Software:</td>
    <td><%=response.write(Request.ServerVariables("SERVER_SOFTWARE"))%></td>
  </tr>
  <tr>
    <td>VBScript Version:</td>
    <td><%=ScriptEngine & " " & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion%></td>
  </tr>
<%	if sqlserver AND NOT mysqlserver then %>
  <tr>
    <td>SQL Server Version:</td>
    <td><%
		on error resume next
		sSQL="SELECT @@VERSION AS serverversion"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then response.write left(rs("serverversion")&"",100)&"..."
		rs.close
		on error goto 0
	%></td>
  </tr>
<%	end if %>
</table>
<%
	sSQL = "SELECT modkey,modtitle,modauthor,modauthorlink,modversion,modectversion,modlink,moddate FROM installedmods ORDER BY moddate"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		print "<table width=""98%"" align=""center"">"
		print "<tr><td align=""center"" colspan=""2"">&nbsp;<br /><strong>---------------| Installed 3rd Party MODs |---------------<br />&nbsp;</strong></td></tr>"
		print "<tr><td align=""center"" colspan=""2""><table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">"
		print "<tr><td align=""left""><strong>Title</strong></td><td align=""left""><strong>Author</strong></td><td align=""left""><strong>MOD Version</strong></td><td align=""left""><strong>ECT Version</strong></td><td align=""left""><strong>Admin Link</strong></td><td align=""left""><strong>Install Date</strong></td></tr>"
		do while NOT rs.EOF
			modauthorlink=trim(rs("modauthorlink")&"")
			print "<tr><td align=""left"">" & rs("modtitle") & "</td>"
			print "<td align=""left"">" & IIfVs(modauthorlink<>"","<a href=""" & IIfVs(left(modauthorlink,7)<>"http://" AND left(modauthorlink,8)<>"https://","http://") & modauthorlink & """ target=""_blank"">") & rs("modauthor") & IIfVs(modauthorlink<>"","</a>") & "</td>"
			print "<td align=""left"">" & rs("modversion") & "</td>"
			print "<td align=""left"">" & rs("modectversion") & "</td>"
			print "<td align=""left""><strong>" & IIfVr(trim(rs("modlink")&"")<>"","<a href=""" & rs("modlink") & """>Admin Page</a>","&nbsp;") & "</strong></td>"
			print "<td align=""left"">" & FormatDateTime(rs("moddate"),2) & "</td>"
			rs.movenext
		loop
		print "</table><br />&nbsp;</td></tr></table>"
	end if
	rs.close
%>
</div>
<%		if nocheckdatabasedownload<>TRUE AND sqlserver<>TRUE AND mysqlserver<>TRUE AND success<=0 then %>
<script>
/* <![CDATA[ */
function getwarnmessage(){
	return('<p><span style="color:#FF0000;font-weight:bold">WARNING!!</span> It may be that your database is downloadable. This may mean that someone could download your database and gain access to your admin username and password. For more details please visit <a href="https://www.ecommercetemplates.com/help/checklist.asp#asp">https://www.ecommercetemplates.com/help/checklist.asp#asp</a></p>');
}
function checkstatechange(){
	if(ckAJAX.readyState==4){
		if(ckAJAX.status==200){
			document.getElementById("testspanid").innerHTML=getwarnmessage();
		}
	}
}
function checkstatechange2(){
	if(ckAJAX2.readyState==4){
		if(ckAJAX2.status==200){
			document.getElementById("testspanid").innerHTML=getwarnmessage();
		}
	}
}
if(window.XMLHttpRequest){
	ckAJAX = new XMLHttpRequest();
	ckAJAX2 = new XMLHttpRequest();
}else{
	ckAJAX = new ActiveXObject("MSXML2.XMLHTTP");
	ckAJAX2 = new ActiveXObject("MSXML2.XMLHTTP");
}
ckAJAX.onreadystatechange = checkstatechange;
ckAJAX.open("GET", "../fpdb/vsproducts.mdb", true);
ckAJAX.send(null);
setTimeout('ckAJAX.abort();',1000);
ckAJAX2.onreadystatechange = checkstatechange2;
ckAJAX2.open("GET", "/fpdb/vsproducts.mdb", true);
ckAJAX2.send(null);
setTimeout('ckAJAX2.abort();',1100);
/* ]]> */
</script>
<%		end if
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>