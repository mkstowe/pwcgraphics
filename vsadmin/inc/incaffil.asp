<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
Dim sSQL,rs,cnn,success,showaccount,addsuccess,alldata,index,allcountries,rowcounter,sd,ed,errmsg
Dim affilName,affilPW,affilAddress,affilCity,affilState,affilZip,affilCountry,affilEmail,affilInform,smonth
addsuccess = true
success = true
showaccount = true
dorefresh = FALSE
if dateadjust="" then dateadjust=0
thedate = DateAdd("h",dateadjust,Now())
thedate = DateSerial(year(thedate),month(thedate),day(thedate))
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
if getpost("act")="quickupdate" then
	for each objItem in request.form
		if left(objItem, 4)="pra_" then
			theid=right(objItem, len(objItem)-4)
			theval=getpost(objItem)
			pract=getpost("pract")
			if pract="del" then
				ect_query("DELETE FROM affiliates WHERE affilID='" & escape_string(theid) & "'")
			end if
		end if
	next
	dorefresh=TRUE
elseif getpost("editaction")="modify" then
	if getpost("affilid")<>getpost("origaffilid") then
		sSQL = "SELECT affilID FROM affiliates WHERE affilID='"&escape_string(getpost("affilid"))&"'"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then errmsg=yyAffDup : success=FALSE
		rs.close
	end if
	if success then
		sSQL = "UPDATE affiliates SET affilID='"&escape_string(getpost("affilid")) & "',"
		if getpost("affilpw")<>"" then sSQL = sSQL & "affilPW='"&escape_string(dohashpw(getpost("affilpw"))) & "',"
		sSQL = sSQL & "affilEmail='"&escape_string(getpost("email")) & "'," & _
			"affilName='"&escape_string(getpost("name")) & "'," & _
			"affilAddress='"&escape_string(getpost("address")) & "'," & _
			"affilCity='"&escape_string(getpost("city")) & "'," & _
			"affilState='"&escape_string(getpost("state")) & "'," & _
			"affilCountry='"&escape_string(getpost("country")) & "'," & _
			"affilZip='"&escape_string(getpost("zip")) & "',"
		if NOT is_numeric(getpost("affilcommision")) then
			sSQL = sSQL & "affilCommision=0,"
		else
			sSQL = sSQL & "affilCommision="&getpost("affilcommision")&","
		end if
		if getpost("affildate")<>"" then
			sSQL = sSQL & "affilDate=" & vsusdate(datevalue(getpost("affildate"))) & ","
		else
			sSQL = sSQL & "affilDate=" & vsusdate(date()) & ","
		end if
		sSQL = sSQL & "affilInform=" & IIfVr(getpost("Inform")="ON", "1 ", "0 ")
		sSQL = sSQL & "WHERE affilID='" & escape_string(getpost("origaffilid")) & "'"
		ect_query(sSQL)
		dorefresh=TRUE
	end if
elseif getpost("editaction")="addnew" then
	sSQL = "SELECT affilID FROM affiliates WHERE affilID='"&escape_string(getpost("affilid"))&"'"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then errmsg=yyAffDup : success=FALSE
	rs.close
	if success then
		sSQL = "INSERT INTO affiliates (affilID,affilPW,affilEmail,affilName,affilAddress,affilCity,affilState,affilCountry,affilZip,affilCommision,affilDate,affilInform) VALUES ("
		sSQL = sSQL & "'"&escape_string(getpost("affilid")) & "'," & _
			"'"&escape_string(dohashpw(getpost("affilpw"))) & "'," & _
			"'"&escape_string(getpost("email")) & "'," & _
			"'"&escape_string(getpost("name")) & "'," & _
			"'"&escape_string(getpost("address")) & "'," & _
			"'"&escape_string(getpost("city")) & "'," & _
			"'"&escape_string(getpost("state")) & "'," & _
			"'"&escape_string(getpost("country")) & "'," & _
			"'"&escape_string(getpost("zip")) & "',"
		if NOT is_numeric(getpost("affilcommision")) then
			sSQL = sSQL & "0,"
		else
			sSQL = sSQL & getpost("affilcommision")&","
		end if
		if getpost("affildate")<>"" then
			sSQL = sSQL & vsusdate(datevalue(getpost("affildate"))) & ","
		else
			sSQL = sSQL & vsusdate(date()) & ","
		end if
		sSQL = sSQL & IIfVr(getpost("Inform")="ON", "1 ", "0 ") & ")"
		ect_query(sSQL)
		dorefresh=TRUE
	end if
elseif getpost("editaction")="delete" then
	sSQL = "DELETE FROM affiliates WHERE affilID='" & escape_string(getpost("affilID")) & "'"
	ect_query(sSQL)
	dorefresh=TRUE
elseif getpost("editaction")="editaffil" then
	sSQL = "UPDATE orders SET ordAffiliate='"&escape_string(getpost("affilid"))&"' WHERE ordID=" & getpost("id")
	ect_query(sSQL)
elseif getpost("editaction")="removeaffil" then
	sSQL = "UPDATE orders SET ordAffiliate='' WHERE ordAffiliate='"&escape_string(getpost("affilid"))&"'"
	ect_query(sSQL)
end if
if dorefresh then
	print "<meta http-equiv=""refresh"" content=""1; url=adminaffil.asp"
	print "?stext=" & urlencode(getpost("stext")) & "&sd=" & request("sd") & "&ed=" & request("ed") & "&stype=" & getpost("stype") & "&resorder=" & getpost("resorder") & "&pg=1"
	print """>"
end if
if getpost("act")="modify" OR getpost("act")="addnew" then
	if getpost("act")="modify" then
		sSQL = "SELECT affilName,affilPW,affilAddress,affilCity,affilState,affilZip,affilCountry,affilEmail,affilInform,affilCommision,affilDate FROM affiliates WHERE affilID='"&escape_string(getpost("id"))&"'"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			affilID = getpost("id")
			affilName = rs("affilName")
			affilPW = ""
			affilAddress = rs("affilAddress")
			affilCity = rs("affilCity")
			affilState = rs("affilState")
			affilZip = rs("affilZip")
			affilCountry = rs("affilCountry")
			affilEmail = rs("affilEmail")
			affilInform = Int(rs("affilInform"))=1
			affilCommision = rs("affilCommision")
			affilDate = rs("affilDate")
		end if
		rs.close
	else
		affilID = ""
		affilName = ""
		affilPW = ""
		affilAddress = ""
		affilCity = ""
		affilState = ""
		affilZip = ""
		affilCountry = ""
		affilEmail = ""
		affilInform = 0
		affilCommision = 0
		affilDate = Date()
	end if
%>
<script>
<!--
function checkform(frm){
if(frm.affilid.value==""){
	alert("<%=jscheck(yyPlsEntr&" """&yyAffId)%>\".");
	frm.affilid.focus();
	return (false);
}
var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
var checkStr = frm.affilid.value;
var allValid = true;
for (i = 0;  i < checkStr.length;  i++){
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
}
if (!allValid){
    alert("<%=jscheck(yyOnlyAl&" """&yyAffId)%>\".");
    frm.affilid.focus();
    return (false);
}
<%	if getpost("act")<>"modify" then %>
if(frm.affilpw.value==""){
	alert("<%=jscheck(yyPlsEntr&" """&yyPass)%>\".");
	frm.affilpw.focus();
	return (false);
}
<%	end if %>
if(frm.name.value==""){
	alert("<%=jscheck(yyPlsEntr&" """&yyName)%>\".");
	frm.name.focus();
	return (false);
}
if(frm.email.value==""){
	alert("<%=jscheck(yyPlsEntr&" """&yyEmail)%>\".");
	frm.email.focus();
	return (false);
}
if(frm.address.value==""){
	alert("<%=jscheck(yyPlsEntr&" """&yyAddress)%>\".");
	frm.address.focus();
	return (false);
}
if(frm.city.value==""){
	alert("<%=jscheck(yyPlsEntr&" """&yyCity)%>\".");
	frm.city.focus();
	return (false);
}
if(frm.state.value==""){
	alert("<%=jscheck(yyPlsEntr&" """&yyState)%>\".");
	frm.state.focus();
	return (false);
}
if(frm.zip.value==""){
	alert("<%=jscheck(yyPlsEntr&" """&yyZip)%>\".");
	frm.zip.focus();
	return (false);
}
var checkOK = "0123456789.";
var checkStr = frm.affilcommision.value;
var allValid = true;
for (i = 0;  i < checkStr.length;  i++){
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
}
if (!allValid){
    alert("<%=jscheck(yyOnlyDec&" """&yyCommis)%>\".");
    frm.affilcommision.focus();
    return (false);
}
return (true);
}
//-->
</script>
		  <form method="post" action="adminaffil.asp" onsubmit="return checkform(this)">
			<input type="hidden" name="origaffilid" value="<%=htmlspecials(affilID)%>" />
			<input type="hidden" name="editaction" value="<%=IIfVr(getpost("act")="modify", "modify", "addnew")%>" />
			<input type="hidden" name="stext" value="<%=getpost("stext")%>" />
			<input type="hidden" name="sd" value="<%=getpost("sd")%>" />
			<input type="hidden" name="ed" value="<%=getpost("ed")%>" />
			<input type="hidden" name="resorder" value="<%=getpost("resorder")%>" />
			<input type="hidden" name="posted" value="1" />
			  <table width="100%" border="0" cellspacing="0" cellpadding="3">
				<tr>
				  <td width="100%" align="center" colspan="4"><strong><%=yyAffAdm%></strong></td>
				</tr>
<% if NOT addsuccess then %>
				<tr>
				  <td width="100%" align="center" colspan="4"><span style="color:#FF0000;font-weight:bold"><%=yyAffDup%></span></td>
				</tr>
<% end if %>
				<tr>
				  <td width="25%" align="right"><strong><%=redasterix&yyAffId%>:</strong></td>
				  <td width="25%" align="left"><input type="text" name="affilid" size="20" value="<%=htmlspecials(affilID)%>" /></td>
				  <td width="25%" align="right"><strong><%=IIfVr(getpost("act")="modify",yyReset&" "&yyPass,redasterix&yyPass)%>:</strong></td>
				  <td width="25%" align="left"><input type="text" name="affilpw" size="20" value="" /></td>
				</tr>
				<tr>
				  <td align="right"><strong><%=redasterix&yyName%>:</strong></td>
				  <td align="left"><input type="text" name="name" size="20" value="<%=htmlspecials(affilName)%>" /></td>
				  <td align="right"><strong><%=redasterix&yyEmail%>:</strong></td>
				  <td align="left"><input type="text" name="email" size="25" value="<%=htmlspecials(affilEmail)%>" /></td>
				</tr>
				<tr>
				  <td align="right"><strong><%=redasterix&yyAddress%>:</strong></td>
				  <td align="left"><input type="text" name="address" size="20" value="<%=htmlspecials(affilAddress)%>" /></td>
				  <td align="right"><strong><%=redasterix&yyCity%>:</strong></td>
				  <td align="left"><input type="text" name="city" size="20" value="<%=htmlspecials(affilCity)%>" /></td>
				</tr>
				<tr>
				  <td align="right"><strong><%=redasterix&yyState%>:</strong></td>
				  <td align="left"><input type="text" name="state" size="20" value="<%=htmlspecials(affilState)%>" /></td>
				  <td align="right"><strong><%=redasterix&yyCountry%>:</strong></td>
				  <td align="left"><select name="country" size="1">
<%
Sub show_countries(tcountry)
	if NOT IsArray(allcountries) then
		sSQL = "SELECT countryName FROM countries ORDER BY countryOrder DESC, countryName"
		rs.open sSQL,cnn,0,1
		allcountries=rs.getrows
		rs.close
	end if
	for rowcounter=0 to UBOUND(allcountries,2)
		print "<option value='" & htmlspecials(allcountries(0,rowcounter)) & "'"
		if tcountry=allcountries(0,rowcounter) then
			print " selected=""selected"""
		end if
		print ">"&allcountries(0,rowcounter)&"</option>"&vbCrLf
	next
End Sub
show_countries(affilCountry)
%>
					</select>
				  </td>
				</tr>
				<tr>
				  <td align="right"><strong><%=redasterix&yyZip%>:</strong></td>
				  <td align="left"><input type="text" name="zip" size="10" value="<%=htmlspecials(affilZip)%>" /></td>
				  <td align="right"><strong>Inform me:</strong></td>
				  <td align="left"><input type="checkbox" name="inform" value="ON" <% if affilInform then print "checked"%> /></td>
				</tr>
				<tr>
				  <td align="right"><strong><%=yyCommis%>:</strong></td>
				  <% session.LCID = 1033 %>
				  <td align="left"><input type="text" name="affilcommision" size="6" value="<%=htmlspecials(affilCommision)%>" />%</td>
				  <% session.LCID = saveLCID %>
				  <td align="right"><strong><%=yyDate%>:</strong></td>
				  <td align="left"><input type="text" name="affildate" size="10" value="<%=affilDate%>" /></td>
				</tr>
				<tr>
				  <td width="100%" colspan="4">
					<span style="font-size:10px"><ul><li><%=yyAffInf%></li></ul></span>
				  </td>
				</tr>
				<tr>
				  <td width="50%" align="center" colspan="4"><input type="submit" value="<%=yySubmit%>" /> <input type="reset" value="<%=yyReset%>" /> </td>
				</tr>
			  </table>
			</form>
<%
elseif getpost("posted")="1" AND success then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <a href="adminaffil.asp<%
							print "?rid="&getpost("rid")&"&stock="&getpost("stock")&"&stext=" & urlencode(getpost("stext")) & "&sd=" & getpost("sd") & "&ed=" & getpost("ed") & "&stype=" & getpost("stype") & "&approved=" & getpost("approved") & "&pg=" & getpost("pg")
						%>"><strong><%=yyClkHer%></strong></a>.<br />
                        <br />&nbsp;
                </td>
			  </tr>
			</table></td>
        </tr>
      </table>
<%
elseif getpost("posted")="1" then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><span style="color:#FF0000;font-weight:bold"><%=yyOpFai%></span><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table></td>
        </tr>
      </table>
<%
else
	pract=request.cookies("practaf")
	hasdaterange=FALSE
	if trim(Request("sd"))<>"" then sd=Request("sd") : hasdaterange=TRUE
	if trim(Request("ed"))="" then ed=thedate else ed=Request("ed")
	on error resume next
	sd = DateValue(sd)
	ed = Datevalue(ed)
	if err.number<>0 then
		hasdaterange=FALSE
		errmsg=yyDatInv
	end if
	on error goto 0
	if hasdaterange then
		tdt = DateValue(sd)
		tdt2 = DateValue(ed)+1
	end if
	
	sText = escape_string(request("stext"))
	findinvalids = (trim(request("stype"))="invalid")
	themask = cStr(DateSerial(2003,12,11))
	themask = replace(themask,"2003","yyyy")
	themask = replace(themask,"12","mm")
	themask = replace(themask,"11","dd")

	numaffiliates=0
	sSQL = "SELECT COUNT(*) AS thecount FROM affiliates"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		if NOT isnull(rs("thecount")) then numaffiliates=rs("thecount")
	end if
	rs.close

	if findinvalids then
		sSQL = "SELECT ordAffiliate,ordID,ordDate,ordReferer,ordQueryStr,ordTotal FROM orders LEFT JOIN affiliates ON orders.ordAffiliate=affiliates.affilID WHERE ordAffiliate<>'' AND NOT (ordAffiliate IS NULL) AND affilID IS NULL"
		if hasdaterange then sSQL = sSQL & " AND ordDate BETWEEN " & vsusdate(tdt)&" AND " & vsusdate(tdt2)
		if sText<>"" then sSQL = sSQL & " AND (ordAffiliate LIKE '%" & sText & "%' OR ordName LIKE '%" & sText & "%')"
		sSQL = sSQL & " ORDER BY ordID DESC"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then alldata=rs.GetRows()
		rs.close
	else
		affillist=""
		if hasdaterange then
			addcomma=""
			sSQL = "SELECT DISTINCT ordAffiliate FROM orders WHERE ordStatus>=3 AND ordAffiliate<>'' AND NOT (ordAffiliate IS NULL) AND ordDate BETWEEN " & vsusdate(tdt)&" AND " & vsusdate(tdt2)
			if sText<>"" then sSQL = sSQL & " AND ordAffiliate LIKE '%" & sText & "%'"
			rs.open sSQL,cnn,0,1
			do while NOT rs.EOF
				affillist=affillist&addcomma&"'"&replace(replace(rs("ordAffiliate"),"'",""),"<","")&"'"
				addcomma=","
				rs.movenext
			loop
			rs.close
		end if
		if affillist<>"" then
			sSQL = "SELECT affilID,affilName,affilPW,affilEmail,affilCommision,SUM(ordTotal-ordDiscount) AS affilQuant,affilDate FROM affiliates LEFT JOIN orders ON affiliates.affilID=orders.ordAffiliate WHERE ordStatus>=3 AND affilID IN ("&affillist&")"
			if hasdaterange then sSQL = sSQL & " AND ordDate BETWEEN " & vsusdate(tdt)&" AND " & vsusdate(tdt2)
			sSQL = sSQL & " GROUP BY affilID,affilName,affilPW,affilEmail,affilCommision,affilDate"
			if request("resorder")="1" then sSQL = sSQL & " ORDER BY affilID" else sSQL = sSQL & " ORDER BY SUM(ordTotal-ordDiscount) DESC"
		else
			sSQL = "SELECT affilID,affilName,affilPW,affilEmail,affilCommision,0 AS affilQuant,affilDate FROM affiliates"
			if sText<>"" then
				sSQL = sSQL & " WHERE affilID LIKE '%" & sText & "%' OR affilName LIKE '%" & sText & "%' OR affilEmail LIKE '%" & sText & "%'"
			end if
			sSQL = sSQL & " ORDER BY affilID"
		end if
		if NOT (hasdaterange AND affillist="") then
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then alldata=rs.GetRows()
			rs.close
		end if
	end if
%>
<script src="popcalendar.js"></script>
<script>
<!--
try{languagetext('<%=adminlang%>');}catch(err){}
function setCookie(c_name,value,expiredays){
	var exdate=new Date();
	exdate.setDate(exdate.getDate()+expiredays);
	document.cookie=c_name+ "=" +escape(value)+((expiredays==null) ? "" : ";expires="+exdate.toGMTString());
}
function mrec(id){
	document.mainform.action="adminaffil.asp";
	document.mainform.id.value=id;
	document.mainform.act.value="modify";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function newrec(id){
	document.mainform.action="adminaffil.asp";
	document.mainform.id.value=id;
	document.mainform.act.value="addnew";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function delrec(id) {
if (confirm("<%=jscheck(yyConDel)%>\n")) {
	document.mainform.affilid.value=id;
	document.mainform.act.value="search";
	document.mainform.editaction.value="delete";
	document.mainform.submit();
}
}
function dumpinventory(){
	document.mainform.action="dumporders.asp";
	document.mainform.act.value="dumpaffiliate";
	document.mainform.submit();
}
function startsearch(){
	document.mainform.action="adminaffil.asp";
	document.mainform.act.value="search";
	document.mainform.stock.value="";
	document.mainform.posted.value="";
	document.mainform.submit();
}
function quickupdate(){
	if(document.mainform.pract.value=="del"){
		if(!confirm("<%=jscheck(yyConDel)%>\n"))
			return;
	}
	document.mainform.action="adminaffil.asp";
	document.mainform.act.value="quickupdate";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function proccod(tmen,ordid,affid){
	theact=tmen[tmen.selectedIndex].value;
	if(theact=="1"){
		newwin=window.open("adminorders.asp?id="+ordid,"Orders","menubar=no, scrollbars=yes, width=800, height=680, directories=no,location=no,resizable=yes,status=no,toolbar=no");
	}else if(theact=="2"){
		if((affid=prompt("Please enter the new affiliate id for this order.",affid))!=null){
			document.mainform.action="adminaffil.asp";
			document.mainform.act.value="search";
			document.mainform.editaction.value="editaffil";
			document.mainform.id.value=ordid;
			document.mainform.affilid.value=affid;
			document.mainform.posted.value="";
			document.mainform.submit();
		}
	}else if(theact=="3"){
		if(confirm("<%=jscheck(yySureCa)%>")){
			document.mainform.action="adminaffil.asp";
			document.mainform.act.value="search";
			document.mainform.editaction.value="editaffil";
			document.mainform.id.value=ordid;
			document.mainform.affilid.value="";
			document.mainform.posted.value="";
			document.mainform.submit();
		}
	}else if(theact=="4"){
		if(confirm("Are you sure you want to remove all instances of affiliate code: "+affid)){
			document.mainform.action="adminaffil.asp";
			document.mainform.act.value="search";
			document.mainform.editaction.value="removeaffil";
			document.mainform.affilid.value=affid;
			document.mainform.posted.value="";
			document.mainform.submit();
		}
	}
	tmen.selectedIndex=0;
}
var currcheck=true;
function checkboxes(){
	if(document.getElementById("resultcounter")){
		maxitems=document.getElementById("resultcounter").value;
		for(index=0;index<maxitems;index++){
			document.getElementById("chkbx"+index).checked=currcheck;
		}
		currcheck=!currcheck;
	}
}
function changepract(obj){
	setCookie('practaf',obj[obj.selectedIndex].value,600);
	startsearch();
}
// -->
</script>
<h2><%=yyAdmAff&" ("&numaffiliates&")"%></h2>
	<form name="mainform" method="post" action="adminaffil.asp">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="" />
			<input type="hidden" name="stock" value="" />
			<input type="hidden" name="id" value="" />
			<input type="hidden" name="editaction" value="" />
			<input type="hidden" name="affilid" value="" />
			<input type="hidden" name="pg" value="<%=IIfVr(getpost("act")="search", "1", getget("pg"))%>" />
			<table class="cobtbl" width="100%" border="0" cellspacing="1" cellpadding="3">
			  <tr> 
				<td class="cobhl" width="20%" align="right"><%=yySrchFr%>:</td>
				<td class="cobll" width="30%"><input type="text" name="stext" size="20" value="<%=request("stext")%>" /></td>
				<td class="cobhl" width="20%" align="right"><%=yyAffBet%>:</td>
				<td class="cobll" width="30%" style="white-space:nowrap"><div style="position:relative;display:inline"><input type="text" name="sd" size="10" value="<%=request("sd")%>" style="vertical-align:middle" />&nbsp;<input type="button" onclick="popUpCalendar(this, document.forms.mainform.sd, '<%=themask%>', -205)" value="DP" />&nbsp;<%=yyAnd%>&nbsp;<input type="text" name="ed" size="10" value="<%=request("ed")%>" style="vertical-align:middle" />&nbsp;<input type=button onclick="popUpCalendar(this, document.forms.mainform.ed, '<%=themask%>', -205)" value="DP" /></div></td>
			  </tr>
			  <tr>
				<td class="cobhl"align="right"><%
					if pract="del" OR pract="app" then %>
						<input type="button" value="<%=yyCheckA%>" onclick="checkboxes(true);" style="float:left" />
<%					end if %><%=yySrchTp%>:</td>
				<td class="cobll"><select name="stype" size="1">
					<option value="">Valid Affiliates</option>
					<option value="invalid"<% if request("stype")="invalid" then print " selected=""selected"""%>>Invalid Affilates</option>
					</select>
				</td>
				<td class="cobhl"align="right"><%=yyResOrd%>:</td>
				<td class="cobll">
				  <select name="resorder" size="1">
				  <option value=""><%=yyTotSal%></option>
				  <option value="1" <% if request("resorder")="1" then print "selected=""selected"""%>><%=yyAffId%></option>
				  </select>
				</td>
			  </tr>
			  <tr>
				<td class="cobhl">&nbsp;</td>
				<td class="cobll" colspan="3"><table width="100%" cellspacing="0" cellpadding="0" border="0">
					<tr>
					  <td class="cobll" align="center"><input type="button" value="<%=yyListRe%>" onclick="startsearch();" /> 
						<input type="button" value="<%=yyNewAff%>" onclick="newrec();" />
						<input type="button" value="<%=yyAffRep%>" onclick="dumpinventory()" />
					  </td>
					  <td class="cobll" height="26" width="20%" align="right">&nbsp;</td>
					</tr>
				  </table></td>
			  </tr>
			</table>
<%	if request("act")="search" OR getget("pg")<>"" then
		resultcounter=0
		hasheader=FALSE
		if findinvalids then
			extcols=6
		else
			if hasdaterange then extcols=7 else extcols=5
		end if
%>
			<table width="100%" class="stackable admin-table-a sta-white">
<%		if findinvalids then %>
				<tr>
				  <th><strong><%=yyAffId%></strong></th>
				  <th align="center"><strong><%=yyOrdId%></strong></th>
				  <th align="center"><strong><%=yyDate%></strong></th>
				  <th align="center"><strong><%=yyWebURL%></strong></th>
				  <th align="right"><strong><%=yyAmount%></strong></th>
				  <th class="minicell"><strong><%=yyAct%></strong></th>
				</tr>
<%		else %>
				<tr>
				  <th class="minicell">
					<select name="pract" id="pract" size="1" onchange="changepract(this)">
					<option value="none">Quick Entry...</option>
					<option value="" disabled="disabled">------------------</option>
					<option value="del"<% if pract="del" then print " selected=""selected"""%>><%=yyDelete%></option>
					</select></th>
				  <th><strong><%=yyAffId%></strong></th>
				  <th><strong><%=yyName%></strong></th>
				  <th><strong><%=yyEmail%></strong></th>
<%			if hasdaterange then %>
				  <th align="right"><strong><%=replace(yyTotSal, " ", "&nbsp;")%></strong></th>
				  <th align="right"><strong><%=yyCommis%></strong></th>
<%			end if %>
				  <th class="minicell"><strong><%=yyDelete%></strong></th>
				</tr>
<%		end if
		if NOT IsArray(alldata) then %>
				<tr>
				  <td width="100%" align="center" colspan="<%=extcols%>"><br />&nbsp;<br /><strong><%=yyItNone%></strong><br />&nbsp;</td>
				</tr>
<%		else
			totsales=0
			totcomission=0
			hasheader=TRUE
			for index=0 to UBOUND(alldata,2)
				if bgcolor="altdark" then bgcolor="altlight" else bgcolor="altdark" %>
				<tr class="<%=bgcolor%>">
<%				if findinvalids then %>
				  <td><strong><%=htmlspecials(alldata(0,index))%></strong></td>
				  <td align="right"><%=htmlspecials(alldata(1,index))%>&nbsp;</td>
				  <td align="right"><%=FormatDateTime(alldata(2,index), 2)%>&nbsp;</td>
				  <td><%
						fullurl = alldata(3,index)&IIFVr(trim(alldata(4,index)&"")<>"", "?"&alldata(4,index), "")
						if fullurl<>"" then print "<a href="""&fullurl&""" title="""&fullurl&""" target=""_blank"">"&left(fullurl, 50)&IIFVr(len(fullurl)>50,"...","")&"</a>"
				%></td>
				  <td align="right"><%=FormatEuroCurrency(alldata(5,index))%></td>
				  <td align="right"><select size="1" onchange="proccod(this,'<%=alldata(1,index)%>','<%=htmlspecials(replace(alldata(0,index),"'","\'"))%>')">
				  <option value=""><%=yySelect%></option>
				  <option value="1"><%=yyVieDet%></option>
				  <option value="2">Edit Code</option>
				  <option value="3">Remove Code</option>
				  <option value="4">Remove All</option>
				  </select></td>
<%				else %>
				  <td class="minicell"><%
					if pract="del" then
						print "<input type=""checkbox"" id=""chkbx"&resultcounter&""" name=""pra_"&htmlspecials(alldata(0,index))&""" value=""del"" tabindex="""&(resultcounter+1)&"""/>"
					else
						print "&nbsp;"
					end if
				%></td><td><a href="javascript:mrec('<%=jsspecials(alldata(0,index))%>')"><strong><%=htmlspecials(alldata(0,index))%></strong></a>
<%					if date()-datevalue(alldata(6,index))<7 then print " <span style=""color:#FF0000"">" & "**"&yyNew&"**" & "</span>"%>
				  </td>
				  <td><%	print htmlspecials(alldata(1,index))
							 %></td>
				  <td><a href="mailto:<%=htmlspecials(alldata(3,index))%>"><%=htmlspecials(alldata(3,index))%></a></td>
<%					if hasdaterange then %>
				  <td align="right"><%if NOT is_numeric(alldata(5,index)) then print "-" else print FormatEuroCurrency(alldata(5,index)) : totsales = totsales + alldata(5,index) %></td>
				  <td align="right"><%if NOT is_numeric(alldata(5,index)) OR alldata(4,index)=0 then print "-" else print FormatEuroCurrency((alldata(4,index)*alldata(5,index)) / 100.0) : totcomission = totcomission + ((alldata(4,index)*alldata(5,index)) / 100.0)%></td>
<%					end if %>
				  <td class="minicell"><input type="button" value="<%=yyDelete%>" onclick="delrec('<%=jsspecials(alldata(0,index))%>')" /></td>
<%				end if %>
				</tr>
<%				resultcounter=resultcounter + 1
			next
 			if totsales>0 OR totcomission>0 then %>
				<tr><td colspan="3">&nbsp;</td><td align="right"><%=FormatEuroCurrency(totsales)%></td><td align="right"><%=FormatEuroCurrency(totcomission)%></td><td>&nbsp;</td></tr>
<%			end if
		end if %>
			  <tr>
<%		if hasheader then %>
				<td align="center" style="white-space:nowrap"><% if resultcounter>0 AND pract<>"" AND pract<>"none" then print "<input type=""hidden"" name=""resultcounter"" id=""resultcounter"" value="""&resultcounter&""" /><input type=""button"" value="""&yyUpdate&""" onclick=""quickupdate()"" /> <input type=""reset"" value="""&yyReset&""" />" else print "&nbsp;"%></td>
<%		end if %>
                <td width="100%" colspan="<%=extcols-IIfVr(hasheader,1,0)%>" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
			  </tr>
			</table>
<%	else
		numitems=0
		sSQL="SELECT COUNT(*) as totcount FROM affiliates"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			numitems=rs("totcount")
		end if
		rs.close
		print "<div class=""itemsdefine"">You have " & numitems & " affiliates defined.</div>"
	end if %>
	</form>
<%
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>