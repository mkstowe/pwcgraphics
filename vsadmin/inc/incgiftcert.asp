<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
Dim aFields(4)
success=true
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs2=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
Session.LCID=1033
dorefresh=FALSE
sub dodeletecert(gcid)
	sSQL="SELECT gcCartID FROM giftcertificate WHERE gcID='" & escape_string(gcid) & "'"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		cartID=rs("gcCartID")
		if cartID<>0 then ect_query("DELETE FROM cart WHERE cartCompleted=0 AND cartID="&cartID)
	end if
	rs.close
	sSQL="DELETE FROM giftcertificate WHERE gcID='" & escape_string(gcid) & "'"
	ect_query(sSQL)
	sSQL="DELETE FROM giftcertsapplied WHERE gcaGCID='" & escape_string(gcid) & "'"
	ect_query(sSQL)
end sub
if getpost("posted")="1" OR getget("act")="deleteassoc" then
	if getpost("act")="confirm" then
		sSQL="UPDATE giftcertificate SET gcAuthorized=1 WHERE gcID='" & escape_string(getpost("id")) & "'"
		ect_query(sSQL)
	elseif getpost("act")="delete" then
		call dodeletecert(getpost("id"))
		dorefresh=TRUE
	elseif getpost("act")="quickupdate" then
		for each objItem in request.form
			if left(objItem, 4)="pra_" then
				theid=right(objItem, len(objItem)-4)
				theval=getpost(objItem)
				pract=getpost("pract")
				sSQL=""
				if pract="del" then
					call dodeletecert(theid)
					sSQL=""
				end if
				if sSQL<>"" then
					sSQL=sSQL & " WHERE rtID="&theid
					ect_query(sSQL)
				end if
			end if
		next
		dorefresh=TRUE
	elseif getget("act")="deleteassoc" then
		if getget("refund")="true" then
			sSQL="SELECT gcaAmount FROM giftcertsapplied WHERE gcaGCID='" & getget("id") & "' AND gcaOrdID=" & getget("ord")
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				sSQL="UPDATE giftcertificate SET gcRemaining=gcRemaining+" & rs("gcaAmount") & " WHERE gcID='" & getget("id") & "'"
				ect_query(sSQL)
			end if
			rs.close
		end if
		sSQL="DELETE FROM giftcertsapplied WHERE gcaGCID='" & getget("id") & "' AND gcaOrdID=" & getget("ord")
		ect_query(sSQL)
	elseif getpost("act")="doaddnew" then
		sSQL="SELECT gcID FROM giftcertificate WHERE gcID='" & ucase(replace(getpost("gcid"),"'","")) & "'"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then success=FALSE : errmsg="Duplicate Gift Certificate ID"
		rs.close
		if success then
			session.lcid=saveLCID
			if getpost("gcdatecreated")<>"" AND isdate(getpost("gcdatecreated")) then datecreated=vsusdate(datevalue(replace(getpost("gcdatecreated"),"'",""))) else datecreated=vsusdate(Date())
			if getpost("gcdateused")<>"" AND isdate(getpost("gcdateused")) then dateused=vsusdate(datevalue(getpost("gcdateused"))) else dateused=""
			session.LCID=1033
			sSQL="INSERT INTO giftcertificate (gcID,gcFrom,gcTo,gcEmail,gcOrigAmount,gcRemaining,gcDateCreated,"
			if dateused<>"" then sSQL=sSQL & "gcDateUsed,"
			sSQL=sSQL & "gcAuthorized,gcMessage) VALUES (" & _
				"'" & trim(ucase(replace(getpost("gcid"),"'",""))) & "'" & _
				",'" & escape_string(getpost("gcfrom")) & "'" & _
				",'" & escape_string(getpost("gcto")) & "'" & _
				",'" & trim(replace(getpost("gcemail"),"'","")) & "'" & _
				"," & trim(replace(getpost("gcorigamount"),"'","")) & _
				"," & trim(replace(getpost("gcremaining"),"'","")) & _
				"," & datecreated
			if dateused<>"" then sSQL=sSQL & "," & dateused
			sSQL=sSQL & "," & replace(getpost("gcauthorized"),"'","") & _
			",'" & escape_string(getpost("gcmessage")) & "')"
			ect_query(sSQL)
			dorefresh=TRUE
			
			if getpost("emailrecipient")="ON" then
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
				emailBody=replace(emailBody, "%toname%", getpost("gcto"))
				emailBody=replace(emailBody, "%fromname%", getpost("gcfrom"))
				emailBody=replace(emailBody, "%value%", FormatEuroCurrency(getpost("gcremaining")))
				emailBody=replaceemailtxt(emailBody, "%message%", getpost("gcmessage"), replaceone)
				emailBody=replace(emailBody, "%storeurl%", storeurl)
				emailBody=replace(emailBody, "%certificateid%", getpost("gcid"))
				emailBody=replace(emailBody, "<br />", emlNl)
				call dosendemaileo(getpost("gcemail"), emailAddr, "", replace(giftcertsubject, "%fromname%", getpost("gcfrom")), emailBody, emailObject, themailhost, theuser, thepass)
			end if
		end if
	elseif getpost("act")="domodify" then
		sSQL="UPDATE giftcertificate SET " & _
			"gcID='" & trim(ucase(replace(getpost("gcid"),"'",""))) & "'" & _
			",gcFrom='" & escape_string(getpost("gcfrom")) & "'" & _
			",gcTo='" & escape_string(getpost("gcto")) & "'" & _
			",gcEmail='" & trim(replace(getpost("gcemail"),"'","")) & "'" & _
			",gcOrigAmount=" & trim(replace(getpost("gcorigamount"),"'","")) & _
			",gcRemaining=" & trim(replace(getpost("gcremaining"),"'",""))
		session.lcid=saveLCID
		sSQL=sSQL & ",gcDateCreated=" & IIfVr(getpost("gcdatecreated")<>"", vsusdate(datevalue(replace(getpost("gcdatecreated"),"'",""))), vsusdate(Date()))
		if getpost("gcdateused")<>"" then sSQL=sSQL & ",gcDateUsed=" & vsusdate(datevalue(replace(getpost("gcdateused"),"'","")))
		session.lcid=1033
		sSQL=sSQL & ",gcAuthorized=" & replace(getpost("gcauthorized"),"'","") & _
			",gcMessage='" & escape_string(getpost("gcmessage")) & "'" & _
			" WHERE gcID='" & ucase(escape_string(getpost("id"))) & "'"
		ect_query(sSQL)
		dorefresh=TRUE
	elseif getpost("act")="purgeunconfirmed" then
		sSQL="DELETE FROM giftcertificate WHERE isconfirmed=0 AND mlConfirmDate<" & vsusdate(date()-mailinglistpurgedays)
		ect_query(sSQL)
		dorefresh=TRUE
	end if
end if
if dorefresh then
	print "<meta http-equiv=""refresh"" content=""1; url=admingiftcert.asp"
	print "?stext=" & urlencode(getpost("stext")) & "&stype=" & getpost("stype") & "&status=" & getpost("status") & "&pg=" & getpost("pg")
	print """>"
end if
if getget("id")<>"" OR (getpost("posted")="1" AND (getpost("act")="modify" OR getpost("act")="addnew" OR getpost("act")="clone")) then
%>
<script>
<!--
function getgcchar(){
	var gcchar='';
	while(gcchar=="" || gcchar=="O" || gcchar=="I" || gcchar=="Q"){
		gcchar=String.fromCharCode('A'.charCodeAt(0)+Math.round(Math.random()*25));
	}
	return(gcchar);
}
function randomgc(){
	var rannum=Math.floor((Math.random()*899999999)+100000000);
	rannum=getgcchar() + getgcchar() + rannum + getgcchar();
	document.getElementById("gcid").value=rannum;
}
function formvalidator(theForm){
if (theForm.gcid.value == ""){
alert("<%=jscheck(yyPlsEntr&" """&yyCerNum)%>\".");
theForm.gcid.focus();
return (false);
}
if (theForm.gcto.value == ""){
alert("<%=jscheck(yyPlsEntr&" """&yyTo)%>\".");
theForm.gcto.focus();
return (false);
}
if (theForm.gcfrom.value == ""){
alert("<%=jscheck(yyPlsEntr&" """&yyFrom)%>\".");
theForm.gcfrom.focus();
return (false);
}
if (theForm.gcemail.value == ""){
alert("<%=jscheck(yyPlsEntr&" """&yyEmail)%>\".");
theForm.gcemail.focus();
return (false);
}
if (theForm.gcorigamount.value == ""){
alert("<%=jscheck(yyPlsEntr&" """&yyOriAmt)%>\".");
theForm.gcorigamount.focus();
return (false);
}
if (theForm.gcremaining.value == ""){
alert("<%=jscheck(yyPlsEntr&" """&yyRemain)%>\".");
theForm.gcremaining.focus();
return (false);
}
return (true);
}
//-->
</script>
<script src="popcalendar.js"></script>
<script>try{languagetext('<%=adminlang%>');}catch(err){}</script>
		  <form name="mainform" method="post" action="admingiftcert.asp" onsubmit="return formvalidator(this)">
<%			call writehiddenvar("posted", "1")
			if getpost("act")="modify" OR getget("id")<>"" then
				call writehiddenvar("act", "domodify")
			else
				call writehiddenvar("act", "doaddnew")
			end if
			call writehiddenvar("stext", getpost("stext"))
			call writehiddenvar("status", getpost("status"))
			call writehiddenvar("stype", getpost("stype"))
			call writehiddenvar("pg", getpost("pg"))
			call writehiddenvar("id", request("id")) %>
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
			  <tr> 
                <td width="100%" colspan="2" align="center"><strong><%=IIfVr(getpost("act")="clone",yyClone&": ",IIfVs(getpost("act")="modify",yyModify&": ")) & yyGCMan & "<br />&nbsp;" %></strong></td>
			  </tr>
<%		if getpost("act")="modify" OR getpost("act")="clone" OR getget("id")<>"" then
			sSQL="SELECT gcID,gcTo,gcFrom,gcEmail,gcOrigAmount,gcRemaining,gcDateCreated,gcDateUsed,gcAuthorized,gcMessage,gcCartID FROM giftcertificate WHERE gcID='" & escape_string(request("id")) & "'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				gcid=rs("gcID")
				gcto=rs("gcTo")
				gcfrom=rs("gcFrom")
				gcemail=rs("gcEmail")
				gcorigamount=rs("gcOrigAmount")
				gcremaining=rs("gcRemaining")
				session.lcid=saveLCID
				gcdatecreated=cstr(rs("gcDateCreated"))
				gcdateused=cstr(rs("gcDateUsed")&"")
				session.lcid=1033
				gcauthorized=rs("gcAuthorized")
				gcmessage=rs("gcMessage")
				gccartid=rs("gcCartID")
			end if
			rs.close %>
<%		else
			gcid=""
			gcto=""
			gcfrom=""
			gcemail=""
			gcorigamount=""
			gcremaining=""
			session.lcid=saveLCID
			gcdatecreated=cstr(date())
			session.lcid=1033
			gcdateused=""
			gcauthorized=0
			gcmessage=""
			gccartid=0
		end if
		session.lcid=saveLCID
		themask=cStr(DateSerial(2003,12,11))
		themask=replace(themask,"2003","yyyy")
		themask=replace(themask,"12","mm")
		themask=replace(themask,"11","dd")
		session.lcid=1033 %>
			  <tr>
				<td align="right"><p><strong><%=yyCerNum%>:</strong></td>
				<td align="left"><%
		if getpost("act")="modify" then
			print "<input type=""hidden"" name=""gcid"" id=""gcid"" value="""&htmlspecials(gcid)&""" /><strong>" & htmlspecials(gcid) & "</strong>"
		else
			print "<input type=""text"" name=""gcid"" id=""gcid"" size=""22"" value="""&htmlspecials(gcid)&""" /> <input type=""button"" value=""Random"" onclick=""randomgc()"" /></td>"
		end if %>
			  </tr>
			  <tr>
				<td align="right"><p><strong><%=yyTo%>:</strong></td>
				<td align="left"><input type="text" name="gcto" size="34" value="<%=htmlspecials(gcto)%>" /></td>
			  </tr>
			  <tr>
				<td align="right"><p><strong><%=yyFrom%>:</strong></td>
				<td align="left"><input type="text" name="gcfrom" size="34" value="<%=htmlspecials(gcfrom)%>" /></td>
			  </tr>
			  <tr>
				<td align="right"><p><strong><%=yyEmail%>:</strong></td>
				<td align="left"><input type="text" name="gcemail" size="34" value="<%=htmlspecials(gcemail)%>" /></td>
			  </tr>
			  <tr>
				<td align="right"><p><strong><%=yyOriAmt%>:</strong></td>
				<td align="left"><input type="text" name="gcorigamount" size="10" value="<%=htmlspecials(gcorigamount)%>" <% if getpost("act")="addnew" then print "onchange=""document.getElementById('gcremaining').value=this.value"" " %>/></td>
			  </tr>
			  <tr>
				<td align="right"><p><strong><%=yyRemain%>:</strong></td>
				<td align="left"><input type="text" id="gcremaining" name="gcremaining" size="10" value="<%=htmlspecials(gcremaining)%>" /></td>
			  </tr>
			  <tr>
				<td align="right"><p><strong><%=yyDatPur%>:</strong></td>
				<td align="left"><div style="position:relative;display:inline"><input type="text" name="gcdatecreated" size="10" value="<%=gcdatecreated%>" style="vertical-align:middle" /> <input type="button" onclick="popUpCalendar(this, document.forms.mainform.gcdatecreated, '<%=themask%>', -200)" value="DP" /></div></td>
			  </tr>
			  <tr>
				<td align="right"><p><strong><%=yyDatUsd%>:</strong></td>
				<td align="left"><div style="position:relative;display:inline"><input type="text" name="gcdateused" size="10" value="<%=gcdateused%>" style="vertical-align:middle" /> <input type="button" onclick="popUpCalendar(this, document.forms.mainform.gcdateused, '<%=themask%>', -200)" value="DP" /></div></td>
			  </tr>
			  <tr>
				<td align="right"><p><strong><%=yyAuthd%>:</strong></td>
				<td align="left"><select name="gcauthorized" size="1">
						<option value="0"><%=yyNo%></option>
						<option value="1" <% if gcauthorized<>0 OR getpost("act")="addnew" then print "selected" %>><%=yyYes%></option></select>
				</td>
			  </tr>
<%		if getpost("act")="addnew" then %>
			  <tr>
				<td align="right"><p><strong>Email Recipient:</strong></td>
				<td align="left"><input type="checkbox" name="emailrecipient" value="ON" /></td>
			  </tr>
<%		end if %>
			  <tr>
				<td align="right"><p><strong><%=yyMessag%>:</strong></td>
				<td align="left"><textarea name="gcmessage" cols="60" rows="5" wrap="virtual"><%=gcmessage%></textarea></td>
			  </tr>
<%	if gccartid<>0 then
		sSQL="SELECT cartOrderID FROM cart WHERE cartID=" & gccartid
		rs.open sSQL,cnn,0,1
			if rs.EOF then gcorderid=0 else gcorderid=rs("cartOrderID")
		rs.close
%>
			  <tr>
				<td align="right"><p><strong><%=yyPurOrd%>:</strong></td>
				<td align="left"><% if gcorderid=0 then print yyUncOrd else print "("&gcorderid&") <a href=""adminorders.asp?id="&gcorderid&""">"&yyClkVw&".</a>"%></td>
			  </tr>
<%	end if
	if gcid<>"" then
		sSQL="SELECT gcaOrdID,gcaAmount FROM giftcertsapplied WHERE gcaGCID='"&gcid&"'"
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF %>
			  <tr>
				<td align="right"><p><strong><%=yyConOrd%>:</strong></td>
				<td align="left"><%=FormatEuroCurrency(rs("gcaAmount")) & " ("&rs("gcaOrdID")&") <input type=""button"" value="""&yyView&""" onclick=""document.location='adminorders.asp?id="&rs("gcaOrdID")&"'"" /> <input type=""button"" value="""&yyDelete&""" onclick=""document.location='admingiftcert.asp?act=deleteassoc&ord="&rs("gcaOrdID")&"&id="&gcid&"'"" /> <input type=""button"" value="""&yyDelRef&""" onclick=""document.location='admingiftcert.asp?act=deleteassoc&refund=true&ord="&rs("gcaOrdID")&"&id="&gcid&"'"" />"%></td>
			  </tr>
<%			rs.MoveNext
		loop
		rs.close
	end if
%>
			  <tr>
                <td width="100%" colspan="2" align="center"><br /><input type="submit" value="<%=yySubmit%>" />&nbsp;<input type="reset" value="<%=yyReset%>" /><br />&nbsp;</td>
			  </tr>
			  <tr>
                <td width="100%" colspan="2" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
                          &nbsp;</td>
			  </tr>
            </table>
		  </form>
<%
elseif getpost("posted")="1" AND getpost("act")<>"confirm" AND success then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <a href="admingiftcert.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br />&nbsp;<br />&nbsp;
                </td>
			  </tr>
			</table></td>
        </tr>
      </table>
<%
elseif getpost("posted")="1" AND getpost("act")<>"confirm" then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><span style="color:#FF0000;font-weight:bold"><%=yyOpFai%></span><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a><p>&nbsp;</p><p>&nbsp;</p></td>
			  </tr>
			</table></td>
        </tr>
      </table>
<%
else
	jscript=""
	modclone=request.cookies("modclone")
	pract=request.cookies("practgc") %>
<script>
<!--
function setCookie(c_name,value,expiredays){
	var exdate=new Date();
	exdate.setDate(exdate.getDate()+expiredays);
	document.cookie=c_name+ "=" +escape(value)+((expiredays==null) ? "" : ";expires="+exdate.toGMTString());
}
function mr(id) {
	document.mainform.id.value=id;
	document.mainform.act.value="modify";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function cr(id) {
	document.mainform.id.value=id;
	document.mainform.act.value="clone";
	document.mainform.posted.value="1";
	document.mainform.submit();
}

function crec(id) {
	document.mainform.id.value=id;
	document.mainform.act.value="confirm";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function newrec(id) {
	document.mainform.id.value=id;
	document.mainform.act.value="addnew";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function sendem(id) {
	document.mainform.act.value="sendem";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function dr(id){
if(confirm("<%=jscheck(yyConDel)%>\n")){
	document.mainform.id.value=id;
	document.mainform.act.value="delete";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
}
function startsearch(){
	document.mainform.action="admingiftcert.asp";
	document.mainform.act.value="search";
	document.mainform.listem.value="";
	document.mainform.posted.value="";
	document.mainform.submit();
}
function quickupdate(){
	if(document.mainform.pract.value=="del"){
		if(!confirm("<%=jscheck(yyConDel)%>\n"))
			return;
	}
	document.mainform.action="admingiftcert.asp";
	document.mainform.act.value="quickupdate";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function listem(thelet){
	document.mainform.action="admingiftcert.asp";
	document.mainform.act.value="search";
	document.mainform.listem.value=thelet;
	document.mainform.posted.value="";
	document.mainform.submit();
}
function removeuncon(){
if(confirm("<%=jscheck(yyConDel)%>\n")) {
	document.mainform.act.value="purgeunconfirmed";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
}
function changepract(obj){
	setCookie('practgc',obj[obj.selectedIndex].value,600);
	startsearch();
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
function changemodclone(modclone){
	setCookie('modclone',modclone[modclone.selectedIndex].value,600);
	startsearch();
}
// -->
</script>
<h2><%=yyAdmGif%></h2>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
		<form name="mainform" method="post" action="admingiftcert.asp">
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="xxxxx" />
			<input type="hidden" name="listem" value="<%=request("listem")%>" />
			<input type="hidden" name="id" value="xxxxx" />
			<input type="hidden" name="pg" value="<%=IIfVr(getpost("act")="search", "1", getget("pg"))%>" />
			<table class="cobtbl" width="100%" border="0" cellspacing="1" cellpadding="3">
			  <tr> 
				<td class="cobhl" width="25%" align="right"><%=yySrchFr%>:</td>
				<td class="cobll" width="25%"><input type="text" name="stext" size="20" value="<%=request("stext")%>" /></td>
				<td class="cobhl" align="right"><%=yyStatus%>:</td>
				<td class="cobll"><select name="status" size="1">
					<option value="any">All Certificates</option>
					<option value="" <% if request("status")="" then print "selected"%>><%=yyActGC%></option>
					<option value="spent" <% if request("status")="spent" then print "selected"%>><%=yyInaGC%></option>
					</select>
				</td>
			  </tr>
			  <tr>
				<td class="cobhl" align="right"><%
					if pract="del" OR pract="app" then %>
						<input type="button" value="<%=yyCheckA%>" onclick="checkboxes(true);" style="float:left" />
<%					end if %><%=yySrchTp%>:</td>
				<td class="cobll"><select name="stype" size="1">
					<option value=""><%=yySrchAl%></option>
					<option value="any" <% if request("stype")="any" then print "selected"%>><%=yySrchAn%></option>
					<option value="exact" <% if request("stype")="exact" then print "selected"%>><%=yySrchEx%></option>
					</select>
				</td>
				<td class="cobll" colspan="2" align="center">
						<input type="button" value="<%=yyListRe%>" onclick="startsearch();" />
						<input type="button" value="New Gift Certificate" onclick="newrec();" />
				</td>
			  </tr>
			</table>
<br />
            <table width="100%" class="stackable admin-table-a sta-white">
<%	hasheader=FALSE
	resultcounter=0
	if getpost("act")="search" OR getget("pg")<>"" OR getpost("act")="confirm" then
		Session.LCID=saveLCID
		sub displayprodrow(xrs)
			if cint(xrs("gcAuthorized"))<>0 then startstyle="" : endstyle="" else startstyle="<span style=""color:#FF0000"">" : endstyle="</span>"
			jscript=jscript&"pa["&resultcounter&"]=["
		%><tr id="tr<%=resultcounter%>"><td class="minicell"><%
				if pract="del" then
					print "<input type=""checkbox"" id=""chkbx"&resultcounter&""" name=""pra_"&xrs("gcID")&""" value=""del"" tabindex="""&(resultcounter+1)&"""/>"
				else
					print "&nbsp;"
				end if
			%></td><td><%=startstyle & htmlspecials(xrs("gcID")) & endstyle%></td>
			<td><%=startstyle & htmlspecials(xrs("gcTo")&"") & endstyle%></td>
			<td><%=startstyle & htmlspecials(xrs("gcFrom")&"") & endstyle%></td>
			<td><%=startstyle & FormatEuroCurrency(xrs("gcOrigAmount")) & endstyle%></td>
			<td><%=startstyle & FormatEuroCurrency(xrs("gcRemaining")) & endstyle%></td>
			<td><%=startstyle & htmlspecials(xrs("gcDateCreated")&"") & endstyle%></td><td>-</td>
		</tr>
<%		end sub
		sub displayheaderrow() %>
			<tr>
				<th class="minicell">
					<select name="pract" id="pract" size="1" onchange="changepract(this)">
					<option value="none">Quick Entry...</option>
					<option value="" disabled="disabled">------------------</option>
					<option value="del"<% if pract="del" then print " selected=""selected"""%>><%=yyDelete%></option>
					</select></th>
				<th class="maincell"><%=yyCerNum%></th>
				<th class="maincell"><%=yyTo%></th>
				<th class="maincell"><%=yyFrom%></th>
				<th class="maincell"><%=yyAmount%></th>
				<th class="maincell"><%=yyRemain%></th>
				<th class="maincell"><%=yyDate%></th>
				<th class="minicell"><%=yyModify%></th>
			</tr>
<%		end sub
		whereand=" WHERE"
		sSQL="SELECT gcID,gcTo,gcFrom,gcEmail,gcOrigAmount,gcRemaining,gcDateCreated,gcDateUsed,gcAuthorized FROM giftcertificate "
		if trim(request("stext"))<>"" then
			sText=escape_string(request("stext"))
			aText=Split(sText)
			aFields(0)="gcID"
			aFields(1)="gcTo"
			aFields(2)="gcFrom"
			aFields(3)="gcEmail"
			if request("stype")="exact" then
				sSQL=sSQL & whereand & " (gcID LIKE '%"&sText&"%' OR gcTo LIKE '%"&sText&"%' OR gcFrom LIKE '%"&sText&"%' OR gcEmail LIKE '%"&sText&"%') "
			else
				if request("stype")="any" then sJoin="OR " else sJoin="AND "
				sSQL=sSQL & whereand & "("
				whereand=" AND "
				for index=0 to 2
					sSQL=sSQL & "("
					for rowcounter=0 to UBOUND(aText)
						sSQL=sSQL & aFields(index) & " LIKE '%"&aText(rowcounter)&"%' "
						if rowcounter<UBOUND(aText) then sSQL=sSQL & sJoin
					next
					sSQL=sSQL & ") "
					if index < 2 then sSQL=sSQL & "OR "
				next
				sSQL=sSQL & ") "
			end if
			whereand=" AND"
		end if
		if trim(request("status"))="" then
			sSQL=sSQL & whereand & " (gcRemaining>0 AND gcAuthorized<>0)"
			whereand=" AND"
		elseif trim(request("status"))="spent" then
			sSQL=sSQL & whereand & " (gcRemaining<=0 OR gcAuthorized=0)"
			whereand=" AND"
		end if
		sSQL=sSQL & " ORDER BY gcDateCreated"
		if admingiftcertsperpage="" then admingiftcertsperpage=100
		rs.CursorLocation=3 ' adUseClient
		rs.CacheSize=admingiftcertsperpage
		rs.open sSQL, cnn
		if rs.eof or rs.bof then
			success=false
			iNumOfPages=0
		else
			success=true
			rs.MoveFirst
			rs.PageSize=admingiftcertsperpage
			CurPage=1
			if is_numeric(getget("pg")) then CurPage=int(getget("pg"))
			iNumOfPages=Int((rs.RecordCount + (admingiftcertsperpage-1)) / admingiftcertsperpage)
			rs.AbsolutePage=CurPage
		end if
		if NOT rs.EOF then
			pblink="<a href=""admingiftcert.asp?status="&request("status")&"&stext="&urlencode(request("stext"))&"&stype="&request("stype")&"&pg="
			if iNumOfPages > 1 then print "<tr><td colspan=""7"" align=""center"">" & writepagebar(CurPage,iNumOfPages,yyPrev,yyNext,pblink,FALSE) & "</td></tr>"
			displayheaderrow()
			addcomma=""
			do while NOT rs.EOF AND resultcounter<rs.PageSize
				hasheader=TRUE
				displayprodrow(rs)
				jscript=jscript&"'"&rs("gcID")&"'];"&vbCrLf
				resultcounter=resultcounter + 1
				rs.MoveNext
			loop
			if iNumOfPages > 1 then print "<tr><td colspan=""7"" align=""center"">" & writepagebar(CurPage,iNumOfPages,yyPrev,yyNext,pblink,FALSE) & "</td></tr>"
		else
			print "<tr><td width=""100%"" colspan=""7"" align=""center""><br />"&yyItNone&"<br />&nbsp;</td></tr>"
		end if
		rs.close
	else
		numitems=0
		sSQL="SELECT COUNT(*) as totcount FROM giftcertificate WHERE gcRemaining>0"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			numitems=rs("totcount")
		end if
		rs.close
		print "<tr><td colspan=""7""><div class=""itemsdefine"">You have " & numitems & " " & yyActGC & " defined.</div></td></tr>"		  
	end if %>
			  <tr>
<%	if hasheader then %>
				<td align="center" style="white-space:nowrap"><% if resultcounter>0 AND pract<>"" AND pract<>"none" then print "<input type=""hidden"" name=""resultcounter"" id=""resultcounter"" value="""&resultcounter&""" /><input type=""button"" value="""&yyUpdate&""" onclick=""quickupdate()"" /> <input type=""reset"" value="""&yyReset&""" />" else print "&nbsp;"%></td>
<%	end if %>
                <td width="100%" colspan="7" align="center"><br /><a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
			  </tr>
            </table></td>
		  </form>
<script>
/* <![CDATA[ */
var pa=[];
<%=jscript%>
for(var pidind in pa){
	var ttr=document.getElementById('tr'+pidind);
	ttr.cells[7].style.textAlign='center';
	ttr.cells[7].style.whiteSpace='nowrap';
	ttr.cells[7].innerHTML='<input type="button" value="M" style="width:30px;margin-right:4px" onclick="mr(\''+pa[pidind][0]+'\')" title="<%=jsescape(htmlspecials(yyModify))%>" />' +
		'<input type="button" value="C" style="width:30px;margin-right:4px" onclick="cr(\''+pa[pidind][0]+'\')" title="<%=jsescape(htmlspecials(yyClone))%>" />' +
		'<input type="button" value="X" style="width:30px" onclick="dr(\''+pa[pidind][0]+'\')" title="<%=jsescape(htmlspecials(yyDelete))%>" />';
}
/* ]]> */
</script>
        </tr>
      </table>
<%
end if
cnn.Close
set rs=nothing
set cnn=nothing
%>
