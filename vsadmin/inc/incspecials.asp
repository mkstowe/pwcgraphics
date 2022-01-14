<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if sDSN="" then response.write "Database connection not set" : response.end
if prodid="" then
	if trim(explicitid)<>"" then prodid=trim(explicitid) else prodid=trim(request("prod"))
end if
if prodid<>giftcertificateid AND prodid<>donationid then prodid=giftcertificateid
WSP=""
OWSP=""
TWSP="pPrice"
iNumOfPages=0
if dateadjust="" then dateadjust=0
if NOT isincluded then
	Set rs=Server.CreateObject("ADODB.RecordSet")
	Set rs2=Server.CreateObject("ADODB.RecordSet")
	Set cnn=Server.CreateObject("ADODB.Connection")
	cnn.open sDSN
end if
alreadygotadmin=getadminsettings()
get_wholesaleprice_sql()
if SESSION("clientLoginLevel")<>"" then minloglevel=SESSION("clientLoginLevel") else minloglevel=0
validitem=TRUE
if getpost("posted")="1" then
	if prodid=giftcertificateid AND recaptchaenabled(64) then validitem=checkrecaptcha(xxAmtNov)
	if validitem then
		validitem=is_numeric(getpost("amount"))
		if validitem then amountdbl=cdbl(getpost("amount"))
		if validitem then validitem=amountdbl>0
	end if
	if validitem then
		session.LCID=1033
		prodname=IIfVr(prodid=giftcertificateid, xxGifCtc, xxDonat)
		sSQL="SELECT "&getlangid("pName",1)&" FROM products WHERE pID='" & escape_string(prodid) & "'"
		rs2.Open sSQL,cnn,0,1
		if NOT rs2.EOF then prodname=rs2(getlangid("pName",1))
		rs2.Close
		rs2.Open "cart",cnn,1,3,&H0002
		rs2.AddNew
		rs2.Fields("cartSessionID")		= getsessionid()
		if SESSION("clientID")<>"" then rs2.Fields("cartClientID")=SESSION("clientID") else rs2.Fields("cartClientID")=0
		rs2.Fields("cartProdID")		= prodid
		rs2.Fields("cartQuantity")		= 1
		rs2.Fields("cartCompleted")		= 0
		rs2.Fields("cartProdName")		= prodname
		rs2.Fields("cartProdPrice")		= amountdbl
		rs2.Fields("cartOrderID")		= 0
		rs2.Fields("cartDateAdded")		= DateAdd("h",dateadjust,Now())
		rs2.Update
		if mysqlserver=true then
			rs2.Close
			rs2.Open "SELECT LAST_INSERT_ID() AS lstIns",cnn,0,1
			cartid=rs2("lstIns")
		else
			cartid=rs2.Fields("cartID")
		end if
		rs2.Close
		session.LCID=saveLCID
		if prodid=giftcertificateid then
			' Create GC id
			randomize
			gotunique=FALSE
			do while NOT gotunique
				sequence=getgcchar() & getgcchar() & (Int(100000000 * Rnd) + 100000000) & getgcchar()
				sSQL="SELECT gcID FROM giftcertificate WHERE gcID='" & sequence & "'"
				rs2.Open sSQL,cnn,0,1
				if rs2.EOF then gotunique=TRUE
				rs2.Close
			loop
			sSQL="INSERT INTO giftcertificate (gcID,gcTo,gcFrom,gcEmail,gcOrigAmount,gcRemaining,gcDateCreated,gcCartID,gcAuthorized,gcMessage) VALUES ("
			sSQL=sSQL & "'" & sequence & "',"
			sSQL=sSQL & "'" & escape_string(getpost("toname")) & "',"
			sSQL=sSQL & "'" & escape_string(getpost("fromname")) & "',"
			sSQL=sSQL & "'" & escape_string(getpost("toemail")) & "',"
			sSQL=sSQL & "0,0,"
			sSQL=sSQL & vsusdate(DateAdd("h",dateadjust,Now())) & ","
			sSQL=sSQL & cartid & ",0,"
			sSQL=sSQL & "'" & escape_string(replace(replace(getpost("gcmessage"),vbCrLf,"<br />"),vbLf,"<br />")) & "')"
			ect_query(sSQL)
		else
			if getpost("fromname")<>"" then
				sSQL="INSERT INTO cartoptions (coCartID,coOptID,coOptGroup,coCartOption,coPriceDiff,coWeightDiff) VALUES ("&cartid&",0,'"&escape_string(xxFrom) & "','"&escape_string(left(getpost("fromname"),255))&"',0,0)"
				ect_query(sSQL)
			end if
			if getpost("gcmessage")<>"" then
				sSQL="INSERT INTO cartoptions (coCartID,coOptID,coOptGroup,coCartOption,coPriceDiff,coWeightDiff) VALUES ("&cartid&",0,'"&escape_string(xxMessag) & "','"&escape_string(left(getpost("gcmessage"),255))&"',0,0)"
				ect_query(sSQL)
			end if
		end if
		response.redirect "cart" & extension
	end if
end if
if getpost("posted")<>"1" OR NOT validitem then
	if prodid=giftcertificateid then ' {
		if giftcertificateminimum="" then giftcertificateminimum=5
%>
<script>
/* <![CDATA[ */
function checkastring(thestr,validchars){
  for (i=0; i < thestr.length; i++){
    ch=thestr.charAt(i);
    for (j=0;  j < validchars.length;  j++)
      if (ch == validchars.charAt(j))
        break;
    if (j == validchars.length)
	  return(false);
  }
  return(true);
}
function formvalECTspecials(frm){
if(frm.amount.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&xxAmount)%>\".");
	frm.amount.focus();
	return(false);
}
if (!checkastring(frm.amount.value,"0123456789<%=replace(cstr(formatnumber(3.33)),"3","")%>")){
	alert("<%=jscheck(xxOnlyDec&" """&xxAmount)%>\".");
	frm.amount.focus();
	return(false);
}
if(frm.amount.value<<%=giftcertificateminimum%>){
	alert("<%=jscheck(xxGCMini & " " & FormatEuroCurrency(giftcertificateminimum))%>.");
	frm.amount.focus();
	return(false);
}
if(frm.toname.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&xxTo)%>\".");
	frm.toname.focus();
	return(false);
}
if(frm.fromname.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&xxFrom)%>\".");
	frm.fromname.focus();
	return(false);
}

if(frm.toemail.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&xxReEmai)%>\".");
	frm.toemail.focus();
	return(false);
}
var regex=/[^@]+@[^@]+\.[a-z]{2,}$/i;
if(!regex.test(frm.toemail.value)){
	alert("<%=jscheck(xxValEm)%>");
	frm.toemail.focus();
	return(false);
}
if(frm.toemail2.value!=frm.toemail.value){
	alert("<%=jscheck(xxEmCNMa)%>.");
	frm.toemail2.focus();
	return(false);
}
<%	if recaptchaenabled(64) then print "if(!giftcertcaptchaok){ alert(""" & jscheck(xxRecapt) & """);return(false); }" %>
return (true);
}
/* ]]> */
</script>
	<form method="post" onsubmit="return formvalECTspecials(this)">
	<input type="hidden" name="posted" value="1" />
	<input type="hidden" name="prod" value="<%=giftcertificateid%>" />
      <div class="ectdiv ectgiftcerts">
        <div class="ectdivhead"><%=xxGCPurc%></div>
<%		if getpost("posted")="1" then %>
        <div class="ectdiv2column ectwarning"><%=xxAmtNov%></div>
<%		end if %>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><label for="amount"><%=xxAmount%></label></div>
			<div class="ectdivright"><input type="text" name="amount" id="amount" size="4" maxlength="10" value="<%=htmlspecials(getpost("amount"))%>" /></div>
        </div>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><label for="toname"><%=xxTo%></label></div>
			<div class="ectdivright"><input type="text" name="toname" id="toname" size="25" maxlength="50" value="<%=htmlspecials(getpost("toname"))%>" /></div>
        </div>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><label for="fromname"><%=xxFrom%></label></div>
			<div class="ectdivright"><input type="text" name="fromname" id="fromname" size="25" maxlength="50" value="<%=htmlspecials(getpost("fromname"))%>" /></div>
        </div>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><label for="toemail"><%=xxReEmai%></label></div>
			<div class="ectdivright"><input type="text" name="toemail" id="toemail" size="25" maxlength="50" value="<%=htmlspecials(getpost("toemail"))%>" /></div>
        </div>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><label for="toemail2"><%=xxCReEma%></label></div>
			<div class="ectdivright"><input type="text" name="toemail2" id="toemail2" size="25" maxlength="50" value="<%=htmlspecials(getpost("toemail2"))%>" /></div>
        </div>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><label for="gcmessage"><%=xxMessag%></label></div>
			<div class="ectdivright"><textarea name="gcmessage" id="gcmessage" cols="35" rows="4"><%=htmlspecials(getpost("gcmessage"))%></textarea></div>
        </div>
<%		if recaptchaenabled(64) then %>
		<div class="ectdivcontainer">
			<div class="ectdivleft">&nbsp;</div>
			<% call displayrecaptchajs("giftcertcaptcha",TRUE,FALSE) %>
			<div id="giftcertcaptcha" class="g-recaptcha ectdivright"></div>
		</div>
<%		end if %>
		<div class="ectdivcontainer">
			<div class="ectdivleft">&nbsp;</div>
			<div class="ectdivright"><%=imageorsubmit(imggcsubmit,xxSubmt,"gcsubmit")%></div>
		</div>
      </div>
	</form>
<%
	else ' }{ donation
%>
<script>
/* <![CDATA[ */
function checkastring(thestr,validchars){
  for (i=0; i < thestr.length; i++){
    ch=thestr.charAt(i);
    for (j=0;  j < validchars.length;  j++)
      if (ch == validchars.charAt(j))
        break;
    if (j == validchars.length)
	  return(false);
  }
  return(true);
}
function formvalECTspecials(frm){
if(frm.amount.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&xxAmount)%>\".");
	frm.amount.focus();
	return(false);
}
if (!checkastring(frm.amount.value,"0123456789<%=replace(cstr(formatnumber(3.33)),"3","")%>")){
	alert("<%=jscheck(xxOnlyDec&" """&xxAmount)%>\".");
	frm.amount.focus();
	return(false);
}
if(frm.gcmessage.value.length>255){
	alert("<%=jscheck(xxPrd255)%>");
	frm.gcmessage.focus();
	return(false);
}
return (true);
}
/* ]]> */
</script>
<%		if NOT isincluded then %>
	<form method="post" onsubmit="return formvalECTspecials(this)">
<%		end if %>
	<input type="hidden" name="posted" value="1" />
	<input type="hidden" name="prod" value="<%=donationid%>" />
      <div class="ectdiv ectdonations">
		<div class="ectdivhead"><%=xxMakDon%></div>
<%		if getpost("posted")="1" then %>
        <div class="ectdiv2column ectwarning"><%=xxAmtNov%></div>
<%		end if %>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><%=redasterix%><label for="amount"><%=xxAmount%></label></div>
			<div class="ectdivright"><input type="text" name="amount" id="amount" size="6" maxlength="10" value="<%=htmlspecials(getpost("amount"))%>" /></div>
        </div>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><label for="fromname"><%=xxFrom%></label></div>
			<div class="ectdivright"><input type="text" name="fromname" id="fromname" size="25" maxlength="50" value="<%=htmlspecials(getpost("fromname"))%>" /></div>
        </div>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><label for="gcmessage"><%=xxMessag%></label></div>
			<div class="ectdivright"><textarea name="gcmessage" id="gcmessage" cols="35" rows="4"><%=htmlspecials(getpost("gcmessage"))%></textarea></div>
        </div>
		<div class="ectdivcontainer">
			<div class="ectdivleft">&nbsp;</div>
			<div class="ectdivright"><%=imageorsubmit(imgdonationsubmit,xxSubmt,"donationsubmit")%></div>
		</div>
      </div>
<%		if NOT isincluded then %>
	</form>
<%		end if
	end if
end if
if NOT isincluded then
	cnn.Close
	set rs=nothing
	set rs2=nothing
	set cnn=nothing
end if
%>