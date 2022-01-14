<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if sDSN="" then response.write "Database connection not set" : response.end
cartisincluded=TRUE
%>
<!--#include file="inccart.asp"-->
<%
if request.totalbytes > 10000 then response.end
success=true
ordGrandTotal=0 : ordTotal=0 : ordStateTax=0 : ordHSTTax=0 : ordCountryTax=0 : ordShipping=0 : ordHandling=0 : ordDiscount=0
affilID="" : ordCity="" : ordState="" : ordCountry="" : ordDiscountText="" : ordEmail=""
nonhomecountries=FALSE
digidownloads=false
allcountries=""
warncheckspamfolder=false
if NOT enableclientlogin then
	success=false
	errmsg="Client login not enabled"
end if
Set rs =Server.CreateObject("ADODB.RecordSet")
Set rs2=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
pagename="clientlogin"&extension
if instrrev(pagename,"/")>0 then pagename=right(pagename,len(pagename)-instrrev(pagename,"/"))
if forceloginonhttps then thisaction=storeurlssl & pagename else thisaction=""
alreadygotadmin=getadminsettings()
if displaysoftlogindone="" then displaysoftlogindone=""
if enableclientlogin then call displaysoftlogin()
%>
<script>
/* <![CDATA[ */
function vieworder(theid){
	document.forms.mainform.action.value="vieworder";
	document.forms.mainform.theid.value=theid;
	document.forms.mainform.submit();
}
function editaddress(theid){
	document.forms.mainform.action.value="editaddress";
	document.forms.mainform.theid.value=theid;
	document.forms.mainform.submit();
}
function newaddress(){
	document.forms.mainform.action.value="newaddress";
	document.forms.mainform.submit();
}
function editaccount(){
	document.forms.mainform.action.value="editaccount";
	document.forms.mainform.submit();
}
function deleteaddress(theid){
	if(confirm("<%=jscheck(xxDelAdd)%>")){
		document.forms.mainform.action.value="deleteaddress";
		document.forms.mainform.theid.value=theid;
		document.forms.mainform.submit();
	}
}
function createlist(){
	if(document.forms.mainform.listname.value.indexOf('<')!=-1){
		alert("<%=jscheck("Illegal Character ""<""")%>.");
		document.forms.mainform.listname.focus();
		return(false);
	}else if(document.forms.mainform.listname.value==''){
		alert("<%=jscheck(xxPlsEntr&" """&xxLisNam)%>\".");
		document.forms.mainform.listname.focus();
		return(false);
	}else{
		document.forms.mainform.action.value="createlist";
		document.forms.mainform.submit();
	}
}
function deletelist(theid){
	if(confirm("<%=jscheck(xxDelLis)%>")){
		document.forms.mainform.action.value="deletelist";
		document.forms.mainform.theid.value=theid;
		document.forms.mainform.submit();
	}
}
/* ]]> */
</script>
<%	if getpost("doresetpw")="1" then ' {
		sSQL="SELECT clID FROM customerlogin WHERE clEmail='"&escape_string(getpost("rst"))&"' AND clPw='"&escape_string(getpost("rsk"))&"'"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then clid=rs("clID") else clid=""
		if getpost("newpw")="" then clid=""
		rs.close
		if clid<>"" then ect_query("UPDATE customerlogin SET clPw='"&escape_string(dohashpw(getpost("newpw")))&"' WHERE clID=" & clid)
%>	  <div class="ectdiv ectclientlogin">
		<div class="ectdivhead"><%=xxCusAcc%></div>
		  <div class="ectdivcontainer">
			<div class="ectdivleft"><%=xxForPas%></div>
			<div class="ectdivright"><%=IIfVr(clid="",xxEmNtFn,xxPasRsS) %></div>
		  </div>
		  <div class="ectdiv2column"><%
		if clid<>"" then
			print imageorbutton(imglogin,xxLogin,"login","return displayloginaccount()",TRUE)
		else
			print imageorbutton(imggoback,xxGoBack,"goback","history.go(-1)",TRUE)
		end if
		%></div>
	  </div>
<%	elseif getget("rst")<>"" AND getget("rsk")<>"" then ' }{
		sSQL="SELECT clID FROM customerlogin WHERE clEmail='"&escape_string(getget("rst"))&"' AND clPw='"&escape_string(getget("rsk"))&"'"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then success=TRUE else success=FALSE
		rs.close
		if NOT success then %>
	  <div class="ectdiv ectclientlogin">
		<div class="ectdivhead"><%=xxCusAcc%></div>
		  <div class="ectdivcontainer">
			<div class="ectdivleft"><%=xxForPas%></div>
			<div class="ectdivright"><%=xxSorRes %></div>
		  </div>
		  <div class="ectdiv2column"><% print imageorbutton(imgcancel,xxCancel,"cancel",storeurl,FALSE) %></div>
	  </div>
<%		else %>
<script>
/* <![CDATA[ */
function checknewpw(frm){
if(frm.newpw.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&xxNewPwd)%>\".");
	frm.newpw.focus();
	return(false);
}
var newpw=frm.newpw.value;
var newpw2=frm.newpw2.value;
if(newpw!=newpw2){
	alert("<%=jscheck(xxPwdMat)%>");
	frm.newpw.focus();
	return(false);
}
return true;
}
/* ]]> */
</script>
	<form method="post" name="mainform" action="<%=thisaction%>" onsubmit="return checknewpw(this)">
	<input type="hidden" name="doresetpw" value="1" />
	<input type="hidden" name="rst" value="<%=replace(getget("rst"),"""","")%>" />
	<input type="hidden" name="rsk" value="<%=replace(getget("rsk"),"""","")%>" />
	  <div class="ectdiv ectclientlogin">
		<div class="ectdivhead"><%=xxCusAcc & " " & xxForPas%></div>
		  <div class="ectdivcontainer">
			<div class="ectdivleft"><%=xxNewPwd%></div>
			<div class="ectdivright"><input type="password" size="20" name="newpw" value="" autocomplete="off" /></div>
		  </div>
		  <div class="ectdivcontainer">
			<div class="ectdivleft"><%=xxRptPwd%></div>
			<div class="ectdivright"><input type="password" size="20" name="newpw2" value="" autocomplete="off" /></div>
		  </div>
		  <div class="ectdiv2column"><%=imageorsubmit(imgsubmit,xxSubmt,"submit")&" "&imageorbutton(imgcancel,xxCancel,"cancel",storeurl,FALSE)%></div>
	  </div>
	</form>
<%		end if
	elseif getget("action")="logout" then ' }{
		SESSION("clientID")=empty
		SESSION("clientUser")=empty
		SESSION("clientActions")=empty
		SESSION("clientLoginLevel")=empty
		SESSION("clientPercentDiscount")=empty
		call setacookie("WRITECLL","", -7)
		call setacookie("WRITECLP","", -7)
		if storeurlssl<>storeurl then print "<script src=""" & storeurlssl & "vsadmin/savecookie.asp?DELCLL=Y""></script>"
		if clientlogoutref<>"" then refURL=clientlogoutref else refURL=storehomeurl
		print "<script>setTimeout(function(){document.location='" & jsescape(refURL) & "'},3000)</script>"
%>
		<div class="ectdiv ectclientlogin">
		  <div class="ectmessagescreen">
			<div><%=xxLOSuc%></div>
			<div><%=xxAutFo%></div>
			<div><%=xxForAut%> <a class="ectlink" href="<%=refURL%>"><%=xxClkHere%></a>.</div>
		  </div>
		</div>
<%	elseif getpost("action")="dolostpassword" then ' }{
		theemail=cleanupemail(getpost("email"))
		dofloodcontrol=FALSE
		sSQL="SELECT clPW FROM customerlogin WHERE clEmail<>'' AND clEmail='" & replace(theemail,"'","") & "'"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			cnn.execute("DELETE FROM ajaxfloodcontrol WHERE afcAction=6 AND afcDate<" & vsusdatetime(dateadd("s",-300,now())))
			sSQL="SELECT afcID FROM ajaxfloodcontrol WHERE afcAction=6 AND (afcIP='" & escape_string(REMOTE_ADDR) & "' OR afcSession='" & escape_string(session.sessionid) & "')"
			rs2.open sSQL,cnn,0,1
			dofloodcontrol=NOT rs2.EOF
			rs2.close

			if NOT dofloodcontrol then
				if htmlemails=TRUE then emlNl="<br />" else emlNl=vbCrLf
				tlink=storeurl & pagename & "?rst=" & theemail & "&rsk=" & rs("clPW")
				if htmlemails=TRUE then tlink="<a href=""" & tlink & """>" & tlink & "</a>"

				call DoSendEmailEO(replace(theemail,"'",""),emailAddr,"",xxForPas,xxLosPw1 & emlNl & storeurl & emlNl & emlNl & xxResPas & emlNl & tlink & emlNl & emlNl & xxLosPw3 & emlNl,emailObject,themailhost,theuser,thepass)
				
				sSQL="INSERT INTO ajaxfloodcontrol (afcAction,afcIP,afcSession,afcDate) VALUES (6,'" & escape_string(REMOTE_ADDR) & "','" & escape_string(session.sessionid) & "'," & vsusdatetime(now()) & ")"
				ect_query(sSQL)
			
				success=TRUE
			end if
		else
			success=FALSE
		end if %>
	  <form method="post" name="mainform" action="<%=thisaction%>">
	  <div class="ectdiv ectclientlogin">
		<div class="ectdivhead"><%=xxCusAcc%></div>
		  <div class="ectdivcontainer">
			<div class="ectdivleft"><%=xxForPas%></div>
			<div class="ectdivright"><% if success then print xxSenPw else print xxSorPw %></div>
		  </div>
		  <div class="ectdivcontainer">
			<div class="ectdivleft"></div>
			<div class="ectdivright"><%
			if dofloodcontrol then
				print xxFldCnt
			else
				if success then print imageorbutton(imglogin,xxLogin,"login","return displayloginaccount()",TRUE) else print imageorbutton(imggoback,xxGoBack,"goback","history.go(-1)",TRUE)
			end if
			%></div>
		  </div>
	  </div>
	  </form>
<%	elseif getget("mode")="lostpassword" then ' }{
%>	  <form method="post" name="mainform" action="<%=thisaction%>">
	  <input type="hidden" name="action" value="dolostpassword" />
	  <div class="ectdiv ectclientlogin ectlostpassword">
		<div class="ectdivhead"><%=xxCusAcc%></div>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><%=xxForPas%></div>
			<div class="ectdivright"><%=xxEntEm%></div>
		</div>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><%=xxEmail%></div>
			<div class="ectdivright"><input type="text" name="email" class="ecttextinput" size="31" placeholder="<%=xxEmail%>" /></div>
		</div>
		<div class="ectdivcontainer">
			<div class="ectdivleft"></div>
			<div class="ectdivright"><%=imageorsubmit(imgsubmit,xxSubmt,"submit")%></div>
		</div>
	  </div>
	  </form>
<%	elseif SESSION("clientID")="" then ' }{
%>	  <div class="ectdiv ectclientlogin">
		<div class="ectdivhead"><%=xxCusAcc%></div>
		<div>
			<div class="clientloginmessage"><%=xxMusLog%></div>
<%			print "<div class=""cartloginbuttons"">"
				if enableclientlogin then
					print "<div class=""cartloginlogin"">" & imageorbutton(imgloginaccount,xxLogAcc,"logintoaccount","displayloginaccount()",TRUE)&"</div>"
					if allowclientregistration then print "<div class=""cartloginnewacct"">" & imageorbutton(imgcreateaccount,xxCreAcc,"createaccount","displaynewaccount()",TRUE)&"</div>"
				else
					print "Customer Login Disabled"
				end if
			print "</div>" %>
		</div>
	  </div>
<script>
displayloginaccount();
</script>
<%	else ' }{ is logged in
		if getpost("action")="vieworder" then ' {
%>	  <div class="ectdiv ectclientlogin">
		  <div class="clientloginvieworder"><%
			ordID=replace(getpost("theid"),"'","")
			if is_numeric(ordID) then success=TRUE else success=FALSE
			if success then
				sSQL="SELECT ordID FROM orders WHERE ordID=" & ordID & " AND ordClientID=" & SESSION("clientID")
				rs.open sSQL,cnn,0,1
				if rs.EOF then success=FALSE
				rs.close
			end if
			if success then
				xxThkYou=imageorbutton(imgbackacct,xxBack,"backacct","history.go(-1)",TRUE)
				xxRecEml=""
				thankspagecontinue="javascript:history.go(-1)"
				xxCntShp=xxBack
				imgcontinueshopping=imgbackacct
				Call do_order_success(ordID,emailAddr,FALSE,TRUE,FALSE,FALSE,FALSE)
			else
				errtext="Sorry, could not find a matching order."
				Call order_failed
			end if %>
		  </div>
	  </div>
<%		elseif getpost("action")="doeditaccount" then ' }{
			oldpw=dohashpw(getpost("oldpw"))
			newpw=getpost("newpw")
			newpw2=getpost("newpw2")
			clientuser=getpost("name")
			clientemail=getpost("email")
			allowemail=getpost("allowemail")
			sSQL="SELECT clPW,clEmail FROM customerlogin WHERE clID=" & SESSION("clientID")
			rs.open sSQL,cnn,0,1
			oldpassword=rs("clPW")
			oldemail=rs("clEmail")
			rs.close
			success=TRUE
			checkhash=sha256(oldemail & "asp know the hash" & adminSecret)
			if checkhash<>getpost("emhash") then
				success=FALSE
				errmsg="Hash check error"
			else
				if newpw<>"" OR newpw2<>"" then
					if oldpw<>oldpassword then
						success=FALSE
						errmsg=xxExNoMa
					end if
				end if
				if oldemail<>clientemail then
					sSQL="SELECT clID FROM customerlogin WHERE clEmail='"&escape_string(clientemail)&"'"
					rs.open sSQL,cnn,0,1
						if NOT rs.EOF then
							success=FALSE
							errmsg=xxEmExi
						end if
					rs.close
				end if
			end if
			if success then
				sSQL="UPDATE customerlogin SET "
				sSQL=sSQL & "clUserName='" & escape_string(clientuser) & "',"
				sSQL=sSQL & "clEmail='" & escape_string(clientemail) & "'"
				if trim(extraclientfield1)<>"" then sSQL=sSQL & ",clientCustom1='" & escape_string(getpost("clientCustom1")) & "'"
				if trim(extraclientfield2)<>"" then sSQL=sSQL & ",clientCustom2='" & escape_string(getpost("clientCustom2")) & "'"
				if newpw<>"" then sSQL=sSQL & ",clPW='" & replace(dohashpw(newpw),"'","") & "'"
				sSQL=sSQL & " WHERE clID=" & SESSION("clientID")
				ect_query(sSQL)
				session("clientPW")=dohashpw(newpw)
				if allowemail="ON" then
					call addtomailinglist(clientemail,clientuser)
					if oldemail<>clientemail then ect_query("DELETE FROM mailinglist WHERE email='" & escape_string(oldemail) & "'")
				else
					ect_query("DELETE FROM mailinglist WHERE email='" & escape_string(clientemail) & "'")
					ect_query("DELETE FROM mailinglist WHERE email='" & escape_string(oldemail) & "'")
				end if
				SESSION("clientUser")=clientuser
				print "<script>var url=[location.protocol,'//',location.host,location.pathname].join('');setTimeout(function(){document.location=url},2000)</script>"
			end if
%>	<form method="post" name="mainform" action="<%=thisaction%>">
	  <div class="ectdiv ectclientlogin">
		<div class="ectdivhead"><%=xxCusAcc%></div>
		<div style="padding:50px 0;text-align:center" class="ectdiv2column<%=IIfVs(NOT success," ectwarning")%>"><% if success then print xxUpdSuc else print errmsg %></div>
		<div style="text-align:center" class="ectdiv2column"><%
		if success then
			print imageorsubmit(imgcustomeracct,xxCusAcc,"customeracct")
		else
			print imageorbutton(imggoback,xxGoBack,"goback","history.go(-1)",TRUE)
		end if %></div>
	  </div>
	</form>
<%		elseif getpost("action")="editaccount" then ' }{
			if forceloginonhttps AND request.servervariables("HTTPS")="off" AND instr(storeurlssl,"https:")>0 then response.redirect storeurlssl & pagename & IIfVr(request.servervariables("QUERY_STRING")<>"","?"&request.servervariables("QUERY_STRING"),"") : response.end
%>
<script>
/* <![CDATA[ */
var checkedfullname=false;
function checknewaccount(){
frm=document.forms.mainform;
if(frm.name.value==""||frm.name.value=="<%=xxFirNam%>"){
	alert("<%=jscheck(xxPlsEntr&" """&IIfVr(usefirstlastname, xxFirNam, xxName))%>\".");
	frm.name.focus();
	return(false);
}
gotspace=false;
var checkStr=frm.name.value;
for (i=0; i < checkStr.length; i++){
	if(checkStr.charAt(i)==" ")
		gotspace=true;
}
if(!checkedfullname && !gotspace){
	alert("<%=jscheck(xxFulNam&" """&xxName)%>\".");
	frm.name.focus();
	checkedfullname=true;
	return(false);
}
if(frm.email.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&xxEmail)%>\".");
	frm.email.focus();
	return(false);
}
<%	if extraclientfield1required then %>
if(frm.clientCustom1.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&extraclientfield1)%>\".");
	frm.clientCustom1.focus();
	return(false);
}
<%	end if
	if extraclientfield2required then %>
if(frm.clientCustom2.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&extraclientfield2)%>\".");
	frm.clientCustom2.focus();
	return(false);
}
<%	end if %>
var regex=/[^@]+@[^@]+\.[a-z]{2,}$/i;
if(!regex.test(frm.email.value)){
	alert("<%=jscheck(xxValEm)%>");
	frm.email.focus();
	return(false);
}
var newpw=frm.newpw.value;
var newpw2=frm.newpw2.value;
if(newpw!='' && newpw!=newpw2){
	alert("<%=jscheck(xxPwdMat)%>");
	frm.newpw.focus();
	return(false);
}
return true;
}
/* ]]> */
</script>
		<form method="post" name="mainform" action="<%=thisaction%>" onsubmit="return checknewaccount()">
		<input type="hidden" name="action" value="doeditaccount" />
		<div class="ectdiv ectclientlogin">
		  <div class="ectdivhead"><%=xxAccDet%></div>
		  <div class="clientlogineditaccount">
<%			sSQL="SELECT clID,clUserName,clActions,clLoginLevel,clPercentDiscount,clEmail,clientCustom1,clientCustom2 FROM customerlogin WHERE clID="&SESSION("clientID")
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				theemail=rs("clEmail")
				clientCustom1=rs("clientCustom1")
				clientCustom2=rs("clientCustom2")
			else
				SESSION("clientID")=""
			end if
			rs.close
			sSQL="SELECT email FROM mailinglist WHERE email='"&escape_string(theemail)&"'"
			rs.open sSQL,cnn,0,1
			if rs.EOF then allowemail=0 else allowemail=1
			rs.close
			print whv("emhash", sha256(theemail & "asp know the hash" & adminSecret)) %>
			<div class="ectdivcontainer">
				<div class="ectdivleft"><%=redstar & xxName%></div>
				<div class="ectdivright"><input type="text" size="30" name="name" value="<%=htmlspecials(SESSION("clientUser"))%>" placeholder="<%=stripnspecials(xxName)%>" /></div>
			</div>
<%			if nounsubscribe<>TRUE then %>
			<div class="ectdivcontainer">
				<div class="ectdivleft"><input type="checkbox" name="allowemail" value="ON"<% if allowemail<>0 then print " checked=""checked"""%> /></div>
				<div class="ectdivright"><div><%=xxAlPrEm%></div><div style="font-size:10px"><%=xxNevDiv%></div></div>
			</div>
<%			end if %>
			<div class="ectdivcontainer">
				<div class="ectdivleft"><%=redstar & xxEmail%></div>
				<div class="ectdivright"><input type="text" size="30" name="email" value="<%=theemail%>" placeholder="<%=stripnspecials(xxEmail)%>" /></div>
			</div>
<%			if trim(extraclientfield1)<>"" then %>
			<div class="ectdivcontainer">
				<div class="ectdivleft"><%=IIfVs(extraclientfield1required,redstar) & extraclientfield1%></div>
				<div class="ectdivright"><input type="text" name="clientCustom1" size="30" value="<%=clientCustom1%>" placeholder="<%=stripnspecials(extraclientfield1)%>" /></div>
			</div>
<%			end if
			if trim(extraclientfield2)<>"" then %>
			<div class="ectdivcontainer">
				<div class="ectdivleft"><%=IIfVs(extraclientfield2required,redstar) & extraclientfield2%></div>
				<div class="ectdivright"><input type="text" name="clientCustom2" size="30" value="<%=clientCustom2%>" placeholder="<%=stripnspecials(extraclientfield2)%>" /></div>
			</div>
<%			end if %>
			<div class="ectdivhead"><%=xxPwdChg%></div>
			<div class="ectdivcontainer">
				<div class="ectdivleft"><%=xxOldPwd%></div>
				<div class="ectdivright"><input type="password" size="20" name="oldpw" value="" placeholder="<%=stripnspecials(xxOldPwd)%>" autocomplete="off" /></div>
			</div>
			<div class="ectdivcontainer">
				<div class="ectdivleft"><%=xxNewPwd%></div>
				<div class="ectdivright"><input type="password" size="20" name="newpw" value="" placeholder="<%=stripnspecials(xxNewPwd)%>" autocomplete="off" /></div>
			</div>
			<div class="ectdivcontainer">
				<div class="ectdivleft"><%=xxRptPwd%></div>
				<div class="ectdivright"><input type="password" size="20" name="newpw2" value="" placeholder="<%=stripnspecials(xxRptPwd)%>" autocomplete="off" /></div>
			</div>
			<div class="ectdiv2column"><%=imageorsubmit(imgsubmit,xxSubmt,"submit")&" "&imageorbutton(imgcancel,xxCancel,"cancel","history.go(-1)",TRUE)%></div>
		  </div>
		</div>
		</form>
<%		elseif getpost("action")="editaddress" OR getpost("action")="newaddress" then ' }{
			addID=replace(getpost("theid"),"'","")
			if NOT is_numeric(addID) then addID=0
			addIsDefault=""
			addName=""
			addLastName=""
			addAddress=""
			addAddress2=""
			addState=""
			addCity=""
			addZip=""
			addPhone=""
			addCountry=""
			addExtra1=""
			addExtra2=""
			havestate=FALSE
			sSQL="SELECT stateID FROM states INNER JOIN countries ON states.stateCountryID=countries.countryID WHERE countryEnabled<>0 AND stateEnabled<>0 AND (loadStates=2 OR countryID=" & origCountryID & ") ORDER BY stateCountryID,stateName"
			rs.open sSQL,cnn,0,1
			hasstates=(NOT rs.EOF)
			rs.close
			sSQL="SELECT countryName,countryOrder,"&getlangid("countryName",8)&" AS cnameshow,countryID,loadStates FROM countries WHERE countryEnabled=1 ORDER BY countryOrder DESC,"&getlangid("countryName",8)
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then allcountries=rs.getrows
			rs.close
			for rowcounter=0 to UBOUND(allcountries,2)
				if allcountries(4,rowcounter)=0 then nonhomecountries=TRUE : exit for
			next
			if NOT nonhomecountries then
				for rowcounter=0 to UBOUND(allcountries,2)
					if allcountries(4,rowcounter)>0 then
						sSQL="SELECT stateID FROM states WHERE stateEnabled<>0 AND stateCountryID=" & allcountries(3,rowcounter)
						rs.open sSQL,cnn,0,1
						if rs.EOF then nonhomecountries=TRUE
						rs.close
						if nonhomecountries then exit for
					end if
				next
			end if
			if getpost("action")="editaddress" then
				sSQL="SELECT addID,addIsDefault,addName,addLastName,addAddress,addAddress2,addState,addCity,addZip,addPhone,addCountry,addExtra1,addExtra2 FROM address WHERE addID=" & addID & " AND addCustID=" & SESSION("clientID") & " ORDER BY addIsDefault"
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					addIsDefault=rs("addIsDefault")
					addName=rs("addName")
					addLastName=rs("addLastName")
					addAddress=rs("addAddress")
					addAddress2=rs("addAddress2")
					addState=rs("addState")
					ordState=addState
					addCity=rs("addCity")
					addZip=rs("addZip")
					addPhone=rs("addPhone")
					addCountry=rs("addCountry")
					addExtra1=rs("addExtra1")
					addExtra2=rs("addExtra2")
				end if
				rs.close
			end if %>
	<form method="post" name="mainform" action="<%=thisaction%>" onsubmit="return checkform(this)">
	<input type="hidden" name="action" value="<% if getpost("action")="editaddress" then print "doeditaddress" else print "donewaddress" %>" />
	<input type="hidden" name="theid" value="<%=addID%>" />
	  <div class="ectdiv ectclientlogin">
		<div class="ectdivhead"><%=xxEdAdd%></div>
		<%	if trim(extraorderfield1)<>"" then %>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><%=IIfVr(extraorderfield1required=true,redstar,"") & extraorderfield1 %></div>
			<div class="ectdivright"><% if extraorderfield1html<>"" then print extraorderfield1html else print "<input type=""text"" name=""ordextra1"" id=""ordextra1"" size=""20"" value="""&htmlspecials(addExtra1&"")&""" autocomplete=""false"" />"%></div>
		</div>
		<%	end if %>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><%=redstar & xxName%></div>
			<div class="ectdivright"><%
		if usefirstlastname then
			print "<input type=""text"" name=""name"" class=""ectinputhalf"" size=""25"" value="""&htmlspecials(addName&"")&""" placeholder="""&stripnspecials(xxFirNam)&""" autocomplete=""given-name"" /> <input type=""text"" name=""lastname"" class=""ectinputhalf"" size=""25"" value="""&htmlspecials(addLastName&"")&""" placeholder="""&stripnspecials(xxLasNam)&""" autocomplete=""family-name"" />"
		else
			print "<input type=""text"" name=""name"" size=""25"" id=""name"" placeholder=""" & stripnspecials(xxName) & """ value="""&htmlspecials(trim(addName&" "&addLastName)&"")&""" />"
		end if %></div>
		</div>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><%=redstar & xxAddress%></div>
			<div class="ectdivright"><input type="text" name="address" id="address" size="25" value="<%=htmlspecials(addAddress&"")%>" placeholder="<%=stripnspecials(xxAddress)%>" /></div>
		</div>
		<%	if useaddressline2=TRUE then %>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><%=xxAddress2%></div>
			<div class="ectdivright"><input type="text" name="address2" id="address2" size="25" value="<%=htmlspecials(addAddress2&"")%>" placeholder="<%=stripnspecials(xxAddress2)%>" /></div>
		</div>
		<%	end if %>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><%=redstar & xxCity%></div>
			<div class="ectdivright"><input type="text" name="city" id="city" size="25" value="<%=htmlspecials(addCity&"")%>" placeholder="<%=stripnspecials(xxCity)%>" /></div>
		</div>
		<%	if hasstates OR nonhomecountries then %>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><%=replace(redstar,"<span","<span id=""statestar""")%><span id="statetxt"><%=xxState%></span></div>
			<div class="ectdivright"><select name="state" id="state" size="1" onchange="dosavestate('')"><% havestate=show_states(addState) %></select><input type="text" name="state2" id="state2" size="25" value="<% if NOT havestate then print htmlspecials(addState&"")%>" placeholder="<%=stripnspecials(xxState)%>" /></div>
		</div>
		<%	end if %>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><%=redstar & xxCountry%></div>
			<div class="ectdivright"><select name="country" id="country" size="1" onchange="checkoutspan('')" ><% call show_countries(addCountry,FALSE) %></select></div>
		</div>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><%=replace(redstar,"<span","<span id=""zipstar""") & "<span id=""ziptxt"">" & xxZip & "</span>"%></div>
			<div class="ectdivright"><input type="text" name="zip" id="zip" size="25" value="<%=htmlspecials(addZip&"")%>" placeholder="<%=stripnspecials(xxZip)%>" /></div>
		</div>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><%=redstar & xxPhone%></div>
			<div class="ectdivright"><input type="text" name="phone" id="phone" size="25" value="<%=htmlspecials(addPhone&"")%>" placeholder="<%=stripnspecials(xxPhone)%>" /></div>
		</div>
		<%	if trim(extraorderfield2)<>"" then %>
		<div class="ectdivcontainer">
			<div class="ectdivleft"><%=IIfVr(extraorderfield2required=true,redstar,"") & extraorderfield2 %></div>
			<div class="ectdivright"><% if extraorderfield2html<>"" then print extraorderfield2html else print "<input type=""text"" name=""ordextra2"" id=""ordextra2"" size=""25"" value="""&htmlspecials(addExtra2&"")&""" autocomplete=""false"" placeholder="""&stripnspecials(extraorderfield2)&""" />"%></div>
		</div>
		<%	end if %>
		<div class="ectdiv2column"><%=imageorsubmit(imgsubmit,xxSubmt,"submit")&" "&imageorbutton(imgcancel,xxCancel,"cancel","history.go(-1)",TRUE)%></div>
	  </div>
	</form>
<script>
/* <![CDATA[ */
var checkedfullname=false;
function zipoptional(cntobj){
var cntid=cntobj[cntobj.selectedIndex].value;
if(cntid==85 || cntid==91 || cntid==154 || cntid==200)return true; else return false;
}
function stateoptional(cntobj){
var cntid=cntobj[cntobj.selectedIndex].value;
if(false<%
rs.open "SELECT countryID FROM countries WHERE countryEnabled<>0 AND loadStates<0",cnn,0,1
do while NOT rs.EOF
	print "||cntid==" & rs("countryID")
	rs.movenext
loop
rs.close
%>)return true; else return false;
}
function checkform(frm)
{
<% if trim(extraorderfield1)<>"" AND extraorderfield1required=true then %>
if(frm.ordextra1.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&extraorderfield1)%>\".");
	frm.ordextra1.focus();
	return (false);
}
<% end if %>
if(frm.name.value==""||frm.name.value=="<%=xxFirNam%>"){
	alert("<%=jscheck(xxPlsEntr&" """&IIfVr(usefirstlastname, xxFirNam, xxName))%>\".");
	frm.name.focus();
	return (false);
}
<%	if usefirstlastname then %>
if(frm.lastname.value==""||frm.lastname.value=="<%=xxLasNam%>"){
	alert("<%=jscheck(xxPlsEntr&" """&xxLasNam)%>\".");
	frm.lastname.focus();
	return(false);
}
<%	else %>
gotspace=false;
var checkStr=frm.name.value;
for (i=0; i < checkStr.length; i++){
	if(checkStr.charAt(i)==" ")
		gotspace=true;
}
if(!checkedfullname && !gotspace){
	alert("<%=jscheck(xxFulNam&" """&xxName)%>\".");
	frm.name.focus();
	checkedfullname=true;
	return (false);
}
<%	end if %>
if(frm.address.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&xxAddress)%>\".");
	frm.address.focus();
	return (false);
}
if(frm.city.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&xxCity)%>\".");
	frm.city.focus();
	return (false);
}
	if(stateoptional(document.getElementById('country'))){
	}else if(stateselectordisabled[0]==false){
<%	if hasstates then %>
	if(frm.state.selectedIndex==0){
		alert("<%=jscheck(xxPlsSlct & " ")%>" + document.getElementById('statetxt').innerHTML);
		frm.state.focus();
		return(false);
	}
<%	end if %>
	}else{
<%	if nonhomecountries then %>
	if(frm.state2.value==""){
		alert("<%=jscheck(xxPlsEntr)%> \"" + document.getElementById('statetxt').innerHTML + "\".");
		frm.state2.focus();
		return(false);
	}
<%	end if %>}
if(frm.zip.value=="" && ! zipoptional(document.getElementById('country'))){
	alert("<%=jscheck(xxPlsEntr&" """&xxZip)%>\".");
	frm.zip.focus();
	return (false);
}
if(frm.phone.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&xxPhone)%>\".");
	frm.phone.focus();
	return (false);
}
<% if trim(extraorderfield2)<>"" AND extraorderfield2required=true then %>
if(frm.ordextra2.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&extraorderfield2)%>\".");
	frm.ordextra2.focus();
	return (false);
}
<% end if %>
return (true);
}
<% if termsandconditions=TRUE then %>
function showtermsandconds(){
newwin=window.open("termsandconditions.asp","Terms","menubar=no, scrollbars=yes, width=420, height=380, directories=no,location=no,resizable=yes,status=no,toolbar=no");
}
<% end if %>
var savestate=0;
var ssavestate=0;
function dosavestate(shp){
	thestate=eval('document.forms.mainform.'+shp+'state');
	eval(shp+'savestate=thestate.selectedIndex');
}
function checkoutspan(shp){
	document.getElementById(shp+'zipstar').style.display=(zipoptional(document.getElementById(shp+'country'))?'none':'');
	document.getElementById(shp+'statestar').style.display=(stateoptional(document.getElementById(shp+'country'))?'none':'');<%
	if hasstates then
		print "thestate=document.getElementById(shp+'state');"&vbCrLf
		print "dynamiccountries(document.getElementById(shp+'country'),shp);" & vbCrLf
	end if
	print "if(stateselectordisabled[shp=='s'?1:0]==false&&!stateoptional(document.getElementById(shp+'country'))){" & vbCrLf
	print "if(document.getElementById(shp+'state2'))document.getElementById(shp+'state2').style.display='none';"&vbCrLf
	if hasstates then
		print "thestate.disabled=false;"&vbCrLf
		print "eval('thestate.selectedIndex='+shp+'savestate');"&vbCrLf
		print "document.getElementById(shp+'state').style.display='';"&vbCrLf
	end if %>
}else{<%
	print "if(document.getElementById(shp+'state2'))document.getElementById(shp+'state2').style.display='';"&vbCrLf
	if hasstates then %>
		document.getElementById(shp+'state').style.display='none';
		if(thestate.disabled==false){
		thestate.disabled=true;
		eval(shp+'savestate=thestate.selectedIndex');
		thestate.selectedIndex=0;}
<%	end if %>
}}
<%	createdynamicstates("SELECT stateID,stateAbbrev,stateName,stateName2,stateName3,stateCountryID,countryName FROM states INNER JOIN countries ON states.stateCountryID=countries.countryID WHERE countryEnabled<>0 AND stateEnabled<>0 AND (loadStates=2 OR countryID=" & origCountryID & ") ORDER BY stateCountryID," & getlangid("stateName",1048576))
	print "checkoutspan('');setinitialstate('');" & vbCrLf
%>/* ]]> */
</script>
<%		elseif (getpost("action")="createlist" AND trim(replace(getpost("listname"),"<",""))<>"") OR getpost("action")="deletelist" OR getpost("action")="deleteaddress" OR getpost("action")="doeditaddress" OR getpost("action")="donewaddress" then ' }{
			viewing=""
			addID=replace(getpost("theid"),"'","")
			if NOT is_numeric(addID) then addID=0
			ordName=strip_tags2(getpost("name"))
			ordLastName=strip_tags2(getpost("lastname"))
			ordAddress=strip_tags2(getpost("address"))
			ordAddress2=strip_tags2(getpost("address2"))
			ordState=strip_tags2(getpost("state2"))
			if getpost("state")<>"" then ordState=strip_tags2(getpost("state"))
			ordState=strip_tags2(getstatefromid(ordState))
			ordCity=strip_tags2(getpost("city"))
			ordZip=strip_tags2(getpost("zip"))
			ordPhone=strip_tags2(getpost("phone"))
			ordCountry=strip_tags2(getcountryfromid(getpost("country")))
			ordExtra1=strip_tags2(getpost("ordextra1"))
			ordExtra2=strip_tags2(getpost("ordextra2"))
			headertext=""
			listname=trim(replace(getpost("listname"),"<",""))
			if getpost("action")="createlist" AND enablewishlists=TRUE AND listname<>"" then
				headertext=xxLisMan
				listaccess=calcmd5(timer() & listname & adminSecret)
				sSQL="INSERT INTO customerlists (listName,listOwner,listAccess) VALUES ('"&escape_string(listname)&"'," & SESSION("clientID") & ",'"&escape_string(listaccess)&"')"
				ect_query(sSQL)
			elseif getpost("action")="deletelist" AND enablewishlists=TRUE then
				headertext=xxLisMan
				sSQL="DELETE FROM customerlists WHERE listID=" & addID & " AND listOwner=" & SESSION("clientID")
				ect_query(sSQL)
				sSQL="DELETE FROM cart WHERE cartListID=" & addID & " AND cartClientID=" & SESSION("clientID")
				ect_query(sSQL)
			elseif getpost("action")="deleteaddress" then
				viewing="add"
				headertext=xxAddMan
				sSQL="DELETE FROM address WHERE addID=" & addID & " AND addCustID=" & SESSION("clientID")
				ect_query(sSQL) 
			elseif getpost("action")="donewaddress" then
				viewing="add"
				headertext=xxAddMan
				sSQL="INSERT INTO address (addCustID,addIsDefault,addName,addLastName,addAddress,addAddress2,addCity,addState,addZip,addCountry,addPhone,addExtra1,addExtra2) VALUES ("&SESSION("clientID")&",0,'"&escape_string(ordName)&"','"&escape_string(ordLastName)&"','"&escape_string(ordAddress)&"','"&escape_string(ordAddress2)&"','"&escape_string(ordCity)&"','"&escape_string(ordState)&"','"&escape_string(ordZip)&"','"&escape_string(ordCountry)&"','"&escape_string(ordPhone)&"','"&escape_string(ordExtra1)&"','"&escape_string(ordExtra2)&"')"
				ect_query(sSQL)
			elseif getpost("action")="doeditaddress" then
				viewing="add"
				headertext=xxAddMan
				sSQL="UPDATE address SET addName='"&escape_string(ordName)&"',addLastName='"&escape_string(ordLastName)&"',addAddress='"&escape_string(ordAddress)&"',addAddress2='"&escape_string(ordAddress2)&"',addCity='"&escape_string(ordCity)&"',addState='"&escape_string(ordState)&"',addZip='"&escape_string(ordZip)&"',addCountry='"&escape_string(ordCountry)&"',addPhone='"&escape_string(ordPhone)&"',addExtra1='"&escape_string(ordExtra1)&"',addExtra2='"&escape_string(ordExtra2)&"' WHERE addCustID="&SESSION("clientID")&" AND addID=" & addID
				ect_query(sSQL)
			end if
			print "<script>var url=[location.protocol,'//',location.host,location.pathname].join('')+'"&IIfVs(viewing<>"","?vw="&viewing)&"';setTimeout(function(){document.location=url},2000)</script>"
%>	  <div class="ectdiv ectclientlogin">
		<div class="ectmessagescreen">
			<div class="ectdivhead"><%=headertext%></div>
			<div style="padding:50px 0;text-align:center"><%=xxUpdSuc%></div>
		</div>
	  </div>
<%		else ' }{
%>
<script>
/* <![CDATA[ */
var currstate=[];
currstate['ad']='none';
currstate['am']='none';
currstate['gr']='none';
currstate['om']='none';
function showhidesection(sect){
	var elem=document.getElementsByTagName('div');
	currstate[sect]=currstate[sect]=='none'?'':'none';
	for(var i=0; i<elem.length; i++){
		var classes=elem[i].className;
		if(classes.indexOf(sect+'formrow')!=-1) elem[i].style.display=currstate[sect];
	}
	document.getElementById('sectimage'+sect).src=currstate[sect]=='none'?'images/arrow-down.png':'images/arrow-up.png';
	return false;
}
/* ]]> */</script>
		  <form method="post" name="mainform" action="<%=thisaction%>">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="action" value="none" />
			<input type="hidden" name="theid" value="" />
			<div class="ectdiv ectclientlogin">
				<div class="ectdivhead clientloginmainheader" style="cursor:pointer" onclick="showhidesection('ad')"><%=xxAccDet%><img id="sectimagead" src="images/arrow-down.png" style="float:right;margin-right:15px" /></div>
				<div class="adformrow" style="display:none">
				  <div class="ectclientloginaccount">
<%			sSQL="SELECT clID,clUserName,clActions,clLoginLevel,clPercentDiscount,clEmail,loyaltyPoints,clientCustom1,clientCustom2 FROM customerlogin WHERE clID="&SESSION("clientID")
			rs.open sSQL,cnn,0,1
			if rs.EOF then
				theemail="ACCOUNT DELETED"
			else
				theemail=rs("clEmail")
				loyaltypointtotal=rs("loyaltyPoints")
				clientCustom1=rs("clientCustom1")
				clientCustom2=rs("clientCustom2")
			end if
			rs.close
			sSQL="SELECT email,isconfirmed FROM mailinglist WHERE email='"&escape_string(theemail)&"'"
			rs.open sSQL,cnn,0,1
			if rs.EOF then allowemail=0 : isconfirmed=FALSE else allowemail=1 : isconfirmed=rs("isconfirmed")
			rs.close %>
					<div class="ectdivcontainer">
						<div class="ectdivleft"><%=xxName%></div>
						<div class="ectdivright"><%=htmlspecials(SESSION("clientUser"))%></div>
					</div>
<%			if nounsubscribe<>TRUE then %>
					<div class="ectdivcontainer">
						<div class="ectdivleft"><%=xxAlPrEm%><div style="font-size:10px"><%=xxNevDiv%></div></div>
						<div class="ectdivright"><% if noconfirmationemail<>TRUE AND allowemail<>0 AND isconfirmed=0 then print xxWaiCon else print "<input type=""checkbox"" name=""allowemail"" value=""ON""" & IIfVr(allowemail<>0, " checked=""checked""", "") & " disabled=""disabled"" />"%></div>
					</div>
<%			end if %>
					<div class="ectdivcontainer">
						<div class="ectdivleft"><%=xxEmail%></div>
						<div class="ectdivright"><%=theemail%></div>
					</div>
<%			if trim(extraclientfield1)<>"" then %>
					<div class="ectdivcontainer">
						<div class="ectdivleft"><%=extraclientfield1%></div>
						<div class="ectdivright"><%=clientCustom1%></div>
					</div>
<%			end if
			if trim(extraclientfield2)<>"" then %>
					<div class="ectdivcontainer">
						<div class="ectdivleft"><%=extraclientfield2%></div>
						<div class="ectdivright"><%=clientCustom2%></div>
					</div>
<%			end if
			if loyaltypoints<>"" then %>
					<div class="ectdivcontainer">
						<div class="ectdivleft"><%=xxLoyPoi%></div>
						<div class="ectdivright"><%=loyaltypointtotal%></div>
					</div>
<%			end if %>
					<div class="ectdiv2column"><%=xxChaAcc%> <a class="ectlink" href="javascript:editaccount()"><%=xxClkHere%></a>.</div>
				  </div>
				</div>
<%				' Address Management
%>			  <div class="ectdivhead clientloginmainheader" style="cursor:pointer" onclick="showhidesection('am')"><%=xxAddMan%><img id="sectimageam" src="images/arrow-<%=IIfVr(getget("vw")="add","up","down")%>.png" style="float:right;margin-right:15px" /></div>
			  <div class="amformrow"<% if getget("vw")<>"add" then print " style=""display:none"""%>>
				  <div class="ectclientloginaddress">
<%			sSQL="SELECT addID,addIsDefault,addName,addLastName,addAddress,addAddress2,addState,addCity,addZip,addPhone,addCountry FROM address WHERE addCustID=" & SESSION("clientID") & " ORDER BY addIsDefault"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				do while NOT rs.EOF
					print "<div class=""ectdivcontainer"">"
						print "<div class=""ectdivleft"">" & strip_tags2(trim(rs("addName")&" "&rs("addLastName"))) & "<br />" & strip_tags2(rs("addAddress")&"") & IIfVr(trim(rs("addAddress2")&"")<>"", "<br />" & strip_tags2(rs("addAddress2")&""), "") & "<br /> " & strip_tags2(rs("addCity")&"") & ", " & strip_tags2(rs("addState")&"") & IIfVr(rs("addZip")<>"", "<br />" & strip_tags2(rs("addZip")), "") & "<br />" & strip_tags2(rs("addCountry")&"") & "</div>"
						print "<div class=""ectdivright""><ul><li><a class=""ectlink"" href=""javascript:editaddress("&rs("addID")&")"">" & xxEdAdd & "</a><br /><br /></li><li><a class=""ectlink"" href=""javascript:deleteaddress("&rs("addID")&")"">" & xxDeAdd & "</a></li></ul></div>"
					print "</div>"
					rs.MoveNext
				loop
			else
				print "<div class=""nosearchresults"">" & xxNoAdd & "</div>"
			end if
			rs.close
%>
					<div class="ectdiv2column"><%=xxPCAdd%> <a class="ectlink" href="javascript:newaddress()"><%=xxClkHere%></a>.</div>
				  </div>
			  </div>
<%				' Gift Registry Management
			if enablewishlists=TRUE then
%>			  <div class="ectdivhead clientloginmainheader" style="cursor:pointer" onclick="showhidesection('gr')"><%=xxLisMan%><img id="sectimagegr" src="images/arrow-down.png" style="float:right;margin-right:15px" /></div>
			  <div class="grformrow" style="display:none">
				  <div class="ectclientlogingiftreg">
<%				sSQL="SELECT listID,listName,listAccess FROM customerlists WHERE listOwner=" & SESSION("clientID") & " ORDER BY listName"
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					do while NOT rs.EOF
						numitems=0
						sSQL="SELECT COUNT(*) AS numitems FROM cart WHERE cartListID="&rs("listID")
						rs2.open sSQL,cnn,0,1
						if NOT rs2.EOF then if NOT isnull(rs2("numitems")) then numitems=rs2("numitems")
						rs2.close
						print "<div class=""ectgiftregname"">" & htmlspecials(trim(rs("listName"))) & " (" & numitems & ")</div>"
						print "<div class=""ectdiv2column""><div class=""ectpublicaccesstext"">" & xxPubAcc & ":</div><div class=""ectpublicaccessinput""><input class=""giftregistrycopy"" id=""publicop" & rs("listID") & """ value=""" & storeurl & "cart" & extension & "?pli=" & rs("listID") & "&pla=" & rs("listAccess") & """ size=""60"" /></div></div>"
						print "<div class=""ectgiftregistrybuttons"">"
							print "<div class=""ectgiftregcopy""><input type=""button"" class=""ectbutton"" onclick=""document.getElementById('publicop" & rs("listID") & "').select();try{document.execCommand('copy')}catch(err){alert('Copy command not supported')}"" value=""Copy To Clipboard"" /></div>"
							if numitems>0 then print "<div class=""ectgiftregview"">" & imageorbutton(imgectgiftregview,xxVieGRe,"ectgiftregview","cart" & extension & "?pli=" & rs("listID"),FALSE) & "</div>"
							print "<div class=""ectgiftregdel"">" & imageorbutton(imgectgiftregdel,xxDelGRe,"ectgiftregdel","deletelist("&rs("listID")&")",TRUE) & "</div>"
						print "</div>"
						rs.MoveNext
					loop
				else
					print "<div class=""nosearchresults"">" & xxNoGRe & "</div>"
				end if
				rs.close
%>					<div>
						<%=imageorbutton(imgcreatelist,"Create New List","createlist","createlist()",TRUE)%>
						<input type="text" class="createlistinput" name="listname" size="40" maxlength="50" placeholder="<%=xxLisNam%>" />
					</div>
				  </div>
			  </div>
<%			end if
			' Order Management
%>			  <div class="ectdivhead clientloginmainheader" style="cursor:pointer" onclick="showhidesection('om')"><%=xxOrdMan%><img id="sectimageom" src="images/arrow-down.png" style="float:right;margin-right:15px" /></div>
			  <div class="omformrow" style="display:none">
				  <div class="ectclientloginorders">
<%			hastracknum=FALSE
			sSQL="SELECT ordID FROM orders WHERE ordClientID=" & SESSION("clientID") & " AND ordTrackNum<>''"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then hastracknum=TRUE
			rs.close %>
					<div class="ectclientloginordersrow ectclientloginordershead">
						<div><%=xxOrdId%></div>
						<div><%=xxDate%></div>
						<div><%=xxStatus%></div>
<%			if hastracknum then print "<div>"&xxTraNum&"</div>" %>
						<div><%=xxGndTot%></div>
						<div><%=xxCODets%></div>
					</div>
<%			success=TRUE
			sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 60 ")&"ordID,ordDate,ordTrackNum,ordTotal,ordStateTax,ordCountryTax,ordShipping,ordHSTTax,ordHandling,ordDiscount,"&getlangid("statPublic",64)&" FROM orders LEFT OUTER JOIN orderstatus ON orders.ordStatus=orderstatus.statID WHERE ordStatus<>1 AND ordClientID=" & SESSION("clientID") & " ORDER BY ordDate DESC"&IIfVs(mysqlserver=TRUE," LIMIT 0,60")
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				do while NOT rs.EOF
					print "<div class=""ectclientloginordersrow"">"
						print "<div>" & rs("ordID") & "</div>"
						print "<div>" & rs("ordDate") & "</div>"
						print "<div>" & rs(getlangid("statPublic",64)) & "</div>"
						if hastracknum then
							print "<div>"
							tracknumarr=split(rs("ordTrackNum")&"",",")
							for uoindex=0 to UBOUND(tracknumarr)
								thecarrier=getcarrierfromtrack(tracknumarr(uoindex),thelink)
								print "<p class=""tracknumline"">"
								if thelink<>"" then print "<a href=""" & thelink & tracknumarr(uoindex) & """ target=""_blank"">" & tracknumarr(uoindex) & "</a>" else print tracknumarr(uoindex)
								print "</p>"
							next
							print "</div>"
						end if
						print "<div>" & FormatEuroCurrency((rs("ordTotal")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordShipping")+rs("ordHSTTax")+rs("ordHandling"))-rs("ordDiscount")) & "</div>"
						print "<div><a class=""ectlink"" href=""javascript:vieworder("&rs("ordID")&")"">" & xxClkHere & "</a></div>"
					print "</div>"
					rs.MoveNext
				loop
			else
				success=FALSE
			end if
			rs.close
%>
				  </div>
<%			if NOT success then print "<div class=""nosearchresults"">" & xxNoOrd & "</div>" %>

			  </div>
			</div>
		  </form>
<script>
/* <![CDATA[ */
if(document.location.hash=='#ord')showhidesection('om');
else if(document.location.hash=='#list')showhidesection('gr');
else if(document.location.hash=='#add')showhidesection('am');
else if(document.location.hash=='#acct')showhidesection('ad');
/* ]]> */</script>
<%		end if ' }
	end if ' }
cnn.Close
set rs=nothing
set cnn=nothing
%>