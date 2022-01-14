<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
cartisincluded=TRUE
%>
<!--#include file="inccart.asp"-->
<%
Dim aFields(3)
success=true
if dateadjust="" then dateadjust=0
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs2=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
sSQL=""
alldata=""
dorefresh=FALSE
if maxloginlevels="" then maxloginlevels=5
resultcounter=0
sub dodeletecl(clid)
	if is_numeric(clid) then
		sSQL="DELETE FROM customerlogin WHERE clID=" & clid
		ect_query(sSQL)
		sSQL="DELETE FROM address WHERE addCustID=" & clid
		ect_query(sSQL)
		sSQL="UPDATE orders SET ordClientID=0 WHERE ordClientID=" & clid
		ect_query(sSQL)
	end if
end sub
if getpost("oldcntryname")<>"" AND getpost("newcntryname")<>"" then
	if getpost("newcntryname")="xxxdeletexxx" then
		ect_query("DELETE FROM address WHERE addCountry='"&escape_string(getpost("oldcntryname"))&"'")
	else
		ect_query("UPDATE address SET addCountry='"&escape_string(getpost("newcntryname"))&"' WHERE addCountry='"&escape_string(getpost("oldcntryname"))&"'")
	end if
elseif getpost("posted")="1" then
	if getpost("act")="quickupdate" then
		clact=getpost("clact")
		for each objItem in request.form
			if left(objItem, 4)="pra_" then
				origid=right(objItem, len(objItem)-4)
				theid=getpost("pid"&origid)
				theval=getpost(objItem)
				sSQL=""
				if clact="dsc" then
					if is_numeric(theval) then sSQL="UPDATE customerlogin SET clPercentDiscount=" & escape_string(theval)
				elseif clact="ste" OR clact="cte" OR clact="she" OR clact="wsp" OR clact="ped" OR clact="hae" then
					fieldnum=1
					if clact="cte" then fieldnum=2
					if clact="she" then fieldnum=4
					if clact="wsp" then fieldnum=8
					if clact="ped" then fieldnum=16
					if clact="hae" then fieldnum=32
					rs.open "SELECT clActions FROM customerlogin WHERE clID="&escape_string(theid),cnn,0,1
					if NOT rs.EOF then theval=rs("clActions") else theval=0
					rs.close
					if getpost("prb_" & origid)="1" then
						if (theval AND fieldnum)=0 then theval=theval + fieldnum
					else
						if (theval AND fieldnum)<>0 then theval=theval - fieldnum
					end if
					sSQL="UPDATE customerlogin SET clActions=" & theval
				elseif clact="lol" then
					sSQL="UPDATE customerlogin SET clLoginLevel='" & escape_string(theval) & "'"
				elseif clact="del" then
					if theval="del" then dodeletecl(theid)
					sSQL=""
				end if
				if sSQL<>"" then
					sSQL=sSQL & " WHERE clID="&int(theid)
					ect_query(sSQL)
				end if
			end if
		next
		dorefresh=TRUE
	elseif getpost("act")="delete" then
		dodeletecl(getpost("id"))
		dorefresh=TRUE
	elseif getpost("act")="deleteaddress" then
		sSQL="DELETE FROM address WHERE addID=" & getpost("theid")
		ect_query(sSQL)
	elseif getpost("act")="doeditaddress" OR getpost("act")="donewaddress" then
		addID=replace(getpost("theid"),"'","")
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
		if getpost("act")="doeditaddress" then
			sSQL="UPDATE address SET addName='"&escape_string(ordName)&"',addLastName='"&escape_string(ordLastName)&"',addAddress='"&escape_string(ordAddress)&"',addAddress2='"&escape_string(ordAddress2)&"',addCity='"&escape_string(ordCity)&"',addState='"&escape_string(ordState)&"',addZip='"&escape_string(ordZip)&"',addCountry='"&escape_string(ordCountry)&"',addPhone='"&escape_string(ordPhone)&"',addExtra1='"&escape_string(ordExtra1)&"',addExtra2='"&escape_string(ordExtra2)&"' WHERE addID=" & addID
		else
			sSQL="INSERT INTO address (addCustID,addIsDefault,addName,addLastName,addAddress,addAddress2,addCity,addState,addZip,addCountry,addPhone,addExtra1,addExtra2) VALUES ("&getpost("id")&",0,'"&escape_string(ordName)&"','"&escape_string(ordLastName)&"','"&escape_string(ordAddress)&"','"&escape_string(ordAddress2)&"','"&escape_string(ordCity)&"','"&escape_string(ordState)&"','"&escape_string(ordZip)&"','"&escape_string(ordCountry)&"','"&escape_string(ordPhone)&"','"&escape_string(ordExtra1)&"','"&escape_string(ordExtra2)&"')"
		end if
		ect_query(sSQL)
	elseif getpost("act")="domodify" then
		sSQL="SELECT clEmail FROM customerlogin WHERE clID<>" & getpost("id") &  " AND clEmail='" & escape_string(getpost("clEmail")) & "'"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			success=false
			errmsg=yyEmReg & "<br />" & htmlspecials(getpost("clEmail"))
		end if
		rs.close
		if getpost("clUserName")="" then
			success=false
			errmsg="Username is NULL"
		end if
		if success then
			sSQL="UPDATE customerlogin SET clUserName='" & escape_string(getpost("clUserName")) & "'"
			if getpost("clPW")<>"" then sSQL=sSQL & ",clPW='" & escape_string(dohashpw(getpost("clPW"))) & "'"
			sSQL=sSQL & ",clLoginLevel=" & getpost("clLoginLevel")
			sSQL=sSQL & ",loyaltyPoints=" & getpostint("loyaltyPoints")
			cpd=getpost("clPercentDiscount")
			sSQL=sSQL & ",clPercentDiscount=" & IIfVr(is_numeric(cpd), cpd, 0)
			if trim(extraclientfield1)<>"" then sSQL=sSQL & ",clientCustom1='" & escape_string(getpost("clientCustom1")) & "'"
			if trim(extraclientfield2)<>"" then sSQL=sSQL & ",clientCustom2='" & escape_string(getpost("clientCustom2")) & "'"
			sSQL=sSQL & ",clientAdminNotes='" & escape_string(getpost("clientAdminNotes")) & "'"
			sSQL=sSQL & ",clEmail='" & escape_string(getpost("clEmail")) & "'"
			clActions=0
			for each objItem in request.form("clActions")
				clActions=clActions + Int(objItem)
			next
			sSQL=sSQL & ",clActions=" & clActions
			sSQL=sSQL & " WHERE clID=" & getpost("id")
			ect_query(sSQL)
			if getpost("clAllowEmail")="ON" then
				on error resume next
				ect_query("INSERT INTO mailinglist (email,isconfirmed,mlConfirmDate,mlIPAddress) VALUES ('" & lcase(escape_string(getpost("clEmail"))) & "',1," & vsusdate(date())&",'"&left(request.servervariables("REMOTE_ADDR"), 48)&"')")
				on error goto 0
			else
				ect_query("DELETE FROM mailinglist WHERE email='" & escape_string(getpost("clEmail")) & "'")
			end if
			dorefresh=TRUE
		end if
	elseif getpost("act")="doaddnew" then
		sSQL="SELECT clEmail FROM customerlogin WHERE clEmail='" & escape_string(getpost("clEmail")) & "'"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			success=false
			errmsg=yyEmReg & "<br />" & htmlspecials(getpost("clEmail"))
		end if
		rs.close
		if getpost("clUserName")="" then
			success=false
			errmsg="Username is NULL"
		end if
		if success then
			sSQL="INSERT INTO customerlogin (clUserName,clPW,clLoginLevel,loyaltyPoints,clPercentDiscount,clientCustom1,clientCustom2,clientAdminNotes,clEmail,clDateCreated,clActions) VALUES ("
			sSQL=sSQL & "'" & escape_string(getpost("clUserName")) & "'"
			sSQL=sSQL & ",'" & escape_string(dohashpw(getpost("clPW"))) & "'"
			sSQL=sSQL & "," & getpost("clLoginLevel")
			sSQL=sSQL & "," & IIfVr(is_numeric(getpost("loyaltyPoints")),getpost("loyaltyPoints"),0)
			cpd=getpost("clPercentDiscount")
			sSQL=sSQL & "," & IIfVr(is_numeric(cpd), cpd, 0)
			sSQL=sSQL & ",'" & escape_string(getpost("clientCustom1")) & "'"
			sSQL=sSQL & ",'" & escape_string(getpost("clientCustom2")) & "'"
			sSQL=sSQL & ",'" & escape_string(getpost("clientAdminNotes")) & "'"
			sSQL=sSQL & ",'" & escape_string(getpost("clEmail")) & "'"
			sSQL=sSQL & "," & vsusdate(DateAdd("h",dateadjust,Now()))
			clActions=0
			for each objItem in request.form("clActions")
				clActions=clActions + Int(objItem)
			next
			sSQL=sSQL & "," & clActions & ")"
			ect_query(sSQL)
			if getpost("clAllowEmail")="ON" then
				on error resume next
				ect_query("INSERT INTO mailinglist (email,isconfirmed,mlConfirmDate,mlIPAddress) VALUES ('" & lcase(escape_string(getpost("clEmail"))) & "',1," & vsusdate(date())&",'"&left(request.servervariables("REMOTE_ADDR"), 48)&"')")
				on error goto 0
			else
				ect_query("DELETE FROM mailinglist WHERE email='" & escape_string(getpost("clEmail")) & "'")
			end if
			dorefresh=TRUE
		end if
	elseif getpost("act")="addorphans" then
		sSQL="SELECT clEmail FROM customerlogin WHERE clID=" & getpost("id")
		rs.open sSQL,cnn,0,1
		theemail=rs("clEmail")
		rs.close
		if loyaltypoints<>"" then
			loyaltypointtotal=0
			sSQL="SELECT SUM(loyaltyPoints) AS pointsSum FROM orders WHERE ordClientID=0 AND ordEmail='" & escape_string(theemail) & "'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if rs("pointsSum")<>NULL then loyaltypointtotal=rs("pointsSum")
			end if
			rs.close
			sSQL="UPDATE customerlogin SET loyaltyPoints=loyaltyPoints+" & loyaltypointtotal & " WHERE clID=" & getpost("id")
			ect_query(sSQL)
		end if
		sSQL="UPDATE orders SET ordClientID=" & getpost("id") & " WHERE ordEmail='" & escape_string(theemail) & "'"
		ect_query(sSQL)
	elseif getpost("act")="addorphan" then
		if loyaltypoints<>"" then
			loyaltypointtotal=0
			sSQL="SELECT loyaltyPoints FROM orders WHERE ordClientID=0 AND ordID=" & getpost("theid")
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then loyaltypointtotal=rs("loyaltyPoints")
			rs.close
			sSQL="UPDATE customerlogin SET loyaltyPoints=loyaltyPoints+" & loyaltypointtotal & " WHERE clID=" & getpost("id")
			ect_query(sSQL)
		end if
		sSQL="UPDATE orders SET ordClientID=" & getpost("id") & " WHERE ordID=" & getpost("theid")
		ect_query(sSQL)
	end if
end if
if dorefresh then
	print "<meta http-equiv=""refresh"" content=""1; url=adminclientlog.asp"
	print "?stext=" & urlencode(getpost("stext")) & "&accdate=" & urlencode(getpost("accdate")) & "&slevel=" & getpost("slevel") & "&stype=" & getpost("stype") & "&daterange=" & getpost("daterange") & "&pg=" & getpost("pg")
	print """>"
end if
%>
<script>
<!--
function formvalidator(theForm){
if (theForm.clUserName.value == ""){
alert("<%=jscheck(yyPlsEntr&" """&yyLiName)%>\".");
theForm.clUserName.focus();
return (false);
}
return (true);
}
function vieworder(theid){
	document.location="adminorders.asp?id="+theid;
}
function editaddress(theid){
	document.forms.mainform.act.value="editaddress";
	document.forms.mainform.theid.value=theid;
	document.forms.mainform.submit();
}
function newaddress(){
	document.forms.mainform.act.value="newaddress";
	document.forms.mainform.submit();
}
function editaccount(){
	document.forms.mainform.act.value="modify";
	document.forms.mainform.submit();
}
function addorphan(theid){
	if(confirm("<%=jscheck(yySureCa)%>")){
		document.forms.mainform.act.value="addorphan";
		document.forms.mainform.theid.value=theid;
		document.forms.mainform.submit();
	}
}
function addorphans(){
	if(confirm("<%=jscheck(yySureCa)%>")){
		document.forms.mainform.act.value="addorphans";
		document.forms.mainform.submit();
	}
}
function deleteaddress(theid){
	if(confirm("<%=jscheck(xxDelAdd)%>")){
		document.forms.mainform.act.value="deleteaddress";
		document.forms.mainform.theid.value=theid;
		document.forms.mainform.submit();
	}
}
//-->
</script>
<%	if getpost("posted")="1" AND (getpost("act")="modify" OR getpost("act")="addnew") then
		if getpost("act")="modify" AND is_numeric(getpost("id")) then
			sSQL="SELECT clUserName,clPW,clLoginLevel,clActions,clPercentDiscount,clEmail,clDateCreated,loyaltyPoints,clientCustom1,clientCustom2,clientAdminNotes FROM customerlogin WHERE clID="&getpost("id")
			rs.open sSQL,cnn,0,1
			clUserName=rs("clUserName")
			clPW=""
			clLoginLevel=rs("clLoginLevel")
			clActions=rs("clActions")
			clPercentDiscount=rs("clPercentDiscount")
			clEmail=rs("clEmail")
			clDateCreated=rs("clDateCreated")
			clLoyaltyPoints=rs("loyaltyPoints")
			clientCustom1=rs("clientCustom1")
			clientCustom2=rs("clientCustom2")
			clientAdminNotes=rs("clientAdminNotes")
			if NOT isdate(clDateCreated) then
				sSQL="UPDATE customerlogin SET clDateCreated=" & vsusdate(DateAdd("h",dateadjust,Now())) & " WHERE clID="&getpost("id")
				ect_query(sSQL)
				clDateCreated=Date()
			end if
			rs.close
			sSQL="SELECT email FROM mailinglist WHERE email='"&escape_string(clEmail)&"'"
			rs.open sSQL,cnn,0,1
			if rs.EOF then clAllowEmail=0 else clAllowEmail=1
			rs.close
		else
			clUserName=""
			clPW=""
			clLoginLevel=0
			clActions=0
			clPercentDiscount=0
			clEmail=""
			clDateCreated=Date()
			clAllowEmail=0
			clLoyaltyPoints=0
			clientCustom1=""
			clientCustom2=""
			clientAdminNotes=""
		end if
%>
	<form name="mainform" method="post" action="adminclientlog.asp" onsubmit="return formvalidator(this)">
<%			call writehiddenvar("posted", "1")
			if getpost("act")="modify" then
				call writehiddenvar("act", "domodify")
			else
				call writehiddenvar("act", "doaddnew")
			end if
			call writehiddenvar("stext", getpost("stext"))
			call writehiddenvar("accdate", getpost("accdate"))
			call writehiddenvar("daterange", getpost("daterange"))
			call writehiddenvar("slevel", getpost("slevel"))
			call writehiddenvar("stype", getpost("stype"))
			call writehiddenvar("pg", getpost("pg"))
			call writehiddenvar("id", getpost("id")) %>
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
			  <tr> 
                <td width="100%" colspan="4" align="center"><strong><%=yyLiAdm%></strong><br /><br /><%="<strong>"&yyDateCr&":</strong> " & FormatDateTime(clDateCreated,2) & "<br /><br />"%></td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyLiName%>:</strong></td>
				<td align="left"><input type="text" name="clUserName" size="20" value="<%=htmlspecials(clUserName)%>" /></td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyEmail%>:</strong></td>
				<td align="left"><input type="text" name="clEmail" size="30" value="<%=htmlspecials(clEmail)%>" /></td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyReset&" "&yyPass%>:</strong></td>
				<td align="left"><input type="text" name="clPW" size="20" value="" /></td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyLiLev%>:</strong></td>
				<td align="left"><select name="clLoginLevel" size="1">
				<%	for rowcounter=0 to maxloginlevels
						print "<option value='"&rowcounter&"'"
						if rowcounter=int(clLoginLevel) then print " selected=""selected"""
						print ">&nbsp; "&rowcounter&" </option>"&vbCrLf
					next
				%>
				</select></td>
			  </tr>
			  <tr>
				<td align="right" valign="top"><strong><%=yyActns%>:</strong></td>
				<td align="left" valign="top"><select name="clActions" size="6" multiple="multiple">
				<option value="1"<% if (clActions AND 1)=1 then print " selected=""selected""" %>><%=yyExStat%></option>
				<option value="2"<% if (clActions AND 2)=2 then print " selected=""selected""" %>><%=yyExCoun%></option>
				<option value="4"<% if (clActions AND 4)=4 then print " selected=""selected""" %>><%=yyExShip%></option>
				<option value="32"<% if (clActions AND 32)=32 then print " selected=""selected""" %>><%=yyExHand%></option>
				<option value="8"<% if (clActions AND 8)=8 then print " selected=""selected""" %>><%=yyWholPr%></option>
				<option value="16"<% if (clActions AND 16)=16 then print " selected=""selected""" %>><%=yyPerDis%></option>
				<option value="64"<% if (clActions AND 64)=64 then print " selected=""selected""" %>><%="Share Carts"%></option>
				</select></td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyPerDis%>:</strong></td>
				<td align="left"><input type="text" name="clPercentDiscount" size="10" value="<%=clPercentDiscount%>" /></td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyAllEml%>:</strong></td>
				<td align="left"><input type="checkbox" name="clAllowEmail" value="ON"<% if clAllowEmail<>0 then print " checked"%> /></td>
			  </tr>
<%			if loyaltypoints<>"" then %>
			  <tr>
				<td align="right" height="22"><strong><%=xxLoyPoi%>:</strong></td>
				<td align="left"><input type="text" name="loyaltyPoints" size="10" value="<%=clLoyaltyPoints%>" /></td>
			  </tr>
<%			end if
			if trim(extraclientfield1)<>"" then %>
			  <tr>
				<td align="right" height="22"><strong><%=extraclientfield1%>:</strong></td>
				<td align="left"><input type="text" name="clientCustom1" size="30" value="<%=clientCustom1%>" /></td>
			  </tr>
<%			end if
			if trim(extraclientfield2)<>"" then %>
			  <tr>
				<td align="right" height="22"><strong><%=extraclientfield2%>:</strong></td>
				<td align="left"><input type="text" name="clientCustom2" size="30" value="<%=clientCustom2%>" /></td>
			  </tr>
<%			end if %>
			  <tr>
				<td align="right" height="22"><strong>Client Admin Notes:</strong></td>
				<td align="left"><textarea name="clientAdminNotes" cols="60" rows="10"><%=htmlspecials(clientAdminNotes)%></textarea></td>
			  </tr>
			  <tr>
                <td width="100%" colspan="4" align="center"><br /><input type="submit" value="<%=yySubmit%>" />&nbsp;<input type="reset" value="<%=yyReset%>" /><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
                          &nbsp;</td>
			  </tr>
            </table>
	</form>
<%	elseif getpost("act")="editaddress" OR getpost("act")="newaddress" then ' }{
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
			if getpost("act")="editaddress" then
				sSQL="SELECT addID,addIsDefault,addName,addLastName,addAddress,addAddress2,addState,addCity,addZip,addPhone,addCountry,addExtra1,addExtra2 FROM address WHERE addID=" & addID
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					addIsDefault=rs("addIsDefault")
					addName=strip_tags2(rs("addName"))
					addLastName=strip_tags2(rs("addLastName"))
					addAddress=strip_tags2(rs("addAddress"))
					addAddress2=strip_tags2(rs("addAddress2"))
					addState=strip_tags2(rs("addState"))
					ordState=addState
					addCity=strip_tags2(rs("addCity"))
					addZip=strip_tags2(rs("addZip"))
					addPhone=strip_tags2(rs("addPhone"))
					addCountry=strip_tags2(rs("addCountry"))
					addExtra1=strip_tags2(rs("addExtra1"))
					addExtra2=strip_tags2(rs("addExtra2"))
				end if
				rs.close
			end if %>
	<form method="post" name="mainform" action="" onsubmit="return checkform(this)">
	<input type="hidden" name="act" value="<% if getpost("act")="editaddress" then print "doeditaddress" else print "donewaddress" %>" />
	<input type="hidden" name="theid" value="<%=addID%>" />
	<input type="hidden" name="id" value="<%=getpost("id")%>" />
	<input type="hidden" name="posted" value="1" />
	  <table class="cobtbl" width="100%" border="0" cellspacing="1" cellpadding="3">
		<tr><td align="center" class="cobhl" colspan="2" height="32"><strong><%=xxEdAdd%></strong></td></tr>
		<%	if trim(extraorderfield1)<>"" then %>
		<tr><td align="right" class="cobhl"><strong><%=IIfVr(extraorderfield1required=true,redstar,"") & extraorderfield1 %>:</strong></td><td class="cobll"><% if extraorderfield1html<>"" then print extraorderfield1html else print "<input type=""text"" name=""ordextra1"" id=""ordextra1"" size=""20"" value="""&htmlspecials(addExtra1&"")&""" />"%></td></tr>
		<%	end if %>
		<tr><td align="right" width="40%" class="cobhl"><strong><%=redstar & xxName%>:</strong></td><td class="cobll"><%
		if usefirstlastname then
			thestyle=""
			if addName="" AND addLastName="" then addName=xxFirNam : addLastName=xxLasNam : thestyle="style=""color:#BBBBBB"" "
			print "<input type=""text"" name=""name"" size=""11"" value="""&htmlspecials(addName&"")&""" alt="""&xxFirNam&""" onfocus=""if(this.value=='"&xxFirNam&"'){this.value='';this.style.color='';}"" "&thestyle&"/> <input type=""text"" name=""lastname"" size=""11"" value="""&htmlspecials(addLastName&"")&""" alt="""&xxLasNam&""" onfocus=""if(this.value=='"&xxLasNam&"'){this.value='';this.style.color='';}"" "&thestyle&"/>"
		else
			print "<input type=""text"" name=""name"" id=""name"" size=""20"" value="""&htmlspecials(addName&"")&""" />"
		end if %></td></tr>
		<tr><td align="right" class="cobhl"><strong><%=redstar & xxAddress%>:</strong></td><td class="cobll"><input type="text" name="address" id="address" size="25" value="<%=htmlspecials(addAddress&"")%>" /></td></tr>
		<%	if useaddressline2=TRUE then %>
		<tr><td align="right" class="cobhl"><strong><%=xxAddress2%>:</strong></td><td class="cobll"><input type="text" name="address2" id="address2" size="25" value="<%=htmlspecials(addAddress2&"")%>" /></td></tr>
		<%	end if %>
		<tr><td align="right" class="cobhl"><strong><%=redstar & xxCity%>:</strong></td><td class="cobll"><input type="text" name="city" id="city" size="20" value="<%=htmlspecials(addCity&"")%>" /></td></tr>
		<%	if hasstates OR nonhomecountries then %>
		<tr><td align="right" class="cobhl"><strong><%=replace(redstar,"<span","<span id=""statestar""")%><span id="statetxt"><%=xxState%></span>:</strong></td><td class="cobll"><select name="state" id="state" size="1" onchange="dosavestate('')"><% havestate=show_states(addState) %></select><input type="text" name="state2" id="state2" size="20" value="<% if NOT havestate then print htmlspecials(addState&"")%>" /></td></tr>
		<%	end if %>
		<tr><td align="right" class="cobhl"><strong><%=redstar & xxCountry%>:</strong></td><td class="cobll"><select name="country" id="country" size="1" onchange="checkoutspan('')" ><% call show_countries(addCountry,FALSE) %></select></td></tr>
		<tr><td align="right" class="cobhl"><strong><%=replace(redstar,"<span","<span id=""zipstar""") & "<span id=""ziptxt"">" & xxZip & "</span>"%>:</strong></td><td class="cobll"><input type="text" name="zip" id="zip" size="10" value="<%=htmlspecials(addZip&"")%>" /></td></tr>
		<tr><td align="right" class="cobhl"><strong><%=redstar & xxPhone%>:</strong></td><td class="cobll"><input type="text" name="phone" id="phone" size="20" value="<%=htmlspecials(addPhone&"")%>" /></td></tr>
		<%	if trim(extraorderfield2)<>"" then %>
		<tr><td align="right" class="cobhl"><strong><%=IIfVr(extraorderfield2required=true,redstar,"") & extraorderfield2 %>:</strong></td><td class="cobll"><% if extraorderfield2html<>"" then print extraorderfield2html else print "<input type=""text"" name=""ordextra2"" id=""ordextra2"" size=""20"" value="""&htmlspecials(addExtra2&"")&""" />"%></td></tr>
		<%	end if %>
		<tr><td align="center" colspan="2" class="cobll"><input type="submit" value="<%=xxSubmt%>" /> <input type="button" value="Cancel" onclick="history.go(-1)" /></td></tr>
	  </table>
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
<%	elseif (getpost("act")="viewacct" OR getpost("act")="deleteaddress" OR getpost("act")="addorphans" OR getpost("act")="addorphan") AND is_numeric(getpost("id")) then
		clID=getpost("id")
		sSQL="SELECT clUserName,clPW,clLoginLevel,clActions,clPercentDiscount,clEmail,clDateCreated,loyaltyPoints,clientCustom1,clientCustom2,clientAdminNotes FROM customerlogin WHERE clID="&clID
		rs.open sSQL,cnn,0,1
		clUserName=rs("clUserName")
		clPW=rs("clPW")
		clLoginLevel=rs("clLoginLevel")
		clActions=rs("clActions")
		clPercentDiscount=rs("clPercentDiscount")
		clEmail=rs("clEmail")
		clDateCreated=rs("clDateCreated")
		clLoyaltyPoints=rs("loyaltyPoints")
		clientCustom1=rs("clientCustom1")
		clientCustom2=rs("clientCustom2")
		clientAdminNotes=rs("clientAdminNotes")
		if NOT isdate(clDateCreated) then
			sSQL="UPDATE customerlogin SET clDateCreated=" & vsusdate(DateAdd("h",dateadjust,Now())) & " WHERE clID="&getpost("id")
			ect_query(sSQL)
			clDateCreated=Date()
		end if
		rs.close
		sSQL="SELECT email FROM mailinglist WHERE email='"&escape_string(clEmail)&"'"
		rs.open sSQL,cnn,0,1
		if rs.EOF then clAllowEmail=0 else clAllowEmail=1
		rs.close
		ordersnotinacct=0
		sSQL="SELECT COUNT(*) AS thecnt FROM orders WHERE ordClientID=0 AND ordEmail='"&escape_string(clEmail)&"'"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			if NOT isnull(rs("thecnt")) then ordersnotinacct=rs("thecnt")
		end if
		rs.close %>
		  <form method="post" name="mainform" action="">
<%			call writehiddenvar("posted", "1")
			call writehiddenvar("act", "none")
			call writehiddenvar("theid", "")
			call writehiddenvar("stext", getpost("stext"))
			call writehiddenvar("accdate", getpost("accdate"))
			call writehiddenvar("daterange", getpost("daterange"))
			call writehiddenvar("slevel", getpost("slevel"))
			call writehiddenvar("stype", getpost("stype"))
			call writehiddenvar("pg", getpost("pg"))
			call writehiddenvar("id", clID) %>
			<table class="cobtbl" width="100%" border="0" cellspacing="1" cellpadding="3">
              <tr> 
                <td class="cobhl" align="center" height="34"><strong><%=xxAccDet%></strong></td>
			  </tr>
			  <tr> 
                <td class="cobll" height="34" align="center">
				  <table class="cobtbl" width="100%" border="0" cellspacing="1" cellpadding="3">
<%			sSQL="SELECT email,isconfirmed FROM mailinglist WHERE email='"&escape_string(clEmail)&"'"
			rs.open sSQL,cnn,0,1
			if rs.EOF then allowemail=0 : isconfirmed=FALSE else allowemail=1 : isconfirmed=rs("isconfirmed")
			rs.close %>
					<tr><td class="cobhl" align="right" width="25%" height="22"><strong><%=xxName%>:</strong></td>
					<td class="cobll" align="left" width="25%"><%=htmlspecials(clUserName)%></td>
					<td class="cobhl" align="right" width="25%"><strong><%=yyActns%>:</strong></td>
					<td class="cobll" align="left" width="25%"><%
						if (clActions AND 1)=1 then print "STE "
						if (clActions AND 2)=2 then print "CTE "
						if (clActions AND 4)=4 then print "SHE "
						if (clActions AND 32)=32 then print "HAE "
						if (clActions AND 8)=8 then print "WSP "
						if (clActions AND 16)=16 then print "PED "
						if (clActions AND 64)=64 then print "SHC "
					%>&nbsp;</td>
					</tr>
					<tr><td class="cobhl" align="right" height="22"><strong><%=xxEmail%>:</strong></td>
					<td class="cobll" align="left"><%=htmlspecials(clEmail&"")%></td>
					<td class="cobhl" align="right"><strong><%=xxAlPrEm%>:</strong></td>
					<td class="cobll" align="left"><% if noconfirmationemail<>TRUE AND allowemail<>0 AND isconfirmed=0 then print xxWaiCon else print "<input type=""checkbox"" name=""allowemail"" value=""ON""" & IIfVr(allowemail<>0, " checked=""checked""", "") & " disabled=""disabled"" />"%></td>
					</tr>
					<tr><td class="cobhl" align="right" height="22"><strong><%=yyPerDis%>:</strong></td>
					<td class="cobll" align="left"><% if (clActions AND 16)=16 then print clPercentDiscount else print "-" %></td>
					<td class="cobhl" align="right"><strong><%=yyLiLev%>:</strong></td>
					<td class="cobll" align="left"><%=clLoginLevel%></td>
					</tr>
<%			if loyaltypoints<>"" then %>
					<tr><td class="cobhl" align="right" height="22"><strong><%=xxLoyPoi%>:</strong></td>
					<td class="cobll" colspan="3" align="left"><%=clLoyaltyPoints%></td>
					</tr>
<%			end if
			if trim(extraclientfield1)<>"" then %>
					<tr><td class="cobhl" align="right" height="22"><strong><%=extraclientfield1%>:</strong></td>
					<td class="cobll" colspan="3" align="left"><%=htmlspecials(clientCustom1)%></td>
					</tr>
<%			end if
			if trim(extraclientfield2)<>"" then %>
					<tr><td class="cobhl" align="right" height="22"><strong><%=extraclientfield2%>:</strong></td>
					<td class="cobll" colspan="3" align="left"><%=htmlspecials(clientCustom2)%></td>
					</tr>
<%			end if %>
					<tr><td class="cobhl" align="right" height="22"><strong>Client Admin Notes:</strong></td>
					<td class="cobll" colspan="3" align="left"><%=IIfVr(clientAdminNotes<>"",clientAdminNotes,"-")%></td>
					</tr>
					<tr><td class="cobll" align="left" colspan="4"><br /><ul><li><%=xxChaAcc%> <a class="ectlink" href="javascript:editaccount()"><strong><%=xxClkHere%></strong></a>.</li>
<%			if ordersnotinacct<>0 then print "<li>" & ordersnotinacct & " orders with this email are not registered to the account. To add them all please" & " <a class=""ectlink"" href=""javascript:addorphans()""><strong>"&xxClkHere&"</strong></a>.</li>" %>
					</ul></td>
					</tr>
				  </table>
				</td>
			  </tr>
			  <tr>
                <td class="cobhl" align="center" height="34"><strong><%=xxAddMan%></strong></td>
			  </tr>
			  <tr> 
                <td class="cobll" height="34" align="center">
				  <table class="cobtbl" width="100%" border="0" cellspacing="1" cellpadding="3">
<%			sSQL="SELECT addID,addIsDefault,addName,addLastName,addAddress,addAddress2,addState,addCity,addZip,addPhone,addCountry FROM address WHERE addCustID=" & clID & " ORDER BY addIsDefault"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				do while NOT rs.EOF
					print "<tr><td width=""50%"" class=""cobll"" align=""left"">" & strip_tags2(trim(rs("addName")&" "&rs("addLastName"))) & "<br />" & strip_tags2(rs("addAddress")&"") & IIfVr(trim(rs("addAddress2")&"")<>"", "<br />" & strip_tags2(rs("addAddress2")&""), "") & "<br /> " & strip_tags2(rs("addCity")&"") & ", " & strip_tags2(rs("addState")&"") & IIfVr(rs("addZip")<>"", "<br />" & strip_tags2(rs("addZip")&""), "") & "<br />" & strip_tags2(rs("addCountry")&"") & "</td>"
					print "<td class=""cobhl"" align=""left""><ul><li><a class=""ectlink"" href=""javascript:editaddress("&rs("addID")&")"">" & xxEdAdd & "</a><br /><br /></li><li><a class=""ectlink"" href=""javascript:deleteaddress("&rs("addID")&")"">" & xxDeAdd & "</a></li></ul></td></tr>"
					rs.MoveNext
				loop
			else
				print "<tr><td class=""cobll"" align=""center"" colspan=""2"" height=""34"">" & xxNoAdd & "</td></tr>"
			end if
			rs.close
%>
					<tr><td class="cobhl" colspan="2" align="left"><br /><ul><li><%=xxPCAdd%> <a class="ectlink" href="javascript:newaddress()"><strong><%=xxClkHere%></strong></a>.</li></ul></td></tr>
				  </table>
				</td>
			  </tr>
			  <tr> 
                <td class="cobhl" align="center" height="34"><strong><%=xxOrdMan%></strong></td>
			  </tr>
			  <tr> 
                <td class="cobll" height="34" align="center">
				  <table class="cobtbl" width="100%" border="0" cellspacing="1" cellpadding="3">
<%			hastracknum=FALSE
			sSQL="SELECT ordID FROM orders WHERE ordClientID=" & clID & " AND ordTrackNum<>''"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then hastracknum=TRUE
			rs.close
			sSQL="SELECT ordID FROM orders WHERE ordClientID=0 AND ordEmail='"&escape_string(clEmail)&"'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then hasorphan=TRUE
			rs.close %>
					<tr><td class="cobhl"><%=xxOrdId%></td>
						<td class="cobhl"><%=xxDate%></td>
						<td class="cobhl"><%=xxStatus%></td>
<%			if hastracknum then print "<td class=""cobhl"">"&xxTraNum&"</td>" %>
						<td class="cobhl"><%=xxGndTot%></td>
<%			if hasorphan then print "<td class=""cobhl"">"&"Account"&"</td>" %>
						<td class="cobhl"><%=xxCODets%></td>
					</tr>
<%			grandtotal=0
			sSQL="SELECT ordID,ordDate,ordTrackNum,ordTotal,ordStateTax,ordCountryTax,ordShipping,ordHSTTax,ordHandling,ordDiscount,"&getlangid("statPublic",64)&",ordClientID FROM orders LEFT OUTER JOIN orderstatus ON orders.ordStatus=orderstatus.statID WHERE ordClientID=" & clID & " OR ordEmail='"&escape_string(clEmail)&"' ORDER BY ordDate"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				do while NOT rs.EOF
					subtotal=(rs("ordTotal")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordShipping")+rs("ordHSTTax")+rs("ordHandling"))-rs("ordDiscount")
					grandtotal=grandtotal + subtotal
					print "<tr><td class=""cobll"">" & rs("ordID") & "</td>"
					print "<td class=""cobll"">" & rs("ordDate") & "</td>"
					print "<td class=""cobll"">" & rs(getlangid("statPublic",64)) & "</td>"
					if hastracknum then print "<td class=""cobll"">" & IIfVr(trim(rs("ordTrackNum")&"")<>"",rs("ordTrackNum"),"&nbsp;") & "</td>"
					print "<td class=""cobll"" align=""right"">" & FormatEuroCurrency(subtotal) & "&nbsp;</td>"
					if hasorphan then
						print "<td class=""cobll"">"
						if rs("ordClientID")=0 then print "<a href=""javascript:addorphan("&rs("ordID")&")"">"&"Link to Account"&"</a>" else print "&nbsp;"
						print "</td>"
					end if
					print "<td class=""cobll""><a class=""ectlink"" href=""javascript:vieworder("&rs("ordID")&")"">" & xxClkHere & "</a></td></tr>"
					rs.MoveNext
				loop
				if subtotal<>grandtotal then print "<tr><td class=""cobll"" colspan="""&IIfVr(hastracknum,"4","3")&""">&nbsp;</td><td class=""cobll"" align=""right"">" & FormatEuroCurrency(grandtotal) & "&nbsp;</td><td class=""cobll"""&IIfVr(hasorphan," colspan=""2""","")&">&nbsp;</td></tr>"
			else
				print "<tr><td class=""cobll"" colspan=""5"" height=""34"" align=""center"">" & xxNoOrd & "</td></tr>"
			end if
			rs.close
%>
				  </table>
				</td>
			  </tr>
			</table>
		  </form>
<% elseif getpost("posted")="1" AND success then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
<%		if getpost("act")="doeditaddress" OR getpost("act")="donewaddress" then %>
					<form action="adminclientlog.asp" method="post" id="postform">
					<input type="hidden" name="act" value="viewacct" />
					<input type="hidden" name="id" value="<%=getpost("id")%>" />
					&nbsp;<br />&nbsp;<br />
					<%=yyNoAuto%><br />&nbsp;<br />
					<input type="submit" value="<%=yyClkHer%>" /><br />&nbsp;<br />&nbsp;
					</form>
<%			print "<script>document.getElementById(""postform"").submit();</script>" & vbCrLf
		else %>
					<%=yyNoAuto%> <a href="adminclientlog.asp"><strong><%=yyClkHer%></strong></a>.<br />&nbsp;</br />
<%		end if %>
                </td>
			  </tr>
			</table></td>
        </tr>
      </table>
<%	elseif is_numeric(getget("loginas")) then
		sSQL="SELECT clID,clUserName,clPW,clActions,clLoginLevel,clPercentDiscount FROM customerlogin WHERE clID="&getget("loginas")
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			SESSION("clientID")=rs("clID")
			SESSION("clientPW")=rs("clPW")
			SESSION("clientUser")=rs("clUserName")
			SESSION("clientActions")=rs("clActions")
			SESSION("clientLoginLevel")=rs("clLoginLevel")
			SESSION("clientPercentDiscount")=(100.0-cdbl(rs("clPercentDiscount")))/100.0
			redirecturl=storeurl
			if request.servervariables("HTTPS")="on" then redirecturl=replace(redirecturl,"http:","https:")
			response.redirect redirecturl & "cart.asp"
		else
			print "Login not found"
		end if
		rs.close
	elseif getpost("posted")="1" then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><span style="color:#FF0000;font-weight:bold"><%=yyOpFai%></span><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a><br />&nbsp;<br />&nbsp;</td>
			  </tr>
			</table></td>
        </tr>
      </table>
<% else
clact=request.cookies("clact")
%>
<script src="popcalendar.js"></script>
<script>
<!--
try{languagetext('<%=adminlang%>');}catch(err){}
function mrec(id){
	document.mainform.id.value=id;
	document.mainform.act.value="viewacct";
	document.mainform.submit();
}
function newrec(id){
	document.mainform.id.value=id;
	document.mainform.act.value="addnew";
	document.mainform.submit();
}
function lrec(id){
	window.open('adminclientlog.asp?loginas='+id,'clientlogin','menubar=no, scrollbars=yes, width=950, height=700, directories=no,location=no,resizable=yes,status=yes,toolbar=no')
}
function drec(id){
if(confirm("<%=jscheck(yyConDel)%>\n")){
	document.mainform.id.value=id;
	document.mainform.act.value="delete";
	document.mainform.submit();
}
}
function startsearch(){
	document.mainform.action="adminclientlog.asp";
	document.mainform.act.value="search";
	document.mainform.posted.value="";
	document.mainform.submit();
}
function quickupdate(){
	if(document.mainform.clact.value=="del"){
		if(!confirm("<%=jscheck(yyConDel)%>\n"))
			return;
	}
	document.mainform.action="adminclientlog.asp";
	document.mainform.act.value="quickupdate";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function setCookie(c_name,value,expiredays){
	var exdate=new Date();
	exdate.setDate(exdate.getDate()+expiredays);
	document.cookie=c_name+ "=" +escape(value)+((expiredays==null) ? "" : ";expires="+exdate.toGMTString());
}
function changeclact(obj){
	setCookie('clact',obj[obj.selectedIndex].value,600);
	startsearch();
}
function chi(x,pid){
	x.name='pra_'+patch_pid(pid);
}
function setsxz(){
	maxitems=document.getElementById("resultcounter").value;
	amnt=document.getElementById("txtset").selectedIndex;
	for(index=0;index<maxitems;index++){
		if(document.getElementById("sxz"+index)){
			document.getElementById("sxz"+index).selectedIndex=amnt;
			document.getElementById("sxz"+index).onchange();
		}
	}
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
function checkboxes(docheck){
	maxitems=document.getElementById("resultcounter").value;
	for(index=0;index<maxitems;index++){
		var thiselem=document.getElementById("chkbx"+index);
		if(thiselem.checked!=docheck){
			thiselem.checked=docheck;
			thiselem.onchange();
		}
	}
}
// -->
</script>
<h2><%=yyAdmCli%></h2>
<%	function dispcountries(ind)
		dispcountries="<select size=""1"" id=""newcntryname"&ind&""" name=""newcntryname"&ind&"""><option value="""">"&yySelect&"</option>"
		for rowcounter=0 to UBOUND(allcountries,2)
			dispcountries=dispcountries&"<option value=""" & htmlspecials(allcountries(0,rowcounter)) & """>"&allcountries(0,rowcounter)&"</option>"&vbCrLf
		next
		dispcountries=dispcountries&"<option value="""" disabled=""disabled"">==============================</option>"
		dispcountries=dispcountries&"<option value=""xxxdeletexxx"">"&"DELETE ADDRESS - Country No Longer Supported"&"</option>"
		dispcountries=dispcountries&"</select>"
	end function
	sSQL="SELECT DISTINCT addCountry,countryID FROM address LEFT JOIN countries ON address.addCountry=countries.countryName WHERE countryID IS NULL"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then %>
		<form id="updcntryid" method="post" action="adminclientlog.asp">
		<input type="hidden" id="oldcntryname" name="oldcntryname" value="" />
		<input type="hidden" id="newcntryname" name="newcntryname" value="" />
		</form>
		<table border="1" cellspacing="3" cellpadding="3">
		  <tr><td colspan="3">There are countries in the client login table that do not now exist. These need to be mapped to actual countries.</td></tr>
<%		sSQL="SELECT countryName,countryID FROM countries WHERE countryEnabled=1 ORDER BY countryOrder DESC,countryName"
		rs2.Open sSQL,cnn,0,1
		if NOT rs2.EOF then allcountries=rs2.getrows
		rs2.Close
		index=0
		do while NOT rs.EOF
			print "<tr><td>" & rs("addCountry") & "</td><td>" & dispcountries(index) & "</td><td>" %>
<input type="button" value="<%=yySubmit%>" onclick="document.getElementById('oldcntryname').value='<%=jsspecials(rs("addCountry"))%>';document.getElementById('newcntryname').value=document.getElementById('newcntryname<%=index%>')[document.getElementById('newcntryname<%=index%>').selectedIndex].value;if(document.getElementById('newcntryname<%=index%>').selectedIndex==0)alert('Please select a country...');else document.getElementById('updcntryid').submit()" />
<%			print "</td></tr>"
			index=index+1
			rs.movenext
		loop %>
		</table>
<%	end if
	rs.close %>
	<form name="mainform" method="post" action="adminclientlog.asp">
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="xxxxx" />
			<input type="hidden" name="id" value="xxxxx" />
			<input type="hidden" name="pg" value="<%=IIfVr(getpost("act")="search", "1", getget("pg"))%>" />
<%		themask=cStr(DateSerial(2003,12,11))
		themask=replace(themask,"2003","yyyy")
		themask=replace(themask,"12","mm")
		themask=replace(themask,"11","dd")
		thelevel=request("slevel")
		if thelevel<>"" then thelevel=int(thelevel)
%>				<table class="cobtbl" width="100%" border="0" cellspacing="1" cellpadding="3">
				  <tr> 
	                <td class="cobhl" width="25%" align="right"><%=yySrchFr%>:</td>
					<td class="cobll" width="25%"><input type="text" name="stext" size="20" value="<%=htmlspecials(request("stext"))%>" /></td>
					<td class="cobhl" width="20%" align="right"><%=yyDate%>:</td>
					<td class="cobll"><div style="position:relative;display:inline">
					<select name="daterange" size="1" style="vertical-align:middle">
					<option value=""><%=yySinc%></option>
					<option value="1"<%if request("daterange")="1" then print " selected=""selected"""%>><%=yyTill%></option>
					<option value="2"<%if request("daterange")="2" then print " selected=""selected"""%>><%=yyOn%></option>
					</select>
					<input type="text" size="14" name="accdate" value="<%=htmlspecials(request("accdate"))%>" style="vertical-align:middle" /> <input type="button" onclick="popUpCalendar(this, document.forms.mainform.accdate, '<%=themask%>', -205)" value="DP" />
					</div></td>
				  </tr>
				  <tr>
				    <td class="cobhl" align="right"><%=yySrchTp%>:</td>
					<td class="cobll"><select name="stype" size="1">
						<option value=""><%=yySrchAl%></option>
						<option value="any" <% if request("stype")="any" then print "selected=""selected"""%>><%=yySrchAn%></option>
						<option value="exact" <% if request("stype")="exact" then print "selected=""selected"""%>><%=yySrchEx%></option>
						</select>
					</td>
					<td class="cobhl" align="right"><%=yyLiLev%>:</td>
					<td class="cobll">
					  <select name="slevel" size="1">
					  <option value=""><%=yyAllLev%></option>
						<%	for rowcounter=0 to maxloginlevels
								print "<option value='"&rowcounter&"'"
								if thelevel<>"" then
									if thelevel=rowcounter then print " selected=""selected"""
								end if
								print ">&nbsp; "&rowcounter&" </option>"
							next %>
					  </select>
					</td>
	              </tr>
				  <tr>
				    <td class="cobhl" align="center">
<%				if getpost("act")="search" OR getget("pg")<>"" then
					if clact="ste" OR clact="cte" OR clact="she" OR clact="wsp" OR clact="ped" OR clact="hae" OR clact="del" then %>
					<div style="margin-top:2px"><input type="button" value="<%=yyCheckA%>" onclick="checkboxes(true);" /> <input type="button" value="<%=yyUCheck%>" onclick="checkboxes(false);" /></div>
<%					elseif clact="dsc" then %>
					<div style="margin-top:2px"><input type="text" name="txtadd" id="txtadd" size="5" value="" style="vertical-align:middle" /> <input type="button" value="Set" onclick="setto()" /></div>
<%					elseif clact="lol" then %>
					<div style="margin-top:2px"><select size="1" id="txtset" style="vertical-align:middle">
						<option value="0">0</option>
						<option value="1">1</option>
						<option value="2">2</option>
						<option value="3">3</option>
						<option value="4">4</option>
						<option value="5">5</option>
						</select> <input type="button" value="Set" onclick="setsxz()" /></div>
<%					end if
				end if %>
					</td>
				    <td class="cobll" colspan="3"><table width="100%" cellspacing="0" cellpadding="0" border="0">
					    <tr>
						  <td class="cobll" align="center"><input type="button" value="<%=yyListRe%>" onclick="startsearch();" />
							<input type="button" value="<%=yyCLNew%>" onclick="newrec();" />
						  </td>
						  <td class="cobll" height="26" width="20%" align="right">&nbsp;</td>
						</tr>
					  </table></td>
				  </tr>
				</table>
		<table width="100%" class="stackable admin-table-a sta-white">
<%	if getpost("act")="search" OR getget("pg")<>"" then
		jscript="" : qetype="" : qesize=""
		sub displayprodrow(xrs)
			jscript=jscript&"pa["&resultcounter&"]=["
			%><tr class="<%=bgcolor%>" id="tr<%=resultcounter%>"><td class="minicell"><%
				qetype="text"
				qesize="18"
				if clact="lol" then
					qetype="special"
					jscript=jscript&"''"
					currpos=xrs("clLoginLevel")
					print "<select id=""sxz"&resultcounter&""" onchange=""chi(this,"&resultcounter&")"">"
					for index=0 to 5
						print "<option value="""&index&""""&IIfVs(index=xrs("clLoginLevel")," selected=""selected""")&">"&index&"</option>"
					next
					print "</select>"
				elseif clact="dsc" then
					jscript=jscript&"'"&jsspecials(xrs("clPercentDiscount"))&"'"
					qesize="5"
				elseif clact="ste" OR clact="cte" OR clact="she" OR clact="wsp" OR clact="ped" OR clact="hae" then
					fieldnum=1
					if clact="cte" then fieldnum=2
					if clact="she" then fieldnum=4
					if clact="wsp" then fieldnum=8
					if clact="ped" then fieldnum=16
					if clact="hae" then fieldnum=32
					jscript=jscript&IIfVr((xrs("clActions") AND fieldnum)<>0,1,0)
					qetype="checkbox"
				elseif clact="del" then
					jscript=jscript&"'del'"
					qetype="delbox"
				else
					qetype=""
				end if %></td><td><%=htmlspecials(xrs("clUserName")&"")%></td><td><%=htmlspecials(xrs("clEmail")&"")%></td>
			<td><%	if (xrs("clActions") AND 1)=1 then print "STE "
					if (xrs("clActions") AND 2)=2 then print "CTE "
					if (xrs("clActions") AND 4)=4 then print "SHE "
					if (xrs("clActions") AND 32)=32 then print "HAE "
					if (xrs("clActions") AND 8)=8 then print "WSP "
					if (xrs("clActions") AND 16)=16 then print "PED "
			%>&nbsp;</td>
			<td class="minicell"><input type="button" value="<%=yyLogin%>" onclick="lrec('<%=xrs("clID")%>',event)" /></td>
			<td class="minicell"><input type="button" value="<%=yyModify%>" onclick="mrec('<%=xrs("clID")%>',event)" /></td>
			<td class="minicell"><input type="button" value="<%=yyDelete%>" onclick="drec('<%=xrs("clID")%>')" /></td></tr>
<%			jscript=jscript&",'"&jsspecials(xrs("clID"))&"'];"&vbCrLf
			resultcounter=resultcounter+1
		end sub
		sub displayheaderrow() %>
			<tr>
				<th class="small minicell">
					<select name="clact" id="clact" size="1" onchange="changeclact(this)" style="width:150px">
				<option value="none">Quick Entry...</option>
				<option value="lol"<% if clact="lol" then print " selected=""selected"""%>><%=yyLiLev%></option>
				<option value="dsc"<% if clact="dsc" then print " selected=""selected"""%>><%=yyDscAmt%></option>
				<option value="" disabled="disabled">---------------------</option>
				<option value="ste"<% if clact="ste" then print " selected=""selected"""%>><%=yyExStat%></option>
				<option value="cte"<% if clact="cte" then print " selected=""selected"""%>><%=yyExCoun%></option>
				<option value="she"<% if clact="she" then print " selected=""selected"""%>><%=yyExShip%></option>
				<option value="hae"<% if clact="hae" then print " selected=""selected"""%>><%=yyExHand%></option>
				<option value="wsp"<% if clact="wsp" then print " selected=""selected"""%>><%=yyWholPr%></option>
				<option value="ped"<% if clact="ped" then print " selected=""selected"""%>><%=yyPerDis%></option>
				<option value="" disabled="disabled">---------------------</option>
				<option value="del"<% if clact="del" then print " selected=""selected"""%>><%=yyDelete%></option>
					</select>
				</th>
				<th class="maincell"><%=yyLiName%></th>
				<th class="maincell"><%=yyEmail%></th>
				<th class="minicell"><%=yyActns%></th>
				<th class="minicell"><%=yyLogin%></th>
				<th class="minicell"><%=yyModify%></th>
				<th class="minicell"><%=yyDelete%></th>
			</tr>
<%		end sub
		whereand=" WHERE "
		sSQL="SELECT clID,clUserName,clActions,clLoginLevel,clPercentDiscount,clEmail,clPW FROM customerlogin "
		if thelevel<>"" then
			sSQL=sSQL & whereand & " clLoginLevel=" & thelevel
			whereand=" AND "
		end if
		accdate=trim(request("accdate"))
		if accdate<>"" then
			on error resume next
			accdate=DateValue(accdate)
			if err.number <> 0 then
				accdate=""
			end if
			on error goto 0
			if accdate<>"" then
				if request("daterange")="1" then ' Till
					sSQL=sSQL & whereand & "clDateCreated <= " & vsusdate(accdate) & " "
				elseif request("daterange")="2" then ' On
					sSQL=sSQL & whereand & "clDateCreated BETWEEN " & vsusdate(accdate) & " AND " & vsusdate(accdate+1) & " "
				else ' Since
					sSQL=sSQL & whereand & "clDateCreated >= " & vsusdate(accdate) & " "
				end if
				whereand=" AND "
			end if
		end if
		if trim(request("stext"))<>"" then
			sText=escape_string(request("stext"))
			aText=Split(sText)
			aFields(0)="clUserName"
			aFields(1)="clPw"
			aFields(2)="clEmail"
			if request("stype")="exact" then
				sSQL=sSQL & whereand & "(clUserName LIKE '%"&sText&"%' OR clPw LIKE '%"&sText&"%' OR clEmail LIKE '%"&sText&"%') "
				whereand=" AND "
			else
				sJoin="AND "
				if request("stype")="any" then sJoin="OR "
				sSQL=sSQL & whereand&"("
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
		end if
		sSQL=sSQL & " ORDER BY clUserName"
		if adminclientloginperpage="" then adminclientloginperpage=200
		rs.CursorLocation=3 ' adUseClient
		rs.CacheSize=adminclientloginperpage
		rs.open sSQL, cnn
		if rs.eof or rs.bof then
			success=false
			iNumOfPages=0
		else
			success=true
			rs.MoveFirst
			rs.PageSize=adminclientloginperpage
			CurPage=1
			if is_numeric(getget("pg")) then CurPage=int(getget("pg"))
			iNumOfPages=Int((rs.RecordCount + (adminclientloginperpage-1)) / adminclientloginperpage)
			rs.AbsolutePage=CurPage
		end if
		Count=0
		haveerrprods=FALSE
		if NOT rs.EOF then
			pblink="<a href=""adminclientlog.asp?rid="&request("rid")&"&stext="&urlencode(request("stext"))&"&stype="&request("stype")&"&slevel="&request("slevel")&"&accdate="&urlencode(request("accdate"))&"&daterange="&request("daterange")&"&pg="
			if iNumOfPages > 1 then print "<tr><td colspan=""6"" align=""center"">" & writepagebar(CurPage,iNumOfPages,yyPrev,yyNext,pblink,FALSE) & "</td></tr>"
			displayheaderrow()
			do while NOT rs.EOF AND Count < rs.PageSize
				if bgcolor="altdark" then bgcolor="altlight" else bgcolor="altdark"
				displayprodrow(rs)
				rs.MoveNext
				Count=Count + 1
			loop
			if haveerrprods then print "<tr><td width=""100%"" colspan=""6""><br />"&redasterix&yySeePr&"</td></tr>"
			if iNumOfPages > 1 then print "<tr><td colspan=""6"" align=""center"">" & writepagebar(CurPage,iNumOfPages,yyPrev,yyNext,pblink,FALSE) & "</td></tr>"
		else
			print "<tr><td width=""100%"" colspan=""6"" align=""center""><br />"&yyItNone&"<br />&nbsp;</td></tr>"
		end if
		rs.close
	else
		numitems=0
		sSQL="SELECT COUNT(*) as totcount FROM customerlogin"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			numitems=rs("totcount")
		end if
		rs.close
		print "<tr><td colspan=""7""><div class=""itemsdefine"">You have " & numitems & " customer accounts defined.</div></td></tr>"
	end if %>
			  <tr>
				<td align="center" style="white-space:nowrap"><% if resultcounter>0 AND clact<>"" AND clact<>"none" then print "<input type=""hidden"" name=""resultcounter"" id=""resultcounter"" value="""&resultcounter&""" /><input type=""button"" value=""" & yyUpdate & """ onclick=""quickupdate()"" /> <input type=""reset"" value=""" & yyReset & """ />" else print "&nbsp;"%></td>
                <td width="100%" colspan="6" align="center"><br /><ul><li><%=yyCLTyp%></li></ul>
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
			  </tr>
            </table>
		  </td>
        </tr>
      </table>
	</form>
<script>
/* <![CDATA[ */
var pa=[];
<%=jscript%>
function patch_pid(pid){
	document.getElementById('pid'+pid).name='pid'+pid;
	document.getElementById('pid'+pid).value=pa[pid][1];
	return pid;
}
for(var pidind in pa){
	var ttr=document.getElementById('tr'+pidind);
	ttr.cells[1].innerHTML+='<input type="hidden" id="pid'+pidind+'" value="" />';
<%		if qetype="text" then %>
	ttr.cells[0].innerHTML=pa[pidind][0]===false?'-':'<input type="text" id="chkbx'+pidind+'" size="<% print qesize%>" onchange="this.name=\'pra_'+patch_pid(pidind)+'\'" value="'+pa[pidind][0].replace('"','&quot;')+'" tabindex="'+(pidind+1)+'" />';
<%		elseif qetype="delbox" then %>
	ttr.cells[0].innerHTML='<input type="checkbox" id="chkbx'+pidind+'" onchange="this.name=\'pra_'+patch_pid(pidind)+'\'" value="del" tabindex="'+(pidind+1)+'" />';
<%		elseif qetype="checkbox" then %>
	ttr.cells[0].innerHTML='<input type="hidden" id="pra_'+pa[pidind][1]+'" value="1" /><input type="checkbox" id="chkbx'+pidind+'" onchange="this.name=\'prb_'+patch_pid(pidind)+'\';document.getElementById(\'pra_'+pa[pidind][1]+'\').name=\'pra_'+patch_pid(pidind)+'\'" value="1" '+(pa[pidind][0]==1?'checked="checked" ':'')+'tabindex="'+(pidind+1)+'" />';
<%		end if %>
}
/* ]]> */
</script>
<% end if
cnn.Close
set rs=nothing
set cnn=nothing
%>