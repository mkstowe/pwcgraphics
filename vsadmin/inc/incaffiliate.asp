<!--#include file="md5.asp"-->
<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if sDSN="" then response.write "Database connection not set" : response.end
Dim sSQL,rs,cnn,success,showaccount,addsuccess
if request.totalbytes > 10000 then response.end
addsuccess = true
success = true
showaccount = true

if forceloginonhttps AND request.servervariables("HTTPS")="off" AND instr(storeurlssl,"https")>0 then response.redirect storeurlssl & "affiliate" & extension & IIfVr(request.servervariables("QUERY_STRING")<>"","?"&request.servervariables("QUERY_STRING"),"") : response.end
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
Set toregexp = new RegExp
toregexp.pattern = "\W"
toregexp.ignorecase=TRUE
toregexp.global=TRUE
theaffilid = toregexp.replace(getpost("affilid")&"","")
Set toregexp = Nothing
if getpost("editaction")<>"" then
	if theaffilid="" then
		addsuccess = FALSE
	elseif getpost("editaction")="modify" then
		sSQL="UPDATE affiliates SET "
		if getpost("affilpw")<>"" then sSQL=sSQL & "affilPW='" & escape_string(dohashpw(getpost("affilpw"))) & "',"
		sSQL=sSQL & "affilEmail='" & escape_string(getpost("email")) & "'," & _
			"affilName='" & escape_string(getpost("name")) & "'," & _
			"affilAddress='" & escape_string(getpost("address")) & "'," & _
			"affilCity='" & escape_string(getpost("city")) & "'," & _
			"affilState='" & escape_string(getpost("state")) & "'," & _
			"affilCountry='" & escape_string(getpost("country")) & "'," & _
			"affilZip='" & escape_string(getpost("zip")) & "',"
		sSQL=sSQL & "affilInform=" & IIfVr(getpost("inform")="ON",1,0) & " WHERE affilID='" & escape_string(theaffilid) & "'"
		err.number=0
		on error resume next
		cnn.execute(sSQL)
		if err.number<>0 then
			addsuccess=FALSE
			xxAffUse="There was a problem updating your affiliate details. Please try again."
		end if
	elseif getpost("editaction")="new" then
		addsuccess=TRUE
		if recaptchaenabled(4) then addsuccess=checkrecaptcha(xxAffUse)
		if addsuccess then
			sSQL="SELECT affilID FROM affiliates WHERE affilID='" & escape_string(theaffilid) & "'"
			rs.open sSQL,cnn,0,1
			addsuccess=rs.EOF
			rs.close
		end if
		if addsuccess then
			sSQL="INSERT INTO affiliates (affilID,affilPW,affilEmail,affilName,affilAddress,affilCity,affilState,affilCountry,affilZip,affilCommision,affilDate,affilInform) VALUES (" & _
				"'" & escape_string(theaffilid) & "'," & _
				"'" & escape_string(dohashpw(getpost("affilpw"))) & "'," & _
				"'" & escape_string(getpost("email")) & "'," & _
				"'" & escape_string(getpost("name")) & "'," & _
				"'" & escape_string(getpost("address")) & "'," & _
				"'" & escape_string(getpost("city")) & "'," & _
				"'" & escape_string(getpost("state")) & "'," & _
				"'" & escape_string(getpost("country")) & "'," & _
				"'" & escape_string(getpost("zip")) & "',"
			if defaultcommission<>"" then
				sSQL=sSQL & defaultcommission & ","
				SESSION("affilCommision") = cdbl(defaultcommission)
			else
				sSQL=sSQL & "0,"
				SESSION("affilCommision") = 0
			end if
			sSQL=sSQL & vsusdate(date()) & "," & IIfVr(getpost("inform")="ON",1,0) & ") "
			err.number=0
			on error resume next
			cnn.execute(sSQL)
			if err.number<>0 then
				addsuccess=FALSE
				xxAffUse="There was a problem entering your affiliate details. Please try again."
			end if
			if addsuccess then
				if (adminEmailConfirm AND 2)=2 then
					emailmessage="There has been a new affiliate signup at your store: " & theaffilid & emlNl & _
						"Email: " & getpost("Email") & emlNl & _
						"Name: " & getpost("Name") & emlNl & _
						"Address: " & getpost("Address") & emlNl & _
						"City: " & getpost("City") & emlNl & _
						"State: " & getpost("State") & emlNl & _
						"Country: " & getpost("Country") & emlNl & _
						"Zip: " & getpost("Zip") & emlNl
					call dosendemaileo(emailAddr,emailAddr,getpost("Email"),"New Affiliate Signup",emailmessage,emailObject,themailhost,theuser,thepass)
				end if
				print "<meta http-equiv=""Refresh"" content=""0"">"
			end if
		end if
	end if
	if addsuccess then
		SESSION("affilid") = theaffilid
		if getpost("affilpw")<>"" then SESSION("affilpw") = replace(dohashpw(getpost("affilpw")),"'","")
		SESSION("affilName") = getpost("Name")
	end if
elseif getpost("act")="affillogin" then
	sSQL = "SELECT affilID,affilName,affilCommision,affilPW FROM affiliates WHERE affilID='"&escape_string(theaffilid)&"' AND affilPW='"&replace(dohashpw(getpost("affilpw")),"'","")&"'"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		SESSION("affilid")=theaffilid
		SESSION("affilpw")=rs("affilPW")
		SESSION("affilName") = rs("affilName")
		SESSION("affilCommision") = cdbl(rs("affilCommision"))
		showaccount=false
	else
		success=false
	end if
	rs.close
	if success then
		print "<meta http-equiv=""Refresh"" content=""3"">"
%>
			<form method="post" action="affiliate<%=extension%>">
			  <div class="ectdiv">
				<div class="ectdivhead"><%=xxAffPrg & " " & xxWelcom & " " & htmlspecials(SESSION("affilName"))%>.</div>
				<div class="ectmessagescreen">
					<div><%=xxAffLog%></div>
					<div><%=xxForAut%> <a class="ectlink" href="affiliate<%=extension%>"><%=xxClkHere%></a>.</div>
				</div>
			  </div>
			</form>
<%
	end if
elseif getpost("act")="logout" then
	SESSION("affilid") = ""
	SESSION("affilpw") = ""
	SESSION("affilName") = ""
end if
if getpost("act")="newaffil" OR (getpost("act")="editaffil" AND trim(SESSION("affilid"))<>"") OR NOT addsuccess then
	showaccount=false %>
<script>
<!--
function checkform(frm){
if(frm.affilid.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&xxAffID)%>\".");
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
	alert("<%=jscheck(xxAlphaNu&" """&xxAffID)%>\".");
	frm.affilid.focus();
	return (false);
}
<%	if getpost("act")<>"editaffil" then %>
if(frm.affilpw.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&xxPwd)%>\".");
	frm.affilpw.focus();
	return (false);
}
<%	end if %>
if(frm.name.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&xxName)%>\".");
	frm.name.focus();
	return (false);
}
if(frm.email.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&xxEmail)%>\".");
	frm.email.focus();
	return (false);
}
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
if(frm.state.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&xxAllSta)%>\".");
	frm.state.focus();
	return (false);
}
if(frm.zip.value==""){
	alert("<%=jscheck(xxPlsEntr&" """&xxZip)%>\".");
	frm.zip.focus();
	return (false);
}
<%	if recaptchaenabled(4) then print "if(!affilcaptchaok){ alert(""" & jscheck(xxRecapt) & """);return(false); }" %>
return (true);
}
//-->
</script>
<%	if NOT addsuccess then
		affilName = getpost("Name")
		affilPW = getpost("affilPW")
		affilID = getpost("affilID")
		affilAddress = getpost("Address")
		affilCity = getpost("City")
		affilState = getpost("State")
		affilZip = getpost("Zip")
		affilCountry = getpost("Country")
		affilEmail = getpost("Email")
		affilInform = getpost("Inform")="ON"
	elseif (getpost("act")="editaffil" AND trim(SESSION("affilid"))<>"") then
		sSQL = "SELECT affilName,affilPW,affilAddress,affilCity,affilState,affilZip,affilCountry,affilEmail,affilInform FROM affiliates WHERE affilID='"&replace(trim(SESSION("affilid")),"'","")&"' AND affilPW='"&replace(trim(SESSION("affilpw")),"'","")&"'"
		rs.open sSQL,cnn,1,3,&H0001
		if NOT rs.EOF then
			affilName = rs("affilName")
			affilPW = ""
			affilAddress = rs("affilAddress")
			affilCity = rs("affilCity")
			affilState = rs("affilState")
			affilZip = rs("affilZip")
			affilCountry = rs("affilCountry")
			affilEmail = rs("affilEmail")
			affilInform = Int(rs("affilInform"))=1
		end if
		rs.close
	end if
%>			<form method="post" action="<%if forceloginonhttps then print storeurlssl%>affiliate<%=extension%>" onsubmit="return checkform(this)">
			  <div class="ectdiv ectaffiliate">
				<div class="ectdivhead"><%=xxAffDts%></div>
<%	if NOT addsuccess then %>
				<div class="ectdiv2column ectwarning"><%=xxAffUse%></div>
<%	end if %>
				<div class="ectdivcontainer">
				  <div class="ectdivleft"><%=redasterix&xxAffID%></div>
				  <div class="ectdivright"><%
					if (getpost("act")="editaffil" AND trim(SESSION("affilid"))<>"") then
						print htmlspecials(trim(SESSION("affilid")))
						%><input type="hidden" name="affilid" value="<%=htmlspecials(trim(SESSION("affilid")))%>" />
						  <input type="hidden" name="editaction" value="modify" /><%
					else
						%><input type="text" name="affilid" size="30" value="<%=htmlspecials(affilid)%>" />
						  <input type="hidden" name="editaction" value="new" /><%
					end if %></div>
				</div>
				<div class="ectdivcontainer">
				  <div class="ectdivleft"><%=IIfVr(getpost("act")="editaffil",xxReset&" "&xxPwd,redasterix&xxPwd)%></div>
				  <div class="ectdivright"><input type="password" name="affilpw" size="20" value="" autocomplete="off" /></div>
				</div>
				<div class="ectdivcontainer">
				  <div class="ectdivleft"><%=redasterix&xxName%></div>
				  <div class="ectdivright"><input type="text" name="name" size="20" value="<%=htmlspecials(affilName)%>" /></div>
				</div>
				<div class="ectdivcontainer">
				  <div class="ectdivleft"><%=redasterix&xxEmail%></div>
				  <div class="ectdivright"><input type="text" name="email" size="25" value="<%=htmlspecials(affilEmail)%>" /></div>
				</div>
				<div class="ectdivcontainer">
				  <div class="ectdivleft"><%=redasterix&xxAddress%></div>
				  <div class="ectdivright"><input type="text" name="address" size="20" value="<%=htmlspecials(affilAddress)%>" /></div>
				</div>
				<div class="ectdivcontainer">
				  <div class="ectdivleft"><%=redasterix&xxCity%></div>
				  <div class="ectdivright"><input type="text" name="city" size="20" value="<%=htmlspecials(affilCity)%>" /></div>
				</div>
				<div class="ectdivcontainer">
				  <div class="ectdivleft"><%=redasterix&xxAllSta%></div>
				  <div class="ectdivright"><input type="text" name="state" size="20" value="<%=htmlspecials(affilState)%>" /></div>
				</div>
				<div class="ectdivcontainer">
				  <div class="ectdivleft"><%=redasterix&xxCountry%></div>
				  <div class="ectdivright"><select name="country" size="1"><%
sub show_countries(tcountry)
	if NOT IsArray(allcountries) then
		sSQL = "SELECT countryName,countryOrder,"&getlangid("countryName",8)&" AS cnameshow FROM countries ORDER BY countryOrder DESC,"&getlangid("countryName",8)
		rs.open sSQL,cnn,0,1
		allcountries=rs.getrows
		rs.close
	end if
	for rowcounter=0 to UBOUND(allcountries,2)
		print "<option value='" & htmlspecials(allcountries(0,rowcounter)) & "'"
		if tcountry=allcountries(0,rowcounter) then
			print " selected"
		end if
		print ">"&allcountries(2,rowcounter)&"</option>"&vbCrLf
	next
end sub
show_countries(affilCountry)
%></select>
				  </div>
				</div>
				<div class="ectdivcontainer">
				  <div class="ectdivleft"><%=redasterix&xxZip%></div>
				  <div class="ectdivright"><input type="text" name="zip" size="10" value="<%=htmlspecials(affilZip)%>" /></div>
				</div>
				<div class="ectdivcontainer">
				  <div class="ectdivleft"><%=xxInfMe%></div>
				  <div class="ectdivright"><input type="checkbox" name="inform" value="ON" <% if affilInform then print "checked"%> /></div>
				</div>
<%		if recaptchaenabled(4) then %>
				<div class="ectdivcontainer">
					<div class="ectdivleft">&nbsp;</div>
					<% call displayrecaptchajs("affilcaptcha",TRUE,FALSE) %>
					<div id="affilcaptcha" class="g-recaptcha ectdivright"></div>
				</div>
<%		end if %>
				<div class="ectdiv2column">
					<ul><li><span style="font-size:10px"><%=xxInform%></span></li></ul>
				</div>
				<div class="ectdiv2column"><%
					print imageorsubmit(imgsubmit,xxSubmt,"submit")
					if getpost("act")="editaffil" AND trim(SESSION("affilid"))<>"" then
						print "<br /><br />" & imageorbutton(imgbackacct,xxBack,"backacct","history.go(-1)",TRUE)
					end if %></div>
			  </div>
			</form>
<%
end if
if showaccount then
	if SESSION("affilid")="" then
%>		<form method="post" name="mainform" action="<%if forceloginonhttps then print storeurlssl%>affiliate<%=extension%>">
		<input type="hidden" name="act" id="act" value="xxx" />
			<div class="ectdiv ectaffiliate">
				<div class="ectdivhead"><%=xxAffPrg%></div>
<%		if NOT success then %>
				<div class="ectdiv2column ectwarning"><%=xxAffNo%></div>
<%		end if %>
				<div class="ectdivcontainer">
				  <div class="ectdivleft"><%=xxAffID%></div>
				  <div class="ectdivright"><input type="text" name="affilid" size="30" value="<%=htmlspecials(getpost("affilid"))%>" /></div>
				</div>
				<div class="ectdivcontainer">
				  <div class="ectdivleft"><%=xxPwd%></div>
				  <div class="ectdivright"><input type="password" name="affilpw" size="30" value="<%=htmlspecials(getpost("affilpw"))%>" autocomplete="off" /></div>
				</div>
				<div class="ectdivcontainer">
				  <div class="ectdivleft">&nbsp;</div>
				  <div class="ectdivright"><%=imageorbutton(imgnewaffiliate,xxNewAct,"newaffiliate","document.getElementById('act').value='newaffil';document.forms.mainform.submit();",TRUE) & " " & imageorsubmit(imgaffiliatelogin,xxAffLI&""" onclick=""document.getElementById('act').value='affillogin'","affiliatelogin")%></div>
				</div>
			</div>
		</form>
<%	else
		totalDay=0.0
		totalYesterday=0.0
		totalMonth=0.0
		totalLastMonth=0.0
		tdt = Date()
		tdt2 = Date()+1
		sSQL = "SELECT Sum(ordTotal-ordDiscount) as theCount FROM orders WHERE ordStatus>=3 AND ordAffiliate='"&trim(replace(SESSION("affilid"),"'",""))&"' AND ordDate BETWEEN " & vsusdate(tdt)&" AND " & vsusdate(tdt2)
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then totalDay = rs("theCount")
		rs.close
		tdt = Date()-1
		tdt2 = Date()
		sSQL = "SELECT Sum(ordTotal-ordDiscount) as theCount FROM orders WHERE ordStatus>=3 AND ordAffiliate='"&trim(replace(SESSION("affilid"),"'",""))&"' AND ordDate BETWEEN " & vsusdate(tdt)&" AND " & vsusdate(tdt2)
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then totalYesterday = rs("theCount")
		rs.close
		tdt = DateSerial(year(Date()),Month(Date()),1)
		tdt2 = Date()+1
		sSQL = "SELECT Sum(ordTotal-ordDiscount) as theCount FROM orders WHERE ordStatus>=3 AND ordAffiliate='"&trim(replace(SESSION("affilid"),"'",""))&"' AND ordDate BETWEEN " & vsusdate(tdt)&" AND " & vsusdate(tdt2)
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then totalMonth = rs("theCount")
		rs.close
		tdt = DateSerial(year(Date()),Month(Date())-1,1)
		tdt2 = DateSerial(year(Date()),Month(Date()),1)
		sSQL = "SELECT Sum(ordTotal-ordDiscount) as theCount FROM orders WHERE ordStatus>=3 AND ordAffiliate='"&trim(replace(SESSION("affilid"),"'",""))&"' AND ordDate BETWEEN " & vsusdate(tdt)&" AND " & vsusdate(tdt2)
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then totalLastMonth = rs("theCount")
		rs.close
		if IsNull(totalDay) then totalDay=0.0
		if IsNull(totalYesterday) then totalYesterday=0.0
		if IsNull(totalMonth) then totalMonth=0.0
		if IsNull(totalLastMonth) then totalLastMonth=0.0
%>		<form method="post" name="mainform" action="affiliate<%=extension%>">
		<input type="hidden" name="act" value="" />
			<div class="ectdiv ectaffiliate">
				<div class="ectdivhead"><%=xxAffPrg & " " & xxWelcom & " " & htmlspecials(SESSION("affilName"))%>.</div>
				<div class="ectdivcontainer">
				  <div class="ectdivleft"><%=xxTotTod%></div>
				  <div class="ectdivright"><% print FormatEuroCurrency(totalDay)
				  if SESSION("affilCommision")<>0 then print " = " & FormatEuroCurrency((totalDay * SESSION("affilCommision")) / 100.0) & " " & xxCommis
				  %></div>
				</div>
				<div class="ectdivcontainer">
				  <div class="ectdivleft"><%=xxTotYes%></div>
				  <div class="ectdivright"><% print FormatEuroCurrency(totalYesterday)
				  if SESSION("affilCommision")<>0 then print " = " & FormatEuroCurrency((totalYesterday * SESSION("affilCommision")) / 100.0) & " " & xxCommis %></div>
				</div>
				<div class="ectdivcontainer">
				  <div class="ectdivleft"><%=xxTotMTD%></div>
				  <div class="ectdivright"><% print FormatEuroCurrency(totalMonth)
				  if SESSION("affilCommision")<>0 then print " = " & FormatEuroCurrency((totalMonth * SESSION("affilCommision")) / 100.0) & " " & xxCommis %></div>
				</div>
				<div class="ectdivcontainer">
				  <div class="ectdivleft"><%=xxTotLM%></div>
				  <div class="ectdivright"><% print FormatEuroCurrency(totalLastMonth)
				  if SESSION("affilCommision")<>0 then print " = " & FormatEuroCurrency((totalLastMonth * SESSION("affilCommision")) / 100.0) & " " & xxCommis %></div>
				</div>
				<div class="ectdiv2column"><%=imageorsubmit(imglogout,xxLogout&""" onclick=""document.forms.mainform.act.value='logout'","logout") & " " & imageorsubmit(imgeditaffiliate,xxEdtAff&""" onclick=""document.forms.mainform.act.value='editaffil'","editaffiliate")%></div>
				<div class="ectdiv2column">
					<ul>
					  <li><%=xxAffLI1%>&nbsp;<%=IIfVr(seocategoryurls,replace(seoprodurlpattern,"%s",""),"products"&extension)%>?PARTNER=<%=htmlspecials(trim(SESSION("affilid")))%></li>
					  <li><%=xxAffLI2%></li>
					  <% if SESSION("affilCommision")=0 then %>
					  <li><%=xxAffLI3%></li>
					  <% end if %>
					</ul>
				</div>
			</div>
		</form>
<%	end if
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>