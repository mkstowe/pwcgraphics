<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
Dim sSQL,rs,alldata,allzones,success,cnn,rowcounter,alloptions,errmsg,index,cena,tax
success=true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
Session.LCID=1033
alternateratesweightbased=FALSE
if adminAltRates>0 then
	sSQL = "SELECT altrateid FROM alternaterates WHERE (altrateid=2 OR altrateid=5) AND (usealtmethod<>0 OR usealtmethodintl<>0)"
	rs.open sSQL,cnn,0,1
	alternateratesweightbased = NOT rs.EOF
	rs.close
end if
editzones = (shipType=2 OR shipType=5 OR adminIntShipping=2 OR adminIntShipping=5 OR alternateratesweightbased)
if getpost("posted")="1" then
	cena=0
	if getpost("ena")<>"" then cena=1
	fsa=0
	if getpost("fsa")<>"" then fsa=1
	tax = getpost("tax")
	taxthreshold = getpost("taxthreshold")
	if NOT is_numeric(taxthreshold) then taxthreshold=0
	if NOT is_numeric(tax) then
		success=false
		errmsg = yyNum100 & " """ & yyTax & """."
	elseif tax > 100 OR tax < 0 then
		success=false
		errmsg = yyNum100 & " """ & yyTax & """."
	else
		sSQL = "UPDATE countries SET countryEnabled="&cena&",countryTax="&tax&",countryTaxThreshold="&taxthreshold&",countryFreeShip="&fsa&",countryOrder="&getpost("pos")&",countryLCID='"&getpost("lcid")&"'"
		if editzones then sSQL=sSQL&",countryZone="&getpost("zon")
		if getpost("countryname")<>"" then sSQL=sSQL&",countryName='"&escape_string(getpost("countryname"))&"'"
		if getpost("countryname2")<>"" then sSQL=sSQL&",countryName2='"&escape_string(getpost("countryname2"))&"'"
		if getpost("countryname3")<>"" then sSQL=sSQL&",countryName3='"&escape_string(getpost("countryname3"))&"'"
		sSQL=sSQL&",currSymbolText='"&escape_string(request.form("currSymbolText"))&"'"
		sSQL=sSQL&",currDecimalSep='"&escape_string(getpost("currDecimalSep"))&"'"
		sSQL=sSQL&",currThousandsSep='"&escape_string(getpost("currThousandsSep"))&"'"
		sSQL=sSQL&",currPostAmount='"&escape_string(getpost("currPostAmount"))&"'"
		sSQL=sSQL&",currDecimals='"&escape_string(getpost("currDecimals"))&"'"
		sSQL=sSQL&",currSymbolHTML='"&escape_string(request.form("currSymbolHTML"))&"'"
		sSQL=sSQL&" WHERE countryID="&getpost("id")
		ect_query(sSQL)
	end if
	if getpost("from")="pz" then print "<meta http-equiv=""refresh"" content=""0; url=adminzones.asp"" />"
elseif getpost("setallact")<>"" then
	setallact = getpost("setallact")
	cena=0
	if getpost("allenable")="ON" then cena=1
	fsa=0
	if getpost("allfsa")<>"" then fsa=1
	tax = getpost("alltax")
	pos = getpost("allpos")
	zone = getpost("allzone")
	if setallact="allenable" then
		sSQL = "UPDATE countries SET countryEnabled="&cena& " WHERE countryID IN (" & getpost("ids") & ")"
	elseif setallact="allfsa" then
		sSQL = "UPDATE countries SET countryFreeShip="&fsa& " WHERE countryID IN (" & getpost("ids") & ")"
	elseif setallact="alltax" then
		if NOT is_numeric(tax) then
			success=false
			errmsg = yyNum100 & " """ & yyTax & """."
		elseif tax > 100 OR tax < 0 then
			success=false
			errmsg = yyNum100 & " """ & yyTax & """."
		else
			sSQL = "UPDATE countries SET countryTax="&tax& " WHERE countryID IN (" & getpost("ids") & ")"
		end if
	elseif setallact="allpos" then
		sSQL = "UPDATE countries SET countryOrder="&pos& " WHERE countryID IN (" & getpost("ids") & ")"
	elseif setallact="allzone" then
		sSQL = "UPDATE countries SET countryZone="&zone& " WHERE countryID IN (" & getpost("ids") & ")"
	end if
	if success then ect_query(sSQL)
end if
sSQL = "SELECT countryID,countryName,countryEnabled,countryTax,countryOrder,countryZone,countryFreeShip FROM countries ORDER BY countryOrder DESC,countryName"
rs.open sSQL,cnn,0,1
alldata=rs.getrows
rs.close
sSQL = "SELECT pzID,pzName FROM postalzones WHERE pzName<>'' AND pzID<100"
rs.open sSQL,cnn,0,1
allzones=""
if NOT rs.EOF then allzones=rs.getrows
rs.close
if (getpost("posted")="1" OR getpost("setallact")<>"") AND NOT success then %>
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><span style="color:#FF0000;font-weight:bold">Some records could not be updated.</span><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table>
<%
elseif getget("id")<>"" AND is_numeric(getget("id")) then %>
		  <form name="mainform" method="post" action="admincountry.asp">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="id" value="<%=getget("id")%>" />
<%			if getget("from")="pz" then call writehiddenvar("from","pz") %>
			<table width="100%" border="0" cellspacing="1" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><strong><%=yyCntAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="2"><ul>
				<li>You should not normally have to edit country names directly apart from regional variations such as &quot;United Kingdom&quot; for &quot;Great Britain&quot;.</li>
				<li>There are scripts available for adding foreign language country names.</li>
				<li>Do not edit country names if using the USPS shipping carrier as this carrier relies on the country name for rates.</li>
				</ul></td>
			  </tr>
<%	sSQL = "SELECT countryID,countryName,countryName2,countryName3,countryEnabled,countryTax,countryTaxThreshold,countryOrder,countryZone,countryFreeShip,countryLCID,currSymbolText,currDecimalSep,currThousandsSep,currPostAmount,currDecimals,currSymbolHTML FROM countries WHERE countryID=" & replace(getget("id"),"'","") & " ORDER BY countryOrder DESC,countryName"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
%>
			  <tr>
				<td align="right" width="50%"><strong><%=yyCntNam%></strong></td>
				<td><input type="text" name="countryname" value="<%=htmlspecials(rs("countryName"))%>" size="30" /></td>
			  </tr>
<%		for index=2 to adminlanguages+1
			if (adminlangsettings AND 8)=8 then %>
			  <tr>
				<td align="right" width="50%"><strong><%=yyCntNam&" "&index%></strong></td>
				<td><input type="text" name="countryname<%=index%>" value="<%=htmlspecials(rs("countryName"&index))%>" size="30" /></td>
			  </tr>
<%			end if
		next %>
			  <tr>
				<td align="right"><strong><%=yyEnable%></strong></td>
				<td><input type="checkbox" name="ena"<% if rs("countryEnabled")=1 then print " checked=""checked""" %> /></td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyTax%></strong></td>
				<td><input type="text" name="tax" value="<%=rs("countryTax")%>" size="4" style="text-align:right" />%</td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyTax&" "&"Threshold"%></strong></td>
				<td><input type="text" name="taxthreshold" value="<%=rs("countryTaxThreshold")%>" size="4" style="text-align:right" /></td>
			  </tr>
			  <tr>
				<td align="right"><strong><acronym title="<%=yyFSApp%>"><% print yyFSApp & " (" & yyFSA & ")"%></acronym></strong></td>
				<td><input type="checkbox" name="fsa"<% if rs("countryFreeShip")=1 then print " checked=""checked"""%> /></td>
			  </tr>
			  <tr>
				<td align="right"><strong><% print yyPosit%></strong></td>
				<td><select name="pos" size="1">
<option value="0"><%=yyAlphab%></option>
<option value="1"<% if rs("countryOrder")=1 then print " selected=""selected""" %>><% print yyTop%></option>
<option value="2"<% if rs("countryOrder")=2 then print " selected=""selected""" %>><% print yyTop&"+"%></option>
<option value="3"<% if rs("countryOrder")=3 then print " selected=""selected""" %>><% print yyTop&"++"%></option></select></td>
			  </tr>
<%		if editzones then %>
			  <tr>
				<td align="right"><strong><%=yyPZone%></strong></td>
<%			foundzone=FALSE
			print "<td><select name=""zon"" size=""1"">"
			if IsArray(allzones) then
				for index=0 to UBOUND(allzones,2)
					print "<option value="""&allzones(0,index)&""""
					if rs("countryZone")=allzones(0,index) then
						print " selected=""selected"""
						foundzone=TRUE
					end if
					print ">"&allzones(1,index)&"</option>"&vbCrLf
				next
			end if
			if NOT foundzone then print "<option value=""0"" selected=""selected"">"&yyUndef&"</option>"
			print "</select></td>"
		end if %>
			  </tr>
			  <tr>
				<td align="right"><strong>Locale ID (Do not change)</strong></td>
				<td><input type="text" name="lcid" value="<%=rs("countryLCID")%>" size="6" /></td>
			  </tr>
			  <tr>
				<td colspan="2" align="center">&nbsp;<br /><strong>Currency Format</strong></td>
			  </tr>
			  <tr>
				<td colspan="2" align="center">
					<table>
					  <tr>
						<td align="center">Post-Amount</td>
						<td align="center">HTML Symbol</td>
						<td align="center">Text Symbol</td>
						<td align="center">Thousands Separator</td>
						<td align="center">Decimals Separator</td>
						<td align="center">Decimal Places</td>
					  </tr>
					  <tr>
						<td align="center"><select name="currPostAmount"><option value="0">No</option><option value="1"<% if rs("currPostAmount")<>0 then print " selected=""selected"""%>>Yes</option></select></td>
						<td align="center"><input type="text" name="currSymbolHTML" size="6" value="<%=htmlspecials(rs("currSymbolHTML")) %>" /></td>
						<td align="center"><input type="text" name="currSymbolText" size="4" value="<%=htmlspecials(rs("currSymbolText")) %>" /></td>
						<td align="center"><input type="text" name="currThousandsSep" size="4" value="<%=htmlspecials(rs("currThousandsSep")) %>" /></td>
						<td align="center"><input type="text" name="currDecimalSep" size="4" value="<%=htmlspecials(rs("currDecimalSep")) %>" /></td>
						<td align="center"><select name="currDecimals"><option value="0">Zero</option><option value="2"<% if rs("currDecimals")=2 then print " selected=""selected"""%>>2</option></select></td>
					  </tr>
					</table>
				</td>
			  </tr>
<%	end if 
	rs.close %>
			  <tr> 
                <td width="100%" colspan="2" align="center">
				  <p>&nbsp;</p>
                  <p><input type="submit" value="<%=yySubmit%>" />&nbsp;&nbsp;<input type="reset" value="<%=yyReset%>" /><br />&nbsp;</p>
                </td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="2" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a></td>
			  </tr>
			</table>
		  </form>
<%
else
	if editzones then colspan="8" else colspan="7"
%>
<script>
/* <![CDATA[ */
function docheckall(){
	allcbs = document.getElementsByName('ids');
	mainidchecked = document.getElementById('xdocheckall').checked;
	for(i=0;i<allcbs.length;i++) {
		allcbs[i].checked=mainidchecked;
	}
	return(true);
}
function setall(theact){
	allcbs = document.getElementsByName('ids');
	var onechecked=false;
	for(i=0;i<allcbs.length;i++) {
		if(allcbs[i].checked)onechecked=true;
	}
	if(onechecked){
		document.getElementById('setallact').value=theact;
		document.forms.mainform.submit();
	}else{
		alert("<%=jscheck(yyNoSelO)%>");
	}
}
/* ]]> */
</script>
		  <form name="mainform" method="post" action="admincountry.asp">
			<input type="hidden" name="setallact" id="setallact" value="" />
            <table width="100%" border="0" cellspacing="1" cellpadding="1">
			  <tr> 
                <td align="center"><strong><%=yyCntAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td align="center">
				  <table border="0" cellspacing="1" cellpadding="3" class="cobtbl">
					<tr><td class="cobhl" colspan="3" align="center"><strong><%=yyWitSel%>...</strong></td></tr>
					<tr><td class="cobhl" align="right"><strong><%=yyEnable%>:</strong></td><td class="cobll" align="left"><select name="allenable" size="1"><option value="ON"><%=yyYes%></option><option value=""><%=yyNo%></option></select></td><td class="cobll"><input type="button" value="<%=yySubmit%>" onclick="setall('allenable')" /></td></tr>
					<tr><td class="cobhl" align="right"><strong><%=yyTax%>:</strong></td><td class="cobll" align="left"><input type="text" name="alltax" size="5" />%</td><td class="cobll"><input type="button" value="<%=yySubmit%>" onclick="setall('alltax')" /></td></tr>
					<tr><td class="cobhl" align="right"><strong><%=yyFSApp%>:</strong></td><td class="cobll" align="left"><select name="allfsa" size="1"><option value="ON"><%=yyYes%></option><option value=""><%=yyNo%></option></select></td><td class="cobll"><input type="button" value="<%=yySubmit%>" onclick="setall('allfsa')" /></td></tr>
					<tr><td class="cobhl" align="right"><strong><%=yyPosit%>:</strong></td><td class="cobll" align="left"><select name="allpos" size="1" >
						<option value="0"><%=yyAlphab%></option>
						<option value="1"><% print yyTop%></option>
						<option value="2"><% print yyTop & "+"%></option>
						<option value="3"><% print yyTop & "++"%></option></select>
					</td><td class="cobll"><input type="button" value="<%=yySubmit%>" onclick="setall('allpos')" /></td></tr>
<%	if editzones then %>
					<tr><td class="cobhl" align="right"><strong><%=yyPZone%>:</strong></td><td class="cobll" align="left"><select name="allzone" size="1">
<%		if IsArray(allzones) then
			for index=0 to UBOUND(allzones,2)
				print "<option value="""&allzones(0,index)&""""
				print ">"&allzones(1,index)&"</option>"&vbCrLf
			next
		end if %>
					</select></td><td class="cobll"><input type="button" value="<%=yySubmit%>" onclick="setall('allzone')" /></td></tr>
<%	end if %>
				  </table><br />
				</td>
			  </tr>
			</table>
<br />
            <table width="100%" class="stackable admin-table-a sta-white">
			  <tr>
				<th class="minicell"><input type="checkbox" id="xdocheckall" value="1" onclick="docheckall()" /></th>
				<th class="maincell"><%=yyCntNam%></th>
				<th class="minicell"><%=yyEnable%></th>
				<th class="minicell"><%=yyTax%></th>
				<th class="minicell"><acronym title="<%=yyFSApp%>"><%=yyFSA%></acronym></th>
				<th class="minicell"><%=yyPosit%></th>
<%	if editzones then print "<th class=""minicell"">" & yyPZone & "</th>" %>
				<th class="minicell"><%=yyModify%></th>
			  </tr><%
	theids = split(getpost("ids"), ",")
	for index=0 to UBOUND(theids)
		theids(index)=int(theids(index))
	next
	bgcolor="cobhl"
	sSQL = "SELECT countryID,countryName,countryEnabled,countryTax,countryTaxThreshold,countryOrder,countryZone,countryFreeShip FROM countries ORDER BY countryEnabled DESC,countryOrder DESC,countryName"
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		if bgcolor="cobhl" then bgcolor="cobll" else bgcolor="cobhl"
		%><tr align="center" class="<%=bgcolor%>">
<td align="center"><input type="checkbox" name="ids" value="<%=rs("countryID")%>" <%
		for index=0 to UBOUND(theids)
			if theids(index)=rs("countryID") then print "checked=""checked"" " : exit for
		next
		%>/></td>
<td align="left"><%
		if rs("countryEnabled")=1 then print "<strong>"
		print rs("countryName")
		if rs("countryEnabled")=1 then print "</strong>"%></td>
<td><%	if rs("countryEnabled")=1 then print yyYes else print "&nbsp;"%></td>
<td><%	if rs("countryTax")<>0 then print rs("countryTax") & "%" & IIfVs(rs("countryTaxThreshold")<>0," / " & rs("countryTaxThreshold")) else print "&nbsp;"%></td>
<td><%	if rs("countryFreeShip")=1 then print yyYes else print "&nbsp;"%></td>
<td><%	if rs("countryEnabled")<>1 then
			print "-"
		elseif rs("countryOrder")=1 then
			print yyTop
		elseif rs("countryOrder")=2 then
			print yyTop & "+"
		elseif rs("countryOrder")=3 then
			print yyTop & "++"
		else
			print yyAlphab
		end if
		print "</td>"
		if editzones then
			if rs("countryEnabled")<>1 then
				print "<td>-</td>"
			else
				foundzone=FALSE
				if IsArray(allzones) then
					for index=0 to UBOUND(allzones,2)
						if rs("countryZone")=allzones(0,index) then
							print "<td>" & allzones(1,index) & "</td>"
							foundzone=TRUE
						end if
					next
				end if
				if NOT foundzone then print "<td>" & yyUndef & "</td>"
			end if
		end if
		print "<td>"
		print "<input type=""button"" onclick=""document.location='admincountry.asp?id=" & rs("countryID") & "'"" value=""" & yyModify & """/>"
		print "</td></tr>"
		rs.MoveNext
	loop
	rs.close
%>
			  <tr> 
                <td class="cobll" width="100%" colspan="<%=colspan%>" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
			  </tr>
            </table>
		  </form>
<%
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>