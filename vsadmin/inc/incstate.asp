<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
Dim sSQL,rs,allzones,success,cnn,rowcounter,alloptions,errmsg,index,cena,tax
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
editzones = ((shipType=2 OR shipType=5 OR adminIntShipping=2 OR adminIntShipping=5 OR alternateratesweightbased) AND splitUSZones)
if getget("setcatax")="true" then
	ect_query("UPDATE states SET stateTax=0 WHERE stateCountryID=2 AND stateAbbrev='AB'")
	ect_query("UPDATE states SET stateTax=7 WHERE stateCountryID=2 AND stateAbbrev='BC'")
	ect_query("UPDATE states SET stateTax=7 WHERE stateCountryID=2 AND stateAbbrev='MB'")
	ect_query("UPDATE states SET stateTax=8 WHERE stateCountryID=2 AND stateAbbrev='NB'")
	ect_query("UPDATE states SET stateTax=8 WHERE stateCountryID=2 AND stateAbbrev='NF'")
	ect_query("UPDATE states SET stateTax=0 WHERE stateCountryID=2 AND stateAbbrev='NT'")
	ect_query("UPDATE states SET stateTax=10 WHERE stateCountryID=2 AND stateAbbrev='NS'")
	ect_query("UPDATE states SET stateTax=0 WHERE stateCountryID=2 AND stateAbbrev='NU'")
	ect_query("UPDATE states SET stateTax=8 WHERE stateCountryID=2 AND stateAbbrev='ON'")
	ect_query("UPDATE states SET stateTax=9 WHERE stateCountryID=2 AND stateAbbrev='PE'")
	ect_query("UPDATE states SET stateTax=9.975 WHERE stateCountryID=2 AND stateAbbrev='QC'")
	ect_query("UPDATE states SET stateTax=5 WHERE stateCountryID=2 AND stateAbbrev='SK'")
	ect_query("UPDATE states SET stateTax=0 WHERE stateCountryID=2 AND stateAbbrev='YT'")
	ect_query("UPDATE countries SET countryTax=5 WHERE countryID=2")
end if
if getpost("posted")="1" then
	cena=0
	if getpost("ena")<>"" then cena=1
	fsa=0
	if getpost("fsa")<>"" then fsa=1
	tax = getpost("tax")
	if NOT is_numeric(tax) then
		success=false
		errmsg = yyNum100 & " """ & yyTax & """."
	elseif tax > 100 OR tax < 0 then
		success=false
		errmsg = yyNum100 & " """ & yyTax & """."
	else
		if editzones then
			sSQL = "UPDATE states SET stateEnabled="&cena&",stateTax="&tax&",stateFreeShip="&fsa&",stateZone="&getpost("zon")&" WHERE stateID="&getpost("id")
		else
			sSQL = "UPDATE states SET stateEnabled="&cena&",stateTax="&tax&",stateFreeShip="&fsa&" WHERE stateID="&getpost("id")
		end if
		ect_query(sSQL)
	end if
	if getpost("from")="pz" then print "<meta http-equiv=""refresh"" content=""0; url=adminzones.asp"" />"
elseif getpost("doeditstates")<>"" then
	nextfreeid=1
	for each objItem in request.form
		if left(objItem,9)="stateName" then
			stateID=right(objItem,len(objItem)-9)
			if getpost(objItem)="" then
				sSQL="DELETE FROM states WHERE stateID=" & stateID
			else
				sSQL="UPDATE states SET stateName='" & escape_string(getpost(objItem)) & "'"
				for index=2 to adminlanguages+1
					if (adminlangsettings AND 1048576)=1048576 then sSQL = sSQL & ",stateName"&index&"='" & escape_string(getpost("state"&index&"Name"&stateID)) & "'"
				next
				sSQL=sSQL & " WHERE stateID=" & stateID
			end if
			ect_query(sSQL)
		elseif left(objItem,12)="stateNewName" then
			rowid=right(objItem,len(objItem)-12)
			stateName=getpost(objItem)
			if stateName<>"" then
				stateName2=getpost("state2Name"&rowid)
				stateName3=getpost("state3Name"&rowid)
				if stateName2="" then stateName2=stateName
				if stateName3="" then stateName3=stateName
				gotstateid=FALSE
				do while NOT gotstateid
					rs.open "SELECT stateID FROM states WHERE stateID=" & nextfreeid,cnn,0,1
					if rs.EOF then gotstateid=TRUE else nextfreeid=nextfreeid+1
					rs.close
				loop
				sSQL = "INSERT INTO states (stateID,stateName,stateName2,stateName3,stateCountryID,stateEnabled) VALUES (" & _
					nextfreeid & "," & _
					"'" & escape_string(stateName) & "'," & _
					"'" & escape_string(stateName2) & "'," & _
					"'" & escape_string(stateName3) & "', " & getpost("thiscountry") & ",1)"
				ect_query(sSQL)
			end if
		end if
	next
	print "<meta http-equiv=""refresh"" content=""1; url=" & IIfVr(getpost("from")="pz","adminzones.asp","adminstate.asp?thiscountry=" & getpost("thiscountry")) & """ />"
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
		sSQL = "UPDATE states SET stateEnabled="&cena& " WHERE stateID IN (" & getpost("ids") & ")"
	elseif setallact="allfsa" then
		sSQL = "UPDATE states SET stateFreeShip="&fsa& " WHERE stateID IN (" & getpost("ids") & ")"
	elseif setallact="alltax" then
		if NOT is_numeric(tax) then
			success=false
			errmsg = yyNum100 & " """ & yyTax & """."
		elseif tax > 100 OR tax < 0 then
			success=false
			errmsg = yyNum100 & " """ & yyTax & """."
		else
			sSQL = "UPDATE states SET stateTax="&tax& " WHERE stateID IN (" & getpost("ids") & ")"
		end if
	elseif setallact="allzone" then
		sSQL = "UPDATE states SET stateZone="&zone& " WHERE stateID IN (" & getpost("ids") & ")"
	end if
	if success then ect_query(sSQL)
end if
if editzones then colspan="7" else colspan="6"
sSQL = "SELECT pzID,pzName FROM postalzones WHERE pzName<>'' AND pzID>100"
rs.open sSQL,cnn,0,1
allzones=""
if NOT rs.EOF then allzones=rs.getrows
rs.close
if getpost("doeditstates")<>"" then %>
			<p align="center"><br />&nbsp;<br />&nbsp;<br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <a href="adminstate.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br /><br /><br /></p>
<%
elseif (getpost("posted")="1" OR getpost("setallact")<>"") AND NOT success then %>
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><span style="color:#FF0000;font-weight:bold">Some records could not be updated.</span><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table>
<%
elseif getget("id")<>"" then %>
		  <form name="mainform" method="post" action="adminstate.asp">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="id" value="<%=getget("id")%>" />
			<input type="hidden" name="thiscountry" value="<%=request("thiscountry")%>" />
<%			if getget("from")="pz" then call writehiddenvar("from","pz") %>
			<table width="100%" border="0" cellspacing="1" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><strong><%=yyStaAdm%></strong><br /><br />
				<span style="font-size:10px"><%=yyFSANot%><br />&nbsp;</span></td>
			  </tr>
<%	sSQL = "SELECT stateID,stateName,stateEnabled,stateTax,stateZone,stateFreeShip FROM states WHERE stateID=" & replace(getget("id"),"'","")
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
%>
			  <tr>
				<td align="right" width="50%"><strong><%=yyStaNam%></strong></td>
				<td><strong><%=rs("stateName")%></strong></td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyEnable%></strong></td>
				<td><input type="checkbox" name="ena"<% if rs("stateEnabled")=1 then print " checked=""checked""" %> /></td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyTax%></strong></td>
				<td><input type="text" name="tax" value="<%=rs("stateTax")%>" size="4" />%</td>
			  </tr>
			  <tr>
				<td align="right"><strong><acronym title="<%=yyFSApp%>"><% print yyFSApp & " (" & yyFSA & ")"%></acronym></strong></td>
				<td><input type="checkbox" name="fsa"<% if rs("stateFreeShip")=1 then print " checked=""checked"""%> /></td>
			  </tr>
<%		if editzones then %>
			  <tr>
				<td align="right"><strong><%=yyPZone%></strong></td>
<%			foundzone=FALSE
			print "<td><select name=""zon"" size=""1"">"
			if IsArray(allzones) then
				for index=0 to UBOUND(allzones,2)
					print "<option value="""&allzones(0,index)&""""
					if rs("stateZone")=allzones(0,index) then
						print " selected=""selected"""
						foundzone=TRUE
					end if
					print ">"&allzones(1,index)&"</option>"&vbCrLf
				next
			end if
			if NOT foundzone then print "<option value=""0"" selected=""selected"">"&yyUndef&"</option>"
			print "</select></td></tr>"
		end if
	end if
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
<%	else
		if NOT is_numeric(request("thiscountry")) then thiscountry=origCountryID else thiscountry=cint(request("thiscountry"))
		forcezonesplit=((thiscountry=1 OR thiscountry=2) AND usandcasplitzones)
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
function doaddrow(){
var rownumber = document.getElementById("maxidvalue").value;
opttable = document.getElementById('statestable');
newrow = opttable.insertRow(opttable.rows.length-1);
if((parseInt(rownumber)%2)==0)newrow.className='cobhl';else newrow.className='cobll';
newcell = newrow.insertCell(0);
newcell.align='center';
newcell.innerHTML = '<input type="text" name="stateNewName'+rownumber+'" size="30" value="" />';
<%			rowcounter=1
			for index=2 to adminlanguages+1
				if (adminlangsettings AND 1048576)=1048576 then %>
newcell = newrow.insertCell(<%=rowcounter%>);
newcell.align='center';
newcell.innerHTML = '<input type="text" name="stateNew<%=index%>Name'+rownumber+'" size="30" value="" />';
<%					rowcounter=rowcounter+1
				end if
			next %>
newcell = newrow.insertCell(<%=rowcounter%>);
newcell.align='center';
newcell.innerHTML = '-';
document.getElementById("maxidvalue").value = parseInt(rownumber)+1;
}
function addmorerows(){
	numextrarows = document.getElementById("numextrarows").value;
	numextrarows = parseInt(numextrarows);
	if(isNaN(numextrarows))numextrarows=1;
	if(numextrarows==0)numextrarows=1;
	if(numextrarows>100)numextrarows=100;
	for(index=0;index<numextrarows;index++){
		doaddrow();
	}
}
/* ]]> */
</script>
	<table width="100%">
	  <tr>
		<td align="center">
		  <form name="mainform" method="post" action="adminstate.asp">
			<input type="hidden" name="setallact" id="setallact" value="" />
            <table width="100%" border="0" cellspacing="3" cellpadding="3">
			  <tr>
                <td align="center" colspan="2"><strong><%=yyStaAdm%></strong><br /><br />
<%	if getget("editstates")<>"1" then %>
				<span style="font-size:10px"><%=yyFSANot%><br />&nbsp;</span>
<%	end if %>
				</td>
			  </tr>
<%	if getget("editstates")<>"1" then %>
			  <tr>
                <td align="right" valign="top">
				  <table width="340" border="0" cellspacing="1" cellpadding="3" class="cobtbl">
					<tr height="30"><td class="cobhl" colspan="3" align="center"><strong><%=yyWitSel%>...</strong></td></tr>
					<tr height="30"><td class="cobhl" align="right"><strong><%=yyEnable%>:</strong></td><td class="cobll" align="left"><select name="allenable" size="1"><option value="ON"><%=yyYes%></option><option value=""><%=yyNo%></option></select></td><td class="cobll" align="center"><input type="button" value="<%=yySubmit%>" onclick="setall('allenable')" /></td></tr>
<%		if thiscountry=origCountryID OR forcezonesplit then
			if thiscountry=origCountryID then %>
					<tr height="30"><td class="cobhl" align="right"><strong><%=yyTax%>:</strong></td><td class="cobll" align="left"><input type="text" name="alltax" size="5" />%</td><td class="cobll" align="center"><input type="button" value="<%=yySubmit%>" onclick="setall('alltax')" /></td></tr>
					<tr height="30"><td class="cobhl" align="right"><strong><%=yyFSApp%>:</strong></td><td class="cobll" align="left"><select name="allfsa" size="1"><option value="ON"><%=yyYes%></option><option value=""><%=yyNo%></option></select></td><td class="cobll" align="center"><input type="button" value="<%=yySubmit%>" onclick="setall('allfsa')" /></td></tr>
<%			end if
			if editzones then %>
					<tr height="30"><td class="cobhl" align="right"><strong><%=yyPZone%>:</strong></td><td class="cobll" align="left"><select name="allzone" size="1">
<%				if IsArray(allzones) then
					for index=0 to UBOUND(allzones,2)
						print "<option value="""&allzones(0,index)&""""
						print ">"&allzones(1,index)&"</option>"&vbCrLf
					next
				end if %>
					</select></td><td class="cobll" align="center"><input type="button" value="<%=yySubmit%>" onclick="setall('allzone')" /></td></tr>
<%			end if
			if thiscountry=origCountryID AND thiscountry=2 then %>
					<tr height="30"><td class="cobll" colspan="3" align="center"><input type="button" value="Please click here to set Canadian tax rates" onclick="if(confirm('We make every effort to keep these Tax Rates up to date, but rates change\nfrequently. Please check the tax rates and inform us of any changes.'))document.location='adminstate.asp?thiscountry=2&setcatax=true'" /></td></tr>
<%			end if
		end if %>
				  </table>
				</td><td align="left" valign="top" width="50%">
				  <table width="340" border="0" cellspacing="1" cellpadding="3" class="cobtbl">
					<tr height="30"><td class="cobhl" colspan="3" align="center"><strong><%=yyStaCou%>...</strong></td></tr>
					<tr height="30"><td class="cobhl" align="right"><strong><%=yyCountry%>:</strong></td>
					<td class="cobll" align="left"><select size="1" name="thiscountry" id="thiscountry" onchange="document.location='adminstate.asp?thiscountry='+this[this.selectedIndex].value"><%
						gotstates=""
						sSQL="SELECT DISTINCT countryID,countryName FROM countries INNER JOIN states ON countries.countryID=states.stateCountryID ORDER BY countryName"
						rs.open sSQL,cnn,0,1
						do while NOT rs.EOF
							print "<option value="""& rs("countryID") & """" & IIfVr(thiscountry=rs("countryID")," selected=""selected""","") & ">" & htmlspecials(rs("countryName")) & "</option>"
							gotstates=gotstates & rs("countryID") & ","
							rs.movenext
						loop
						rs.close
						sSQL="SELECT countryID,countryName FROM countries " & IIfVr(gotstates<>"","WHERE countryID NOT IN (" & left(gotstates,len(gotstates)-1) & ") ","") & "ORDER BY countryName"
						print "<option value="""" disabled=""disabled"">----------------------</option>"
						rs.open sSQL,cnn,0,1
						do while NOT rs.EOF
							print "<option value="""& rs("countryID") & """" & IIfVr(thiscountry=rs("countryID")," selected=""selected""","") & ">" & htmlspecials(rs("countryName")) & "</option>"
							rs.movenext
						loop
						rs.close
						if is_numeric(getget("loadstates")) then
							ect_query("UPDATE countries SET loadStates=" & getget("loadstates") & " WHERE countryID=" & thiscountry)
						end if
						sSQL = "SELECT loadStates FROM countries WHERE countryID=" & thiscountry
						rs.open sSQL,cnn,0,1
						if NOT rs.EOF then loadstates=rs("loadStates") else loadstates=0
						rs.close
					%></select></td></tr>
					<tr height="30"><td class="cobhl" align="right"><strong><%=yyLoadSt%>:</strong></td>
					<td class="cobll" align="left"><select size="1" name="loadstates" onchange="document.location='adminstate.asp?thiscountry='+document.getElementById('thiscountry')[document.getElementById('thiscountry').selectedIndex].value+'&loadstates='+this[this.selectedIndex].value">
					<option value="0"><%=yyNo%></option>
<% '					<option value="1"<% if loadstates=1 then print " selected=""selected"""%>>Dynamically</option>
%>
					<option value="2"<% if loadstates=2 then print " selected=""selected"""%>><%=yyYes%></option>
					<option value="-1"<% if loadstates=-1 then print " selected=""selected"""%>>Not Required</option>
					</select>
					</td></tr>
					<tr height="30"><td class="cobhl" align="right"><strong><%=yyEdiSta%>:</strong></td>
					<td class="cobll" align="left"><input type="button" value="<%=yySubmit%>" onclick="document.location='adminstate.asp?thiscountry='+document.getElementById('thiscountry')[document.getElementById('thiscountry').selectedIndex].value+'&editstates=1'" /></td></tr>
				  </table>
				</td>
			  </tr>
<%	end if %>
			</table>
<%	if getget("editstates")="1" AND is_numeric(getget("thiscountry")) then %>
		  <input type="hidden" name="doeditstates" value="1" />
		  <input type="hidden" name="thiscountry" value="<%=getget("thiscountry")%>" />
<%		sSQL="SELECT countryName FROM countries WHERE countryID=" & getget("thiscountry")
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then countryName=rs("countryName") else countryName="UNDEFINED"
		rs.close
%>
			<p align="center">You are editing states for the country: <strong><%=countryName%></strong><br />&nbsp;</p>
			<table border="0" cellspacing="1" cellpadding="3" class="cobtbl" id="statestable">
			  <tr height="30">
				<td align="center" class="cobhl"><strong><%=yyStaNam%></strong></td>
<%		colspan=2
		for index=2 to adminlanguages+1
			if (adminlangsettings AND 1048576)=1048576 then
				colspan=colspan+1
%><td align="center" class="cobhl"><strong><%=yyStaNam&" "&index%></strong></td><%
			end if
		next %>
				<td class="cobhl" align="center"><strong>&nbsp;<%=yyEnable%>&nbsp;</strong></td>
			  </tr>
			
<%		sSQL = "SELECT stateID,stateName,stateName2,stateName3,stateEnabled FROM states WHERE stateCountryID=" & thiscountry & " ORDER BY stateName"
		rs.open sSQL,cnn,0,1
		hasrows=NOT rs.EOF
		do while NOT rs.EOF
			if bgcolor="cobhl" then bgcolor="cobll" else bgcolor="cobhl"
			%><tr align="center" class="<%=bgcolor%>">
<td align="center"><input type="text" size="30" name="stateName<%=rs("stateID")%>" value="<%=htmlspecials(rs("stateName"))%>" /></td>
<%				for index=2 to adminlanguages+1
					if (adminlangsettings AND 1048576)=1048576 then
%><td align="center"><input type="text" size="30" name="state<%=index%>Name<%=rs("stateID")%>" size="25" value="<%=htmlspecials(rs("stateName"&index))%>" /></td><%
					end if
				next %>
<td><%		if rs("stateEnabled")=1 then print yyYes else print "&nbsp;"%></td></tr>
<%			rs.movenext
		loop
		rs.close %>
			  <tr height="30">
				<td class="cobll" colspan="<%=colspan%>" align="center"><input type="hidden" name="maxidvalue" id="maxidvalue" value="1" /><input type="text" name="numextrarows" id="numextrarows" value="10" size="4" /> <input type="button" value="<%=yyMore & " " & yyLLStat%>" onclick="addmorerows()" />&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" value="<%=yySubmit%>" /> <input type="button" value="<%=yyCancel%>" onclick="document.location='adminstate.asp?thiscountry=<%=getget("thiscountry")%>'" /></td>
			  </tr>
			</table>
<%		if NOT hasrows then
			print "<script>addmorerows();</script>" & vbCrLf
		end if
	else %>
<br />
            <table width="100%" class="stackable admin-table-a sta-white">
			  <tr>
				<th class="minicell"><input type="checkbox" id="xdocheckall" value="1" onclick="docheckall()" /></th>
				<th class="maincell"><%=yyStaNam%></th>
				<th class="minicell"><%=yyEnable%></th>
				<th class="minicell"><%=yyTax%></th>
				<th class="minicell"><acronym title="<%=yyFSApp%>"><%=yyFSA%></acronym></th>
<%		if editzones then print "<th class=""minicell"">" & yyPZone & "</th>"
		if thiscountry=origCountryID OR forcezonesplit then print "<th class=""minicell"">" & yyModify & "</th>" %>
			  </tr><%
		theids = split(getpost("ids"), ",")
		for index=0 to UBOUND(theids)
			theids(index)=int(theids(index))
		next
		bgcolor="cobhl"
		sSQL = "SELECT stateID,stateName,stateEnabled,stateTax,stateZone,stateFreeShip FROM states WHERE stateCountryID=" & thiscountry & " ORDER BY stateEnabled DESC, stateName"
		rs.open sSQL,cnn,0,1
		if rs.EOF then
			print "<tr><td align=""center"" class=""cobll"" colspan=""" & IIfVr(editzones,7,6) & """><p>No states have been defined. To create states for this country please click on the button for Edit States above.</p></td></tr>"
		end if
		do while NOT rs.EOF
			if bgcolor="cobhl" then bgcolor="cobll" else bgcolor="cobhl"
			%><tr align="center" class="<%=bgcolor%>">
<td align="center"><input type="checkbox" name="ids" value="<%=rs("stateID")%>" <%
			for index=0 to UBOUND(theids)
				if theids(index)=rs("stateID") then print "checked=""checked"" " : exit for
			next
		%>/></td>
<td align="left"><%
			if rs("stateEnabled")=1 then print "<strong>"
			print rs("stateName")
			if rs("stateEnabled")=1 then print "</strong>" %></td>
<td><%		if rs("stateEnabled")=1 then print yyYes else print "&nbsp;"%></td>
<%			if thiscountry<>origCountryID AND NOT forcezonesplit then
				print "<td>-</td><td>-</td>" & IIfVs(editzones,"<td>-</td>")
			else %>
<td><%			if rs("stateTax")<>0 then print rs("stateTax") & "%" else print "&nbsp;"%></td>
<td><%			if rs("stateFreeShip")=1 AND rs("stateEnabled")=1 then print yyYes else print "&nbsp;"%></td>
<%				if editzones then
					if rs("stateEnabled")<>1 then
						print "<td>-</td>"
					else
						foundzone=FALSE
						if IsArray(allzones) then
							for index=0 to UBOUND(allzones,2)
								if rs("stateZone")=allzones(0,index) then
									print "<td>" & allzones(1,index) & "</td>"
									foundzone=TRUE
								end if
							next
						end if
						if NOT foundzone then print "<td>" & yyUndef & "</td>"
					end if
				end if
				print "<td><input type=""button"" onclick=""document.location='adminstate.asp?id=" & rs("stateID") & "'"" value=""" & yyModify & """/></td>"
			end if
			print "</tr>"
			rs.movenext
		loop
		rs.close %>
            </table>
<%	end if %>
		  <p align="center"><br /><a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</p>
		  </form>
		</td>
	  </tr>
	</table>
<%
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>