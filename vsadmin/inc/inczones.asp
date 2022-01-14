<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
Dim sSQL,rs,alldata,success,cnn,rowcounter,alloptions,errmsg,index,zoneName,foundmatch,upperbound,hasMultiShip,methodnames(10),hishipvals(10)
success=true
maxshippingmethods=5
alldata=""
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
if getpost("posted")="1" then
	for index=1 to 200
		if getpost("id"&index)="1" then
			sSQL = "UPDATE postalzones SET pzName='"&escape_string(getpost("zon"&index))&"' WHERE pzID="&index
			ect_query(sSQL)
		end if
	next
	print "<meta http-equiv=""refresh"" content=""1; url=adminzones.asp"">"
elseif getpost("posted")="2" then
	numshipmethods=getpost("numshipmethods")
	zone = getpost("zone")
	ect_query("DELETE FROM zonecharges WHERE zcZone="&getpost("zone"))
	if is_numeric(getpost("highweight")) then
		if cdbl(getpost("highweight")) > 0 then
			sSQL = "INSERT INTO zonecharges (zcZone,zcWeight,zcRate,zcRate2,zcRate3,zcRate4,zcRate5) VALUES ("&zone&","&cStr(0.0-cdbl(getpost("highweight")))
			for index=0 to maxshippingmethods-1
				if is_numeric(getpost("highvalue"&index)) then
					sSQL = sSQL & "," & getpost("highvalue"&index)
				else
					sSQL = sSQL & ",0"
				end if
			next
			ect_query(sSQL & ")")
		end if
	end if
	for index=0 to 59
		if is_numeric(getpost("weight"&index)) then
			if cdbl(getpost("weight"&index)) > 0 then
				sSQL = "INSERT INTO zonecharges (zcZone,zcWeight,zcRate,zcRatePC,zcRate2,zcRatePC2,zcRate3,zcRatePC3,zcRate4,zcRatePC4,zcRate5,zcRatePC5) VALUES ("&zone&","&getpost("weight"&index)
				for index2=0 to maxshippingmethods-1
					thecharge = getpost("charge"&index2&"x"&index)
					if is_numeric(replace(thecharge,"%","")) then
						sSQL = sSQL & "," & replace(thecharge,"%","")
					elseif LCase(thecharge)="x" then
						sSQL = sSQL & ",-99999.0"
					else
						sSQL = sSQL & ",0"
					end if
					if InStr(thecharge, "%") > 0 then sSQL = sSQL & ",1" else sSQL = sSQL & ",0"
				next
				ect_query(sSQL & ")")
			end if
		end if
	next
	sSQL = "UPDATE postalzones SET "
	addcomma=""
	pzFSA = 0
	for index=0 to maxshippingmethods-1
		sSQL = sSQL & addcomma & "pzMethodName" & (index+1) & "='" & escape_string(getpost("methodname"&index)) & "'"
		if trim(getpost("methodfsa"&index))="ON" then pzFSA = (pzFSA OR (2 ^ index))
		addcomma=","
	next
	sSQL = sSQL & ",pzFSA=" & pzFSA
	ect_query(sSQL & " WHERE pzID=" & zone)
	print "<meta http-equiv=""refresh"" content=""1; url=adminzones.asp"">"
elseif request.querystring("id")<>"" then
	if getget("shippingmethods")<>"" then
		sSQL = "UPDATE postalzones SET pzMultiShipping=" & request.querystring("shippingmethods") & " WHERE pzID=" & request.querystring("id")
		ect_query(sSQL)
	end if
	sSQL = "SELECT pzName,pzMultiShipping,pzFSA,pzMethodName1,pzMethodName2,pzMethodName3,pzMethodName4,pzMethodName5 FROM postalzones WHERE pzID="&request.querystring("id")
	rs.open sSQL,cnn,0,1
	zoneName=""
	if NOT rs.EOF then
		zoneName=rs("pzName")
		hasMultiShip=rs("pzMultiShipping")
		pzFSA=rs("pzFSA")
		for rowcounter=1 to maxshippingmethods
			methodnames(rowcounter-1)=rs("pzMethodName"&rowcounter)
		next
	end if
	rs.close
	sSQL = "SELECT zcID,zcWeight,zcRate,zcRate2,zcRate3,zcRate4,zcRate5,zcRatePC,zcRatePC2,zcRatePC3,zcRatePC4,zcRatePC5 FROM zonecharges WHERE zcZone="&request.querystring("id")&" ORDER BY zcWeight"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then alldata=rs.getrows
	rs.close
else
	if request.querystring("oneuszone")="yes" then
		sSQL = "UPDATE admin SET adminUSZones=0"
		ect_query(sSQL)
		splitUSZones=FALSE
	end if
	if request.querystring("oneuszone")="no" then
		sSQL = "UPDATE admin SET adminUSZones=1"
		ect_query(sSQL)
		splitUSZones=TRUE
	end if
	sSQL = "SELECT pzID,pzName FROM postalzones ORDER BY pzID"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then alldata=rs.getrows
	rs.close
end if
alreadygotadmin = getadminsettings()
Session.LCID=1033
if getpost("posted")="2" AND success then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <a href="adminzones.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br /><br />&nbsp;
                </td>
			  </tr>
			</table></td>
        </tr>
      </table>
<%
elseif getpost("posted")="2" then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><span style="color:#FF0000;font-weight:bold"><%=yyErrUpd%></span><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table></td>
        </tr>
      </table>
<%
elseif request.querystring("id")<>"" then %>
<script>
<!--
function formvalidator(theForm)
{
	var emptyentries=false;
<% for index=0 to hasMultiShip %>
	if (theForm.methodname<%=index%>.value == ""){
		alert("<%=jscheck(yyAllShp)%>");
		theForm.methodname<%=index%>.focus();
		return (false);
	}
<% next %>
	var checkOK = "0123456789.";
	var checkStr = theForm.highweight.value;
	var allValid = true;
	for (i = 0;  i < checkStr.length;  i++){
		ch = checkStr.charAt(i);
		for (j = 0;  j < checkOK.length;  j++)
			if (ch == checkOK.charAt(j))
				break;
		if (j == checkOK.length){
			allValid = false;
				break;
		}
	}
	if (!allValid){
		alert("<%=jscheck(yyDecFld)%>");
		theForm.highweight.focus();
		return (false);
	}
	for(index=0; index<<%=maxshippingmethods%>;index++){
		var theobj = eval("theForm.highvalue"+index);
		var checkStr = theobj.value;
		var allValid = true;
		for (i = 0;  i < checkStr.length;  i++){
			ch = checkStr.charAt(i);
			for (j = 0;  j < checkOK.length;  j++)
				if (ch == checkOK.charAt(j))
					break;
			if (j == checkOK.length){
				allValid = false;
					break;
			}
		}
		if (!allValid){
			alert("<%=jscheck(yyDecFld)%>");
			theobj.focus();
			return (false);
		}
	}
	for(index=0;index<60;index++){
		var theobj = eval("theForm.weight"+index);
		var checkStr = theobj.value;
		var allValid = true;
		var hasweight = (theobj.value != "");
		for (i = 0;  i < checkStr.length;  i++){
			ch = checkStr.charAt(i);
			for (j = 0;  j < checkOK.length;  j++)
			  if (ch == checkOK.charAt(j))
				break;
			if (j == checkOK.length){
				allValid = false;
				break;
			}
		}
		if (!allValid){
			alert("<%=jscheck(yyDecFld)%>");
			theobj.focus();
			return (false);
		}
		for(index2=0; index2<=<%=hasMultiShip%>;index2++){
			var theobj = eval("theForm.charge"+index2+"x"+index);
			var checkOK = "0123456789.%";
			var checkStr = theobj.value;
			var allValid = true;
			if(hasweight && checkStr==""){
				emptyentries=true;
				emptyobj=theobj;
			}
			for (i = 0;  i < checkStr.length;  i++){
				ch = checkStr.charAt(i);
				for (j = 0;  j < checkOK.length;  j++)
					if (ch == checkOK.charAt(j))
						break;
				if (j == checkOK.length && checkStr.toLowerCase()!="x"){
					allValid = false;
					break;
				}
			}
			if (!allValid){
				alert("<%=jscheck(yyDecFld)%>");
				theobj.focus();
				return (false);
			}
		}
	}
	if(emptyentries){
		if(!confirm("<%=jscheck(yyNoMeth)%> <%if shipType=5 then print jscheck(yyMaxPri) else print jscheck(yyMaxWei)%><%=jscheck(yyNoMet2)%> <%=jscheck(yyNoInt)%>\n\n<%=jscheck(yyOkCan)%>")){
			emptyobj.focus();
			return(false);
		}
	}
	return (true);
}
function setnummethods(){
setto=document.forms.mainform.numshipmethods.selectedIndex;
document.location="adminzones.asp?shippingmethods="+setto+"&id=<%=request.querystring("id")%>";
}
//-->
</script>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
		  <td width="100%" align="center">
		  <form name="mainform" method="post" action="adminzones.asp" onsubmit="return formvalidator(this)">
			<input type="hidden" name="posted" value="2" />
			<input type="hidden" name="zone" value="<%=request.querystring("id")%>" />
            <table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" align="center"><strong><%=yyModRul%> <%
				if zoneName<>"" then
					print chr(34)&zoneName&chr(34)
				else
					print "(unnamed)"
				end if%>.</strong><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" align="center">
				  <span style="font-size:10px"><%=yyZonUse%> 
					<select name="numshipmethods" size="1" onchange="setnummethods()"><% 
						for rowcounter=1 to 5
							print "<option value=""" & rowcounter & """"
							if rowcounter = (hasMultiShip+1) then print " selected=""selected"""
							print ">" & rowcounter & "</option>"
						next %></select> <%=yyZonUs2%></span>
				</td>
			  </tr>
			  <tr> 
                <td width="100%" align="center">
				<table width="80%" cellspacing="2" cellpadding="0">
				  <tr>
					<td align="right" width="45%"><%=yyForEv%></td>
					<td width="10%"><input type="text" name="highweight" value="<%
				foundmatch=0
				if IsArray(alldata) then
					for rowcounter=0 to UBOUND(alldata,2)
						if alldata(1,rowcounter) < 0 then
							foundmatch=abs(alldata(1,rowcounter))
							for index=0 to maxshippingmethods-1
								hishipvals(index)=alldata(2+index,rowcounter)
							next
						end if
					next
				end if
				print foundmatch
				%>" size="5" /></td>
					<td width="45%" align="left"><%=yyAbvHg & " "%> <%if shipType=5 then print yyPrice else print yyWeigh%>...</td>
				  </tr>
<%					for index=0 to hasMultiShip %>
				  <tr>
					<td align="right"><%=yyAddExt%></td>
					<td><input type="text" name="highvalue<%=index%>" value="<%=hishipvals(index) %>" size="5" /></td><td align="left"><%=yyFor%> <strong><% if methodnames(index)<>"" then print methodnames(index) else print yyShipMe & " " & index+1%></strong></td>
				  </tr>
<%					next %>
				</table>
<%					for index=hasMultiShip+1 to maxshippingmethods-1 %>
				  <input type="hidden" name="highvalue<%=index%>" value="<%=hishipvals(index) %>" />
<%					next %>
				</td>
			  </tr>
			  <tr> 
                <td width="100%" align="center">
                  <p><input type="submit" value="<%=yySubmit%>" />&nbsp;&nbsp;<input type="reset" value="<%=yyReset%>" /><br />&nbsp;</p>
                </td>
			  </tr>
			</table>
			<table width="120" border="0" cellspacing="0" cellpadding="1">
			  <tr>
				<td width="<%=Int(100/(2+hasMultiShip))%>%" align="center">&nbsp;</td>
				<%	for index=0 to hasMultiShip
						print "<td width="""&Int(100/(2+hasMultiShip))&"%"" align=""center""><acronym title="""&yyFSApp&"""><strong>"&yyFSA&"</strong></acronym>: <input type=""checkbox"" value=""ON"" name=""methodfsa"&index&""" "&IIfVr((pzFSA AND (2 ^ index)) <> 0,"checked=""checked""","")&" /></td>" & vbCrLf
					next
					for index=hasMultiShip+1 to maxshippingmethods-1
						print "<input type=""hidden"" name=""methodfsa"&index&""" value="""&IIfVr((pzFSA AND (2 ^ index)) <> 0,"ON","")&""" />" & vbCrLf
					next %>
			  </tr>
			  <tr>
				<td align="center"><strong><%if shipType=5 then print yyMaxPri else print yyMaxWgt%></strong></td>
				<%	for index=0 to hasMultiShip
						print "<td align=""center""><input class=""darkborder"" type=""text"" name=""methodname"&index&""" value=""" & htmlspecials(methodnames(index)) & """ size=""14"" /></td>" & vbCrLf
					next
					for index=hasMultiShip+1 to maxshippingmethods-1
						print "<input type=""hidden"" name=""methodname"&index&""" value=""" & htmlspecials(methodnames(index)) & """ />" & vbCrLf
					next %>
			  </tr>
<%	rowcounter=0
	index=0
	if IsArray(alldata) then
		upperbound = UBOUND(alldata,2)
	else
		upperbound = -1
	end if
	do while index < 60
		if rowcounter <= upperbound then
			if alldata(1,rowcounter) > 0 then %>
			  <tr>
				<td align="center"><input class="darkborder" type="text" name="weight<%=index%>" value="<%=alldata(1,rowcounter)%>" size="10" /></td>
				<%	for index2=0 to maxshippingmethods-1
						if index2 <= hasMultiShip then
							print "<td align=""center""><input type=""text"" name=""charge"&index2&"x"&index&""" value="""&IIfVr(alldata(2+index2,rowcounter)<>-99999,alldata(2+index2,rowcounter)&IIfVr(cint(alldata(7+index2,rowcounter))<>0,"%",""),"x")&""" size=""14"" /></td>" & vbCrLf
						else
							print "<input type=""hidden"" name=""charge"&index2&"x"&index&""" value="""&alldata(2+index2,rowcounter)&""" />"
						end if
					next %>
			  </tr>
<%				index=index+1
			end if
		else %>
			  <tr>
				<td align="center"><input class="darkborder" type="text" name="weight<%=index%>" value="" size="10" /></td>
				<%	for index2=0 to maxshippingmethods-1
						if index2 <= hasMultiShip then
							print "<td align=""center""><input type=""text"" name=""charge"&index2&"x"&index&""" size=""14"" /></td>" & vbCrLf
						end if
					next %>
			  </tr>
<%			index=index+1
		end if
		rowcounter=rowcounter+1
	loop %>
			  <tr> 
                <td width="100%" colspan="<%=2+hasMultiShip%>" align="center">
                  <p><input type="submit" value="<%=yySubmit%>" />&nbsp;&nbsp;<input type="reset" value="<%=yyReset%>" /><br />&nbsp;</p>
                </td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="<%=2+hasMultiShip%>" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
			  </tr>
            </table>
		  </form>
		  </td>
        </tr>
      </table>
<%
elseif getpost("posted")="1" AND success then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <a href="adminzones.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br /><br />&nbsp;
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
                <td width="100%" colspan="2" align="center"><br /><span style="color:#FF0000;font-weight:bold"><%=yyErrUpd%></span><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table></td>
        </tr>
      </table>
<%
else %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
		  <td width="100%">
		  <form name="mainform" method="post" action="adminzones.asp">
			<input type="hidden" name="posted" value="1" />
            <table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" <%if splitUSZones then print "colspan='2'"%> align="center"><strong><%=yyModPZo%></strong><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" <%if splitUSZones then print "colspan='2'"%> align="left">
				  <ul>
					<li><span style="font-size:10px"><strong><%=yyPZEx1%></strong></span></li>
				  <% if splitUSZones then %>
					<li><span style="font-size:10px"><%=yyPZEx2%> <a href="adminzones.asp?oneuszone=yes"><strong><%=yyClkHer%></strong></a>.</span></li>
				  <% else %>
				    <li><span style="font-size:10px"><%=yyPZEx3%> <a href="adminzones.asp?oneuszone=no"><strong><%=yyClkHer%></strong></a>.</span></li>
				  <% end if %>
					<li><span style="font-size:10px"><%=yyPZEx4%></span></li>
				  </ul>
				</td>
			  </tr>
<%		if splitUSZones then
			sSQL="SELECT stateID,stateName,stateCountryID,pzID FROM states LEFT JOIN (SELECT pzID FROM postalzones WHERE pzName<>'' AND pzID>100) AS pz_table ON states.stateZone=pz_table.pzID WHERE stateCountryID=" & origCountryID & " AND pzID IS NULL"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				print "<tr><td><div class=""ectred"" style=""margin-bottom:5px"">WARNING The following states have an undefined postal zone...</div>"
				do while NOT rs.EOF
					print "<div><a href=""adminstate.asp?from=pz&amp;id=" & rs("stateID") & """>" & rs("stateName") & " (click to repair)</a></div>"
					rs.movenext
				loop
				print "</td></tr>"
			end if
			rs.close
		end if
		
		sSQL="SELECT countryID,countryName,pzID FROM countries LEFT JOIN (SELECT pzID FROM postalzones WHERE pzName<>'' AND pzID<100) AS pz_table ON countries.countryZone=pz_table.pzID WHERE countryEnabled<>0 AND pzID IS NULL"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			print "<tr><td><div class=""ectred"" style=""margin-bottom:5px"">WARNING The following countries have an undefined postal zone...</div>"
			do while NOT rs.EOF
				print "<div><a href=""admincountry.asp?from=pz&amp;id=" & rs("countryID") & """>" & rs("countryName") & " (click to repair)</a></div>"
				rs.movenext
			loop
			print "</td></tr>"
		end if
		rs.close
%>
			  <tr>
				<td valign="top">
				  <table width="100%" cellspacing="1" cellpadding="1" border="0">
					<tr> 
					  <td width="100%" colspan="3" align="center"><strong><%=yyPZWor%></strong><br /><hr width="70%" /></td>
					</tr>
					 <tr>
					  <td width="40%" align="right">&nbsp;</td>
					  <td width="20%" align="center"><strong><%=yyPZNam%></strong></td>
					  <td width="40%" align="left"><strong><%=yyPZRul%></strong></td>
					</tr>
<%	for rowcounter=0 to UBOUND(alldata,2)
		if alldata(0,rowcounter)<100 then ' First 100 are for world zones
%>					<tr>
					  <td align="right"><strong><%=alldata(0,rowcounter)%> : <input type="hidden" name="id<%=alldata(0,rowcounter)%>" value="1" /></strong></td>
					  <td align="center"><input type="text" name="zon<%=alldata(0,rowcounter)%>" value="<%=alldata(1,rowcounter)%>" size="20" /></td>
					  <td align="left"><% if trim(alldata(1,rowcounter))<>"" then %><a href="adminzones.asp?id=<%=alldata(0,rowcounter)%>"><strong><%=yyEdRul%></strong></a><% else %>&nbsp;<% end if %></td>
					</tr>
<%		end if
	next %>
				  </table>
				</td>
<%	if splitUSZones then %>
				<td width="50%" valign="top">
				  <table width="100%" cellspacing="1" cellpadding="1" border="0">
					<tr> 
					  <td width="100%" colspan="3" align="center"><strong><%=yyPZSta%></strong><br /><hr width="70%" /></td>
					</tr>
					 <tr>
					  <td width="40%" align="right">&nbsp;</td>
					  <td width="20%" align="center"><strong><%=yyPZNam%></strong></td>
					  <td width="40%" align="left"><strong><%=yyPZRul%></strong></td>
					</tr>
<%		index = 0
		for rowcounter=0 to UBOUND(alldata,2)
			if alldata(0,rowcounter)>100 then ' Next 100 are for world zones
				index=index+1 %>
					<tr>
					  <td align="right"><strong><%=alldata(0,rowcounter)-100%> : <input type="hidden" name="id<%=alldata(0,rowcounter)%>" value="1" /></strong></td>
					  <td align="center"><input type="text" name="zon<%=alldata(0,rowcounter)%>" value="<%=alldata(1,rowcounter)%>" size="20" /></td>
					  <td align="left"><% if trim(alldata(1,rowcounter))<>"" then %><a href="adminzones.asp?id=<%=alldata(0,rowcounter)%>"><strong><%=yyEdRul%></strong></a><% else %>&nbsp;<% end if %></td>
					</tr>
<%			end if
		next %>
				  </table>
				</td>
<%	end if %>
			  </tr>
			  <tr> 
                <td width="100%" <%if splitUSZones then print "colspan='2'"%> align="center">
                  <p><input type="submit" value="<%=yySubmit%>" />&nbsp;&nbsp;<input type="reset" value="<%=yyReset%>" /><br />&nbsp;</p>
                </td>
			  </tr>
			  <tr> 
                <td width="100%" <%if splitUSZones then print "colspan='2'"%> align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
			  </tr>
            </table>
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
