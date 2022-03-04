<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
Dim sSQL,rs,cnn,success,showaccount,addsuccess,alldata,index,allcountries,rowcounter,sd,ed,errmsg
addsuccess=true
success=true
showaccount=true
dorefresh=false
if dateadjust="" then dateadjust=0
thedate=DateAdd("h",dateadjust,Now())
thedate=DateSerial(year(thedate),month(thedate),day(thedate))
Set rs=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
dsEmailHeader=getpost("dsEmailHeader")
dsEmailHeader=replace(dsEmailHeader, "<br>", "%nl%")
dsEmailHeader=replace(dsEmailHeader, "<br/>", "%nl%")
dsEmailHeader=replace(dsEmailHeader, "<br />", "%nl%")
if getpost("act")="domodify" then
	sSQL="UPDATE dropshipper SET dsEmail='"&escape_string(getpost("Email")) & "'," & _
		"dsName='"&escape_string(getpost("Name")) & "'," & _
		"dsAddress='"&escape_string(getpost("Address")) & "'," & _
		"dsCity='"&escape_string(getpost("City")) & "'," & _
		"dsState='"&escape_string(getpost("State")) & "'," & _
		"dsCountry='"&escape_string(getpost("Country")) & "'," & _
		"dsZip='"&escape_string(getpost("Zip")) & "'," & _
		"dsAction="&escape_string(getpost("dsAction")) & "," & _
		"dsEmailHeader='"&escape_string(dsEmailHeader) & "' " & _
		"WHERE dsID=" & replace(getpost("dsID"),"'","")
	ect_query(sSQL)
	dorefresh=true
elseif getpost("act")="doaddnew" then
	sSQL="INSERT INTO dropshipper (dsEmail,dsName,dsAddress,dsCity,dsState,dsCountry,dsZip,dsAction,dsEmailHeader) VALUES (" & _
		"'"&escape_string(getpost("Email")) & "'," & _
		"'"&escape_string(getpost("Name")) & "'," & _
		"'"&escape_string(getpost("Address")) & "'," & _
		"'"&escape_string(getpost("City")) & "'," & _
		"'"&escape_string(getpost("State")) & "'," & _
		"'"&escape_string(getpost("Country")) & "'," & _
		"'"&escape_string(getpost("Zip")) & "'," & _
		""&escape_string(getpost("dsAction")) & "," & _
		"'"&escape_string(dsEmailHeader) & "')"
	ect_query(sSQL)
	dorefresh=true
elseif getpost("act")="delete" then
	sSQL="DELETE FROM dropshipper WHERE dsID=" & getpost("id")
	ect_query(sSQL)
	sSQL="UPDATE products SET pDropship=0 WHERE pDropship=" & getpost("id")
	ect_query(sSQL)
	dorefresh=true
end if
if dorefresh then
	print "<script>setTimeout(function(){document.location='admindropship.asp'},2000)</script>"
end if
if dorefresh then
%>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <a href="admindropship.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br /><br />&nbsp;
                </td>
			  </tr>
			</table></td>
        </tr>
      </table>
<%
elseif getpost("act")="modify" OR getpost("act")="addnew" then
	if getpost("act")="modify" then
		dsID=getpost("id")
		sSQL="SELECT dsName,dsAddress,dsCity,dsState,dsZip,dsCountry,dsEmail,dsAction,dsEmailHeader FROM dropshipper WHERE dsID="&dsID
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			dsName=rs("dsName")
			dsAddress=rs("dsAddress")
			dsCity=rs("dsCity")
			dsState=rs("dsState")
			dsZip=rs("dsZip")
			dsCountry=rs("dsCountry")
			dsEmail=rs("dsEmail")
			dsAction=rs("dsAction")
			dsEmailHeader=trim(rs("dsEmailHeader")&"")
		end if
		rs.close
	end if
%>
<script>
<!--
function checkform(frm){
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
return (true);
}
//-->
</script>
	  <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr> 
          <td width="100%">
		    <form method="post" action="admindropship.asp" onsubmit="return checkform(this)">
		<%	if getpost("act")="modify" then %>
			<input type="hidden" name="act" value="domodify" />
		<%	else %>
			<input type="hidden" name="act" value="doaddnew" />
		<%	end if %>
			<input type="hidden" name="dsID" value="<%=dsID%>" />
			  <table width="100%" border="0" cellspacing="0" cellpadding="3">
				<tr>
				  <td width="100%" align="center" colspan="4"><strong><%=yyDSAdm%></strong><br /></td>
				</tr>
				<tr>
				  <td width="20%" align="right"><strong><%=redasterix&yyName%>:</strong></td>
				  <td width="30%" align="left"><input type="text" name="name" size="20" value="<%=dsName%>" /></td>
				  <td width="20%" align="right"><strong><%=redasterix&yyEmail%>:</strong></td>
				  <td width="30%" align="left"><input type="text" name="email" size="25" value="<%=dsEmail%>" /></td>
				</tr>
				<tr>
				  <td align="right"><strong><%=redasterix&yyAddress%>:</strong></td>
				  <td align="left"><input type="text" name="address" size="20" value="<%=dsAddress%>" /></td>
				  <td align="right"><strong><%=redasterix&yyCity%>:</strong></td>
				  <td align="left"><input type="text" name="city" size="20" value="<%=dsCity%>" /></td>
				</tr>
				<tr>
				  <td align="right"><strong><%=redasterix&yyState%>:</strong></td>
				  <td align="left"><input type="text" name="state" size="20" value="<%=dsState%>" /></td>
				  <td align="right"><strong><%=redasterix&yyCountry%>:</strong></td>
				  <td align="left"><select name="country" size="1">
<%
sub show_countries(tcountry)
	if NOT IsArray(allcountries) then
		sSQL="SELECT countryName FROM countries ORDER BY countryOrder DESC, countryName"
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
end Sub
show_countries(dsCountry)
%>
					</select>
				  </td>
				</tr>
				<tr>
				  <td align="right"><strong><%=redasterix&yyZip%>:</strong></td>
				  <td align="left"><input type="text" name="zip" size="10" value="<%=dsZip%>" /></td>
				  <td align="right"><strong><%=yyActns%>:</strong></td>
				  <td align="left"><select name="dsAction" size="1">
					<option value="0"><%=yyNoAct%></option>
					<option value="1"<% if dsAction=1 then print " selected=""selected"""%>><%=yySendEM%></option>
					</select>
				  </td>
				</tr>
				<tr>
				  <td align="right"><strong><%=yyEmlHea%>:</strong></td>
				  <td align="left" colspan="3"><input type="text" name="dsEmailHeader" size="60" value="<%=replace(dsEmailHeader, """", "&quot;")%>" /></td>
				</tr>
				<tr>
				  <td width="100%" colspan="4">&nbsp;<br />
					<span style="font-size:10px"><ul><li><%=yyDSInf%></li><li><%=yyDSIn2%></li></ul></span>
				  </td>
				</tr>
				<tr>
				  <td width="50%" align="center" colspan="4"><input type="submit" value="<%=yySubmit%>" /> <input type="reset" value="<%=yyReset%>" /> </td>
				</tr>
			  </table>
			</form>
		  </td>
        </tr>
      </table>
<%
else
	if Request("sd")="" then
		sd=DateSerial(DatePart("yyyy",thedate),DatePart("m",thedate),1)
	else
		sd=Request("sd")
	end if
	if Request("ed")="" then
		ed=thedate
	else
		ed=Request("ed")
	end if
	on error resume next
	sd=DateValue(sd)
	ed=Datevalue(ed)
	if err.number <> 0 then
		sd=DateSerial(DatePart("yyyy",thedate),DatePart("m",thedate),1)
		ed=thedate
		success=false
		errmsg=yyDatInv
	end if
	on error goto 0
	tdt=DateValue(sd)
	tdt2=DateValue(ed)+1
%>
<script>
<!--
function modrec(id) {
	document.mainform.id.value=id;
	document.mainform.act.value="modify";
	document.mainform.submit();
}
function newrec(id) {
	document.mainform.id.value=id;
	document.mainform.act.value="addnew";
	document.mainform.submit();
}
function delrec(id) {
if (confirm("<%=jscheck(yyConDel)%>\n")) {
	document.mainform.id.value=id;
	document.mainform.act.value="delete";
	document.mainform.submit();
}
}
// -->
</script>
<%		if NOT success then %>
		  <p style="text-align:center"><%="<span style=""color:#FF0000"">"&errmsg&"</span>" %></p>
<%		end if %>
			<table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr>
				<td width="100%" align="center" ><h2><%=yyDSAdm%></h2></td>
			  </tr>
			  <tr> 
                <td width="100%" align="center">
					<form method="post" action="admindropship.asp">
						<strong><%=yyAffBet%>:</strong> <input type="text" size="12" name="sd" value="<%=sd%>" /> <strong><%=yyAnd%>:</strong> <input type="text" size="12" name="ed" value="<%=ed%>" /> <input type="submit" value="Go" /><br />&nbsp;
					</form>
				</td>
			  </tr>
			  <tr> 
                <td width="100%" align="center">
					<form method="post" action="admindropship.asp">
				<p><strong><%=yyAffFrm%>:</strong> <select name="sd" size="1"><%
					For rowcounter=0 to Day(thedate)-1
						print "<option value='"&thedate-rowcounter&"'"
						if thedate-rowcounter=sd then print " selected"
						print ">"&thedate-rowcounter&"</option>"&vbCrLf
						smonth=thedate-rowcounter
					Next
					For rowcounter=1 to 12
						print "<option value='"&DateAdd("m",0-rowcounter,smonth)&"'"
						if DateAdd("m",0-rowcounter,smonth)=sd then print " selected"
						print ">"&DateAdd("m",0-rowcounter,smonth)&"</option>"&vbCrLf
					Next
				%></select> <strong><%=yyTo%>:</strong> <select name="ed" size="1"><%
					For rowcounter=0 to Day(thedate)-1
						print "<option value='"&thedate-rowcounter&"'"
						if thedate-rowcounter=ed then print " selected"
						print ">"&thedate-rowcounter&"</option>"&vbCrLf
						smonth=thedate-rowcounter
					Next
					For rowcounter=1 to 12
						print "<option value='"&DateAdd("m",0-rowcounter,smonth)&"'"
						if DateAdd("m",0-rowcounter,smonth)=ed then print " selected"
						print ">"&DateAdd("m",0-rowcounter,smonth)&"</option>"&vbCrLf
					Next
				%></select> <input type="submit" value="Go" /><br />&nbsp;</p>
					</form>
				</td>
			  </tr>
			</table>
		  </form>
<br />
		  <form name="mainform" method="post" action="admindropship.asp">
			<input type="hidden" name="id" value="xxx" />
			<input type="hidden" name="act" value="xxxxx" />
			<input type="hidden" name="ed" value="<%=DateValue(ed)%>" />
			<input type="hidden" name="sd" value="<%=DateValue(sd)%>" />
            <table width="100%" class="stackable admin-table-a sta-white">
				<tr>
				  <th class="minicell"><%=yyID%></th>
				  <th class="maincell"><%=yyName%></th>
				  <th class="maincell"><%=yyEmail%></th>
				  <th class="aright"><%=yyTotSal%></th>
				  <th class="minicell"><%=yyModify%></th>
				  <th class="minicell"><%=yyDelete%></th>
				</tr>
<%	sSQL="SELECT dsID,dsName,dsEmail,0 FROM dropshipper ORDER BY dsName"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then alldata=rs.GetRows()
	rs.close
	if NOT IsArray(alldata) then %>
				<tr>
				  <td width="100%" align="center" colspan="6"><br />&nbsp;<br /><strong><%=yyNoAff%></strong><br />&nbsp;</td>
				</tr>
<%	else
		for index=0 to UBOUND(alldata,2)
				sSQL="SELECT SUM(cartProdPrice*cartQuantity) AS sumSale FROM cart INNER JOIN products ON cart.cartProdID=products.pID WHERE pDropship=" & alldata(0,index) & " AND cartCompleted=1 AND cartDateAdded BETWEEN " & vsusdate(tdt)&" AND " & vsusdate(tdt2)
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then alldata(3,index)=rs("sumSale") else alldata(3,index)=0
				rs.close
				if mysqlserver=TRUE then
					sSQL="SELECT SUM(coPriceDiff*cartQuantity) AS sumSale FROM cartoptions INNER JOIN cart ON cartoptions.coCartID=cart.cartID INNER JOIN products ON cart.cartProdID=products.pID WHERE pDropship=" & alldata(0,index) & " AND cartCompleted=1 AND cartDateAdded BETWEEN " & vsusdate(tdt)&" AND " & vsusdate(tdt2)
				else
					sSQL="SELECT SUM(coPriceDiff*cartQuantity) AS sumSale FROM cartoptions INNER JOIN (cart INNER JOIN products ON cart.cartProdID=products.pID) ON cartoptions.coCartID=cart.cartID WHERE pDropship=" & alldata(0,index) & " AND cartCompleted=1 AND cartDateAdded BETWEEN " & vsusdate(tdt)&" AND " & vsusdate(tdt2)
				end if
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					if not isnull(rs("sumSale")) then alldata(3,index)=alldata(3,index)+rs("sumSale")
				end if
				rs.close
				if bgcolor="altdark" then bgcolor="altlight" else bgcolor="altdark" %>
				<tr class="<%=bgcolor%>">
				  <td class="minicell"><%=alldata(0,index)%></td>
				  <td><%=alldata(1,index)%></td>
				  <td><a href="mailto:<%=alldata(2,index)%>"><%=alldata(2,index)%></a></td>
				  <td class="aright"><%if NOT is_numeric(alldata(3,index)) then print "-" else print FormatEuroCurrency(alldata(3,index))%></td>
				  <td class="minicell"><input type="button" value="Modify" onclick="modrec('<%=alldata(0,index)%>')" /></td>
				  <td class="minicell"><input type="button" value="Delete" onclick="delrec('<%=alldata(0,index)%>')" /></td>
				</tr><%
		next
	end if
%>
				<tr> 
				  <td width="100%" colspan="6" align="center"><br /><input type="button" value="<%=yyAddNew%>" onclick="newrec()" /><br />&nbsp;</td>
				</tr>
				<tr> 
				  <td width="100%" colspan="6" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br /><br />&nbsp;</td>
				</tr>
			</table>
		  </form>
<%
end if
cnn.Close
set rs=nothing
set cnn=nothing
%>