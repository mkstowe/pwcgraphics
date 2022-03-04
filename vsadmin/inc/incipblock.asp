<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
success=TRUE
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
if getpost("posted")="1" then
	for each objItem in request.form
		if left(objItem,4) = "idxx" then
			ip1 = ip2long(getpost(objItem))
			if getpost(replace(objItem,"xx","yy"))<>"" then
				ip2 = ip2long(getpost(replace(objItem,"xx","yy")))
			else
				ip2 = 0
			end if
			if ip1 <> -1 AND ip2 <> -1 then
				sSQL = "UPDATE ipblocking SET dcip1=" & ip1 & ",dcip2=" & ip2 & " WHERE dcid=" & mid(objItem,5)
				ect_query(sSQL)
			end if
		elseif left(objItem,7) = "newidxx" then
			ip1 = ip2long(getpost(objItem))
			if getpost(replace(objItem,"xx","yy"))<>"" then
				ip2 = ip2long(getpost(replace(objItem,"xx","yy")))
			else
				ip2 = 0
			end if
			if ip1 <> -1 AND ip2 <> -1 then
				sSQL = "INSERT INTO ipblocking (dcip1,dcip2) VALUES (" & ip1 & "," & ip2 & ")"
				ect_query(sSQL)
			end if
		elseif left(objItem,5) = "delip" then
			sSQL = "DELETE FROM ipblocking WHERE dcid=" & mid(objItem,6)
			ect_query(sSQL)
		elseif left(objItem,5) = "delss" then
			sSQL = "DELETE FROM multibuyblock WHERE ssdenyid=" & mid(objItem,6)
			ect_query(sSQL)
		end if
	next
	if success then print "<meta http-equiv=""refresh"" content=""1; url=adminipblock.asp"">"
end if
%>
<%	if getpost("posted") = "1" AND success then %>
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <a href="adminipblock.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br /><br />&nbsp;
                </td>
			  </tr>
			</table>
<%	elseif getpost("posted") = "1" then %>
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><span style="color:#FF0000;font-weight:bold"><%=yyErrUpd%></span><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table>
<%	else %>
<script>
<!--
var currinputid=1;
function addnewinput(tid){
	if(tid==currinputid){
		currinputid++;
		var node = document.createElement("DIV");
		node.className='ecttablerow';
		node.innerHTML='<div><input type="text" size="15" name="newidxx' + currinputid + '" onfocus="addnewinput(' + currinputid + ')" /></div>' +
				'<div><input type="text" size="15" name="newidyy' + currinputid + '" onfocus="addnewinput(' + currinputid + ')" /></div>' +
				'<div>n/a</div>';
		document.getElementById("ipblocktable").appendChild(node);
	}
}
function updatemultiblock(tobj){
document.location='adminipblock.asp?multiblock='+tobj[tobj.selectedIndex].value;
}
function updatecartblock(tobj){
document.location='adminipblock.asp?cartblock='+tobj[tobj.selectedIndex].value;
}
//-->
</script>
<%
	if is_numeric(getget("multiblock")) then
		sSQL="UPDATE admin SET blockMultiPurchase='" & escape_string(getget("multiblock")) & "' WHERE adminID=1"
		blockmultipurchase=int(getget("multiblock"))
		cnn.execute(sSQL)
	end if
	if is_numeric(getget("cartblock")) then
		sSQL="UPDATE admin SET blockMaxCartAdds='" & escape_string(getget("cartblock")) & "' WHERE adminID=1"
		blockmaxcartadds=int(getget("cartblock"))
		cnn.execute(sSQL)
	end if
%>
	<h2><%=yyUsIPBl%></h2>
	
	<div style="padding:10px;border:1px solid lightgrey;margin-bottom:10px">
		<h3>Excessive Cart Additions</h3>
		<div style="padding:20px 0">
			<div style="float:left;padding-left:40px"><select size="1" name="multicartadd" onchange="updatecartblock(this)"><option value="0">Disabled</option>
<%
			hasdisplayed=blockmaxcartadds=0
			sub domcblockoption(tind)
				if tind=blockmaxcartadds then hasdisplayed=TRUE
				if tind>blockmaxcartadds AND NOT hasdisplayed then domcblockoption(blockmaxcartadds)
				print "<option value=""" & tind & """" & IIfVs(blockmaxcartadds=tind," selected=""selected""") & ">" & tind & "</option>"
			end sub
			for index=70 to 250 step 10
				domcblockoption(index)
			next
%>
			</select></div>
			<div style="float:left;padding:3px">You can limit the maximum items a customer can add to cart.</div>
		</div>
		<div style="clear:both;padding:40px">Hackers can use the trick of adding lots of different items to the cart to hog the server resources in what is called a Denial of Service (DOS) attack. You can set a limit to the items that can be added to the cart 
		before it starts to look like suspicious activity. You don't want to set this value too low however as you don't want to block your best customers. Blocked IP's will appear below.
		</div>
	</div>

<form name="mainform" method="post" action="adminipblock.asp">
	<input type="hidden" name="posted" value="1" />
	<div style="padding:10px;border:1px solid lightgrey;margin-bottom:10px">
		<h3>IP Blocking</h3>
		<div id="ipblocktable" class="ecttable" style="max-width:600px;margin:20px auto">
			<div class="ecttablerow ecttablehead">
				<div><%=yySinIP%></div>
				<div><%=yyLasIP%></div>
				<div><%=yyDelete%></div>
			</div><%
		sSQL="SELECT dcid,dcip1,dcip2 FROM ipblocking ORDER BY dcip1"
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF %>
			<div class="ecttablerow">
				<div><input type="text" size="15" name="idxx<%=rs("dcid")%>" value="<%=long2ip(int(rs("dcip1")))%>" /></div>
				<div><input type="text" size="15" name="idyy<%=rs("dcid")%>" value="<% if rs("dcip2")<>0 then print long2ip(int(rs("dcip2")))%>" /></div>
				<div><input type="checkbox" name="delip<%=rs("dcid")%>"></div>
			</div>
<%			rs.movenext
		loop
		rs.close %>
			<div class="ecttablerow">
				<div><input type="text" size="15" name="newidxx1" onfocus="addnewinput(1)" /></div>
				<div><input type="text" size="15" name="newidyy1" onfocus="addnewinput(1)" /></div>
				<div>n/a</div>
			</div>
		</div>
		<div style="padding:20px 0;text-align:center"><input type="submit" value="<%=yySubmit%>" />&nbsp;&nbsp;<input type="reset" value="<%=yyReset%>" /></div>
	</div>
</form>

<form name="mainform" method="post" action="adminipblock.asp">
	<input type="hidden" name="posted" value="1" />
	<div style="padding:10px;border:1px solid lightgrey;margin-bottom:10px">
		<h3>Multi Purchase Blocking</h3>
		<div style="padding:20px 0">
			<div style="float:left;padding-left:40px"><select size="1" name="multipurchase" onchange="updatemultiblock(this)"><option value="0">Disabled</option>
<%
			hasdisplayed=(blockmultipurchase=0)
			sub dompblockoption(tind)
				if tind=blockmultipurchase then hasdisplayed=TRUE
				if tind>blockmultipurchase AND NOT hasdisplayed then call dompblockoption(blockmultipurchase)
				print "<option value=""" & tind & """" & IIfVs(blockmultipurchase=tind," selected=""selected""") & ">" & tind & "</option>"
			end sub
			for index=10 to 50 step 5
				call dompblockoption(index)
			next
%>
			</select></div>
			<div style="float:left;padding:3px">You can limit the number of transactions a customer can make in a 24 hour period with this setting.</div>
		</div>
<%			sSQL="SELECT ssdenyid,ssdenyip,sstimesaccess,lastaccess FROM multibuyblock WHERE sstimesaccess>=" & blockmultipurchase & " ORDER BY ssdenyip"
			rs.open sSQL,cnn,0,1
			if rs.EOF then
				print "<div class=""nosearchresults"">" & yyNoIPBl & "</div>"
			else %>
		<div class="ecttable" style="max-width:600px;margin:20px auto">
			<div class="ecttablerow ecttablehead">
				<div>IP Address</div>
				<div>Checkout Attempts</div>
				<div>Delete</div>
			</div>
<%				do while NOT rs.EOF %>
			<div class="ecttablerow">
				<div><%=rs("ssdenyip")%></div>
				<div><%=(rs("sstimesaccess")+1)%></div>
				<div><input type="checkbox" name="delss<%=rs("ssdenyid")%>"></div>
			</div>
<%					rs.movenext
				loop %>
		</div>
		<div style="padding:20px 0;text-align:center"><input type="submit" value="<%=yySubmit%>" />&nbsp;&nbsp;<input type="reset" value="<%=yyReset%>" /></div>
<%			end if %>
	</div>
</form>
	
	<div class="adminhome"><a href="admin.php"><%=yyAdmHom%></a></div>
<%
end if
%>