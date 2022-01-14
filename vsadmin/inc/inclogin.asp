<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
success=true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
errmsg=""
sub addtopasswordhistory(loginid,hpw)
	sSQL2="INSERT INTO passwordhistory (liID,pwhPwd,datePWChanged) VALUES ("&loginid&",'"& hpw&"',"&vsusdatetime(Now())&")"
	ect_query(sSQL2)
	sSQL2 = "SELECT COUNT(*) AS pwhcount FROM passwordhistory WHERE liID="&loginid
	rs.open sSQL2,cnn,0,1
		pwhcount=rs("pwhcount")
	rs.close
	if pwhcount>4 then
		sSQL2="SELECT "&IIfVr(mysqlserver<>TRUE,"TOP "&(pwhcount-4),"")&" pwhID FROM passwordhistory WHERE liID="&loginid&" ORDER BY datePWChanged"&IIfVs(mysqlserver=TRUE," LIMIT 0,"&(pwhcount-4))
		rs.open sSQL2,cnn,0,1
		do while NOT rs.EOF
			ect_query("DELETE FROM passwordhistory WHERE pwhID="&rs("pwhID"))
			rs.movenext
		loop
		rs.close
	end if
end sub
if getpost("posted")="1" then
	if getpost("act")="changeprimary" then
		if getpost("pass") <> getpost("pass2") then
			success = FALSE
			errmsg=yyNoMat
		elseif getpost("pass")="changeme" AND padssfeatures=TRUE then
			success = FALSE
			errmsg=yyLIErr1
		elseif ectdemostore<>TRUE then
			hashedpw=dohashpw(getpost("pass"))
			if padssfeatures=TRUE then
				sSQL="SELECT pwhID FROM passwordhistory WHERE liID="&SESSION("loginid")&" AND pwhPwd='"& hashedpw&"'"
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					success=FALSE
					errmsg="You cannot use the same password as any of your last 4 previous passwords."
				end if
				rs.close
			end if
			if success then
				if Mid(SESSION("loggedonpermissions"),20,1)<>"X" then
					sSQL = "UPDATE adminlogin SET adminLoginName='"&getpost("user")&"'"
					if getpost("pass")<>"" then
						sSQL = sSQL & ",adminLoginPassword='"& hashedpw&"',adminLoginLastChange=" & vsusdate(Now())
						SESSION("mustchangepw")=empty
					end if
					sSQL = sSQL & " WHERE adminLoginID="&SESSION("loginid")
				else
					sSQL = "UPDATE admin SET adminUser='"&getpost("user")&"'"
					if getpost("pass")<>"" then
						sSQL = sSQL & ",adminPassword='"& hashedpw&"',adminPWLastChange=" & vsusdate(Now())
						SESSION("mustchangepw")=empty
					end if
					sSQL = sSQL & " WHERE adminID=1"
				end if
				ect_query(sSQL)
				call addtopasswordhistory(SESSION("loginid"),hashedpw)
				print "<meta http-equiv=""refresh"" content=""1; url=admin.asp"">"
			end if
		end if
		call logevent(SESSION("loginuser"),"CHANGEPASSWORD",success,"adminlogin.asp",getpost("user"))
	elseif Mid(SESSION("loggedonpermissions"),20,1)<>"X" then
		success = FALSE
		errmsg="No Permissions"
	elseif getpost("act")="doaddnew" OR getpost("act")="domodify" then
		permissions = ""
		if getpost("main")="ON" then permissions = permissions & "X" else permissions = permissions & "O"
		if getpost("orders")="ON" then permissions = permissions & "X" else permissions = permissions & "O"
		if getpost("payprov")="ON" then permissions = permissions & "X" else permissions = permissions & "O"
		if getpost("affiliate")="ON" then permissions = permissions & "X" else permissions = permissions & "O"
		if getpost("clientlogin")="ON" then permissions = permissions & "X" else permissions = permissions & "O"
		if getpost("products")="ON" then permissions = permissions & "X" else permissions = permissions & "O"
		if getpost("categories")="ON" then permissions = permissions & "X" else permissions = permissions & "O"
		if getpost("discounts")="ON" then permissions = permissions & "X" else permissions = permissions & "O"
		if getpost("regions")="ON" then permissions = permissions & "X" else permissions = permissions & "O"
		if getpost("shipping")="ON" then permissions = permissions & "X" else permissions = permissions & "O"
		if getpost("ordstatus")="ON" then permissions = permissions & "X" else permissions = permissions & "O"
		if getpost("dropship")="ON" then permissions = permissions & "X" else permissions = permissions & "O"
		if getpost("ipblock")="ON" then permissions = permissions & "X" else permissions = permissions & "O"
		if getpost("maillist")="ON" then permissions = permissions & "X" else permissions = permissions & "O"
		if getpost("statistics")="ON" then permissions = permissions & "X" else permissions = permissions & "O"
		if getpost("ratings")="ON" then permissions = permissions & "X" else permissions = permissions & "O"
		if getpost("contentregion")="ON" then permissions = permissions & "X" else permissions = permissions & "O"
		if getpost("act")="doaddnew" then
			sSQL="SELECT adminloginid FROM adminlogin WHERE adminloginname='" & escape_string(getpost("user")) & "'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then success=FALSE
			rs.close
			sSQL="SELECT adminID FROM admin WHERE adminUser='" & escape_string(getpost("user")) & "'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then success=FALSE
			rs.close
			if NOT success then errmsg="That login is already in use. Please choose another."
			sSQL = "INSERT INTO adminlogin (adminloginname,adminloginpassword,adminloginpermissions) VALUES ('"&escape_string(getpost("user"))&"','"&escape_string(dohashpw(getpost("pass")))&"','"&permissions&"')"
		else
			sSQL = "UPDATE adminlogin SET adminLoginLock=0,adminloginname='"&escape_string(getpost("user"))&"',"
			if getpost("pass")<>"" then
				hashedpw=dohashpw(getpost("pass"))
				sSQL = sSQL & "adminloginpassword='"&escape_string(hashedpw)&"',"
				if padssfeatures=TRUE then
					rs.open "SELECT pwhID FROM passwordhistory WHERE liID="&SESSION("loginid")&" AND pwhPwd='"& hashedpw&"'",cnn,0,1
					if NOT rs.EOF then
						success=FALSE
						errmsg="You cannot use the same password as any of your last 4 previous passwords."
					end if
					rs.close
				end if
				if success then call addtopasswordhistory(SESSION("loginid"),hashedpw)
			end if
			sSQL = sSQL & "adminloginpermissions='"&permissions&"' WHERE adminloginid="&getpost("id")
		end if
		if success then
			ect_query(sSQL)
			print "<meta http-equiv=""refresh"" content=""2; url=adminlogin.asp"">"
		end if
		call logevent(SESSION("loginuser"),"ALTERLOGIN",success,"adminlogin.asp",getpost("user"))
	elseif getpost("act")="delete" then
		sSQL = "DELETE FROM adminlogin WHERE adminloginid="&getpost("id")
		ect_query(sSQL)
		print "<meta http-equiv=""refresh"" content=""2; url=adminlogin.asp"">"
	end if
end if
%>
<script>
/* <![CDATA[ */
function checkform(frm){
	if(frm.pass.value!=""||frm.pass2.value!=""){
		if(frm.pass.value!=frm.pass2.value){
			alert("Your password does not match the confirm password.");
			frm.pass.focus();
			return(false);
		}
<%	if padssfeatures=TRUE then %>
		if(frm.pass.value.length<7){
			alert("Your password must be at least 7 characters.");
			frm.pass.focus();
			return(false);
		}
		if(frm.pass.value=="changeme"){
			alert("That password is illegal.");
			frm.pass.focus();
			return(false);
		}
		var regexn = /[0-9]/;
		var regexa = /[a-z]/i;
		if(!(regexn.test(frm.pass.value)&&regexa.test(frm.pass.value))){
			alert("Your password must contain at least one numeric and one alphabetic character.");
			frm.pass.focus();
			return(false);
		}
<%	end if %>
	}
	return(true);
}
function modrec(id){
	document.mainform.id.value = id;
	document.mainform.act.value = "modify";
	document.mainform.submit();
}
function clone(id){
	document.mainform.id.value = id;
	document.mainform.act.value = "clone";
	document.mainform.submit();
}
function newrec(){
	document.mainform.act.value = "addnew";
	document.mainform.submit();
}
function delrec(id){
if(confirm("<%=jscheck(yyConDel)%>\n")) {
	document.mainform.id.value = id;
	document.mainform.act.value = "delete";
	document.mainform.submit();
}
}
/* ]]> */
</script>
<%
if getpost("act")="addnew" OR getpost("act")="modify" then
	if getpost("act")="modify" AND is_numeric(getpost("id")) then
		sSQL = "SELECT adminloginid,adminloginname,adminloginpassword,adminloginpermissions FROM adminlogin WHERE adminloginid=" & getpost("id")
		rs.open sSQL,cnn,0,1
			adminloginname=rs("adminloginname")
			adminloginpassword=""
			permissions=rs("adminloginpermissions")
		rs.close
	else
		adminloginname=""
		adminloginpassword=""
		permissions=""
	end if
%>
		  <form method="post" action="adminlogin.asp" onsubmit="return checkform(this)">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="<%=IIfVr(getpost("act")="modify", "domodify", "doaddnew") %>" />
			<input type="hidden" name="id" value="<%=getpost("id")%>" />
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr>
				<td colspan="2" align="center"><br /><strong><%=yyPlDfPr%></strong></td>
			  </tr>
			  <tr> 
				<td width="50%" align="right"><strong><%=redasterix&yyUname%>:</strong></td>
				<td align="left"><input type="text" name="user" size="20" value="<%=adminloginname%>" /></td>
			  </tr>
			  <tr> 
				<td align="right"><strong><%=IIfVr(getpost("act")="modify",yyReset&" "&yyPass,redasterix&yyPass)%>:</strong></td>
				<td align="left"><input type="password" name="pass" size="20" value="" autocomplete="off" /></td>
			  </tr>
			  <tr> 
				<td align="right"><strong><%=yyPassCo%>: </strong></td>
				<td align="left"><input type="password" name="pass2" size="20" value="" autocomplete="off" /></td>
			  </tr>
			  <tr>
				<td colspan="2" align="center"><hr /></td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyLLMain%></strong></td><td align="left"><input type="checkbox" name="main" value="ON" <% if Mid(permissions,1,1)="X" then print "checked=""checked""" %>/></td>
			  </tr>
			  <tr> 
				<td align="right"><strong><%=yyLLOrds%></strong></td><td align="left"><input type="checkbox" name="orders" value="ON" <% if Mid(permissions,2,1)="X" then print "checked=""checked""" %>/></td>
			  </tr>
			  <tr> 
				<td align="right"><strong><%=yyLLPayP%></strong></td><td align="left"><input type="checkbox" name="payprov" value="ON" <% if Mid(permissions,3,1)="X" then print "checked=""checked""" %>/></td>
			  </tr>
			  <tr> 
				<td align="right"><strong><%=yyLLAffl%></strong></td><td align="left"><input type="checkbox" name="affiliate" value="ON" <% if Mid(permissions,4,1)="X" then print "checked=""checked""" %>/></td>
			  </tr>
			  <tr> 
				<td align="right"><strong><%=yyLLClLo%></strong></td><td align="left"><input type="checkbox" name="clientlogin" value="ON" <% if Mid(permissions,5,1)="X" then print "checked=""checked""" %>/></td>
			  </tr>
			  <tr> 
				<td align="right"><strong><%=yyLLProA & " + " & yyLLProO%></strong></td><td align="left"><input type="checkbox" name="products" value="ON" <% if Mid(permissions,6,1)="X" then print "checked=""checked""" %>/></td>
			  </tr>
			  <tr> 
				<td align="right"><strong><%=yyLLCats & " + " & yySeaCri%></strong></td><td align="left"><input type="checkbox" name="categories" value="ON" <% if Mid(permissions,7,1)="X" then print "checked=""checked""" %>/></td>
			  </tr>
			  <tr> 
				<td align="right"><strong><%=yyLLDisc & " + " & yyLLQuan & " + " & yyLLGftC%></strong></td><td align="left"><input type="checkbox" name="discounts" value="ON" <% if Mid(permissions,8,1)="X" then print "checked=""checked""" %>/></td>
			  </tr>
			  <tr> 
				<td align="right"><strong><%=yyLLStat & " + " & yyLLCoun%></strong></td><td align="left"><input type="checkbox" name="regions" value="ON" <% if Mid(permissions,9,1)="X" then print "checked=""checked""" %>/></td>
			  </tr>
			  <tr> 
				<td align="right"><strong><%=yyLLZone & " + " & yyLLShpM%></strong></td><td align="left"><input type="checkbox" name="shipping" value="ON" <% if Mid(permissions,10,1)="X" then print "checked=""checked""" %>/></td>
			  </tr>
			  <tr> 
				<td align="right"><strong><%=yyLLOrSt %></strong></td><td align="left"><input type="checkbox" name="ordstatus" value="ON" <% if Mid(permissions,11,1)="X" then print "checked=""checked""" %>/></td>
			  </tr>
			  <tr> 
				<td align="right"><strong><%=yyDrShpr %></strong></td><td align="left"><input type="checkbox" name="dropship" value="ON" <% if Mid(permissions,12,1)="X" then print "checked=""checked""" %>/></td>
			  </tr>
			  <tr> 
				<td align="right"><strong><%=yyIPBlock %></strong></td><td align="left"><input type="checkbox" name="ipblock" value="ON" <% if Mid(permissions,13,1)="X" then print "checked=""checked""" %>/></td>
			  </tr>
			  <tr> 
				<td align="right"><strong><%=yyMaLiMa %></strong></td><td align="left"><input type="checkbox" name="maillist" value="ON" <% if Mid(permissions,14,1)="X" then print "checked=""checked""" %>/></td>
			  </tr>
			  <tr> 
				<td align="right"><strong><%=yyStatis %></strong></td><td align="left"><input type="checkbox" name="statistics" value="ON" <% if Mid(permissions,15,1)="X" then print "checked=""checked""" %>/></td>
			  </tr>
			  <tr> 
				<td align="right"><strong><%=yyRating %></strong></td><td align="left"><input type="checkbox" name="ratings" value="ON" <% if Mid(permissions,16,1)="X" then print "checked=""checked""" %>/></td>
			  </tr>
			  <tr> 
				<td align="right"><strong><%=yyContReg %></strong></td><td align="left"><input type="checkbox" name="contentregion" value="ON" <% if Mid(permissions,17,1)="X" then print "checked=""checked""" %>/></td>
			  </tr>
			  <tr> 
				<td colspan="2" align="center"><br /><input type="submit" value="<%=yySubmit%>" />  <input type="reset" value="<%=yyReset%>" /></td>
			  </tr>
			</table>
		  </form>
<%
elseif getpost("posted")="1" AND success then %>
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
				<td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
						<%=yyNoAuto%><a href="admin.asp"><strong><%=yyClkHer%></strong></a>.<br /><br />&nbsp;<br />&nbsp;</td>
			  </tr>
			</table>
<%
elseif getpost("posted")="1" then %>
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><span style="color:#FF0000;font-weight:bold"><%=yyOpFai%></span><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a><br /><br />&nbsp;<br />&nbsp;</td>
			  </tr>
			</table>
<%
else
	if Mid(SESSION("loggedonpermissions"),20,1)="X" then
		sSQL = "SELECT adminUser FROM admin WHERE adminID=1"
		rs.open sSQL,cnn,0,1
		theuser=rs("adminUser")
		rs.close
	else
		sSQL = "SELECT adminLoginName FROM adminlogin WHERE adminLoginID="&SESSION("loginid")
		rs.open sSQL,cnn,0,1
		theuser=rs("adminLoginName")
		rs.close
	end if
	if SESSION("mustchangepw")<>"" then
		if SESSION("mustchangepw")="A" then errmsg=yyLIErr1&"<br />"&errmsg
		if SESSION("mustchangepw")="B" then errmsg=yyLIErr2&"<br />"&errmsg
		success=FALSE
	end if
%>
		  <form method="post" action="adminlogin.asp" onsubmit="return checkform(this)">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="changeprimary" />
			<table style="border:1px dotted #194C7F" width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
				<td colspan="2" align="center"><br /><strong><%=yyNewUN%></strong></td>
			  </tr>
<%	if not success then %>
			  <tr> 
				<td colspan="2" align="center"><br /><span style="color:#FF0000"><%=errmsg%></span></td>
			  </tr>
<%	end if %>
			  <tr> 
				<td width="50%" align="right"><strong><%=redasterix&yyUname%>: </strong></td>
				<td width="50%" align="left"><input type="text" name="user" size="20" value="<%=theuser%>" /></td>
			  </tr>
			  <tr> 
				<td align="right"><strong><%=yyReset&" "&yyPass%>: </strong></td>
				<td align="left"><input type="password" name="pass" size="20" value="" autocomplete="off" /></td>
			  </tr>
			  <tr> 
				<td align="right"><strong><%=yyPassCo%>: </strong></td>
				<td align="left"><input type="password" name="pass2" size="20" value="" autocomplete="off" /></td>
			  </tr>
			  <tr> 
				<td colspan="2" align="center"><br /><input type="submit" value="<%=yySubmit%>" /></td>
			  </tr>
			</table>
		  </form>
<%	if Mid(SESSION("loggedonpermissions"),20,1)="X" then %>
		  <form method="post" name="mainform" action="adminlogin.asp">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="" />
			<input type="hidden" name="id" value="" />
			<table style="border:1px dotted #194C7F" width="100%" border="0" cellspacing="0" cellpadding="3">
<%		sSQL = "SELECT adminloginid,adminloginname,adminloginpermissions FROM adminlogin"
		rs.open sSQL,cnn,0,1
		if rs.EOF then
%>
			  <tr> 
				<td colspan="20" align="center"><br /><strong><%=yyNoSecL%></strong></td>
			  </tr>
<%		else %>
			  <tr> 
				<td colspan="20" align="center"><br /><strong><%=yySecLog%></strong></td>
			  </tr>
			  <tr> 
				<td><strong><%=yyLiName%></strong></td>
				<td><acronym title="<%=yyLLMain%>">MAI</acronym></td>
				<td><acronym title="<%=yyLLOrds%>">ORD</acronym></td>
				<td><acronym title="<%=yyLLPayP%>">PAY</acronym></td>
				<td><acronym title="<%=yyLLAffl%>">AFF</acronym></td>
				<td><acronym title="<%=yyLLClLo%>">LOG</acronym></td>
				<td><acronym title="<%=yyLLProA & " + " & yyLLProO%>">PRO</acronym></td>
				<td><acronym title="<%=yyLLCats & " + " & yySeaCri%>">CAT</acronym></td>
				<td><acronym title="<%=yyLLDisc & " + " & yyLLQuan%>">DSC</acronym></td>
				<td><acronym title="<%=yyLLStat & " + " & yyLLCoun%>">REG</acronym></td>
				<td><acronym title="<%=yyLLZone & " + " & yyLLShpM%>">SHI</acronym></td>
				<td><acronym title="<%=yyLLOrSt %>">ORS</acronym></td>
				<td><acronym title="<%=yyDrShpr %>">DRP</acronym></td>
				<td><acronym title="<%=yyIPBlock %>">IPB</acronym></td>
				<td><acronym title="<%=yyMaLiMa %>">MLM</acronym></td>
				<td><acronym title="<%=yyStatis %>">STA</acronym></td>
				<td><acronym title="<%=yyRating %>">RAT</acronym></td>
				<td><acronym title="<%=yyContReg %>">CRG</acronym></td>
				<td><strong><%=yyModify%></strong></td>
				<td><strong><%=yyDelete%></strong></td>
			  </tr>
<%			do while NOT rs.EOF
				if bgcolor="altdark" then bgcolor="altlight" else bgcolor="altdark" %>
			  <tr class="<%=bgcolor%>">
				<td align="left"> &nbsp; <%=rs("adminloginname")%></td>
<%				for index=1 to 17
					print "<td>"
					if Mid(rs("adminloginpermissions"),index,1)="X" then print "X" else print "&nbsp;"
					print "</td>"
				next %>
				<td align="center"><input type="button" value="<%=yyModify%>" onclick="modrec('<%=rs("adminloginid")%>')" /></td>
				<td align="center"><input type="button" value="<%=yyDelete%>" onclick="delrec('<%=rs("adminloginid")%>')" /></td>
			  </tr>
<%				rs.Movenext
			loop
		end if %>
			  <tr> 
				<td colspan="20" align="center">&nbsp;<br /><input type="button" value="<%=yyNewSec%>" onclick="newrec()" /><br />&nbsp;</td>
			  </tr>
			</table>
		  </form>
<%	end if
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>