<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
success=true
Set rs=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
if getpost("act")="domodify" then
	for index=0 to 70
		statusid=getpost("statusid" & index)
		if statusid<>"" then
			statPrivate = escape_string(getpost("privstatus" & index))
			statPublic = escape_string(getpost("pubstatus" & index))
			if statPublic="" then statPublic = statPrivate
			sSQL = "UPDATE orderstatus SET statPrivate='" & statPrivate & "',statPublic='" & statPublic & "'"
			if getpost("emailstatus" & index)="1" then sSQL = sSQL & ",emailstatus=1" else sSQL = sSQL & ",emailstatus=0"
			for index2=2 to adminlanguages+1
				if (adminlangsettings AND 64)=64 then sSQL = sSQL & ",statPublic" & index2 & " ='" & escape_string(getpost("pubstatus" & index & "x" & index2)) & "'"
			next
			sSQL = sSQL & " WHERE statID="&statusid
			ect_query(sSQL)
		end if
	next
	print "<meta http-equiv=""refresh"" content=""3; url=admin.asp"">"
else
	sSQL = "SELECT statID,statPrivate,statPublic,statPublic2,statPublic3,emailstatus FROM orderstatus ORDER BY statID"
	rs.open sSQL,cnn,0,1
	alldata=rs.getrows
	rs.close
end if
%>
<script>
<!--
function formvalidator(theForm){
for(index=0;index<=3;index++){
theelm=eval('theForm.privstatus'+index);
if(theelm.value == ""){
alert("Please enter a value in the field \"Private Text (Status " + (index+1) + ")\".");
theelm.focus();
return (false);
}
}
return (true);
}
//-->
</script>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
<% if getpost("act")="domodify" AND success then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
				<%=yyNoAuto%> <a href="admin.asp"><strong><%=yyClkHer%></strong></a>.<br /><br />&nbsp;
                </td>
			  </tr>
			</table></td>
        </tr>
<% elseif getpost("act")="domodify" then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><span style="color:#FF0000;font-weight:bold"><%=yyOpFai%></span><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table></td>
        </tr>
<% else
	if (adminlangsettings AND 64)<>64 then numcols=6 else numcols=6+adminlanguages
%>
        <tr>
          <td width="100%" align="center">
		  <form name="mainform" method="post" action="adminordstatus.asp" onsubmit="return formvalidator(this)">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="domodify" />
            <table width="500" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="<%=numcols%>" align="center"><strong><%=yyOSAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td align="center" valign="top" width="50"><strong>&nbsp;</strong></td>
				<td align="center" valign="top"><strong>&nbsp;</strong></td>
				<td align="center" valign="top"><strong><%=yyPrTxt%></strong></td>
				<td align="center" valign="top"><strong><%=yyPubTxt%></strong></td>
<%	for index2=2 to adminlanguages+1
		if (adminlangsettings AND 64)=64 then print "<td align=""center"" valign=""top""><strong>" & yyPubTxt & " " & index2 & "</strong></td>"
	next %>
				<td align="center" valign="top"><strong><%=replace(yySendEM," ","&nbsp;")%></strong></td>
				<td align="center" valign="top" width="50"><strong>&nbsp;</strong></td>
			  </tr>
<%	for rowcounter=0 to UBOUND(alldata,2)
		if alldata(0,rowcounter)=4 then %>
			  <tr> 
                <td width="100%" colspan="<%=numcols%>" align="center"><span style="font-size:10px"><%=yyOSExp1%></span></td>
			  </tr>
<%		end if %>
			  <tr>
				<td align="center" valign="top"><strong>&nbsp;&nbsp;&nbsp;&nbsp;</strong></td>
				<td align="right"><input type="hidden" name="statusid<%=rowcounter%>" value="<%=alldata(0,rowcounter) %>" /><%=yyStatus%>&nbsp;<%=rowcounter%>:</td>
				<td align="center"><input type="text" size="30" name="privstatus<%=rowcounter%>" value="<%=htmlspecials(alldata(1,rowcounter)) %>" /></td>
				<td align="center"><input type="text" size="30" name="pubstatus<%=rowcounter%>" value="<%=htmlspecials(alldata(2,rowcounter)) %>" /></td>
<%	for index2=2 to adminlanguages+1
		if (adminlangsettings AND 64)=64 then print "<td align=""center""><input type=""text"" size=""20"" name=""pubstatus" & rowcounter & "x" & index2 & """ value=""" & htmlspecials(alldata(1 + index2,rowcounter)) & """ /></td>"
	next %>
				<td align="center"><input type="checkbox" name="emailstatus<%=rowcounter%>" value="1" <%if alldata(5,rowcounter)<>0 then print "checked=""checked"" "%>/></td>
				<td align="center" valign="top"><strong>&nbsp;&nbsp;&nbsp;&nbsp;</strong></td>
			  </tr>
<%	next %>
			  <tr> 
                <td width="100%" colspan="<%=numcols%>" align="center"><input type="submit" value="<%=yySubmit%>" /></td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="<%=numcols%>" align="center"><br /><a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
			  </tr>
            </table>
		  </form>
		  </td>
        </tr>

<%
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>
      </table>