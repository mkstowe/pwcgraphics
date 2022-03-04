<%
if Session("loggedon") <> "virtualstore" then response.end
Dim sSQL,rs,alldata,alladmin,success,cnn,rowcounter,netnav,errmsg,subCats
netnav = true
if instr(Request.ServerVariables("HTTP_USER_AGENT"), "compatible") > 0 OR instr(Request.ServerVariables("HTTP_USER_AGENT"), "Gecko") > 0 then netnav = false
function atb(size)
	if netnav then
		atb = CInt(size / 2 + 1)
	else
		atb = size
	end if
end function
success=true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
sSQL = "SELECT adminSubCats FROM admin"
rs.Open sSQL,cnn,0,1
subCats=(Int(rs("adminSubCats"))=1)
rs.Close
sSQL = ""
if request.form("posted")="1" then
	if request.form("act")="delete" then
		sSQL = "DELETE FROM topsections WHERE tsID=" & request.form("id")
	elseif request.form("act")="domodify" then
		sSQL = "UPDATE topsections SET tsName='"&replace(request.form("secname"),"'","''")&"',tsDescription='"&replace(request.form("secdesc"),"'","''")&"',tsImage='"&replace(request.form("secimage"),"'","''")&"' WHERE tsID="&Request.Form("id")
	elseif request.form("act")="doaddnew" then
		sSQL = "INSERT INTO topsections (tsName,tsDescription,tsImage) VALUES ('"&replace(request.form("secname"),"'","''")&"','"&replace(request.form("secdesc"),"'","''")&"','"&replace(request.form("secimage"),"'","''")&"')"
	end if
	on error resume next
	cnn.Execute(sSQL)
	if err.number<>0 then
		success=false
		errmsg = "There was an error writing to the database.<br>"
		if err.number = -2147467259 then
			errmsg = errmsg & "Your database does not have write permissions."
		else
			errmsg = errmsg & err.description
		end if
	else
		response.write "<meta http-equiv=""refresh"" content=""3; url=admintopcats.asp"">"
	end if
	on error goto 0
end if
%>
<script Language="JavaScript">
<!--
function formvalidator(theForm)
{
  if (theForm.secname.value == "")
  {
    alert("Please enter a value in the field \"Top Category Name\".");
    theForm.secname.focus();
    return (false);
  }
  if (theForm.secdesc.value.length > 255)
  {
    alert("A maximum of 255 characters are allowed in the field \"Top Category Description\".");
    theForm.secdesc.focus();
    return (false);
  }
  return (true);
}
//-->
</script>
      <table border="0" cellspacing="<%=maintablespacing%>" cellpadding="<%=maintablepadding%>" width="<%=maintablewidth%>" bgcolor="<%=maintablebg%>" align="center">
<% if request.form("posted")="1" AND request.form("act")="modify" then 
		sSQL = "SELECT tsID,tsName,tsDescription,tsImage FROM topsections WHERE tsID="&Request.Form("id")
		rs.Open sSQL,cnn,0,1
		alldata=rs.getrows
		rs.Close
%>
        <tr>
		        <form name="mainform" method="POST" action="admintopcats.asp" onSubmit="return formvalidator(this)">
                  <td width="100%">
			<input type="hidden" name="posted" value="1">
			<input type="hidden" name="act" value="domodify">
			<input type="hidden" name="id" value="<%=Request.Form("id")%>">
            <table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br><b>Use this page to update your top categories.</b><br>&nbsp;</td>
			  </tr>
			  <tr>
				<td width="40%" align="center" valign="top"><b>Top Category Name</b>
                </td>
				<td width="60%" align="center" valign="top"><b>Top Category Description</b>
                </td>
			  </tr>
			  <tr>
				<td width="40%" align="center" valign="top"><input type="text" name="secname" size="<%=atb(30)%>" value="<%=replace(alldata(1,0),"""","&quot;")%>">
                </td>
				<td width="60%" rowspan="3" align="center" valign="top"><textarea name="secdesc" cols="<%=atb(30)%>" rows="5" wrap=virtual><%=alldata(2,0)%></textarea>
                </td>
			  </tr>
			  <tr>
				<td align="center" valign="top"><b>Top Category Image</b>
                </td>
			  </tr>
			  <tr>
				<td align="center" valign="top"><input type="text" name="secimage" size="<%=atb(30)%>" value="<%if NOT IsNull(alldata(3,0)) then response.write replace(alldata(3,0),"""","&quot;")%>">
                </td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="2" align="center"><br><input type="submit" value="Submit"><br>&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="2" align="center"><br>
                          <a href="admin.asp"><b>Admin Home</b></a><br>
                          &nbsp;</td>
			  </tr>
            </table></td>
		  </form>
        </tr>
<% elseif request.form("posted")="1" AND request.form("act")="addnew" then %>
        <tr>
		        <form name="mainform" method="POST" action="admintopcats.asp" onSubmit="return formvalidator(this)">
                  <td width="100%">
			<input type="hidden" name="posted" value="1">
			<input type="hidden" name="act" value="doaddnew">
            <table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br><b>Enter your new top category details here.</b><br>&nbsp;</td>
			  </tr>
			  <tr>
				<td width="40%" align="center" valign="top"><b>Top Category Name</b>
                </td>
				<td width="60%" align="center" valign="top"><b>Top Category Description</b>
                </td>
			  </tr>
			  <tr>
				<td width="40%" align="center" valign="top"><input type="text" name="secname" size="<%=atb(30)%>" value="">
                </td>
				<td width="60%" rowspan="3" align="center" valign="top"><textarea name="secdesc" cols="<%=atb(30)%>" rows="5" wrap=virtual></textarea>
                </td>
			  </tr>
			  <tr>
				<td align="center" valign="top"><b>Top Category Image</b>
                </td>
			  </tr>
			  <tr>
				<td align="center" valign="top"><input type="text" name="secimage" size="<%=atb(30)%>" value="">
                </td>
			  </tr>
			  <tr>
                <td width="100%" colspan="2" align="center"><br><input type="submit" value="Submit"><br>&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="2" align="center"><br>
                          <a href="admin.asp"><b>Admin Home</b></a><br>
                          &nbsp;</td>
			  </tr>
            </table></td>
		  </form>
        </tr>
<% elseif request.form("posted")="1" AND success then %>
        <tr>
          <td width="100%">
			<table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br><b>Update Successful !</b><br><br>You will now be forwarded  
				to the top categories admin page.<br><br>
                        If that does not happen automatically then please <A href="admintopcats.asp"><b>click 
                        here</b></a>.<br>
                        <br>
				<img src="../images/clearpixel.gif" width="350" height="3">
                </td>
			  </tr>
			</table></td>
        </tr>
<% elseif request.form("posted")="1" then %>
        <tr>
          <td width="100%">
			<table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br><font color="#FF0000"><b>Operation Failed ! !</b></font><br><br><%=errmsg%><br><br>
				<a href="javascript:history.go(-1)"><b>Click here to go back.</b></a></td>
			  </tr>
			</table></td>
        </tr>
<% else 
		sSQL = "SELECT tsID,tsName,tsDescription FROM topsections"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then alldata=rs.getrows
		rs.Close
%>
<script language="JavaScript">
<!--
function modrec(id) {
	document.mainform.id.value = id;
	document.mainform.act.value = "modify";
	document.mainform.submit();
}
function newrec(id) {
	document.mainform.id.value = id;
	document.mainform.act.value = "addnew";
	document.mainform.submit();
}
function delrec(id) {
cmsg = "Sure you want to delete this record?\n"
if (confirm(cmsg)) {
	document.mainform.id.value = id;
	document.mainform.act.value = "delete";
	document.mainform.submit();
}
}
// -->
</script>
        <tr>
		        <form name="mainform" method="POST" action="admintopcats.asp">
                  <td width="100%">
			<input type="hidden" name="posted" value="1">
			<input type="hidden" name="act" value="xxxxx">
			<input type="hidden" name="id" value="xxxxx">
            <table width="<%=innertablewidth%>" border="0" cellspacing="<%=innertablespacing%>" cellpadding="<%=innertablepadding%>" bgcolor="<%=innertablebg%>">
			  <tr> 
                <td width="100%" colspan="3" align="center"><br><b>Use this page to update your top categories.</b><br>&nbsp;</td>
			  </tr>

			  <tr> 
                <td width="100%" colspan="3" align="left"><br>Your admin settings <b>are<% if NOT subCats then response.write " NOT"%></b> set to use top categories and sub categories.<br>
				<% if NOT subCats then %>
				This means that these settings will be ignored.<br>
                          <% end if %> To change this please <a href="adminmain.asp">click 
                          here</a>.<br>
                          &nbsp;</td>
			  </tr>
			  <tr>
				<td width="60%" align="center" valign="top"><b>Top Category Name</b>
                </td>
				<td width="20%" align="center" valign="top"><b>Modify</b>
                </td>
				<td width="20%" align="center" valign="top"><b>Delete</b>
                </td>
			  </tr>
<%
	if IsArray(alldata) then
		for rowcounter=0 to UBOUND(alldata,2)
%>
			  <tr>
				<td width="60%" align="center" valign="top"><%=alldata(1,rowcounter)%>
                </td>
				<td width="20%" align="center" valign="top"><input type=button name=modify value="Modify" onClick="modrec('<%=alldata(0,rowcounter)%>')">
                </td>
				<td width="20%" align="center" valign="top"><input type=button name=delete value="Delete" onClick="delrec('<%=alldata(0,rowcounter)%>')">
                </td>
			  </tr>
<%		next
	else
%>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br><b>There are no top categories currently configured.<br>&nbsp;</td>
			  </tr>
<%
	end if
%>
			  <tr> 
                <td width="100%" colspan="3" align="center"><br><b>Click here to add a new top category</b>&nbsp;&nbsp;<input type="button" value="New Top Category" onClick="newrec()"><br>&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="3" align="center"><br>
                          <a href="admin.asp"><b>Admin Home</b></a><br>
				<img src="../images/clearpixel.gif" width="350" height="3"></td>
			  </tr>
            </table></td>
		  </form>
        </tr>
<% end if
cnn.Close
%>
      </table>