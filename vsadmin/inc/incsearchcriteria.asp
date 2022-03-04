<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
Dim sSQL,rs,cnn,success,showaccount,addsuccess,alldata,index,allcountries,rowcounter,sd,ed,errmsg
addsuccess = TRUE
success = TRUE
maxcatsperpage = 500
showaccount = TRUE
dorefresh = FALSE
set rs=Server.CreateObject("ADODB.RecordSet")
set rs2=Server.CreateObject("ADODB.RecordSet")
set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
if defaultcatimages="" then defaultcatimages="images/"
if getget("act")="repair" then
	sSQL = "SELECT mSCpID,mSCscID FROM (multisearchcriteria LEFT JOIN products ON multisearchcriteria.mSCpID=products.pID) LEFT JOIN searchcriteria ON multisearchcriteria.mSCscID=searchcriteria.scID WHERE pID IS NULL OR scID IS NULL"
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		ect_query("DELETE FROM multisearchcriteria WHERE mSCpID='"&rs("mSCpID")&"' AND mSCscID="&rs("mSCscID"))
		rs.movenext
	loop
	rs.close
elseif getpost("act")="newattribute" AND is_numeric(getpost("group")) then
	haveuniqueindex=FALSE
	uniqueindex=1
	do while NOT haveuniqueindex
		rs.open "SELECT scID FROM searchcriteria WHERE scID="&uniqueindex,cnn,0,1
		if rs.EOF then haveuniqueindex=TRUE else uniqueindex=uniqueindex+1
		rs.close
	loop
	sSQL="SELECT MAX(scOrder) AS maxorder FROM searchcriteria WHERE scGroup="&getpost("group")
	rs.open sSQL,cnn,0,1
	if rs.EOF then maxorder=1 else maxorder=rs("maxorder")
	rs.close
	if isnull(maxorder) then
		maxorder=1
	else
		if maxorder>0 then maxorder=maxorder+1
	end if
	sSQL="INSERT INTO searchcriteria (scID,scWorkingName,scName,scName2,scName3,scGroup,scOrder) VALUES (" & uniqueindex & "," & _
		"'" & escape_string(IIfVr(getpost("newwn")<>"",getpost("newwn"),getpost("newname"))) & "'," & _
		"'" & escape_string(getpost("newname")) & "'," & _
		"'" & escape_string(IIfVr(getpost("newname2")<>"",getpost("newname2"),getpost("newname"))) & "'," & _
		"'" & escape_string(IIfVr(getpost("newname3")<>"",getpost("newname3"),getpost("newname"))) & "'," & _
		getpost("group") & "," & maxorder & ")"
	ect_query(sSQL)
	dorefresh=TRUE
elseif getpost("act")="dodiscounts" then
	sSQL="INSERT INTO cpnassign (cpaCpnID,cpaType,cpaAssignment) VALUES ("&getpost("assdisc")&",3,'"&getpost("id")&"')"
	ect_query(sSQL)
	dorefresh=TRUE
elseif getpost("act")="deletedisc" then
	sSQL="DELETE FROM cpnassign WHERE cpaType=3 AND cpaID="&getpost("id")
	ect_query(sSQL)
	dorefresh=TRUE
elseif getpost("act")="newgroup" then
	haveuniqueindex=FALSE
	uniqueindex=0
	do while NOT haveuniqueindex
		rs.open "SELECT scgID FROM searchcriteriagroup WHERE scgID="&uniqueindex,cnn,0,1
		if rs.EOF then haveuniqueindex=TRUE else uniqueindex=uniqueindex+1
		rs.close
	loop
	sSQL="SELECT MAX(scgOrder) AS maxorder FROM searchcriteriagroup"
	rs.open sSQL,cnn,0,1
	if rs.EOF then maxorder=1 else maxorder=rs("maxorder")
	rs.close
	if isnull(maxorder) then maxorder=1 else maxorder=maxorder+1
	sSQL="INSERT INTO searchcriteriagroup (scgID,scgWorkingName,scgTitle,scgTitle2,scgTitle3,scgOrder) VALUES (" & uniqueindex & "," & _
		"'" & escape_string(IIfVr(getpost("newwn")<>"",getpost("newwn"),getpost("newname"))) & "'," & _
		"'" & escape_string(getpost("newname")) & "'," & _
		"'" & escape_string(IIfVr(getpost("newname2")<>"",getpost("newname2"),getpost("newname"))) & "'," & _
		"'" & escape_string(IIfVr(getpost("newname3")<>"",getpost("newname3"),getpost("newname"))) & "'," & _
		maxorder & ")"
	ect_query(sSQL)
	dorefresh=TRUE
elseif getpost("act")="changepos" then
	theid = int(getpost("id"))
	neworder = int(getpost("newval"))-1
	rc=0
	if getget("act")="modifyatts" then
		if getpost("alphabetically")="1" then
			sSQL = "UPDATE searchcriteria SET scOrder=0 WHERE scGroup="&getpost("group")
			ect_query(sSQL)
		else
			sSQL="SELECT scID,scOrder FROM searchcriteria WHERE scGroup IN ("&getpost("group")&") ORDER BY scOrder"
			rs.open sSQL,cnn,0,1
			do while NOT rs.EOF
				if rs("scID")=theid then
					sSQL = "UPDATE searchcriteria SET scOrder="&neworder&" WHERE scID="&theid
				else
					sSQL = "UPDATE searchcriteria SET scOrder="&IIfVr(rc<neworder,rc,rc+1)&" WHERE scID="&rs("scID")
				end if
				ect_query(sSQL)
				rc=rc+1
				rs.movenext
			loop
			rs.close
		end if
	else
		sSQL="SELECT scgID,scgOrder FROM searchcriteriagroup ORDER BY scgOrder"
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			if rs("scgID")=theid then
				sSQL = "UPDATE searchcriteriagroup SET scgOrder="&neworder&" WHERE scgID="&theid
			else
				sSQL = "UPDATE searchcriteriagroup SET scgOrder="&IIfVr(rc<neworder,rc,rc+1)&" WHERE scgID="&rs("scgID")
			end if
			ect_query(sSQL)
			rc=rc+1
			rs.movenext
		loop
		rs.close
	end if
	print "<meta http-equiv=""refresh"" content=""0; url=adminsearchcriteria.asp?pg="& getpost("pg") & IIfVs(getget("act")="modifyatts","&act=modifyatts&id="&getpost("group")) &""">"
elseif getpost("act")="domodifygroup" then
	scworkingname=getpost("scworkingname")
	if scworkingname="" then scworkingname=getpost("scname")
	sSQL = "UPDATE searchcriteriagroup SET " & _
		"scgTitle='"&escape_string(getpost("scname")) & "',"
	for index=2 to adminlanguages+1
		if (adminlangsettings AND 131072)=131072 then
			sSQL = sSQL & "scgTitle"&index&"='"&escape_string(getpost("scname" & index)) & "',"
		end if
	next
	sSQL = sSQL & "scgWorkingName='"&escape_string(scworkingname) & "' " & _
		"WHERE scgID=" & replace(getpost("scID"),"'","")
	ect_query(sSQL)
	dorefresh=TRUE
elseif getpost("act")="domodify" then
	scworkingname=getpost("scworkingname")
	if scworkingname="" then scworkingname=getpost("scname")
	sSQL = "UPDATE searchcriteria SET scName='"&escape_string(getpost("scname")) & "'," & _
		"scURL='"&escape_string(getpost("scurl")) & "'," & _
		"scDescription='"&escape_string(getpost("scDescription")) & "'," & _
		"scHeader='"&escape_string(getpost("scHeader")) & "',"
	for index=2 to adminlanguages+1
		if (adminlangsettings AND 131072)=131072 then sSQL = sSQL & "scName"&index&"='"&escape_string(getpost("scname" & index)) & "',"
		if (adminlangsettings AND 8192)=8192 then sSQL = sSQL & "scURL"&index&"='"&escape_string(getpost("scurl" & index)) & "',"
		if (adminlangsettings AND 16384)=16384 then sSQL = sSQL & "scDescription"&index&"='"&escape_string(getpost("scDescription" & index)) & "',"
		if (adminlangsettings AND 524288)=524288 then sSQL=sSQL & "scHeader"&index&"='"& escape_string(getpost("scHeader"&index)) &"',"
	next
	sSQL = sSQL & "scWorkingName='"&escape_string(scworkingname) & "'," & _
		"scNotes='"&escape_string(getpost("scnotes")) & "'," & _
		"scLogo='"&escape_string(getpost("sclogo")) & "'," & _
		"scEmail='"&escape_string(getpost("scemail")) & "' " & _
		"WHERE scID=" & replace(getpost("scID"),"'","")
	ect_query(sSQL)
	dorefresh=TRUE
elseif getpost("act")="doaddnew" then
	' Not used as you have to create it first.
elseif getpost("act")="delete" then
	sSQL = "DELETE FROM multisearchcriteria WHERE mSCscID=" & getpost("id")
	ect_query(sSQL)
	sSQL = "DELETE FROM searchcriteria WHERE scID=" & getpost("id")
	ect_query(sSQL)
	dorefresh=TRUE
elseif getpost("act")="deletegroup" then
	sSQL = "DELETE FROM multisearchcriteria WHERE mSCscID IN (SELECT scID FROM searchcriteria WHERE scGroup=" & getpost("id") & ")"
	ect_query(sSQL)
	sSQL = "DELETE FROM searchcriteria WHERE scGroup=" & getpost("id")
	ect_query(sSQL)
	sSQL = "DELETE FROM searchcriteriagroup WHERE scgID=" & getpost("id")
	ect_query(sSQL)
	dorefresh=TRUE
end if
if dorefresh then
	if getpost("act")="newattribute" OR getpost("act")="domodify" OR getpost("act")="delete" then
		print "<meta http-equiv=""refresh"" content=""1; url=adminsearchcriteria.asp?act=modifyatts&id="&getpost("group")&""">"
	else
		print "<meta http-equiv=""refresh"" content=""1; url=adminsearchcriteria.asp"">"
	end if
end if
if dorefresh OR getpost("act")="changepos" OR getpost("act")="newattribute" then
%>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <a href="adminsearchcriteria.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br /><br />&nbsp;
                </td>
			  </tr>
			</table></td>
        </tr>
      </table>
<%
elseif (getpost("act")="modify" OR getpost("act")="addnew") AND is_numeric(getpost("group")) then
	Dim scaName(3),scURL(3)
	sSQL = "SELECT scgTitle FROM searchcriteriagroup WHERE scgID="&getpost("group")
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then groupname=rs("scgTitle") else groupname=""
	rs.close
	if getpost("act")="modify" then
		scID=getpost("id")
		sSQL = "SELECT scName,scName2,scName3,scWorkingName,scGroup,scLogo,scURL,scURL2,scURL3,scEmail,scNotes,scHeader FROM searchcriteria WHERE scID="&scID
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			scworkingname = rs("scWorkingName")
			scgroup = rs("scGroup")
			for index=1 to 3
				scaName(index) = rs("scName"&IIfVr(index=1,"",index))
				scURL(index) = rs("scURL"&IIfVr(index=1,"",index))
			next
			scEmail = rs("scEmail")
			scLogo = rs("scLogo")
			scNotes = rs("scNotes")
			scHeader=rs("scHeader")
		end if
		rs.close
	end if
		if htmleditor="ckeditor" then %>
<script src="ckeditor/ckeditor.js"></script>
<%		end if %>
<script>
<!--
function checkform(frm){
if(frm.name.value==""){
	alert("<%=jscheck(yyPlsEntr&" """&yyName)%>\".");
	frm.scname.focus();
	return (false);
}
return (true);
}
function uploadimage(imfield){
	var winwid=360; var winhei=240;
	var prnttext = '<html><head><link rel="stylesheet" type="text/css" href="adminstyle.css"/></head><body>\n';
	prnttext += '<form name="mainform" method="post" action="doupload.asp?defimagepath=<%=defaultcatimages%>" enctype="multipart/form-data">';
	prnttext += '<input type="hidden" name="defimagepath" value="<%=defaultcatimages%>" />';
	prnttext += '<input type="hidden" name="imagefield" value="'+imfield+'" />';
	prnttext += '<table border="0" cellspacing="1" cellpadding="3" width="100%">';
	prnttext += '<tr><td align="center" colspan="2">&nbsp;<br /><%=replace(yyUplIma,"'","\'")%><br />&nbsp;</td></tr>';
	prnttext += '<tr><td align="center" colspan="2"><%=replace(yyPlsSUp,"'","\'")%><br />&nbsp;</td></tr>';
	prnttext += '<tr><td align="right"><%=replace(yyLocIma,"'","\'")%>:</td><td><input type="file" name="imagefile" /></td></tr>';
	prnttext += '<tr><td colspan="2" align="center">&nbsp;<br /><input type="submit" value="<%=replace(yySubmit,"'","\'")%>" /></td></tr>';
	prnttext += '</table></form>';
	prnttext += '<p align="center"><a href="javascript:window.close()"><%=replace(yyClsWin,"'","\'")%></a></p>';
	prnttext += '</body></'+'html>';
	scrwid=screen.width; scrhei=screen.height;
	var newwin = window.open("","printlicense",'menubar=no,scrollbars=yes,width='+winwid+',height='+winhei+',left='+((scrwid-winwid)/2)+',top=100,directories=no,location=no,resizable=yes,status=no,toolbar=no');
	newwin.document.open();
	newwin.document.write(prnttext);
	newwin.document.close();
}
function expandckeditor(divid){
	document.getElementById(divid).style.border='none';
	document.getElementById(divid).style.padding=0;
}
//-->
</script>
	  <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr> 
          <td width="100%">
		    <form method="post" action="adminsearchcriteria.asp" onsubmit="return checkform(this)">
			<input type="hidden" name="group" value="<%=getpost("group")%>" />
		<%	if getpost("act")="modify" then %>
			<input type="hidden" name="act" value="domodify" />
		<%	else %>
			<input type="hidden" name="act" value="doaddnew" />
		<%	end if %>
			<input type="hidden" name="scID" value="<%=scID%>" />
			  <table width="100%" border="0" cellspacing="0" cellpadding="3">
				<tr>
				  <td width="100%" align="center" colspan="2"><strong><%=yySeaCri&" : Group - " & groupname%></strong><br />&nbsp;</td>
				</tr>
				<tr>
				  <td align="right"><%=redasterix&yyName%>:</td>
				  <td align="left"><input type="text" name="scname" size="30" value="<%=htmlspecials(scaName(1))%>" /></td>
				</tr>
<%			for index=2 to adminlanguages+1
				if (adminlangsettings AND 131072)=131072 then %>
				<tr>
				  <td align="right"><%=redasterix&yyName%> <%=index%></td>
				  <td align="left"><input type="text" name="scname<%=index%>" size="30" value="<%=htmlspecials(scaName(index))%>" />
				  </td>
				</tr>
<%				end if
			next %>
				<tr>
				  <td align="right"><%=yyWrkNam%>:</td>
				  <td align="left"><input type="text" name="scworkingname" size="30" value="<%=htmlspecials(scworkingname)%>" /></td>
				</tr>
				<tr>
				  <td align="right"><%=yyEmail%>:</td>
				  <td align="left"><input type="text" name="scemail" size="25" value="<%=htmlspecials(scEmail)%>" /></td>
				</tr>
				<tr>
				  <td align="right">Attribute Logo:</td>
				  <td align="left"><input type="text" name="sclogo" id="sclogo" size="30" value="<%=htmlspecials(scLogo)%>" /> <input type="button" name="smallimup" value="..." onclick="uploadimage('sclogo')" /></td>
				</tr>
<%			if getpost("act")="modify" then
				sSQL = "SELECT scDescription FROM searchcriteria WHERE scID="&scID
				rs.open sSQL,cnn,0,1
				scDescription = rs("scDescription")
				rs.close
			else
				scDescription = ""
			end if %>
				<tr>
				  <td align="right">Description</td>
				  <td align="left">
<%	if htmleditor="froala" then print "<div id=""descdiv"" class=""htmleditorcontainer"">" %>
					<textarea name="scDescription" id="scDescription" cols="38" rows="8" wrap=virtual><%=htmlspecials(scDescription)%></textarea>
<%	if htmleditor="froala" then print "</div>" %>
				  </td>
				</tr>
<%			for index=2 to adminlanguages+1
				if (adminlangsettings AND 16384)=16384 then
					if getpost("act")="modify" then
						sSQL = "SELECT scDescription"&index&" FROM searchcriteria WHERE scID="&scID
						rs.open sSQL,cnn,0,1
							scDescription = rs("scDescription"&index)
						rs.close
					else
						scDescription = ""
					end if
%>
				<tr>
				  <td align="right">Description <%=index%></td>
				  <td align="left">
<%	if htmleditor="froala" then print "<div id=""descdiv"&index&""" class=""htmleditorcontainer"">" %>
					<textarea name="scDescription<%=index%>" id="scDescription<%=index%>" cols="38" rows="8"><%=htmlspecials(scDescription)%></textarea>
<%	if htmleditor="froala" then print "</div>" %>
				  </td>
				</tr>
<%				end if
			next %>
				<tr>
				  <td align="right">Static Page URL (Optional)</td>
				  <td align="left"><input type="text" name="scurl" size="50" value="<%=htmlspecials(scURL(1))%>" /></td>
				</tr>
<%			for index=2 to adminlanguages+1
				if (adminlangsettings AND 8192)=8192 then %>
				<tr>
				  <td align="right">Static Page URL <%=index%> (Optional)</td>
				  <td align="left"><input type="text" name="scurl<%=index%>" size="50" value="<%=htmlspecials(scURL(index))%>" /></td>
				</tr>
<%				end if
			next %>
<%			if getpost("group")="0" then %>
			  <tr>
				<td align="right">Attribute Header:</td>
				<td>
<%	if htmleditor="froala" then print "<div id=""headdiv"" class=""htmleditorcontainer"">" %>
					<textarea name="scHeader" id="scHeader" cols="48" rows="8"><%=scHeader%></textarea>
<%	if htmleditor="froala" then print "</div>" %>
				  </td>
			  </tr>
<%				for index=2 to adminlanguages+1
					if (adminlangsettings AND 524288)=524288 then
						if getpost("act")<>"addnew" then
							sSQL="SELECT scHeader" & index & " FROM searchcriteria WHERE scID="&getpost("id")
							rs2.Open sSQL,cnn,0,1
							scHeader=rs2("scHeader" & index)
							rs2.Close
						end if
					%>
			  <tr>
				<td align="right"><%="Attribute Header" & " " & index%>:</td>
                <td>
<%	if htmleditor="froala" then print "<div id=""headdiv"&index&""" class=""htmleditorcontainer"">" %>
					<textarea name="scHeader<%=index%>" id="scHeader<%=index%>" cols="55" rows="8"><%=htmlspecials(scHeader)%></textarea>
<%	if htmleditor="froala" then print "</div>" %>
				  </td>
			  </tr>
<%					end if
				next
			end if %>
				<tr>
				  <td align="right">Notes:</td>
				  <td align="left"><textarea name="scnotes" cols="50" rows="8" wrap=virtual><%=htmlspecials(scNotes)%></textarea></td>
				</tr>
				<tr>
				  <td align="center" colspan="2"><input type="submit" value="<%=yySubmit%>" /> <input type="reset" value="<%=yyReset%>" /> </td>
				</tr>
				<tr><td align="center" colspan="2">&nbsp;</td></tr>
			  </table>
			</form>
		  </td>
        </tr>
      </table>
<%	if htmleditor="ckeditor" then
		print "<script>"
		pathtovsadmin=request.servervariables("URL")
		slashpos=instrrev(pathtovsadmin, "/")
		if slashpos>0 then pathtovsadmin=left(pathtovsadmin, slashpos-1)
		print "function loadeditors(){"
		streditor="var scDescription=CKEDITOR.replace('scDescription',{extraPlugins : 'stylesheetparser,autogrow',autoGrow_maxHeight : 800,removePlugins : 'resize', toolbarStartupExpanded : false, toolbar : 'Basic', filebrowserBrowseUrl :'ckeditor/filemanager/browser/default/browser.html?Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserImageBrowseUrl : 'ckeditor/filemanager/browser/default/browser.html?Type=Image&Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserFlashBrowseUrl :'ckeditor/filemanager/browser/default/browser.html?Type=Flash&Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=File',filebrowserImageUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=Image',filebrowserFlashUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=Flash'});" & vbCrLf
		streditor=streditor & "scDescription.on('instanceReady',function(event){var myToolbar='Basic';event.editor.on( 'beforeMaximize', function(){if(event.editor.getCommand('maximize').state == CKEDITOR.TRISTATE_ON && myToolbar != 'Basic'){scDescription.setToolbar('Basic');myToolbar='Basic';scDescription.execCommand('toolbarCollapse');}else if(event.editor.getCommand('maximize').state == CKEDITOR.TRISTATE_OFF && myToolbar != 'Full'){scDescription.setToolbar('Full');myToolbar='Full';scDescription.execCommand('toolbarCollapse');}});event.editor.on('contentDom', function(e){event.editor.document.on('blur', function(){if(!scDescription.isToolbarCollapsed){scDescription.execCommand('toolbarCollapse');scDescription.isToolbarCollapsed=true;}});event.editor.document.on('focus',function(){if(scDescription.isToolbarCollapsed){scDescription.execCommand('toolbarCollapse');scDescription.isToolbarCollapsed=false;}});});scDescription.fire('contentDom');scDescription.isToolbarCollapsed=true;});"
		print streditor
		if getpost("group")="0" then print replace(streditor,"scDescription","scHeader")
		for index=2 to adminlanguages+1
			if getpost("group")="0" AND (adminlangsettings AND 524288)=524288 then print replace(streditor, "scDescription", "scHeader" & index)
			if (adminlangsettings AND 16384)=16384 then print replace(streditor, "scDescription", "scDescription" & index)
		next
		print "}window.onload=function(){loadeditors();}"
		print "</script>" & vbCrLf
	elseif htmleditor="froala" then
		call displayfroalaeditor("scDescription","Description",".on('froalaEditor.focus',function(){expandckeditor(""descdiv"");})",FALSE,FALSE,1,FALSE)
		call displayfroalaeditor("scHeader","Attribute Header",".on('froalaEditor.focus',function(){expandckeditor(""headdiv"");})",FALSE,FALSE,1,FALSE)
		for index=2 to adminlanguages+1
			if (adminlangsettings AND 524288)=524288 then call displayfroalaeditor("scHeader"&index,"Attribute Header"&" (language "&index&")",".on('froalaEditor.focus',function(){expandckeditor(""headdiv"&index&""");})",FALSE,FALSE,1,FALSE)
			if (adminlangsettings AND 16384)=16384 then call displayfroalaeditor("scDescription"&index,"Description"&" (language "&index&")",".on('froalaEditor.focus',function(){expandckeditor(""descdiv"&index&""");})",FALSE,FALSE,1,FALSE)
		next
	end if
elseif getpost("act")="modifygroup" then
	dim scagName(3)
	scID=getpost("id")
	sSQL = "SELECT scgID,scgTitle,scgTitle2,scgTitle3,scgWorkingName FROM searchcriteriagroup WHERE scgID="&scID
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		scworkingname = rs("scgWorkingName")
		for index=1 to 3
			scagName(index) = rs("scgTitle"&IIfVr(index=1,"",index))
		next
	end if
	rs.close
%>
<script>
<!--
function checkform(frm){
if(frm.scname.value==""){
	alert("<%=jscheck(yyPlsEntr&" """&yyName)%>\".");
	frm.scname.focus();
	return(false);
}
<%			for index=2 to adminlanguages+1
				if (adminlangsettings AND 131072)=131072 then %>
if(frm.scname<%=index%>.value==""){
	alert("<%=jscheck(yyPlsEntr&" """&yyName&" "&index)%>\".");
	frm.scname<%=index%>.focus();
	return(false);
}
<%				end if
			next %>
return (true);
}
//-->
</script>
<h2><%=yyAdmSeC%></h2>
	  <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr> 
          <td width="100%">
		    <form method="post" action="adminsearchcriteria.asp" onsubmit="return checkform(this)">
		<%	if getpost("act")="modifygroup" then %>
			<input type="hidden" name="act" value="domodifygroup" />
		<%	else %>
			<input type="hidden" name="act" value="doaddnewgroup" />
		<%	end if %>
			<input type="hidden" name="scID" value="<%=scID%>" />
			  <table width="100%" border="0" cellspacing="0" cellpadding="3">
				<tr>
				  <td width="100%" align="center" colspan="2"><strong><%=yyPrAtGr%></strong><br />&nbsp;</td>
				</tr>
				<tr>
				  <td align="right"><strong><%=redasterix&yyName%>:</strong></td>
				  <td align="left"><input type="text" name="scname" size="30" value="<%=scagName(1)%>" /></td>
				</tr>
<%			for index=2 to adminlanguages+1
				if (adminlangsettings AND 131072)=131072 then %>
				<tr>
				  <td align="right"><strong><%=redasterix&yyName%> <%=index%></strong></td>
				  <td align="left"><input type="text" name="scname<%=index%>" size="30" value="<%=htmlspecials(scagName(index))%>" />
				  </td>
				</tr>
<%				end if
			next %>
				<tr>
				  <td width="50%" align="right"><strong><%=yyWrkNam%>:</strong></td>
				  <td align="left"><input type="text" name="scworkingname" size="30" value="<%=scworkingname%>" /></td>
				</tr>
				<tr><td align="center" colspan="2">&nbsp;</td></tr>
				<tr>
				  <td align="center" colspan="2"><input type="submit" value="<%=yySubmit%>" /> <input type="reset" value="<%=yyReset%>" /> </td>
				</tr>
				<tr><td align="center" colspan="2">&nbsp;</td></tr>
			  </table>
			</form>
		  </td>
        </tr>
      </table>
<%
elseif getpost("act")="discounts" then
	sSQL="SELECT scWorkingName FROM searchcriteria WHERE scID="&getpost("id")
	rs.open sSQL,cnn,0,1
	thisname=rs("scWorkingName")
	rs.close
	alldata=""
	sSQL="SELECT cpaID,cpaCpnID,cpnWorkingName,cpnSitewide,cpnEndDate,cpnType FROM cpnassign INNER JOIN coupons ON cpnassign.cpaCpnID=coupons.cpnID WHERE cpaType=3 AND cpaAssignment='" & getpost("id") & "'"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then alldata=rs.GetRows
	rs.close
	alldata2=""
	tdt=Date()
	sSQL="SELECT cpnID,cpnWorkingName,cpnSitewide FROM coupons WHERE (cpnSitewide=0 OR cpnSitewide=3) AND cpnEndDate>=" & vsusdate(tdt)
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then alldata2=rs.GetRows
	rs.close
%>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
		  <td width="100%">
<script>
/* <![CDATA[ */
function delrec(id) {
if(confirm("<%=jscheck(yyConAss)%>\n")){
	document.mainform.id.value=id;
	document.mainform.act.value="deletedisc";
	document.mainform.submit();
}
}
/* ]]> */
</script>
		  <form name="mainform" method="post" action="adminsearchcriteria.asp">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="dodiscounts" />
			<input type="hidden" name="id" value="<%=getpost("id")%>" />
			<input type="hidden" name="pg" value="<%=getpost("pg")%>" />
            <table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="4" align="center"><strong>Assign Discounts to Attribute &quot;<%=thisname%>&quot;.</strong><br />&nbsp;</td>
			  </tr>
<%	gotone=false
	if IsArray(alldata2) then
		thestr="<tr><td colspan='4' align='center'>"&yyAsDsCp&": <select name='assdisc' size='1'>"
		for index=0 to UBOUND(alldata2,2)
			alreadyassign=false
			if IsArray(alldata) then
				for index2=0 to UBOUND(alldata,2)
					if alldata2(0,index)=alldata(1,index2) then alreadyassign=true
				next
			end if
			if NOT alreadyassign then
				thestr=thestr & "<option value='"&alldata2(0,index)&"'>"&alldata2(1,index)&"</option>" & vbCrLf
				gotone=true
			end if
		next
		thestr=thestr & "</select> <input type=""submit"" value="""&yyGo&""" /></td></tr>"
	end if
	if gotone then
		print thestr
	else
%>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br /><strong><%=yyNoDis%></td>
			  </tr>
<%
	end if
	if IsArray(alldata) then
%>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br /><strong><%=yyCurDis%> &quot;<%=thisname%>&quot;.</strong><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td><strong><%=yyWrkNam%></strong></td>
				<td><strong><%=yyDisTyp%></strong></td>
				<td><strong><%=yyExpire%></strong></td>
				<td align="center"><strong><%=yyDelete%></strong></td>
			  </tr>
<%
		for index=0 to UBOUND(alldata,2)
			prefont=""
			postfont=""
			if alldata(3,index)=1 OR alldata(4,index)-Date() < 0 then
				prefont="<span style=""color:#FF0000"">"
				postfont="</span>"
			end if
%>
			  <tr> 
                <td><%=prefont & alldata(2,index) & postfont %></td>
				<td><%	if alldata(5,index)=0 then
							print prefont & yyFrSShp & postfont
						elseif alldata(5,index)=1 then
							print prefont & yyFlatDs & postfont
						elseif alldata(5,index)=2 then
							print prefont & yyPerDis & postfont
						end if %></td>
				<td><%	if alldata(4,index)=DateSerial(3000,1,1) then
							print yyNever
						elseif alldata(4,index)-Date() < 0 then
							print "<span style=""color:#FF0000"">"&yyExpird&"</span>"
						else
							print prefont & alldata(4,index) & postfont
						end if %></td>
				<td align="center"><input type="button" name="discount" value="Delete Assignment" onclick="delrec('<%=alldata(0,index)%>')" /></td>
			  </tr>
<%
		next
	else
%>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br /><strong><%=yyNoAss%></strong></td>
			  </tr>
<%
	end if
%>
			  <tr>
                <td width="100%" colspan="4" align="center"><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
			  </tr>
            </table>
		  </form>
		  </td>
        </tr>
	  </table>
<%
elseif getrequest("act")="modifyatts" AND is_numeric(getrequest("id")) then %>
<script>
<!--
function popsel(x,theid,grpid){
	if(x.length>1) return;
	for(index=theid-1; index>0; index--){
		var y=document.createElement('option');
		y.text=index;
		y.value=index;
		var sel=x.options[0];
		try{
			x.add(y, sel); // FF etc
		}
		catch(ex){
			x.add(y, 0); // IE
		}
	}
	for(index=theid+1; index<=totrows; index++){
		var y=document.createElement('option');
		y.text=index;
		y.value=index;
		try{
			x.add(y, null); // FF etc
		}
		catch(ex){
			x.add(y); // IE
		}
	}
}
function chi(id,obj){
	document.mainform.action="adminsearchcriteria.asp?act=modifyatts";
	document.mainform.newval.value = obj.selectedIndex+1;
	document.mainform.id.value = id;
	document.mainform.act.value = "changepos";
	document.mainform.alphabetically.value = "0";
	document.mainform.submit();
}
function sortalphabetically(){
	if(confirm("<%=jscheck(yySureCa)%>")){
		document.mainform.action="adminsearchcriteria.asp?act=modifyatts";
		document.mainform.newval.value=0;
		document.mainform.id.value=0;
		document.mainform.act.value = "changepos";
		document.mainform.alphabetically.value = "1";
		document.mainform.submit();
	}
}
function modrec(id){
	document.mainform.id.value = id;
	document.mainform.act.value = "modify";
	document.mainform.submit();
}
function newrec(id){
	document.mainform.id.value = id;
	document.mainform.act.value = "addnew";
	document.mainform.submit();
}
function dsc(id) {
	document.mainform.action="adminsearchcriteria.asp";
	document.mainform.id.value=id;
	document.mainform.act.value="discounts";
	document.mainform.submit();
}
function delrec(id){
if(confirm("<%=jscheck(yyConDel)%>\n")){
	document.mainform.id.value = id;
	document.mainform.act.value = "delete";
	document.mainform.submit();
}
}
function addnewcriteria(){
	if(document.getElementById('newname').value==''){
		alert("<%=jscheck(yyPlsEntr&" """&yyName)%>\".");
		document.getElementById('newname').focus();
	}else{
		document.mainform.action="adminsearchcriteria.asp?group=<%=getrequest("id")%>";
		document.mainform.id.value='';
		document.mainform.act.value="newattribute";
		document.mainform.submit();
	}
}
// -->
</script>
<h2><%=yyAdmSeC%></h2>
<%	sSQL = "SELECT scgWorkingName FROM searchcriteriagroup WHERE scgID="&getrequest("id")
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then thegroupname=rs("scgWorkingName") else thegroupname=""
	rs.close
%>
	  <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr> 
          <td width="100%" align="center">
		  <div style="text-align:center"><span style="font-weight:bold">Modify Attributes for Group:</span> <%=thegroupname%><br />&nbsp;</div>
			<form name="mainform" method="post" action="adminsearchcriteria.asp">
			<input type="hidden" name="id" value="xxx" />
			<input type="hidden" name="act" value="xxxxx" />
			<input type="hidden" name="selectedq" value="1" />
			<input type="hidden" name="newval" value="1" />
			<input type="hidden" name="alphabetically" value="" />
			<input type="hidden" name="group" value="<%=getrequest("id")%>" />
			  <table width="80%" border="0" cellspacing="0" cellpadding="2">
				<tr>
				  <th width="5%">&nbsp;<strong><%=yyOrder%></strong></th>
				  <th width="10%"><strong><%=yyID%></strong></th>
				  <th align="left"><strong><%=yyWrkNam%></strong></th>
				  <th align="left"><strong><%=yyName%></strong></th>
				  <th class="minicell"><%=yyDiscnt%></th>
				  <th class="minicell"><%=yyModify%></th>
				  <th class="minicell"><%=yyDelete%></th>
				</tr>
<%	allcoupon=""
	sSQL="SELECT DISTINCT cpaAssignment FROM cpnassign WHERE cpaType=3"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then allcoupon=rs.getrows
	rs.close
	rowcounter=1
	sSQL = "SELECT scID,scWorkingName,scGroup,scName,scOrder,COUNT(mSCscID) AS tcount FROM searchcriteria LEFT JOIN multisearchcriteria ON searchcriteria.scID=multisearchcriteria.mSCscID WHERE scGroup="&getrequest("id")&" GROUP BY scID,scWorkingName,scGroup,scName,scOrder ORDER BY scOrder,scName"
	rs.open sSQL,cnn,0,1
	ordingroup=1
	rowsingrp=""
	showalphabetbutton=FALSE
	do while NOT rs.EOF
		if bgcolor="altdark" then bgcolor="altlight" else bgcolor="altdark"
		if rs("scOrder")<>0 then showalphabetbutton=TRUE
		hascoupon=FALSE
		if isarray(allcoupon) then
			for index=0 to UBOUND(allcoupon,2)
				if int(allcoupon(0,index))=rs("scID") then hascoupon=TRUE : exit for
			next
		end if
%>		<tr class="<%=bgcolor%>">
			<td>&nbsp;<%
		print "<select name=""newpos"" onchange=""chi("&rs("scID")&",this)"" onmouseover=""popsel(this,"&ordingroup&","&rs("scGroup")&")"">"
		print "<option value="""" selected=""selected"">"&ordingroup&IIfVr(ordingroup<100,"&nbsp;","")&"</option>"
		print "</select>" %></td>
			<td class="minicell"><%=rs("scID")%></td>
			<td align="left"><%=rs("scWorkingName")%>&nbsp;</td>
			<td align="left"><%=rs("scName")&" ("&rs("tcount")&")"%></td>
			<td class="minicell"><input type="button" <%=IIfVs(hascoupon,"style=""color:#F4E64B"" ")%>value="<%=htmlspecials(yyAssign)%>" onclick="dsc(<%=rs("scID")%>)" /></td>
			<td class="minicell"><input type="button" value="<%=yyModify%>" onclick="modrec(<%=rs("scID")%>)" /></td>
			<td class="minicell"><input type="button" value="<%=yyDelete%>" onclick="delrec(<%=rs("scID")%>)" /></td>
		</tr><%
		rowcounter=rowcounter+1
		ordingroup=ordingroup+1
		rs.movenext
	loop
%>
				<tr class="<%=bgcolor%>">
				  <td colspan="2"><% if showalphabetbutton then print "<input type=""button"" value=""Sort Alphabetically"" onclick=""sortalphabetically()"" />" else print "&nbsp;" %></td>
				  <td align="left"><input type="text" name="newwn" size="24" value="" />&nbsp;</td>
				  <td align="left"><input type="text" id="newname" name="newname" size="24" value="" placeholder="Attribute Name" />
<%	for index=2 to adminlanguages+1
		if (adminlangsettings AND 131072)=131072 then
			print "<br /><input type=""text"" name=""newname"&index&""" value="""" size=""24"" placeholder=""Language "&index&""" />"
		end if
	next %>
				  </td>
				  <td colspan="2"><input type="button" value="<%=yyAddNew%>" onclick="addnewcriteria()" /></td>
				</tr>
				<tr> 
				  <td width="100%" colspan="6" align="center"><br /><input type="button" value="Back to Attribute Groups" onclick="document.location='adminsearchcriteria.asp'" /><br />&nbsp;</td>
				</tr>
				<tr> 
				  <td width="100%" colspan="6" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
				</tr>
			  </table>
			</form>
<script>
/* <![CDATA[ */
var totrows=<%=rowcounter-1%>
/* ]]> */
</script>
		  </td>
        </tr>
      </table>
<%
else %>
<script>
<!--
function popsel(x,theid){
	if(x.length>1) return;
	for(index=theid-1; index>0; index--){
		var y=document.createElement('option');
		y.text=index;
		y.value=index;
		var sel=x.options[0];
		try{
			x.add(y, sel); // FF etc
		}
		catch(ex){
			x.add(y, 0); // IE
		}
	}
	for(index=theid+1; index<=totrows; index++){
		var y=document.createElement('option');
		y.text=index;
		y.value=index;
		try{
			x.add(y, null); // FF etc
		}
		catch(ex){
			x.add(y); // IE
		}
	}
}
function chi(id,obj){
	document.mainform.action="adminsearchcriteria.asp";
	document.mainform.newval.value = obj.selectedIndex+1;
	document.mainform.id.value = id;
	document.mainform.act.value = "changepos";
	document.mainform.submit();
}
function modrec(id) {
	document.mainform.id.value = id;
	document.mainform.act.value = "modifygroup";
	document.mainform.submit();
}
function modatts(id) {
	document.mainform.id.value = id;
	document.mainform.act.value = "modifyatts";
	document.mainform.submit();
}
function newrec(id) {
	document.mainform.id.value = id;
	document.mainform.act.value = "addnewgroup";
	document.mainform.submit();
}
function delrec(id){
if(confirm("<%=jscheck(yyConDel)%>\n")){
	document.mainform.id.value = id;
	document.mainform.act.value = "deletegroup";
	document.mainform.submit();
}
}
function addnewgroup(){
	if(document.getElementById('newname').value==''){
		alert("<%=jscheck(yyPlsEntr&" """&yyName)%>\".");
		document.getElementById('newname').focus();
	}else{
		document.mainform.action="adminsearchcriteria.asp";
		document.mainform.id.value='';
		document.mainform.act.value="newgroup";
		document.mainform.submit();
	}
}
// -->
</script>
<h2><%=yyAdmSeC%></h2>
	  <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr> 
          <td width="100%" align="center"><%
	sSQL = "SELECT COUNT(*) as countpid FROM (multisearchcriteria LEFT JOIN products ON multisearchcriteria.mSCpID=products.pID) LEFT JOIN searchcriteria ON multisearchcriteria.mSCscID=searchcriteria.scID WHERE pID IS NULL OR scID IS NULL"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then numberorphans=rs("countpid") else numberorphans=0
	rs.close
	if numberorphans>0 then print "<div style=""text-align:center;color:red;margin-bottom:10px""><input type=""button"" onclick=""document.location='adminsearchcriteria.asp?act=repair'"" value=""There are "&numberorphans&" orphaned entries in the Product Attributes database table. Please click here to repair this"" /></div>"
%>
			<form name="mainform" method="post" action="adminsearchcriteria.asp">
			<input type="hidden" name="id" value="xxx" />
			<input type="hidden" name="act" value="xxxxx" />
			<input type="hidden" name="selectedq" value="1" />
			<input type="hidden" name="newval" value="1" />
			  <table width="80%" border="0" cellspacing="0" cellpadding="2">
				<tr>
				  <td width="5%"><strong><%=yyOrder%></strong></td>
				  <td width="10%"><strong><%=yyID%></strong></td>
				  <td align="left"><strong><%=yyWrkNam%></strong></td>
				  <td align="left"><strong>Attribute Group <%=yyName%></strong></td>
				  <td width="14%"><strong><%=yyModify&" "&yyGroup%></strong></td>
				  <td width="15%"><strong><%=yyModify&" Attributes"%></strong></td>
				  <td width="10%"><strong><%=yyDelete%></strong></td>
				</tr>
<%	rowcounter=1
	sSQL = "SELECT scgID,scgWorkingName,scgTitle FROM searchcriteriagroup ORDER BY scgOrder"
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		sSQL="SELECT COUNT(*) AS thecount FROM searchcriteria WHERE scGroup="&rs("scgID")
		rs2.open sSQL,cnn,0,1
		if NOT rs2.EOF then numcriteria=rs2("thecount") else numcriteria=0
		rs2.close
		if bgcolor="altdark" then bgcolor="altlight" else bgcolor="altdark" %>
			<tr class="<%=bgcolor%>">
			  <td><%
				print "<select name=""newpos"" onchange=""chi("&rs("scgID")&",this)"" onmouseover=""popsel(this,"&rowcounter&")"">"
				print "<option value="""" selected=""selected"">"&rowcounter&IIfVr(rowcounter<100,"&nbsp;","")&"</option>"
				print "</select>" %></td>
			  <td><%=rs("scgID")%></td>
			  <td align="left"><%=rs("scgWorkingName")%></td>
			  <td align="left"><%=rs("scgTitle")%>&nbsp;</td>
			  <td><input type="button" value="<%=yyModify&" "&yyGroup%>" onclick="modrec('<%=rs("scgID")%>')" /></td>
			  <td><input type="button" value="<%=yyModify&" Attributes ("&numcriteria&")"%>" onclick="modatts('<%=rs("scgID")%>')" /></td>
			  <td><input type="button" value="<%=yyDelete%>" onclick="delrec('<%=rs("scgID")%>')" /></td>
			</tr>
<%		rowcounter=rowcounter+1
		rs.movenext
	loop
%>
				<tr class="<%=bgcolor%>">
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				  <td align="left"><input type="text" name="newwn" size="24" value="" />&nbsp;</td>
				  <td align="left"><input type="text" id="newname" name="newname" size="24" value="" placeholder="Group Name" />
<%	for index=2 to adminlanguages+1
		if (adminlangsettings AND 131072)=131072 then
			print "<br /><input type=""text"" name=""newname"&index&""" value="""" size=""24"" placeholder=""Language "&index&""" />"
		end if
	next %>
				  </td>
				  <td colspan="3"><input type="button" value="<%=yyAddNew%>" onclick="addnewgroup()" /></td>
				</tr>
				<tr> 
				  <td width="100%" colspan="7" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
				</tr>
			  </table>
			</form>
<script>
/* <![CDATA[ */
var totrows=<%=rowcounter-1%>
/* ]]> */
</script>
		  </td>
        </tr>
      </table>
<%
end if
cnn.close
set rs=nothing
set rs2=nothing
set cnn=nothing
%>