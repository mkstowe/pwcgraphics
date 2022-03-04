<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
Dim sSQL,rs,alldata,alladmin,success,cnn,rowcounter,errmsg,aFields(3)
success=true
maxcatsperpage = 100
dorefresh=FALSE
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
sSQL = ""
if getpost("posted")="1" then
	if getpost("act")="delete" then
		sSQL = "DELETE FROM contentregions WHERE contentID="&getpost("id")
		ect_query(sSQL)
		dorefresh=TRUE
	elseif getpost("act")="domodify" then
		sSQL = "UPDATE contentregions SET contentName='"&escape_string(getpost("contentname"))&"',contentX="&getpost("contentX")&",contentData='"&escape_string(getpost("contentdata"))&"'"
		for index=2 to adminlanguages+1
			if (adminlangsettings AND 32768)=32768 then
				sSQL = sSQL & ",contentData" & index & "='"&escape_string(getpost("contentdata" & index))&"'"
			end if
		next
		sSQL = sSQL & " WHERE contentID="&getpost("id")
		ect_query(sSQL)
		dorefresh=TRUE
	elseif getpost("act")="doaddnew" then
		sSQL = "INSERT INTO contentregions (contentName,contentX,contentData"
		for index=2 to adminlanguages+1
			if (adminlangsettings AND 32768)=32768 then
				sSQL = sSQL & ",contentData" & index
			end if
		next
		sSQL = sSQL & ") VALUES ("
		sSQL = sSQL & "'"&escape_string(getpost("contentname"))&"'"
		sSQL = sSQL & ","&getpost("contentX")
		sSQL = sSQL & ",'"&escape_string(getpost("contentdata"))&"'"
		for index=2 to adminlanguages+1
			if (adminlangsettings AND 32768)=32768 then
				sSQL = sSQL & ",'"&escape_string(getpost("contentdata" & index))&"'"
			end if
		next
		sSQL = sSQL & ")"
		ect_query(sSQL)
		dorefresh=TRUE
	end if
end if
if dorefresh then
	print "<meta http-equiv=""refresh"" content=""1; url=admincontent.asp"
	print "?stext=" & urlencode(request("stext")) & "&stype=" & request("stype") & "&pg=" & request("pg")
	print """>"
end if
%>
<script>
/* <![CDATA[ */
function formvalidator(theForm){
  if (theForm.contentname.value == ""){
    alert("<%=jscheck(yyPlsEntr)%> \"Region Name\".");
    theForm.contentname.focus();
    return (false);
  }
  return (true);
}
function froalafocus(editornumber){
	document.getElementById('contenttable'+(editornumber==1?'':editornumber)).style.border='none';
}
function editsize(dir){
	var contentX=document.getElementById('contentX').value;
<%		if htmleditor="froala" OR htmleditor="ckeditor" then %>
			var wid=(contentX*6);
			var amt=6;
			if(dir=='++'||dir=='--') amt=60;
			if(dir=='+'||dir=='++')
				contentX=(Math.min(parseInt(wid)+amt,900));
			else
				contentX=(Math.max(parseInt(wid)-amt,180));
<%		else %>
			var wid=contentX;
			var amt=1;
			if(dir=='++'||dir=='--') amt=10;
			if(dir=='+'||dir=='++')
				contentX=(Math.min(parseInt(wid)+amt,150));
			else
				contentX=(Math.max(parseInt(wid)-amt,30));
<%		end if %>
	for(var ix=1;ix<=3;ix++){
		if(ix==1)ixt='';else ixt=ix;
		if(contab = document.getElementById('<%=IIfVr(htmleditor="froala" OR htmleditor="ckeditor","contenttable","contentdata")%>'+ixt)){
<%		if htmleditor="froala" OR htmleditor="ckeditor" then %>
			contab.style.width=contentX+'px';
			document.getElementById('contentX').value=parseInt(contentX/6.0);
<%		else %>
			contab.cols=contentX;
			document.getElementById('contentX').value=contentX;
<%		end if %>
		}
	}
}
/* ]]> */
</script>
<%		if getpost("posted")="1" AND (getpost("act")="modify" OR getpost("act")="addnew") then
			if htmleditor="ckeditor" then %>
<script src="ckeditor/ckeditor.js"></script>
<%			end if
		end if %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
<% if getpost("posted")="1" AND (getpost("act")="modify" OR getpost("act")="addnew") then
		if getpost("act")="modify" then
			sSQL = "SELECT contentID,contentName,contentX,contentData FROM contentregions WHERE contentID="&getpost("id")
			rs.open sSQL,cnn,0,1
			contentID = rs("contentID")
			contentName = rs("contentName")
			contentX = rs("contentX")
			contentData = rs("contentData")
			rs.close
		else
			contentID = ""
			contentName = ""
			contentX = 0
			contentData = ""
		end if
		if contentX=0 then contentX=100
%>
        <tr>
		  <td width="100%">
		  <form name="mainform" method="post" action="admincontent.asp" onsubmit="return formvalidator(this)">
			<input type="hidden" name="posted" value="1" />
			<% if getpost("act")="modify" then %>
			<input type="hidden" name="act" value="domodify" />
			<% else %>
			<input type="hidden" name="act" value="doaddnew" />
			<% end if
			call writehiddenvar("stext", getpost("stext"))
			call writehiddenvar("stype", getpost("stype"))
			call writehiddenvar("pg", getpost("pg")) %>
			<input type="hidden" name="id" value="<%=getpost("id")%>" />
			<input type="hidden" id="contentX" name="contentX" value="<%=contentX%>" />
            <table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><strong>Use this page to manage your CMS Content Regions</strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><strong>Region Name</strong></td><td><input type="text" name="contentname" size="30" value="<%=htmlspecials(contentName)%>" /></td>
			  </tr>
			  <tr>
				<td align="center" colspan="2">&nbsp;<br /><input type="button" value=" -- " onclick="editsize('--')" /> <input type="button" value=" - " onclick="editsize('-')" /> <strong>Content Data</strong> <input type="button" value=" + " onclick="editsize('+')" /> <input type="button" value=" ++ " onclick="editsize('++')" /><br />
				
<%				if htmleditor="froala" OR htmleditor="ckeditor" then print "<div style="""&IIfVs(htmleditor="froala","border:1px solid grey;")&"margin-top:20px;width:"&(contentX*6)&"px;text-align:left"" id=""contenttable"">" %>
				<textarea name="contentdata" id="contentdata" cols="<%=contentX%>" rows="30"><%=htmlspecials(contentData)%></textarea>
<%				if htmleditor="froala" OR htmleditor="ckeditor" then print "</div>" %>
				</td>
			  </tr>
<%			for index=2 to adminlanguages+1
				if (adminlangsettings AND 32768)=32768 then
					if getpost("act")="modify" then
						sSQL = "SELECT contentData"&index&" FROM contentregions WHERE contentID="&getpost("id")
						rs.open sSQL,cnn,0,1
						contentData = trim(rs("contentData"&index)&"")
						rs.close
					end if
%>
			  <tr>
				<td align="center" colspan="2">
					<div><strong>Content Data <%=index%></strong></div>
<%				if htmleditor="froala" OR htmleditor="ckeditor" then print "<div style="""&IIfVs(htmleditor="froala","border:1px solid grey;")&"margin-top:20px;width:"&(contentX*6)&"px;text-align:left"" id=""contenttable" & index & """>" %>
					<textarea name="contentdata<%=index%>" id="contentdata<%=index%>" cols="<%=contentX%>" rows="30"><%=htmlspecials(contentData)%></textarea>
<%				if htmleditor="froala" OR htmleditor="ckeditor" then print "</div>" %>
				</td>
			  </tr>
<%				end if
			next %>
			  <tr>
			    <td colspan="2" align="center">&nbsp;<br /><input type="submit" value="<%=yySubmit%>" /></td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="2" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
                          &nbsp;</td>
			  </tr>
            </table>
		  </form>
<%	if htmleditor="ckeditor" then
		pathtovsadmin=request.servervariables("URL")
		slashpos=instrrev(pathtovsadmin, "/")
		if slashpos>0 then pathtovsadmin=left(pathtovsadmin, slashpos-1)
		print "<script>function loadeditors(){"
		streditor = "var contentdata=CKEDITOR.replace('contentdata',{extraPlugins : 'stylesheetparser,autogrow',autoGrow_maxHeight : 800,removePlugins : 'resize', toolbarStartupExpanded : false, toolbar : 'Basic', filebrowserBrowseUrl :'ckeditor/filemanager/browser/default/browser.html?Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserImageBrowseUrl : 'ckeditor/filemanager/browser/default/browser.html?Type=Image&Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserFlashBrowseUrl :'ckeditor/filemanager/browser/default/browser.html?Type=Flash&Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=File',filebrowserImageUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=Image',filebrowserFlashUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=Flash'});" & vbCrLf
		streditor = streditor & "contentdata.on('instanceReady',function(event){var myToolbar = 'Basic';event.editor.on( 'beforeMaximize', function(){if(event.editor.getCommand('maximize').state == CKEDITOR.TRISTATE_ON && myToolbar != 'Basic'){contentdata.setToolbar('Basic');myToolbar = 'Basic';contentdata.execCommand('toolbarCollapse');}else if(event.editor.getCommand('maximize').state == CKEDITOR.TRISTATE_OFF && myToolbar != 'Full'){contentdata.setToolbar('Full');myToolbar = 'Full';contentdata.execCommand('toolbarCollapse');}});event.editor.on('contentDom', function(e){event.editor.document.on('blur', function(){if(!contentdata.isToolbarCollapsed){contentdata.execCommand('toolbarCollapse');contentdata.isToolbarCollapsed=true;}});event.editor.document.on('focus',function(){if(contentdata.isToolbarCollapsed){contentdata.execCommand('toolbarCollapse');contentdata.isToolbarCollapsed=false;}});});contentdata.fire('contentDom');contentdata.isToolbarCollapsed=true;});"
		print streditor
		for index=2 to adminlanguages+1
			if (adminlangsettings AND 32768)=32768 then print replace(streditor, "contentdata", "contentdata" & index)
		next
		print "}window.onload=function(){loadeditors();}</script>"
	elseif htmleditor="froala" then
		call displayfroalaeditor("contentdata","Content Data",".on('froalaEditor.focus',function(){froalafocus(1);})",FALSE,FALSE,1,FALSE)
		for index=2 to adminlanguages+1
			if (adminlangsettings AND 32768)=32768 then call displayfroalaeditor("contentdata"&index,"Content Data (language "&index&")",".on('froalaEditor.focus',function(){froalafocus("&index&");})",FALSE,FALSE,1,FALSE)
		next
	end if %>
		  </td>
        </tr>
<% elseif getpost("posted")="1" AND success then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%><a href="admincontent.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br />&nbsp;
                </td>
			  </tr>
			</table></td>
        </tr>
<% elseif getpost("posted")="1" then %>
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><span style="color:#FF0000;font-weight:bold"><%=yyOpFai%></span><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table></td>
        </tr>
<% else %>
        <tr>
		  <td width="100%">
<script>
/* <![CDATA[ */
function mrk(id) {
	document.mainform.id.value = id;
	document.mainform.act.value = "modify";
	document.mainform.submit();
}
function newrec(id){
	document.mainform.id.value = id;
	document.mainform.act.value = "addnew";
	document.mainform.submit();
}
function drk(id) {
if (confirm("<%=jscheck(yyConDel)%>\n")) {
	document.mainform.id.value = id;
	document.mainform.act.value = "delete";
	document.mainform.submit();
}
}
function startsearch(){
	document.mainform.action="admincontent.asp";
	document.mainform.act.value = "search";
	document.mainform.posted.value = "";
	document.mainform.submit();
}
/* ]]> */
</script>
<h2><%=yyAdmCon%></h2>
		  <form name="mainform" method="post" action="admincontent.asp">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="xxxxx" />
			<input type="hidden" name="id" value="xxxxx" />
			<input type="hidden" name="pg" value="<%=IIfVr(getpost("act")="search", "1", getget("pg"))%>" />
			<table class="cobtbl" width="100%" border="0" cellspacing="1" cellpadding="3">
				  <tr> 
	                <td class="cobhl" width="25%" align="right"><%=yySrchFr%>:</td>
					<td class="cobll" width="25%"><input type="text" name="stext" size="20" value="<%=request("stext")%>" /></td>
					<td class="cobhl" width="25%" align="right"><%=yySrchTp%>:</td>
					<td class="cobll" width="25%"><select name="stype" size="1">
						<option value=""><%=yySrchAl%></option>
						<option value="any"<% if request("stype")="any" then print " selected=""selected"""%>><%=yySrchAn%></option>
						<option value="exact"<% if request("stype")="exact" then print " selected=""selected"""%>><%=yySrchEx%></option>
						</select>
					</td>
				  </tr>
				  <tr>
				    <td class="cobhl" align="center">&nbsp;</td>
				    <td class="cobll" colspan="3"><table width="100%" cellspacing="0" cellpadding="0" border="0">
					    <tr>
						  <td class="cobll" align="center" style="white-space:nowrap">
							<input type="submit" value="List Content Regions" onclick="startsearch();" />
							<input type="button" value="New Content Region" onclick="newrec()" />
						  </td>
						  <td class="cobll" height="26" width="20%" align="right" style="white-space:nowrap">
						&nbsp;
						  </td>
						</tr>
					  </table></td>
				  </tr>
				</table>
<br />
            <table width="100%" class="stackable admin-table-a sta-white">
<%
if getpost("act")="search" OR getget("pg")<>"" then
	CurPage = 1
	if is_numeric(getget("pg")) then CurPage=int(getget("pg"))
	sSQL = "SELECT contentID,contentName,contentData FROM contentregions"
	whereand=" WHERE "
	if trim(request("stext"))<>"" then
		Xstext = escape_string(request("stext"))
		aText = Split(Xstext)
		if nosearchadmindescription then maxsearchindex=0 else maxsearchindex=1
		aFields(0)="contentName"
		aFields(1)="contentData"
		if request("stype")="exact" then
			sSQL=sSQL & whereand & "(contentName LIKE '%"&Xstext&"%' OR contentData LIKE '%"&Xstext&"%')"
			whereand=" AND "
		else
			sJoin="AND "
			if request("stype")="any" then sJoin="OR "
			sSQL=sSQL & whereand&"("
			whereand=" AND "
			for index=0 to maxsearchindex
				sSQL=sSQL & "("
				for rowcounter=0 to UBOUND(aText)
					sSQL=sSQL & aFields(index) & " LIKE '%"&aText(rowcounter)&"%' "
					if rowcounter<UBOUND(aText) then sSQL=sSQL & sJoin
				next
				sSQL=sSQL & ") "
				if index < maxsearchindex then sSQL=sSQL & "OR "
			next
			sSQL=sSQL & ") "
		end if
	end if
	sSQL = sSQL & " ORDER BY contentName"
	rs.CursorLocation = 3 ' adUseClient
	rs.CacheSize = maxcatsperpage
	rs.open sSQL,cnn
	if NOT rs.EOF then
		rs.MoveFirst
		rs.PageSize = maxcatsperpage
		rs.AbsolutePage = CurPage
		rowcounter=0
		totnumrows=rs.RecordCount
		iNumOfPages = Int((totnumrows + (maxcatsperpage-1)) / maxcatsperpage)
		if iNumOfPages > 1 then print "<tr><td align=""center"" colspan=""6"">" & writepagebar(CurPage,iNumOfPages,yyPrev,yyNext,"<a href=""admincontent.asp?pg=",FALSE) & "<br /><br /></td></tr>"
%>
			  <tr>
				<th><strong>Region ID</strong></th>
				<th class="maincell"><strong>Region Name</strong></th>
				<th><strong>Example URL</strong></th>
				<th class="small minicell"><%=yyModify%></th>
				<th class="small minicell"><%=yyDelete%></th>
			  </tr>
<%		do while NOT rs.EOF AND rowcounter < maxcatsperpage
			if bgcolor="altdark" then bgcolor="altlight" else bgcolor="altdark"%>
<tr class="<%=bgcolor%>">
<td><%=rs("contentID")%></td>
<td><%=rs("contentName")%></td>
<td>default.asp?region=<%=rs("contentID")%></td>
<td class="minicell"><input type="button" value="<%=yyModify%>" onclick="mrk('<%=rs("contentID")%>')" /></td>
<td class="minicell"><input type="button" value="<%=yyDelete%>" onclick="drk('<%=rs("contentID")%>')" /></td>
</tr><%		rowcounter=rowcounter+1
			rs.MoveNext
		loop
		if iNumOfPages > 1 then print "<tr><td align=""center"" colspan=""6""><br />" & writepagebar(CurPage,iNumOfPages,yyPrev,yyNext,"<a href=""admincontent.asp?pg=",FALSE) & "</td></tr>"
	else %>
			  <tr><td width="100%" colspan="6" align="center"><br /><strong><%=yyItNone%><br />&nbsp;</td></tr>
<%	end if
	rs.close
else
	numitems=0
	sSQL="SELECT COUNT(*) as totcount FROM contentregions"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		numitems=rs("totcount")
	end if
	rs.close
	print "<tr><td colspan=""6""><div class=""itemsdefine"">You have " & numitems & " content regions defined.</div></td></tr>"
end if %>
			  <tr> 
                <td width="100%" colspan="6" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
			  </tr>
            </table>
		  </form>
		  </td>
        </tr>
<% end if
cnn.Close
set rs = nothing
set rs2 = nothing
set cnn = nothing
%>
      </table>