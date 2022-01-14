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
Set rs2=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
Session.LCID=1033
dorefresh=FALSE
if getpost("posted")="1"  then
	if getpost("act")="domodify" then
		if (adminlangsettings AND 4096)=4096 AND getpost("id")<>"invoiceheaders" then maxlangs=adminlanguages else maxlangs=0
		for index=0 to maxlangs
			if index=0 then mesgid="" else mesgid=index+1
			sSQL="UPDATE emailmessages SET "
			themessage=trim(getpost("emtextarea" & (index+1)))
			if getpost("id")="orderstatusemail" then
				sSQL=sSQL & "orderstatussubject"&mesgid&"='" & escape_string(getpost("eminputtext" & (index+1))) & "',"
				sSQL=sSQL & "orderstatusemail"&mesgid&"='" & escape_string(themessage) & "'"
			elseif getpost("id")="emailheaders" then
				sSQL=sSQL & "emailsubject"&mesgid&"='" & escape_string(getpost("eminputtext" & (index+1))) & "',"
				sSQL=sSQL & "emailheaders"&mesgid&"='" & escape_string(themessage) & "'"
			elseif getpost("id")="receiptheaders" then
				sSQL=sSQL & "receiptheaders"&mesgid&"='" & escape_string(themessage) & "'"
			elseif getpost("id")="invoiceheaders" then
				cnn.execute("UPDATE admin SET packingslipuseinvoice="&getpost("packingslipuseinvoice")&" WHERE adminID=1")
				sSQL=sSQL & "invoiceheader='" & escape_string(getpost("invoiceheader")) & "',"
				sSQL=sSQL & "invoiceaddress='" & escape_string(getpost("invoiceaddress")) & "',"
				sSQL=sSQL & "invoicefooter='" & escape_string(getpost("invoicefooter")) & "',"
				sSQL=sSQL & "packingslipheader='" & escape_string(getpost("packingslipheader")) & "',"
				sSQL=sSQL & "packingslipaddress='" & escape_string(getpost("packingslipaddress")) & "',"
				sSQL=sSQL & "packingslipfooter='" & escape_string(getpost("packingslipfooter")) & "'"
			elseif getpost("id")="dropshipheaders" then
				sSQL=sSQL & "dropshipsubject"&mesgid&"='" & escape_string(getpost("eminputtext" & (index+1))) & "',"
				sSQL=sSQL & "dropshipheaders"&mesgid&"='" & escape_string(themessage) & "'"
			elseif getpost("id")="giftcertificate" then
				sSQL=sSQL & "giftcertsubject"&mesgid&"='" & escape_string(getpost("eminputtext" & (index+1))) & "',"
				sSQL=sSQL & "giftcertemail"&mesgid&"='" & escape_string(themessage) & "'"
			elseif getpost("id")="giftcertsender" then
				sSQL=sSQL & "giftcertsendersubject"&mesgid&"='" & escape_string(getpost("eminputtext" & (index+1))) & "',"
				sSQL=sSQL & "giftcertsender"&mesgid&"='" & escape_string(themessage) & "'"
			elseif getpost("id")="notifybackinstock" then
				sSQL=sSQL & "notifystocksubject"&mesgid&"='" & escape_string(getpost("eminputtext" & (index+1))) & "',"
				sSQL=sSQL & "notifystockemail"&mesgid&"='" & escape_string(themessage) & "'"
			elseif getpost("id")="abandonedcart" then
				sSQL=sSQL & "abandonedcartsubject"&mesgid&"='" & escape_string(getpost("eminputtext" & (index+1))) & "',"
				sSQL=sSQL & "abandonedcartemail"&mesgid&"='" & escape_string(themessage) & "'"
			end if
			sSQL=sSQL & " WHERE emailID=1"
			ect_query(sSQL)
		next
		dorefresh=TRUE
	end if
end if
if dorefresh then
	print "<meta http-equiv=""refresh"" content=""1; url=adminemailmsgs.asp"
	print "?id=" & urlencode(getpost("id"))
	print """>"
end if
if getpost("id")<>"" AND getpost("act")="modify" then
	if htmlemails<>TRUE then htmleditor=""
	if htmleditor="ckeditor" then %>
<script src="ckeditor/ckeditor.js"></script>
<%	end if %>
<script>
<!--
function formvalidator(theForm){
return (true);
}
function setpsvisibility(theobj){
if(theobj[theobj.selectedIndex].value=='1'){
	for(var index=1;index<=3;index++)
		document.getElementById('ps'+index).style.display='none';
}else{
	for(var index=1;index<=3;index++)
		document.getElementById('ps'+index).style.display='';
}
}
function expandckeditor(divid){
	document.getElementById(divid).style.border='none';
	document.getElementById(divid).style.padding=0;
}
//-->
</script>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
		  <td width="100%" align="center">
		  <form name="mainform" method="post" action="adminemailmsgs.asp" onsubmit="return formvalidator(this)">
<%		call writehiddenvar("posted", "1")
		call writehiddenvar("act", "domodify")
		call writehiddenvar("id", getpost("id"))
%>
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
			  <tr> 
                <td width="100%" colspan="2" align="center"><strong><% print yyEmlAdm & ": " & getpost("id") & "<br />&nbsp;" %></strong></td>
			  </tr>
<%		theid=getpost("id")
		if (adminlangsettings AND 4096)=4096 AND theid<>"invoiceheaders" then maxlangs=adminlanguages else maxlangs=0
		for index=0 to maxlangs
			replacementfields=""
			subjectreplacementfields=""
			hassubject=FALSE
			languageid=index+1
			if theid="orderstatusemail" then
				fieldlist=getlangid("orderstatussubject",4096)&","&getlangid("orderstatusemail",4096)
				replacementfields="%orderid% %ordername% %orderdate% %oldstatus% %newstatus% %date% {%statusid<span style=""color:#F00"">X</span>%} {%statusinfo%} {%trackingnum%} {%invoicenum%} {%reviewlinks%}"
				subjectreplacementfields="%orderid%"
				hassubject=TRUE
			elseif theid="emailheaders" then
				fieldlist =getlangid("emailsubject",4096)&","&getlangid("emailheaders",4096)
				replacementfields="%messagebody% %ordername% %orderdate% {%reviewlinks%}"
				subjectreplacementfields="%orderid% %ordername%"
				hassubject=TRUE
			elseif theid="receiptheaders" then
				fieldlist =getlangid("receiptheaders",4096)
				replacementfields="%messagebody% %reviewlinks% %ordername% %orderdate%"
				hassubject=FALSE
			elseif theid="invoiceheaders" then
				fieldlist="invoiceheader,invoiceaddress,invoicefooter,packingslipheader,packingslipaddress,packingslipfooter"
				replacementfields=""
				hassubject=FALSE
			elseif theid="dropshipheaders" then
				fieldlist =getlangid("dropshipsubject",4096)&","&getlangid("dropshipheaders",4096)
				replacementfields="%messagebody% %ordername% %orderdate%"
				subjectreplacementfields="%orderid%"
				hassubject=TRUE
			elseif theid="giftcertificate" then
				fieldlist =getlangid("giftcertsubject",4096)&","&getlangid("giftcertemail",4096)
				replacementfields="%toname% %fromname% %value% %certificateid% %storeurl% {%message%}"
				subjectreplacementfields="%fromname%"
				hassubject=TRUE
			elseif theid="giftcertsender" then
				fieldlist =getlangid("giftcertsendersubject",4096)&","&getlangid("giftcertsender",4096)
				replacementfields="%toname%"
				subjectreplacementfields="%toname%"
				hassubject=TRUE
			elseif theid="notifybackinstock" then
				fieldlist =getlangid("notifystocksubject",4096)&","&getlangid("notifystockemail",4096)
				replacementfields="%pid% %pname% %link% %storeurl%"
				subjectreplacementfields="%pid %pname%"
				hassubject=TRUE
			elseif theid="abandonedcart" then
				fieldlist =getlangid("abandonedcartsubject",4096)&","&getlangid("abandonedcartemail",4096)
				replacementfields="%ordername% %orderdate% %abandonedcartid% {%email1%} {%email2%} {%email3%}"
				subjectreplacementfields=""
				hassubject=TRUE
			end if
			sSQL="SELECT "&fieldlist&" FROM emailmessages WHERE emailID=1"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if theid="orderstatusemail" then
					thesubject=trim(rs(getlangid("orderstatussubject",4096))&"")
					themessage=trim(rs(getlangid("orderstatusemail",4096))&"")
				elseif theid="emailheaders" then
					thesubject=trim(rs(getlangid("emailsubject",4096))&"")
					themessage=trim(rs(getlangid("emailheaders",4096))&"")
				elseif theid="receiptheaders" then
					themessage=trim(rs(getlangid("receiptheaders",4096))&"")
				elseif theid="invoiceheaders" then
					themessage=""
					invoiceheader=trim(rs("invoiceheader")&"")
					invoiceaddress=trim(rs("invoiceaddress")&"")
					invoicefooter=trim(rs("invoicefooter")&"")
					packingslipheader=trim(rs("packingslipheader")&"")
					packingslipaddress=trim(rs("packingslipaddress")&"")
					packingslipfooter=trim(rs("packingslipfooter")&"")
				elseif theid="dropshipheaders" then
					thesubject=trim(rs(getlangid("dropshipsubject",4096))&"")
					themessage=trim(rs(getlangid("dropshipheaders",4096))&"")
				elseif theid="giftcertificate" then
					thesubject=trim(rs(getlangid("giftcertsubject",4096))&"")
					themessage=trim(rs(getlangid("giftcertemail",4096))&"")
				elseif theid="giftcertsender" then
					thesubject=trim(rs(getlangid("giftcertsendersubject",4096))&"")
					themessage=trim(rs(getlangid("giftcertsender",4096))&"")
				elseif theid="notifybackinstock" then
					thesubject=trim(rs(getlangid("notifystocksubject",4096))&"")
					themessage=trim(rs(getlangid("notifystockemail",4096))&"")
				elseif theid="abandonedcart" then
					thesubject=trim(rs(getlangid("abandonedcartsubject",4096))&"")
					themessage=trim(rs(getlangid("abandonedcartemail",4096))&"")
				else
					print "id not set"
				end if
			end if
			rs.close
			themessage=replace(themessage, "<br>", "<br />")
			themessage=replace(themessage, "%nl%", "<br />")
			if adminlanguages > 0 then %>
			  <tr>
				<td align="center" colspan="2"><strong><%=yyLanID&": " & (index+1)%></strong></td>
			  </tr>
<%			end if
			if hassubject then
				if subjectreplacementfields<>"" then %>
			  <tr>
				<td align="right"><strong><%=yyRepFld%>:</strong></td>
				<td align="left"><%=subjectreplacementfields%></td>
			  </tr>
<%				end if %>
			  <tr>
				<td align="right"><strong><%=yySubjc%>:</strong></td>
				<td align="left"><input type="text" name="eminputtext<%=(index+1)%>" size="55" maxlength="255" value="<%=thesubject%>" /></td>
			  </tr>
<%			end if
			if replacementfields<>"" then %>
			  <tr>
				<td align="right"><strong><%=yyRepFld%>:</strong></td>
				<td align="left"><%=replacementfields%></td>
			  </tr>
<%			end if
			if htmlemails AND NOT (htmleditor="ckeditor" OR htmleditor="froala") then %>
			  <tr>
				<td align="right">&nbsp;</td>
				<td align="left">Remember to use <strong>&lt;br /&gt;</strong> for a new line.</td>
			  </tr>
<%			end if
			if theid="invoiceheaders" then
				sSQL="SELECT packingslipuseinvoice FROM admin WHERE adminID=1"
				rs.open sSQL,cnn,0,1
				packingslipuseinvoice=rs("packingslipuseinvoice")
				rs.close %>
			  <tr>
				<td colspan="2">
			<table>
			  <tr>
				<td align="right"><strong>Separate Packing Slip Headers:</strong></td>
				<td align="left"><select name="packingslipuseinvoice" size="1" onchange="setpsvisibility(this)"><option value="0">Set Packing Slip Headers Separately</option><option value="1"<% if packingslipuseinvoice<>0 then print " selected=""selected"""%>>Use same headers for Packing Slips</option></select></td>
			  </tr>
			  <tr>
				<td align="right"><strong>Invoice Header:</strong></td>
				<td align="left">
<%	if htmleditor="froala" then print "<div id=""invoiceheaderdiv"" class=""htmleditorcontainer"">" %>
					<textarea id="invoiceheader" name="invoiceheader" cols="90" rows="10"><%=htmlspecials(invoiceheader)%></textarea>
<%	if htmleditor="froala" then print "</div>" %>
				</td>
			  </tr>
			  <tr>
				<td align="right"><strong>Invoice Address:</strong></td>
				<td align="left">
<%	if htmleditor="froala" then print "<div id=""invoiceaddressdiv"" class=""htmleditorcontainer"">" %>
					<textarea id="invoiceaddress" name="invoiceaddress" cols="60" rows="10"><%=htmlspecials(invoiceaddress)%></textarea>
<%	if htmleditor="froala" then print "</div>" %>
				</td>
			  </tr>
			  <tr>
				<td align="right"><strong>Invoice Footer:</strong></td>
				<td align="left">
<%	if htmleditor="froala" then print "<div id=""invoicefooterdiv"" class=""htmleditorcontainer"">" %>
					<textarea id="invoicefooter" name="invoicefooter" cols="90" rows="10"><%=htmlspecials(invoicefooter)%></textarea>
<%	if htmleditor="froala" then print "</div>" %>
				</td>
			  </tr>
			  <tr id="ps1"<% if packingslipuseinvoice<>0 then print " style=""display:none"""%>>
				<td align="right"><strong>Packing Slip Header:</strong></td>
				<td align="left">
<%	if htmleditor="froala" then print "<div id=""packingslipheaderdiv"" class=""htmleditorcontainer"">" %>
					<textarea id="packingslipheader" name="packingslipheader" cols="90" rows="10"><%=htmlspecials(packingslipheader)%></textarea>
<%	if htmleditor="froala" then print "</div>" %>
				</td>
			  </tr>
			  <tr id="ps2"<% if packingslipuseinvoice<>0 then print " style=""display:none"""%>>
				<td align="right"><strong>Packing Slip Address:</strong></td>
				<td align="left">
<%	if htmleditor="froala" then print "<div id=""packingslipaddressdiv"" class=""htmleditorcontainer"">" %>
					<textarea id="packingslipaddress" name="packingslipaddress" cols="60" rows="10"><%=htmlspecials(packingslipaddress)%></textarea>
<%	if htmleditor="froala" then print "</div>" %>
				</td>
			  </tr>
			  <tr id="ps3"<% if packingslipuseinvoice<>0 then print " style=""display:none"""%>>
				<td align="right"><strong>Packing Slip Footer:</strong></td>
				<td align="left">
<%	if htmleditor="froala" then print "<div id=""packingslipfooterdiv"" class=""htmleditorcontainer"">" %>
					<textarea id="packingslipfooter" name="packingslipfooter" cols="90" rows="10"><%=htmlspecials(packingslipfooter)%></textarea>
<%	if htmleditor="froala" then print "</div>" %>
				</td>
			  </tr>
			</table>
				</td>
			  </tr>
<%			else %>
			  <tr>
				<td align="right"><strong><%=yyMessag%>:</strong></td>
				<td align="left">
<%	if htmleditor="froala" then print "<div id=""textareadiv"&(index+1)&""" class=""htmleditorcontainer"">" %>
					<textarea id="emtextarea<%=(index+1)%>" name="emtextarea<%=(index+1)%>" cols="90" rows="15"><%=htmlspecials(themessage)%></textarea>
<%	if htmleditor="froala" then print "</div>" %>
				</td>
			  </tr>
<%			end if
		next %>
			  <tr>
                <td width="100%" colspan="2" align="center"><br /><input type="submit" value="<%=yySubmit%>" />&nbsp;<input type="reset" value="<%=yyReset%>" />&nbsp;<input type="button" value="<%=yyCancel%>" onclick="document.location='adminemailmsgs.asp?id=<%=getpost("id")%>'" /><br />&nbsp;</td>
			  </tr>
			  <tr>
                <td width="100%" colspan="2" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
                          &nbsp;</td>
			  </tr>
            </table>
		  </form>
		  </td>
        </tr>
      </table>
<%	if htmleditor="ckeditor" then
		pathtovsadmin=request.servervariables("URL")
		slashpos=instrrev(pathtovsadmin, "/")
		if slashpos>0 then pathtovsadmin=left(pathtovsadmin, slashpos-1)
		print "<script>function loadeditors(){"
		streditor="var emtextarea=CKEDITOR.replace('emtextarea',{extraPlugins : 'stylesheetparser,autogrow',autoGrow_maxHeight : 800,removePlugins : 'resize', toolbarStartupExpanded : false, toolbar : 'Basic', filebrowserBrowseUrl :'ckeditor/filemanager/browser/default/browser.html?Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserImageBrowseUrl : 'ckeditor/filemanager/browser/default/browser.html?Type=Image&Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserFlashBrowseUrl :'ckeditor/filemanager/browser/default/browser.html?Type=Flash&Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=File',filebrowserImageUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=Image',filebrowserFlashUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=Flash'});" & vbCrLf
		streditor=streditor & "emtextarea.on('instanceReady',function(event){var myToolbar='Basic';event.editor.on( 'beforeMaximize', function(){if(event.editor.getCommand('maximize').state == CKEDITOR.TRISTATE_ON && myToolbar != 'Basic'){emtextarea.setToolbar('Basic');myToolbar='Basic';emtextarea.execCommand('toolbarCollapse');}else if(event.editor.getCommand('maximize').state == CKEDITOR.TRISTATE_OFF && myToolbar != 'Full'){emtextarea.setToolbar('Full');myToolbar='Full';emtextarea.execCommand('toolbarCollapse');}});event.editor.on('contentDom', function(e){event.editor.document.on('blur', function(){if(!emtextarea.isToolbarCollapsed){emtextarea.execCommand('toolbarCollapse');emtextarea.isToolbarCollapsed=true;}});event.editor.document.on('focus',function(){if(emtextarea.isToolbarCollapsed){emtextarea.execCommand('toolbarCollapse');emtextarea.isToolbarCollapsed=false;}});});emtextarea.fire('contentDom');emtextarea.isToolbarCollapsed=true;});"
		if theid="invoiceheaders" then
			print replace(streditor, "emtextarea", "invoiceheader")
			print replace(streditor, "emtextarea", "invoiceaddress")
			print replace(streditor, "emtextarea", "invoicefooter")
			print replace(streditor, "emtextarea", "packingslipheader")
			print replace(streditor, "emtextarea", "packingslipaddress")
			print replace(streditor, "emtextarea", "packingslipfooter")
		else
			if (adminlangsettings AND 4096)=4096 then maxlangs=adminlanguages else maxlangs=0
			for index=1 to maxlangs+1
				print replace(streditor, "emtextarea", "emtextarea" & index)
			next
		end if
		print "}window.onload=function(){loadeditors();}</script>"
	elseif htmleditor="froala" then
		if (adminlangsettings AND 4096)=4096 then maxlangs=adminlanguages else maxlangs=0
		if theid="invoiceheaders" then
			call displayfroalaeditor("invoiceheader",yyMessag,".on('froalaEditor.focus',function(){expandckeditor(""invoiceheaderdiv"");})",FALSE,FALSE,1,FALSE)
			call displayfroalaeditor("invoiceaddress",yyMessag,".on('froalaEditor.focus',function(){expandckeditor(""invoiceaddressdiv"");})",FALSE,FALSE,1,FALSE)
			call displayfroalaeditor("invoicefooter",yyMessag,".on('froalaEditor.focus',function(){expandckeditor(""invoicefooterdiv"");})",FALSE,FALSE,1,FALSE)
			call displayfroalaeditor("packingslipheader",yyMessag,".on('froalaEditor.focus',function(){expandckeditor(""packingslipheaderdiv"");})",FALSE,FALSE,1,FALSE)
			call displayfroalaeditor("packingslipaddress",yyMessag,".on('froalaEditor.focus',function(){expandckeditor(""packingslipaddressdiv"");})",FALSE,FALSE,1,FALSE)
			call displayfroalaeditor("packingslipfooter",yyMessag,".on('froalaEditor.focus',function(){expandckeditor(""packingslipfooterdiv"");})",FALSE,FALSE,1,FALSE)
		else
			for index=1 to maxlangs+1
				call displayfroalaeditor("emtextarea"&index,yyMessag,".on('froalaEditor.focus',function(){expandckeditor(""textareadiv"&index&""");})",FALSE,FALSE,1,FALSE)
			next
		end if
	end if
elseif getpost("posted")="1" AND success then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <a href="adminemailmsgs.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br />&nbsp;<br />&nbsp;
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
			<table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><span style="color:#FF0000;font-weight:bold"><%=yyOpFai%></span><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a><p>&nbsp;</p><p>&nbsp;</p></td>
			  </tr>
			</table></td>
        </tr>
      </table>
<%
else
%>
<script>
<!--
function mrec(id) {
	// document.mainform.id.value=id;
}
// -->
</script>
<h2><%=yyAdmEmm%></h2>
		  <form name="mainform" method="post" action="adminemailmsgs.asp">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="modify" />
			<input type="hidden" name="id" id="idset" value="" />
			<table class="cobtbl" width="100%" border="0" cellspacing="1" cellpadding="3">
			  <tr>
				<td class="cobhl" colspan="3" align="center"><strong><%
					print yyHTMLic & " "
					if htmlemails then print yyOn else print yyNot & " " & yyOn %></strong></td>
			  </tr>
			  <tr>
				<td class="cobhl" align="right" width="30%">Order Status Email:</td>
				<td class="cobll"><input type="button" value="Edit Email" onclick="document.getElementById('idset').value='orderstatusemail';document.forms.mainform.submit()" /></td>
				<td class="cobll">The email message that customers receive when you change the status of an order. The different status types that receive an email can be configured on the <a href="adminordstatus.asp">Admin Order Status</a> page.</td>
			  </tr>
			  <tr>
				<td class="cobhl" align="right"><%=yyEmlHdr%>:</td>
				<td class="cobll"><input type="button" value="Edit Email" onclick="document.getElementById('idset').value='emailheaders';document.forms.mainform.submit()" /></td>
				<td class="cobll">The email message that is sent to the customer after a successful transaction. The parameter &quot;%messagebody%&quot; will be replaced by the products ordered.</td>
			  </tr>
			  <tr>
				<td class="cobhl" align="right">Receipt Headers / Footers:</td>
				<td class="cobll"><input type="button" value="Edit Email" onclick="document.getElementById('idset').value='receiptheaders';document.forms.mainform.submit()" /></td>
				<td class="cobll">This is the message that is displayed on the thanks page after a successful transaction. The parameter &quot;%messagebody%&quot; will be replaced by the products ordered.</td>
			  </tr>
			  <tr>
				<td class="cobhl" align="right">Invoice &amp; Packing Slip Headers / Footers:</td>
				<td class="cobll"><input type="button" value="Edit Email" onclick="document.getElementById('idset').value='invoiceheaders';document.forms.mainform.submit()" /></td>
				<td class="cobll">This is where you can set the headers, footers and company address for the printable receipt and packing slips.</td>
			  </tr>
			  <tr>
				<td class="cobhl" align="right"><%=yyDrSppr & " " & yyEmlHdr%>:</td>
				<td class="cobll"><input type="button" value="Edit Email" onclick="document.getElementById('idset').value='dropshipheaders';document.forms.mainform.submit()" /></td>
				<td class="cobll">This is the email message that a drop shipper will receive after a successful order. The %messagebody% parameter will be replaced by the line items that relate to that drop shipper.</td>
			  </tr>
			  <tr>
				<td class="cobhl" align="right">Gift Certificate Email:</td>
				<td class="cobll"><input type="button" value="Edit Email" onclick="document.getElementById('idset').value='giftcertificate';document.forms.mainform.submit()" /></td>
				<td class="cobll">This is the email message that someone will receive when a Gift Certificate is purchased for them.</td>
			  </tr>
			  <tr>
				<td class="cobhl" align="right">Gift Certificate Sender:</td>
				<td class="cobll"><input type="button" value="Edit Email" onclick="document.getElementById('idset').value='giftcertsender';document.forms.mainform.submit()" /></td>
				<td class="cobll">This is the email message that someone will receive when they send someone a Gift Certificate.</td>
			  </tr>
			  <tr>
				<td class="cobhl" align="right">Notify Back In Stock Email:</td>
				<td class="cobll"><input type="button" value="Edit Email" onclick="document.getElementById('idset').value='notifybackinstock';document.forms.mainform.submit()" /></td>
				<td class="cobll">This is the email message that someone will receive when an item goes back in stock and they have requested notification.</td>
			  </tr>
			  <tr>
				<td class="cobhl" align="right">Abandoned Cart Email:</td>
				<td class="cobll"><input type="button" value="Edit Email" onclick="document.getElementById('idset').value='abandonedcart';document.forms.mainform.submit()" /></td>
				<td class="cobll">This is the email message that you can send to prompt a customer to complete an abandoned checkout.</td>
			  </tr>
			  <tr>
                <td class="cobll" colspan="3" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
			  </tr>
			</table>
		  </form>
<%
end if
cnn.Close
set rs=nothing
set cnn=nothing
%>
