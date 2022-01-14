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
if admincatsperpage="" then admincatsperpage=200
if maxloginlevels="" then maxloginlevels=5
dorefresh=FALSE
currentdiscount=""
if lcase(adminencoding)="iso-8859-1" then raquo="»" else raquo=">"
sub writemenulevel(id,itlevel)
	Dim wmlindex
	if itlevel<10 then
		for wmlindex=0 TO ubound(alldata,2)
			if alldata(2,wmlindex)=id then
				print "<option value='"&alldata(0,wmlindex)&"'"
				if thecat=alldata(0,wmlindex) then print " selected=""selected"">" else print ">"
				for index=0 to itlevel-2
					print raquo & " "
				next
				print alldata(1,wmlindex)&"</option>" & vbCrLf
				if alldata(3,wmlindex)=0 then call writemenulevel(alldata(0,wmlindex),itlevel+1)
			end if
		next
	end if
end sub
sub dodeletecat(cid)
	sSQL="DELETE FROM cpnassign WHERE cpaType=1 AND cpaAssignment='"&cid&"'"
	ect_query(sSQL)
	sSQL="DELETE FROM sections WHERE sectionID=" & cid
	ect_query(sSQL)
	sSQL="DELETE FROM multisections WHERE pSection=" & cid
	ect_query(sSQL)
end sub
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs2=Server.CreateObject("ADODB.RecordSet")
Set rsCats=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
sSQL=""
if defaultcatimages="" then defaultcatimages="images/"
if getpost("act")="changepos" then
	theid=int(getpost("id"))
	neworder=int(getpost("newval"))-1
	sSQL="SELECT sectionOrder,topSection FROM sections WHERE sectionID=" & theid
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then topsection=rs("topSection")
	rs.close
	rc=0
	if menucategoriesatroot AND catalogroot<>0 then
		sSQL="SELECT sectionID,topSection FROM sections WHERE (sectionID="&topsection&" OR topSection="&topsection&") AND sectionID="&catalogroot
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then topsection=rs("sectionID")&","&rs("topSection")
		rs.close
	end if
	sSQL="SELECT sectionID,sectionOrder FROM sections WHERE topSection IN ("&topsection&") ORDER BY sectionOrder"
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		if rs("sectionID")=theid then
			sSQL="UPDATE sections SET sectionOrder="&neworder&" WHERE sectionID="&theid
		else
			sSQL="UPDATE sections SET sectionOrder="&IIfVr(rc<neworder,rc,rc+1)&" WHERE sectionID="&rs("sectionID")
		end if
		ect_query(sSQL)
		rc=rc+1
		rs.movenext
	loop
	rs.close
	dorefresh=TRUE
elseif getpost("posted")="1" then
	if getpost("act")="delete" then
		dodeletecat(getpost("id"))
		dorefresh=TRUE
	elseif getpost("act")="quickupdate" then
		for each objItem in request.form
			if left(objItem, 4)="pra_" then
				origid=right(objItem, len(objItem)-4)
				theid=getpost("pid"&origid)
				theval=getpost(objItem)
				cract=getpost("cract")
				sSQL=""
				if cract="can" then
					if trim(theval)<>"" then sSQL="UPDATE sections SET sectionName='" & escape_string(theval) & "'"
				elseif cract="can2" OR cract="can3" then
					sSQL="UPDATE sections SET sectionName" & IIfVr(cract="can2","2","3") & "='" & escape_string(theval) & "'"
				elseif cract="cwn" then
					if theval<>"" then sSQL="UPDATE sections SET sectionWorkingName='" & escape_string(theval) & "'"
				elseif cract="cim" then
					sSQL="UPDATE sections SET sectionImage='" & escape_string(theval) & "'"
				elseif cract="dis" AND getpost("currentdiscount")<>"" then
					ect_query("DELETE FROM cpnassign WHERE cpaType=1 AND cpaAssignment='"&escape_string(theid)&"' AND cpaCpnID="&getpost("currentdiscount"))
					if getpost("prb_" & origid)="1" then
						ect_query("INSERT INTO cpnassign (cpaType,cpaAssignment,cpaCpnID) VALUES (1,'"&escape_string(theid)&"',"&getpost("currentdiscount")&")")
					end if
				elseif cract="ptt" then
					sSQL="UPDATE sections SET sTitle='" & escape_string(theval) & "'"
				elseif cract="ptt2" then
					sSQL="UPDATE sections SET sTitle2='" & escape_string(theval) & "'"
				elseif cract="ptt3" then
					sSQL="UPDATE sections SET sTitle3='" & escape_string(theval) & "'"
				elseif cract="med" then
					sSQL="UPDATE sections SET sMetaDesc='" & escape_string(theval) & "'"
				elseif cract="med2" then
					sSQL="UPDATE sections SET sMetaDesc2='" & escape_string(theval) & "'"
				elseif cract="med3" then
					sSQL="UPDATE sections SET sMetaDesc3='" & escape_string(theval) & "'"
				elseif cract="cur" then
					sSQL="UPDATE sections SET sectionURL='" & escape_string(theval) & "'"
				elseif cract="lol" then
					if is_numeric(theval) then sSQL="UPDATE sections SET sectionDisabled=" & theval
				elseif cract="rec" then
					sSQL="UPDATE sections SET sRecommend=" & IIfVr(getpost("prb_" & origid)="1","1","0")
				elseif cract="cur2" OR cract="cur3" then
					sSQL="UPDATE sections SET sectionURL" & IIfVr(cract="cur2","2","3") & "='" & escape_string(theval) & "'"
				elseif cract="css" then
					sSQL="UPDATE sections SET sCustomCSS='" & escape_string(theval) & "'"
				elseif cract="del" then
					if theval="del" then dodeletecat(theid)
					sSQL=""
				end if
				if sSQL<>"" then
					sSQL=sSQL & " WHERE sectionID="&int(theid)
					ect_query(sSQL)
				end if
			end if
		next
		dorefresh=TRUE
	elseif getpost("act")="domodify" then
		olddisabled=0
		sSQL="SELECT sectionDisabled FROM sections WHERE sectionID="&getpost("id")
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then olddisabled=rs("sectionDisabled")
		rs.close
		sSQL="UPDATE sections SET sectionName='"&escape_string(getpost("secname"))&"',sectionDescription='"&escape_string(getpost("secdesc"))&"',sectionImage='"&escape_string(getpost("secimage"))&"',topSection="&getpost("tsTopSection")&",rootSection="&getpost("catfunction")
		workname=escape_string(getpost("secworkname"))
		if workname<>"" then
			sSQL=sSQL & ",sectionWorkingName='"&workname&"'"
		else
			sSQL=sSQL & ",sectionWorkingName='"&escape_string(getpost("secname"))&"'"
		end if
		for index=2 to adminlanguages+1
			if (adminlangsettings AND 256)=256 then
				sSQL=sSQL & ",sectionName" & index & "='"&escape_string(getpost("secname" & index))&"'"
			end if
			if (adminlangsettings AND 512)=512 then
				sSQL=sSQL & ",sectionDescription" & index & "='"&escape_string(getpost("secdesc" & index))&"'"
			end if
			if (adminlangsettings AND 2048)=2048 then
				sSQL=sSQL & ",sectionurl" & index & "='"&escape_string(getpost("sectionurl" & index))&"'"
			end if
		next
		sSQL=sSQL & ",sectionDisabled=" & getpost("sectionDisabled")
		sSQL=sSQL & ",sectionHeader='" & escape_string(getpost("sectionHeader")) & "'"
		for index=2 to adminlanguages+1
			if (adminlangsettings AND 524288)=524288 then sSQL=sSQL & ",sectionHeader"&index&"='"& escape_string(getpost("sectionHeader"&index)) &"'"
			if (adminlangsettings AND 2097152)=2097152 then sSQL=sSQL & ",sTitle"&index&"='"& escape_string(getpost("sTitle"&index)) &"'"
			if (adminlangsettings AND 2097152)=2097152 then sSQL=sSQL & ",sMetaDesc"&index&"='"& escape_string(getpost("sMetaDesc"&index)) &"'"
		next
		sSQL=sSQL & ",sectionurl='" & escape_string(getpost("sectionurl")) & "',sTitle='" & escape_string(getpost("sTitle")) & "',sMetaDesc='" & escape_string(getpost("sMetaDesc")) & "'"
		sSQL=sSQL & " WHERE sectionID="&getpost("id")
		ect_query(sSQL)
		if getpost("catalogroot")="ON" then
			if catalogroot<>int(getpost("id")) then
				ect_query("UPDATE admin SET catalogRoot="&getpost("id")&" WHERE adminID=1")
			end if
		else
			if catalogroot=int(getpost("id")) then
				ect_query("UPDATE admin SET catalogRoot=0 WHERE adminID=1")
			end if
		end if
		if (olddisabled<>int(getpost("sectionDisabled")) OR getpost("forcesubsection")="1") AND getpost("forcesubsection")<>"2" then
			idlist=getpost("id")
			ect_query("UPDATE sections SET sectionDisabled=" & getpost("sectionDisabled") & " WHERE topSection=" & idlist)
			for index=1 to 10
				sSQL="SELECT sectionID,sectionDisabled,rootSection FROM sections WHERE rootSection=0 AND topSection IN (" & idlist & ")"
				idlist=""
				rs.open sSQL,cnn,0,1
				do while NOT rs.EOF
					sSQL="UPDATE sections SET sectionDisabled=" & getpost("sectionDisabled") & " WHERE topSection=" & rs("sectionID")
					ect_query(sSQL)
					idlist=idlist&rs("sectionID")&","
					rs.movenext
				loop
				rs.close
				if idlist<>"" then idlist=left(idlist,len(idlist)-1) else exit for
			next
		end if
		dorefresh=TRUE
	elseif getpost("act")="doaddnew" then
		if getpost("catfunction")="2" then
			uniqueindex=0
			mxOrder=0
		else
			haveuniqueindex=FALSE
			uniqueindex=1
			do while NOT haveuniqueindex
				rs.open "SELECT sectionID FROM sections WHERE sectionID="&uniqueindex,cnn,0,1
				if rs.EOF then haveuniqueindex=TRUE else uniqueindex=uniqueindex+1
				rs.close
			loop
			sSQL="SELECT MAX(sectionOrder) AS mxOrder FROM sections"
			rs.open sSQL,cnn,0,1
			mxOrder=rs("mxOrder")
			rs.close
			if IsNull(mxOrder) OR mxOrder="" then mxOrder=1 else mxOrder=mxOrder+1
		end if
		sSQL="INSERT INTO sections (sectionID,sectionName,sectionName2,sectionName3,sectionDescription,sectionDescription2,sectionDescription3,sectionImage,sectionOrder,topSection,rootSection,sectionWorkingName"
		sSQL=sSQL & ",sectionDisabled,sectionHeader,sectionHeader2,sectionHeader3"
		sSQL=sSQL & ",sectionurl,sectionurl2,sectionurl3,sTitle,sTitle2,sTitle3,sMetaDesc,sMetaDesc2,sMetaDesc3) VALUES ("&uniqueindex&",'"&escape_string(getpost("secname"))&"','"&escape_string(getpost("secname2"))&"','"&escape_string(getpost("secname3"))&"','"&escape_string(getpost("secdesc"))&"','"&escape_string(getpost("secdesc2"))&"','"&escape_string(getpost("secdesc3"))&"','"&escape_string(getpost("secimage"))&"',"&mxOrder&","&getpost("tsTopSection")&","&getpost("catfunction")
		workname=escape_string(getpost("secworkname"))
		if workname<>"" then
			sSQL=sSQL & ",'"&workname&"'"
		else
			sSQL=sSQL & ",'"&escape_string(getpost("secname"))&"'"
		end if
		sSQL=sSQL & "," & getpost("sectionDisabled")
		sSQL=sSQL & ",'" & escape_string(getpost("sectionHeader")) & "','" & escape_string(getpost("sectionHeader2")) & "','" & escape_string(getpost("sectionHeader3")) & "'"
		sSQL=sSQL & ",'" & escape_string(getpost("sectionurl")) & "','" & escape_string(getpost("sectionurl2")) & "','" & escape_string(getpost("sectionurl3")) & "','" & escape_string(getpost("sTitle")) & "','" & escape_string(getpost("sTitle2")) & "','" & escape_string(getpost("sTitle3")) & "','" & escape_string(getpost("sMetaDesc")) & "','" & escape_string(getpost("sMetaDesc2")) & "','" & escape_string(getpost("sMetaDesc3")) & "')"
		ect_query(sSQL)
		dorefresh=TRUE
	elseif getpost("act")="dodiscounts" then
		sSQL="INSERT INTO cpnassign (cpaCpnID,cpaType,cpaAssignment) VALUES ("&getpost("assdisc")&",1,'"&getpost("id")&"')"
		ect_query(sSQL)
		dorefresh=TRUE
	elseif getpost("act")="deletedisc" then
		sSQL="DELETE FROM cpnassign WHERE cpaType=1 AND cpaID="&getpost("id")
		ect_query(sSQL)
		dorefresh=TRUE
	elseif getpost("act")="sort" then
		response.cookies("catsort")=getpost("sort")
		response.cookies("catsort").Expires=Date()+365
		if request.servervariables("HTTPS")="on" then response.cookies("catsort").secure=TRUE
	end if
elseif getget("catorman")<>"" then
	response.cookies("ccatorman")=getget("catorman")
	response.cookies("ccatorman").Expires=Date()+926
	if request.servervariables("HTTPS")="on" then response.cookies("ccatorman").secure=TRUE
end if
if dorefresh then
	print "<meta http-equiv=""refresh"" content="""&IIfVr(getpost("act")="changepos",0,1)&"; url=admincats.asp"
	print "?stext=" & urlencode(request("stext")) & "&catfun=" & request("catfun") & "&stype=" & request("stype") & "&scat=" & request("scat") & "&pg=" & request("pg")
	print """>"
else
%>
<script>
/* <![CDATA[ */
function formvalidator(theForm){
  if (theForm.secname.value == ""){
    alert("<%=jscheck(yyPlsEntr&" """&yyCatNam)%>\".");
    theForm.secname.focus();
    return (false);
  }
  if (theForm.tsTopSection[theForm.tsTopSection.selectedIndex].value == ""){
    alert("<%=jscheck(yyPlsSel&" """&yyCatSub)%>\".");
    theForm.tsTopSection.focus();
    return (false);
  }
  return (true);
}
function uploadimage(imfield){
	var winwid=360; var winhei=220;
	var prnttext='<html><head><link rel="stylesheet" type="text/css" href="adminstyle.css"/></head><body>\n';
	prnttext += '<form name="mainform" method="post" action="doupload.asp?defimagepath=<%=defaultcatimages%>" enctype="multipart/form-data">';
	prnttext += '<input type="hidden" name="defimagepath" value="<%=defaultcatimages%>" />';
	prnttext += '<input type="hidden" name="imagefield" value="'+imfield+'" />';
	prnttext += '<table border="0" cellspacing="1" cellpadding="3" width="100%">';
	prnttext += '<tr><td align="center" colspan="2">&nbsp;<br /><strong><%=replace(yyUplIma,"'","\'")%></strong><br />&nbsp;</td></tr>';
	prnttext += '<tr><td align="center" colspan="2"><%=replace(yyPlsSUp,"'","\'")%><br />&nbsp;</td></tr>';
	prnttext += '<tr><td align="right"><%=replace(yyLocIma,"'","\'")%>:</td><td><input type="file" name="imagefile" /></td></tr>';
	prnttext += '<tr><td colspan="2" align="center">&nbsp;<br /><input type="submit" value="<%=replace(yySubmit,"'","\'")%>" /></td></tr>';
	prnttext += '</table></form>';
	prnttext += '<p align="center"><a href="javascript:window.close()"><strong><%=replace(yyClsWin,"'","\'")%></strong></a></p>';
	prnttext += '</body></'+'html>';
	scrwid=screen.width; scrhei=screen.height;
	var newwin=window.open("","printlicense",'menubar=no,scrollbars=yes,width='+winwid+',height='+winhei+',left='+((scrwid-winwid)/2)+',top=100,directories=no,location=no,resizable=yes,status=no,toolbar=no');
	newwin.document.open();
	newwin.document.write(prnttext);
	newwin.document.close();
}
function expandckeditor(divid){
	document.getElementById(divid).style.border='none';
	document.getElementById(divid).style.padding=0;
}
function displaymultilangname(elem){
	for(var index=2;index<=3;index++){
		if(document.getElementById(elem+index))document.getElementById(elem+index).style.display='block';
	}
}
/* ]]> */
</script>
<%
end if %>
<% if getpost("posted")="1" AND (getpost("act")="modify" OR getpost("act")="addnew" OR getpost("act")="clone") then
		if htmleditor="ckeditor" then %>
<script src="ckeditor/ckeditor.js"></script>
<%		end if
		alltopsections=""
		sectionID=""
		sectionName=""
		sectionName2=""
		sectionName3=""
		rootSection=1
		sectionImage=""
		sectionWorkingName=""
		topSection=0
		sectionDisabled=0
		sectionurl=""
		sectionurl2=""
		sectionurl3=""
		sTitle=""
		sTitle2=""
		sTitle3=""
		sMetaDesc=""
		sMetaDesc2=""
		sMetaDesc3=""
		sectionDescription=""
		if (getpost("act")="modify" OR getpost("act")="clone") AND is_numeric(getpost("id")) then
			sSQL="SELECT sectionID,sectionName,sectionName2,sectionName3,rootSection,sectionImage,sectionWorkingName,topSection,sectionDisabled,sectionurl,"
			if (adminlangsettings AND 2048)=2048 then
				if adminlanguages>=1 then sSQL=sSQL & "sectionurl2,"
				if adminlanguages>=2 then sSQL=sSQL & "sectionurl3,"
			end if
			sSQL=sSQL & "sTitle,sTitle2,sTitle3,sMetaDesc,sMetaDesc2,sMetaDesc3,sectionDescription,sectionHeader FROM sections WHERE sectionID="&getpost("id")
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				sectionID=rs("sectionID")
				sectionName=trim(rs("sectionName")&"")
				sectionName2=trim(rs("sectionName2")&"")
				sectionName3=trim(rs("sectionName3")&"")
				rootSection=rs("rootSection")
				sectionImage=rs("sectionImage")
				sectionWorkingName=rs("sectionWorkingName")
				topSection=rs("topSection")
				sectionDisabled=rs("sectionDisabled")
				sectionurl=trim(rs("sectionurl")&"")
				sectionurl2=""
				sectionurl3=""
				if (adminlangsettings AND 2048)=2048 then
					if adminlanguages>=1 then sectionurl2=trim(rs("sectionurl2")&"")
					if adminlanguages>=2 then sectionurl3=trim(rs("sectionurl3")&"")
				end if
				sTitle=rs("sTitle")
				sTitle2=rs("sTitle2")
				sTitle3=rs("sTitle3")
				sMetaDesc=rs("sMetaDesc")
				sMetaDesc2=rs("sMetaDesc2")
				sMetaDesc3=rs("sMetaDesc3")
				sectionDescription=rs("sectionDescription")
				sectionHeader=rs("sectionHeader")
			end if
			rs.close
		end if
		sSQL="SELECT sectionID,sectionWorkingName FROM sections WHERE rootSection=0 ORDER BY sectionWorkingName"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			alltopsections=rs.getrows
		end if
		rs.close
%>
		  <form name="mainform" method="post" action="admincats.asp" onsubmit="return formvalidator(this)">
			<input type="hidden" name="posted" value="1" />
			<% if getpost("act")="modify" then %>
			<input type="hidden" name="act" value="domodify" />
			<% else %>
			<input type="hidden" name="act" value="doaddnew" />
			<% end if
			call writehiddenvar("stext", getpost("stext"))
			call writehiddenvar("stype", getpost("stype"))
			call writehiddenvar("catfun", getpost("catfun"))
			call writehiddenvar("scat", getpost("scat"))
			call writehiddenvar("pg", getpost("pg"))
			call writehiddenvar("id", getpost("id")) %>
            <table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><strong><%=IIfVr(getpost("act")="clone",yyClone&": ",IIfVs(getpost("act")="modify",yyModify&": ")) & yyCatAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td align="right"><%=redasterix&yyCatNam%>:</td><td><input type="text" name="secname" size="30" value="<%=htmlspecialsucode(sectionName)%>" /></td>
			  </tr>
<%		for index=2 to adminlanguages+1
			if (adminlangsettings AND 256)=256 then %>
			  <tr>
				<td align="right"><%=yyCatNam & " " & index %>:</td><td><input type="text" name="secname<%=index%>" size="30" value="<%=htmlspecialsucode(IIfVr(index=2,sectionName2,sectionName3))%>" /></td>
			  </tr>
<%			end if
		next %>
			  <tr>
				<td align="right"><%=yyCatWrNa%>:</td>
				<td><input type="text" name="secworkname" size="30" value="<%=htmlspecialsucode(sectionWorkingName)%>" /></td>
			  </tr>
			  <tr>
				<td align="right"><%=yyCatSub%>:</td>
				<td><select name="tsTopSection" size="1"><option value="0"><%=yyCatHom%></option>
				<%	foundcat=(topSection=0)
					if IsArray(alltopsections) then
						for index=0 to UBOUND(alltopsections,2)
							if alltopsections(0, index)<>sectionID then
								print "<option value=""" & alltopsections(0, index) & """"
								if topSection=alltopsections(0, index) then
									print " selected=""selected"""
									foundcat=true
								end if
								print ">" & alltopsections(1, index) & "</option>" & vbCrLf
							end if
						next
					end if
					if NOT foundcat then print "<option value="""" selected=""selected"">**undefined**</option>"
					%></select>
                </td>
			  </tr>
			  <tr>
				<td align="right"><%=yyCatFn%>:</td>
				<td><select name="catfunction" id="catfunction" size="1">
				  <option value="1"><%=yyCatPrd%></option>
				  <option value="0"<% if rootSection=0 then print " selected=""selected"""%>><%=yyCatCat%></option>
<%	if getpost("id")="0" OR NOT is_numeric(getpost("id")) then %>
				  <option value="2"<% if rootSection=2 then print " selected=""selected"""%>>Root Category</option>
<%	end if %>
				  </select></td>
			  </tr>
			  <tr>
				<td align="right"><%=yyCatImg%>:</td>
				<td><input type="text" name="secimage" id="secimage" size="30" value="<%=htmlspecials(sectionImage)%>" /> <input type="button" name="smallimup" value="..." onclick="uploadimage('secimage')" /></td>
			  </tr>
			  <tr>
				<td align="right"><%=yyCatDes%>:</td><td>
<%	if htmleditor="froala" then print "<div id=""secdescdiv"" class=""htmleditorcontainer"">" %>
					<textarea name="secdesc" id="sectionDescription" cols="48" rows="8"><%=sectionDescription%></textarea>
<%	if htmleditor="froala" then print "</div>" %>
				</td>
			  </tr>
<%		for index=2 to adminlanguages+1
			if (adminlangsettings AND 512)=512 then
				if getpost("act")="modify" then
					sSQL="SELECT sectionDescription" & index & " FROM sections WHERE sectionID="&getpost("id")
					rs.open sSQL,cnn,0,1
					sectionDescription=rs("sectionDescription" & index)
					rs.close
				else
					sectionDescription=""
				end if %>
			  <tr>
				<td align="right"><%=yyCatDes & " " & index %>:</td>
				<td>
<%	if htmleditor="froala" then print "<div id=""secdescdiv"&index&""" class=""htmleditorcontainer"">" %>
					<textarea name="secdesc<%=index%>" id="sectionDescription<%=index%>" cols="48" rows="8"><%=sectionDescription%></textarea>
<%	if htmleditor="froala" then print "</div>" %>
				</td>
			  </tr>
<%			end if
		next %>
			  <tr>
				<td align="right">Restrictions:</td>
				<td><select name="sectionDisabled" size="1">
				<option value="0"><%=yyNoRes%></option>
<%		for index=1 to maxloginlevels
			print "<option value="""&index&""""
			if sectionDisabled=index then print " selected=""selected"""
			print ">" & yyLiLev & " " & index & "</option>"
		next%>
				<option value="127"<% if sectionDisabled=127 then print " selected=""selected"""%>><%=yyDisCat%></option>
				</select>
<%		if getpost("act")="modify" then %>
				<select name="forcesubsection" size="1">
				<option value="0"><%=yySSForM%></option>
				<option value="1"><%=yySSForF%></option>
				<option value="2"><%=yySSForN%></option>
				</select>
<%		end if %>
				</td>
			  </tr>
			  <tr>
				<td align="right">Page Title Tag<%=" ("&yyOptnl&")"%>:</td>
				<td><input type="text" name="sTitle" size="40" value="<%=htmlspecials(sTitle)%>" onfocus="displaymultilangname('sTitle')" /><%
				for index=2 to adminlanguages+1
					if (adminlangsettings AND 2097152)=2097152 then
			%><input type="text" style="display:none;margin-top:2px" name="sTitle<%=index%>" id="sTitle<%=index%>" size="40" placeholder="Page Title Language <%=index%>" value="<%=htmlspecialsucode(IIfVr(index=2,sTitle2,sTitle3))%>" /><%
					end if
				next %></td>
			  </tr>
			  <tr>
				<td align="right">Meta Description<%=" ("&yyOptnl&")"%>:</td>
				<td><input type="text" name="sMetaDesc" size="40" value="<%=htmlspecials(sMetaDesc)%>" onfocus="displaymultilangname('sMetaDesc')" maxlength="250" /><%
				for index=2 to adminlanguages+1
					if (adminlangsettings AND 2097152)=2097152 then
			%><input type="text" style="display:none;margin-top:2px" name="sMetaDesc<%=index%>" id="sMetaDesc<%=index%>" size="40" placeholder="Meta Description Language <%=index%>" value="<%=htmlspecialsucode(IIfVr(index=2,sMetaDesc2,sMetaDesc3))%>" /><%
					end if
				next %></td>
			  </tr>
			  <tr>
				<td align="right"><%=yyCatURL&" ("&yyOptnl&")"%>:</td>
				<td><input type="text" name="sectionurl" size="40" value="<%=htmlspecials(sectionurl)%>" /></td>
			  </tr>
<%		for index=2 to adminlanguages+1
			if (adminlangsettings AND 2048)=2048 then %>
			  <tr>
				<td align="right"><%=yyCatURL&" "&index&" ("&yyOptnl&")"%>:</td>
				<td><input type="text" name="sectionurl<%=index%>" size="40" value="<%=htmlspecials(IIfVr(index=2,sectionurl2,sectionurl3))%>" /></td>
			  </tr>
<%			end if
		next %>
			  <tr>
				<td align="right">Catalog Root (Optional):</td>
				<td><input type="checkbox" name="catalogroot" value="ON" <%if catalogroot=sectionID then print "checked=""checked"" "%>/> Check to make this category the product catalog root.</td>
			  </tr>
			  <tr>
				<td align="right">Category Header:</td>
				<td>
<%	if htmleditor="froala" then print "<div id=""secheaddiv"" class=""htmleditorcontainer"">" %>
					<textarea name="sectionHeader" id="sectionHeader" cols="48" rows="8"><%=sectionHeader%></textarea>
<%	if htmleditor="froala" then print "</div>" %>
				</td>
			  </tr>
<%				for index=2 to adminlanguages+1
					if (adminlangsettings AND 524288)=524288 then
						if getpost("act")<>"addnew" then
							sSQL="SELECT sectionHeader" & index & " FROM sections WHERE sectionID="&getpost("id")
							rs2.Open sSQL,cnn,0,1
							sectionheader=rs2("sectionHeader" & index)
							rs2.Close
						end if
					%>
			  <tr>
				<td align="right"><%="Category Header" & " " & index%>:</td>
                <td>
<%	if htmleditor="froala" then print "<div id=""secheaddiv"&index&""" class=""htmleditorcontainer"">" %>
					<textarea name="sectionHeader<%=index%>" id="sectionHeader<%=index%>" cols="55" rows="8"><%=htmlspecials(sectionheader)%></textarea>
<%	if htmleditor="froala" then print "</div>" %>
				</td>
			  </tr>
<%					end if
				next %>
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><input type="submit" value="<%=yySubmit%>" /></td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="2"><br /><ul>
				  <li><%=yyCatEx1%></li>
				  <li><%=yyCatEx2%></li>
				  </ul></td>
			  </tr>
			  <tr>
                <td width="100%" colspan="2" align="center"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
                          &nbsp;</td>
			  </tr>
            </table>
		  </form>
<%	if htmleditor="ckeditor" then
		print "<script>"
		pathtovsadmin=request.servervariables("URL")
		slashpos=instrrev(pathtovsadmin, "/")
		if slashpos>0 then pathtovsadmin=left(pathtovsadmin, slashpos-1)
		print "function loadeditors(){"
		streditor="var sectionHeader=CKEDITOR.replace('sectionHeader',{extraPlugins : 'stylesheetparser,autogrow',autoGrow_maxHeight : 800,removePlugins : 'resize', toolbarStartupExpanded : false, toolbar : 'Basic', filebrowserBrowseUrl :'ckeditor/filemanager/browser/default/browser.html?Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserImageBrowseUrl : 'ckeditor/filemanager/browser/default/browser.html?Type=Image&Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserFlashBrowseUrl :'ckeditor/filemanager/browser/default/browser.html?Type=Flash&Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=File',filebrowserImageUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=Image',filebrowserFlashUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=Flash'});" & vbCrLf
		streditor=streditor & "sectionHeader.on('instanceReady',function(event){var myToolbar='Basic';event.editor.on( 'beforeMaximize', function(){if(event.editor.getCommand('maximize').state == CKEDITOR.TRISTATE_ON && myToolbar != 'Basic'){sectionHeader.setToolbar('Basic');myToolbar='Basic';sectionHeader.execCommand('toolbarCollapse');}else if(event.editor.getCommand('maximize').state == CKEDITOR.TRISTATE_OFF && myToolbar != 'Full'){sectionHeader.setToolbar('Full');myToolbar='Full';sectionHeader.execCommand('toolbarCollapse');}});event.editor.on('contentDom', function(e){event.editor.document.on('blur', function(){if(!sectionHeader.isToolbarCollapsed){sectionHeader.execCommand('toolbarCollapse');sectionHeader.isToolbarCollapsed=true;}});event.editor.document.on('focus',function(){if(sectionHeader.isToolbarCollapsed){sectionHeader.execCommand('toolbarCollapse');sectionHeader.isToolbarCollapsed=false;}});});sectionHeader.fire('contentDom');sectionHeader.isToolbarCollapsed=true;});"
		print streditor
		print replace(streditor, "sectionHeader", "sectionDescription")
		for index=2 to adminlanguages+1
			if (adminlangsettings AND 512)=512 then print replace(streditor, "sectionHeader", "sectionDescription" & index)
			if (adminlangsettings AND 524288)=524288 then print replace(streditor, "sectionHeader", "sectionHeader" & index)
		next
		print "}window.onload=function(){loadeditors();}"
		print "</script>" & vbCrLf
	elseif htmleditor="froala" then
		call displayfroalaeditor("sectionDescription",yyCatDes,".on('froalaEditor.focus',function(){expandckeditor(""secdescdiv"");})",FALSE,FALSE,1,FALSE)
		call displayfroalaeditor("sectionHeader","Category Header",".on('froalaEditor.focus',function(){expandckeditor(""secheaddiv"");})",FALSE,FALSE,1,FALSE)
		for index=2 to adminlanguages+1
			if (adminlangsettings AND 512)=512 then call displayfroalaeditor("sectionDescription"&index,yyCatDes&" (language "&index&")",".on('froalaEditor.focus',function(){expandckeditor(""secdescdiv"&index&""");})",FALSE,FALSE,1,FALSE)
			if (adminlangsettings AND 524288)=524288 then call displayfroalaeditor("sectionHeader"&index,"Category Header"&" (language "&index&")",".on('froalaEditor.focus',function(){expandckeditor(""secheaddiv"&index&""");})",FALSE,FALSE,1,FALSE)
		next
	end if
elseif getpost("act")="discounts" then
		sSQL="SELECT sectionName FROM sections WHERE sectionID="&getpost("id")
		rs.open sSQL,cnn,0,1
		thisname=rs("sectionName")
		rs.close
		alldata=""
		sSQL="SELECT cpaID,cpaCpnID,cpnWorkingName,cpnSitewide,cpnEndDate,cpnType FROM cpnassign INNER JOIN coupons ON cpnassign.cpaCpnID=coupons.cpnID WHERE cpaType=1 AND cpaAssignment='" & getpost("id") & "'"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then alldata=rs.GetRows
		rs.close
		alldata2=""
		tdt=Date()
		sSQL="SELECT cpnID,cpnWorkingName,cpnSitewide FROM coupons WHERE (cpnSitewide=0 OR cpnSitewide=3) AND cpnEndDate >=" & vsusdate(tdt)
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
		  <form name="mainform" method="post" action="admincats.asp">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="dodiscounts" />
			<input type="hidden" name="id" value="<%=getpost("id")%>" />
			<input type="hidden" name="pg" value="<%=getpost("pg")%>" />
            <table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="4" align="center"><strong><%=yyAssDis%> &quot;<%=thisname%>&quot;.</strong><br />&nbsp;</td>
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
<% elseif getpost("act")="changepos" then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%" align="center">
			<p>&nbsp;</p>
			<p>&nbsp;</p>
			<p>&nbsp;</p>
			<p><strong><%=yyUpdat%> . . . . . . . </strong></p>
			<p>&nbsp;</p>
			<p><%=yyNoFor%> <a href="admincats.asp?pg=<%=getpost("pg")%>"><%=yyClkHer%></a>.</p>
			<p>&nbsp;</p>
			<p>&nbsp;</p>
		  </td>
		</tr>
	  </table>
<% elseif getpost("posted")="1" AND getpost("act")<>"sort" AND success then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%><a href="admincats.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br /><br />&nbsp;
                </td>
			  </tr>
			</table></td>
        </tr>
	  </table>
<% elseif getpost("posted")="1" AND getpost("act")<>"sort" then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><span style="color:#FF0000;font-weight:bold"><%=yyOpFai%></span><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table></td>
        </tr>
	  </table>
<% else
cract=request.cookies("cract")
sortorder=request.cookies("catsort")
catorman=request.cookies("ccatorman")
allcoupon=""
sSQL="SELECT DISTINCT cpaAssignment FROM cpnassign WHERE cpaType=1"
rs.open sSQL,cnn,0,1
if NOT rs.EOF then allcoupon=rs.getrows
rs.close
modclone=request.cookies("modclone")
%>
<script>
/* <![CDATA[ */
function setCookie(c_name,value,expiredays){
	var exdate=new Date();
	exdate.setDate(exdate.getDate()+expiredays);
	document.cookie=c_name+ "=" +escape(value)+((expiredays==null) ? "" : ";expires="+exdate.toGMTString());
}
var rowsingrp=[];
function cpu(x,theid,grpid,secid){
	if(x.length>1) return;
	x.onchange= function(){chi(secid,x);};
	var totrows=rowsingrp[grpid];
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
	document.mainform.action="admincats.asp?catfun=<%=request("catfun")%>&stext=<%=urlencode(request("stext"))%>&sprice=<%=urlencode(request("sprice"))%>&stype=<%=request("stype")%>&scat=<%=request("scat")%>&pg=<%=IIfVr(getget("pg")="", 1, getget("pg"))%>";
	document.mainform.newval.value=obj.selectedIndex+1;
	document.mainform.id.value=id;
	document.mainform.act.value="changepos";
	document.mainform.submit();
}
function mr(id){
	document.mainform.action="admincats.asp";
	document.mainform.id.value=id;
	document.mainform.act.value="modify";
	document.mainform.submit();
}
function cr(id){
	document.mainform.action="admincats.asp";
	document.mainform.id.value=id;
	document.mainform.act.value="clone";
	document.mainform.submit();
}
function newrec(id){
	document.mainform.action="admincats.asp";
	document.mainform.id.value=id;
	document.mainform.act.value="addnew";
	document.mainform.submit();
}
function quickupdate(){
	if(document.mainform.cract.value=="del"){
		if(!confirm("<%=jscheck(yyConDel)%>\n"))
			return;
	}
	document.mainform.action="admincats.asp";
	document.mainform.act.value="quickupdate";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function changecract(obj){
	setCookie('cract',obj[obj.selectedIndex].value,600);
	startsearch();
}
function dsc(id) {
	document.mainform.action="admincats.asp";
	document.mainform.id.value=id;
	document.mainform.act.value="discounts";
	document.mainform.submit();
}
function dr(id){
if(confirm("<%=jscheck(yyConDel)%>\n")){
	document.mainform.action="admincats.asp";
	document.mainform.id.value=id;
	document.mainform.act.value="delete";
	document.mainform.submit();
}
}
function startsearch(){
	document.mainform.action="admincats.asp";
	document.mainform.act.value="search";
	document.mainform.posted.value="";
	document.mainform.submit();
}
function inventorymenu(){
	themenuitem=document.mainform.inventoryselect.options[document.mainform.inventoryselect.selectedIndex].value;
	if(themenuitem=="1") document.mainform.act.value="catinventory";
	document.mainform.action="dumporders.asp";
	document.mainform.submit();
}
function changesortorder(men){
	var selectedopt=men.options[men.selectedIndex].value;
	if(selectedopt=='act'){
		setCookie('cract','cor',600);
		setCookie('catsort','act',600);
	}
	document.mainform.action="admincats.asp<% if getpost("act")="search" OR getget("pg")<>"" then print "?pg=1"%>";
	document.mainform.id.value=selectedopt;
	document.mainform.act.value="sort";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function checkboxes(docheck){
	maxitems=document.getElementById("resultcounter").value;
	for(index=0;index<maxitems;index++){
		var thiselem=document.getElementById("chkbx"+index);
		if(thiselem.checked!=docheck){
			thiselem.checked=docheck;
			thiselem.onchange();
		}
	}
}
function switchcatorman(obj){
	document.location="admincats.asp?catorman="+obj[obj.selectedIndex].value+"&stext=<%=urlencode(request("stext"))%>&stype=<%=request("stype")%>&pg=<%=IIfVr(getget("pg")="" AND getpost("act")="search", 1, getget("pg"))%>";
}
function changemodclone(modclone){
	setCookie('modclone',modclone[modclone.selectedIndex].value,600);
	startsearch();
}
/* ]]> */
</script>
<%
thecat=request("scat")
if thecat<>"" then thecat=int(thecat)
if noadmincategorysearch<>TRUE then
	sSQL="SELECT sectionID,sectionWorkingName,topSection,rootSection FROM sections WHERE rootSection=0 ORDER BY sectionWorkingName"
	rs.open sSQL,cnn,0,1
	if rs.eof then
		success=false
	else
		alldata=rs.getrows
		success=true
	end if
	rs.close
end if %>
<h2><%=yyAdmCat%></h2>
		  <form name="mainform" method="post" action="admincats.asp">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="xxxxx" />
			<input type="hidden" name="id" value="xxxxx" />
			<input type="hidden" name="pg" value="<%=IIfVr(getpost("act")="search", "1", getget("pg"))%>" />
			<input type="hidden" name="newval" value="1" />
			<table class="cobtbl" width="100%" border="0" cellspacing="1" cellpadding="3">
			  <tr> 
				<td class="cobhl" width="25%" align="right"><%=yySrchFr%>:</td>
				<td class="cobll" width="25%"><input type="text" name="stext" size="20" value="<%=request("stext")%>" /></td>
				<td class="cobhl" width="25%" align="right"><%=replace(yyCatFn,"...","")%>:</td>
				<td class="cobll" width="25%"><select name="catfun" size="1">
					<option value=""><%=yySrchAC%></option>
					<option value="1"<% if request("catfun")="1" then print " selected=""selected"""%>><%=yyCatPrd%></option>
					<option value="2"<% if request("catfun")="2" then print " selected=""selected"""%>><%=yyCatCat%></option>
					<option value="3"<% if request("catfun")="3" then print " selected=""selected"""%>>Restricted Categories</option>
					<option value="4"<% if request("catfun")="4" then print " selected=""selected"""%>>Disabled Categories</option>
					<option value="5"<% if request("catfun")="5" then print " selected=""selected"""%>>Recommended Categories</option>
				</select></td>
			  </tr>
			  <tr>
				<td class="cobhl" width="25%" align="right"><%=yySrchTp%>:</td>
				<td class="cobll" width="25%"><select name="stype" size="1">
					<option value=""><%=yySrchAl%></option>
					<option value="any"<% if request("stype")="any" then print " selected=""selected"""%>><%=yySrchAn%></option>
					<option value="exact"<% if request("stype")="exact" then print " selected=""selected"""%>><%=yySrchEx%></option>
					</select>
				</td>
				<td class="cobhl" width="25%" align="right"><select size="1" name="catorman" onchange="switchcatorman(this)">
					<option value="cat"><%=yySrchCt%></option>
					<option value="non"<% if catorman="non" then print " selected=""selected"""%>><%=yyNone%></option>
					</select></td>
				<td class="cobll" width="25%">
<%					if catorman="non" then
						print "&nbsp;"
					else %>
					  <select name="scat" size="1">
					  <option value=""><%=yySrchAC%></option>
						<%	if IsArray(alldata) then
								call writemenulevel(0,1)
							end if %>
					  </select>
<%					end if %></td>
			  </tr>
			  <tr>
				<td class="cobhl" align="center"><%
				if getpost("act")="search" OR getget("pg")<>"" then
					if cract="del" OR cract="dis" OR cract="rec" then %>
					<input type="button" value="<%=yyCheckA%>" onclick="checkboxes(true)" /> <input type="button" value="<%=yyUCheck%>" onclick="checkboxes(false)" />
<%					end if
				end if %></td>
				<td class="cobll" colspan="3"><table width="100%" cellspacing="0" cellpadding="0" border="0">
					<tr>
					  <td class="cobll" align="center" style="white-space:nowrap">
						<select name="sort" size="1" onchange="changesortorder(this)">
						<option value="can"<% if sortorder="can" then print " selected=""selected"""%>>Sort - Cat Name</option>
						<option value="cwn"<% if sortorder="cwn" then print " selected=""selected"""%>>Sort - Working Name</option>
						<option value="act"<% if sortorder="act" then print " selected=""selected"""%>>Sort - Actual Order</option>
						<option value="pra"<% if sortorder="pra" then print " selected=""selected"""%>>Sort - Products Assigned</option>
<%		if useStockManagement then %>
						<option value="sta"<% if sortorder="sta" then print " selected=""selected"""%>>Sort - Stock Assigned</option>
<%		end if %>
						<option value="nsf"<% if sortorder="nsf" then print " selected=""selected"""%>>No Sort (Fastest)</option>
						</select>
						<input type="submit" value="List Categories" onclick="startsearch();" />
						<input type="button" value="<%=yyNewCat%>" onclick="newrec()" />
					  </td>
					  <td class="cobll" height="26" width="20%" align="right" style="white-space:nowrap">
					<select name="inventoryselect" size="1">
						<option value="1">Category Inventory</option>
					</select>&nbsp;<input type="button" value="<%=yyGo%>" onclick="inventorymenu();" />
					  </td>
					</tr>
				  </table></td>
			  </tr>
			</table>
<br />
            <table width="100%" class="stackable admin-table-a sta-white" id="catstable">
<%
sub displayheaderrow() %>
	<tr>
		<th class="small minicell">
	  <select name="cract" id="cract" size="1" onchange="changecract(this)" style="width:150px">
				<option value="none">Quick Entry...</option>
				<option value="" disabled="disabled">---------------------</option>
				<option value="cor"<% if cract="cor" OR cract="" then print " selected=""selected"""%>>Actual Order</option>
				<option value="" disabled="disabled">---------------------</option>
				<option value="can"<% if cract="can" then print " selected=""selected"""%>><%=yyCatNam%></option>
<%	for index=2 to adminlanguages+1
		if (adminlangsettings AND 256)=256 then print "<option value=""can" & index & """" & IIfVs(cract=("can"&index)," selected=""selected""") & ">" & yyCatNam & " " & index & "</option>"
	next %>
				<option value="cwn"<% if cract="cwn" then print " selected=""selected"""%>><%=yyCatWrNa%></option>
				<option value="cim"<% if cract="cim" then print " selected=""selected"""%>><%=yyCatImg%></option>
				<option value="dis"<% if cract="dis" then print " selected=""selected"""%>><%=yyDiscnt%></option>
				<option value="rec"<% if cract="rec" then print " selected=""selected"""%>><%=yyRecomd%></option>
				<option value="lol"<% if cract="lol" then print " selected=""selected"""%>><%="Restrictions"%></option>
				<option value="css"<% if cract="css" then print " selected=""selected"""%>><%="Custom CSS Class"%></option>
				<option value="cur"<% if cract="cur" then print " selected=""selected"""%>><%=yyCatURL%></option>
<%	for index=2 to adminlanguages+1
		if (adminlangsettings AND 2048)=2048 then print "<option value=""cur" & index & """" & IIfVs(cract=("cur"&index)," selected=""selected""") & ">" & yyCatURL & " " & index & "</option>"
	next %>
				<option value="" disabled="disabled">---------------------</option>
				<option value="ptt"<% if cract="ptt" then print " selected=""selected"""%>><%="Page Title Tag"%></option>
<%	for index=2 to adminlanguages+1
		if (adminlangsettings AND 2097152)=2097152 then print "<option value=""ptt" & index & """" & IIfVs(cract=("ptt"&index)," selected=""selected""") & ">" & "Page Title Tag" & " " & index & "</option>"
	next %>
				<option value="med"<% if cract="med" then print " selected=""selected"""%>><%="Meta Description"%></option>
<%	for index=2 to adminlanguages+1
		if (adminlangsettings AND 2097152)=2097152 then print "<option value=""med" & index & """" & IIfVs(cract=("med"&index)," selected=""selected""") & ">" & "Meta Description" & " " & index & "</option>"
	next %>
				<option value="" disabled="disabled">---------------------</option>
				<option value="vim"<% if cract="vim" then print " selected=""selected"""%>>View Category Image</option>
				<option value="" disabled="disabled">---------------------</option>
				<option value="del"<% if cract="del" then print " selected=""selected"""%>><%=yyDelete%></option>
				</select><%
	if cract="dis" then
		if is_numeric(request.cookies("ccurrdisc")) then
			currentdiscount=int(request.cookies("ccurrdisc"))
			rs2.open "SELECT cpnID FROM coupons WHERE (cpnSitewide=0 OR cpnSitewide=3) AND cpnID="&currentdiscount,cnn,0,1
			if rs2.EOF then currentdiscount=""
			rs2.close
		else
			currentdiscount=""
		end if
		print "<div style=""margin-top:2px""><select style=""width:150px"" name=""currentdiscount"" size=""1"" onchange=""setCookie('ccurrdisc',this[this.selectedIndex].value,600);changecract(document.getElementById('cract'))"">"
		sSQL="SELECT cpnID,cpnWorkingName FROM coupons WHERE cpnSitewide=0 OR cpnSitewide=3 ORDER BY cpnWorkingName"
		rs2.open sSQL,cnn,0,1
		if rs2.EOF then print "<option value="""" disabled=""disabled"">== No Assignable Discounts Defined ==</option>" & vbCrLf
		do while NOT rs2.EOF
			print "<option value=""" & rs2("cpnID") & """" & IIfVs(currentdiscount=rs2("cpnID")," selected=""selected""") & ">" & rs2("cpnWorkingName") & "</option>" & vbCrLf
			if currentdiscount="" then currentdiscount=rs2("cpnID")
			rs2.movenext
		loop
		rs2.close
		print "</select></div>"
	end if %>
		</th>
		<th class="maincell"><%=yyCatPat%></th>
		<th class="maincell"><%=yyCatNam%></th>
		<th class="minicell"><%="Products"%></th>
<%	if useStockManagement then %>
		<th class="minicell"><%=yyStck%></th>
<%	end if %>
		<th class="minicell"><%=yyDiscnt%></th>
		<th class="minicell"><%=yyModify%></th>
	</tr>
<%
end sub
	rowsingrp="" : jscript="" : qetype="" : qesize=""
	rowcounter=0
	if catalogroot<>0 then
		sSQL="SELECT sectionID FROM sections WHERE sectionID=" & catalogroot
		rs.open sSQL,cnn,0,1
		if rs.EOF then cnn.execute("UPDATE admin SET catalogRoot=0 WHERE adminID=1")
		rs.close
	end if
	if getpost("act")="search" OR getget("pg")<>"" then
		CurPage=1
		roottopsection=0
		if menucategoriesatroot then
			sSQL="SELECT topSection FROM sections WHERE sectionID="&catalogroot
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then roottopsection=rs("topSection")
			rs.close
		end if
		if is_numeric(getget("pg")) then CurPage=int(getget("pg"))
		if NOT menucategoriesatroot then
			sSQL="SELECT sec1.sectionID,sec1.sectionName,sec1.sectionName2,sec1.sectionName3,sec1.sCustomCSS,sec1.sectionWorkingName,sec1.sectionImage,sec1.sRecommend,sec1.sTitle,sec1.sTitle2,sec1.sTitle3,sec1.sMetaDesc,sec1.sMetaDesc2,sec1.sMetaDesc3,sec1.sectionurl,sec1.sectionurl2,sec1.sectionurl3,sec1.sectionDescription,sec1.topSection AS topSection,sec1.rootSection,sec1.sectionDisabled,sec1.sectionOrder FROM sections AS sec1 LEFT JOIN sections AS sec2 ON sec1.topSection=sec2.sectionID"
		elseif mysqlserver=TRUE then
			sSQL="SELECT sec1.sectionID,sec1.sectionName,sec1.sectionName2,sec1.sectionName3,sec1.sCustomCSS,sec1.sectionWorkingName,sec1.sectionImage,sec1.sRecommend,sec1.sTitle,sec1.sTitle2,sec1.sTitle3,sec1.sMetaDesc,sec1.sMetaDesc2,sec1.sMetaDesc3,sec1.sectionurl,sec1.sectionurl2,sec1.sectionurl3,sec1.sectionDescription,IF(sec1.topSection="&catalogroot&","&roottopsection&",sec1.topSection) AS topSection,sec1.rootSection,sec1.sectionDisabled,sec1.sectionOrder FROM sections AS sec1 LEFT JOIN sections AS sec2 ON IF(sec1.topSection="&catalogroot&","&roottopsection&",sec1.topSection)=sec2.sectionID"
		elseif sqlserver=TRUE then
			sSQL="SELECT sec1.sectionID,sec1.sectionName,sec1.sectionName2,sec1.sectionName3,sec1.sCustomCSS,sec1.sectionWorkingName,sec1.sectionImage,sec1.sRecommend,sec1.sTitle,sec1.sTitle2,sec1.sTitle3,sec1.sMetaDesc,sec1.sMetaDesc2,sec1.sMetaDesc3,sec1.sectionurl,sec1.sectionurl2,sec1.sectionurl3,sec1.sectionDescription,CASE WHEN sec1.topSection="&catalogroot&" THEN "&roottopsection&" ELSE sec1.topSection END AS topSection,sec1.rootSection,sec1.sectionDisabled,sec1.sectionOrder FROM sections sec1 LEFT JOIN sections sec2 ON CASE WHEN sec1.topSection="&catalogroot&" THEN "&roottopsection&" ELSE sec1.topSection END=sec2.sectionID"
		else
			sSQL="SELECT sec1.sectionID,sec1.sectionName,sec1.sectionName2,sec1.sectionName3,sec1.sCustomCSS,sec1.sectionWorkingName,sec1.sectionImage,sec1.sRecommend,sec1.sTitle,sec1.sTitle2,sec1.sTitle3,sec1.sMetaDesc,sec1.sMetaDesc2,sec1.sMetaDesc3,sec1.sectionurl,sec1.sectionurl2,sec1.sectionurl3,sec1.sectionDescription,IIF(sec1.topSection="&catalogroot&","&roottopsection&",sec1.topSection) AS topSection,sec1.rootSection,sec1.sectionDisabled,sec1.sectionOrder FROM sections AS sec1 LEFT JOIN sections AS sec2 ON IIf(sec1.topSection="&catalogroot&","&roottopsection&",sec1.topSection)=sec2.sectionID"
		end if
		whereand=" WHERE "
		if sortorder="pra" OR sortorder="sta" then
			sSQL="SELECT sectionID,sectionName,sectionName2,sectionName3,sCustomCSS,sectionWorkingName,sectionImage,sRecommend,sTitle,sTitle2,sTitle3,sMetaDesc,sMetaDesc2,sMetaDesc3,sectionurl,sectionurl2,sectionurl3,sectionDescription,topSection,rootSection,sectionDisabled,"&IIfVs(useStockManagement,"SUM(pInStock) AS sumStock,")&"COUNT(pSection) AS sectionCount FROM sections AS sec1 LEFT JOIN products ON sec1.sectionID=products.pSection WHERE rootSection=1"
			whereand=" AND "
		end if
		if thecat<>"" then
			returnalltopsections=TRUE
			sectionids=getsectionids(thecat, TRUE)
			if sectionids<>"" then
				sSQL=sSQL & whereand & "sec1.sectionID IN (" & sectionids & ") "
			end if
			whereand=" AND "
		end if
		if trim(request("stext"))<>"" then
			Xstext=escape_string(request("stext"))
			aText=Split(Xstext)
			if nosearchadmindescription then maxsearchindex=0 else maxsearchindex=1
			aFields(0)=getlangid("sectionName",256)
			aFields(1)=getlangid("sectionDescription",512)
			if request("stype")="exact" then
				sSQL=sSQL & whereand & "(sec1.sectionWorkingName LIKE '%"&Xstext&"%' OR "
				for index=1 to adminlanguages+1
					sSQL=sSQL & "sec1.sectionName"&IIfVr(index=1,"",index)&" LIKE '%"&Xstext&"%' OR sec1.sectionDescription"&IIfVr(index=1,"",index)&" LIKE '%"&Xstext&"%'"
					if index<adminlanguages+1 then sSQL=sSQL & " OR "
				next
				sSQL=sSQL & ") "
				whereand=" AND "
			else
				sJoin="AND "
				if request("stype")="any" then sJoin="OR "
				sSQL=sSQL & whereand&"("
				whereand=" AND "
				for index=0 to maxsearchindex
					sSQL=sSQL & "("
					for rowcounter=0 to UBOUND(aText)
						sSQL=sSQL & "sec1." & aFields(index) & " LIKE '%"&aText(rowcounter)&"%' "
						if rowcounter<UBOUND(aText) then sSQL=sSQL & sJoin
					next
					sSQL=sSQL & ") "
					if index < maxsearchindex then sSQL=sSQL & "OR "
				next
				sSQL=sSQL & ") "
			end if
		end if
		rs.open "SELECT COUNT(*) AS totcats FROM sections",cnn,0,1
		totalcats=rs("totcats")
		rs.close
		if isnull(totalcats) then totalcats=0
		if request("catfun")="1" then sSQL=sSQL & whereand & "sec1.rootSection=1 " : whereand=" AND "
		if request("catfun")="2" then sSQL=sSQL & whereand & "sec1.rootSection=0 " : whereand=" AND "
		if request("catfun")="3" then sSQL=sSQL & whereand & "sec1.sectionDisabled<>0 " : whereand=" AND "
		if request("catfun")="4" then sSQL=sSQL & whereand & "sec1.sectionDisabled=127 " : whereand=" AND "
		if request("catfun")="5" then sSQL=sSQL & whereand & "sec1.sRecommend<>0 " : whereand=" AND "
		if sortorder="can" then
			sSQL=sSQL & " ORDER BY sec1.sectionName"
		elseif sortorder="cwn" then
			sSQL=sSQL & " ORDER BY sec1.sectionWorkingName"
		elseif sortorder="nsf" then
			' Nothing
		elseif sortorder="pra" then
			sSQL=sSQL & " GROUP BY sectionID,sectionName,sectionName2,sectionName3,sCustomCSS,sectionWorkingName,sectionImage,sRecommend,sTitle,sTitle2,sTitle3,sMetaDesc,sMetaDesc2,sMetaDesc3,sectionurl,sectionurl2,sectionurl3,sectionDescription,topSection,rootSection,sectionDisabled ORDER BY COUNT(pSection)"
		elseif sortorder="sta" then
			sSQL=sSQL & " GROUP BY sectionID,sectionName,sectionName2,sectionName3,sCustomCSS,sectionWorkingName,sectionImage,sRecommend,sTitle,sTitle2,sTitle3,sMetaDesc,sMetaDesc2,sMetaDesc3,sectionurl,sectionurl2,sectionurl3,sectionDescription,topSection,rootSection,sectionDisabled ORDER BY SUM(pInStock)"
		else
			sSQL=sSQL & " ORDER BY sec2.sectionOrder,sec2.sectionID,sec1.sectionOrder"
		end if
		currgroup=-1
		rsCats.CursorLocation=3 ' adUseClient
		rsCats.CacheSize=admincatsperpage
		rsCats.Open sSQL,cnn
		if NOT rsCats.EOF then
			rsCats.MoveFirst
			rsCats.PageSize=admincatsperpage
			rsCats.AbsolutePage=CurPage
			islooping=false
			noproducts=false
			hascatinprodsection=false
			totnumrows=rsCats.RecordCount
			iNumOfPages=int((totnumrows + (admincatsperpage-1)) / admincatsperpage)
			pblink="<a href=""admincats.asp?scat="&request("scat")&"&stext="&server.urlencode(request("stext"))&"&stype="&request("stype")&"&catfun="&request("catfun")&"&pg="
			if iNumOfPages > 1 then print "<tr><td align=""center"" colspan=""7"">" & writepagebar(CurPage,iNumOfPages,yyPrev,yyNext,pblink,FALSE) & "<br /><br /></td></tr>"
			displayheaderrow()
			currgroup=rsCats("topSection")
			ordingroup=1
			rowsingrp="" : jscript=""
			checkfirstgroup=(CurPage<>1)
			do while NOT rsCats.EOF AND rowcounter < admincatsperpage
				if currgroup<>rsCats("topSection") then
					if checkfirstgroup then
						rs.open "SELECT COUNT(*) AS catcnt FROM sections WHERE topSection="&currgroup,cnn,0,1
						if NOT rs.EOF then ordingroup=rs("catcnt")+1
						rs.close
						checkfirstgroup=FALSE
					end if
					if sortorder="" OR sortorder="act" then print "<tr><td colspan=""7"">&nbsp;</td></tr>"
					rowsingrp=rowsingrp&"rowsingrp["&currgroup&"]="&(ordingroup-1)&";"
					currgroup=rsCats("topSection")
					ordingroup=1
				end if
				jscript=jscript&"pa["&rowcounter&"]=[" %>
<tr id="tr<%=rowcounter%>"><td class="minicell"><%
				qetype="text"
				qesize="18"
				if cract="can" then
					jscript=jscript&"'"&jsspecials(rsCats("sectionName"))&"'"
				elseif cract="can2" then
					jscript=jscript&"'"&jsspecials(rsCats("sectionName2"))&"'"
				elseif cract="can3" then
					jscript=jscript&"'"&jsspecials(rsCats("sectionName3"))&"'"
				elseif cract="cwn" then
					jscript=jscript&"'"&jsspecials(rsCats("sectionWorkingName"))&"'"
				elseif cract="cim" then
					jscript=jscript&"'"&jsspecials(rsCats("sectionImage"))&"'"
				elseif cract="ptt" then
					jscript=jscript&"'"&jsspecials(rsCats("sTitle"))&"'"
				elseif cract="ptt2" then
					jscript=jscript&"'"&jsspecials(rsCats("sTitle2"))&"'"
				elseif cract="ptt3" then
					jscript=jscript&"'"&jsspecials(rsCats("sTitle3"))&"'"
				elseif cract="med" then
					jscript=jscript&"'"&jsspecials(rsCats("sMetaDesc"))&"'"
				elseif cract="med2" then
					jscript=jscript&"'"&jsspecials(rsCats("sMetaDesc2"))&"'"
				elseif cract="med3" then
					jscript=jscript&"'"&jsspecials(rsCats("sMetaDesc3"))&"'"
				elseif cract="rec" then
					jscript=jscript&IIfVr(cint(rsCats("sRecommend"))<>0,1,0)
					qetype="checkbox"
				elseif cract="cur" then
					jscript=jscript&"'"&jsspecials(rsCats("sectionurl"))&"'"
				elseif cract="cur2" then
					jscript=jscript&"'"&jsspecials(rsCats("sectionurl2"))&"'"
				elseif cract="cur3" then
					jscript=jscript&"'"&jsspecials(rsCats("sectionurl3"))&"'"
				elseif cract="cor" OR cract="" then
					qetype="special"
					jscript=jscript&"''"
					if sortorder="" OR sortorder="act" then
						currpos=vrmin(totalcats,vrmax(1,rsCats("sectionOrder")))
						print "<select onmouseover=""cpu(this,"&ordingroup&","&rsCats("topSection")&","&rsCats("sectionID")&")"">"
						print "<option value="""&currpos&""">"&ordingroup&IIfVr(ordingroup<100,"&nbsp;","")&"</option>"
						print "</select>"
					else
						print "&nbsp;"
					end if
				elseif cract="del" then
					jscript=jscript&"'del'"
					qetype="delbox"
				elseif cract="dis" AND currentdiscount<>"" then
					sSQL="SELECT cpaID FROM cpnassign WHERE cpaType=1 AND cpaAssignment='"&escape_string(rsCats("sectionID"))&"' AND cpaCpnID="&currentdiscount
					rs2.open sSQL,cnn,0,1
					jscript=jscript&IIfVr(rs2.EOF,0,1)
					rs2.close
					qetype="checkbox"
				elseif cract="vim" then
					jscript=jscript&"'"
					if trim(rsCats("sectionImage")&"")<>"" then jscript=jscript&IIfVs(lcase(left(rsCats("sectionImage"),5))<>"http:" AND lcase(left(rsCats("sectionImage"),6))<>"https:" AND left(rsCats("sectionImage"),1)<>"/","../") & jsspecials(rsCats("sectionImage"))
					jscript=jscript&"'"
					qetype="image"
				elseif cract="lol" then
					jscript=jscript&rsCats("sectionDisabled")
					qetype="loginlevel"
				elseif cract="css" then
					jscript=jscript&"'"&jsspecials(rsCats("sCustomCSS"))&"'"
					qesize="16"
				else
					qetype=""
				end if %></td><%
				if cint(rsCats("rootSection"))=0 then
					sumStock="-"
					sectionCount="-"
				elseif sortorder="pra" OR sortorder="sta" then
					if useStockManagement then sumStock=rsCats("sumStock") else sumStock=0
					sectionCount=rsCats("sectionCount")
				else
					sSQL="SELECT "&IIfVs(useStockManagement,"SUM(pInStock) AS sumStock,")&"COUNT(pSection) AS sectionCount FROM products WHERE pSection=" & rsCats("sectionID")
					rs2.open sSQL,cnn,0,1
					if NOT rs2.EOF then
						if useStockManagement then sumStock=rs2("sumStock") else sumStock=0
						sectionCount=rs2("sectionCount")
					end if
					rs2.close
				end if %></td><td class="maincell" style="font-size:10px"><%
				tslist=""
				thetopts=rsCats("topSection")
				for index=0 to 10
					if thetopts=0 then
						if len(tslist)>3 then tslist=right(tslist,len(tslist)-3)
						exit for
					elseif index=10 then
						tslist="<span style=""color:#FF0000;font-weight:bold"">"&yyLoop&"</span>" & tslist
						islooping=true
					else
						sSQL="SELECT sectionID,topSection,sectionWorkingName,rootSection FROM sections WHERE sectionID=" & thetopts
						rs.open sSQL,cnn,0,1
						if NOT rs.EOF then
							errstart=""
							errend=""
							if rs("rootSection")=1 then
								errstart="<span style=""color:#FF0000;font-weight:bold"">"
								errend="</span>"
								hascatinprodsection=true
							end if
							tslist=" " & raquo & " " & errstart & rs("sectionWorkingName") & errend & tslist
							thetopts=rs("topSection")
						else
							tslist="<span style=""color:#FF0000;font-weight:bold"">"&yyTopDel&"</span>" & tslist
							rs.close
							exit for
						end if
						rs.close
					end if
				next
				print tslist & "</td><td>"
				if int(CStr(rsCats("rootSection")))=1 then print "<strong>"
				if int(CStr(rsCats("sectionDisabled")))=127 then print "<span style=""color:#FF0000;text-decoration:line-through"">"
				if catalogroot=rsCats("sectionID") then print "<span title=""Catalog Root"" style=""padding-left:10px;text-decoration:underline overline;font-size:larger;"">"
				print rsCats("sectionWorkingName") & " (" & rsCats("sectionID") & ")"
				if catalogroot=rsCats("sectionID") then print "</span>"
				if int(CStr(rsCats("sectionDisabled")))=127 then print "</span>"
				if int(CStr(rsCats("rootSection")))=1 then print "</strong>"
				hascoupon="0"
				if isarray(allcoupon) then
					for index=0 to UBOUND(allcoupon,2)
						if int(allcoupon(0,index))=rsCats("sectionID") then
							hascoupon="1"
							exit for
						end if
					next
				end if
		%></td><td class="minicell"><%=sectionCount %></td><%
		if useStockManagement then print "<td class=""minicell"">" & sumStock & "</td>"
		%><td>-</td><td>-</td></tr>
<%				jscript=jscript&","&rsCats("sectionID")&"," & hascoupon & "];"&vbCrLf
				rowcounter=rowcounter+1
				ordingroup=ordingroup+1
				rsCats.MoveNext
			loop
			if iNumOfPages > 1 then print "<tr><td align=""center"" colspan=""7""><br />" & writepagebar(CurPage,iNumOfPages,yyPrev,yyNext,pblink,FALSE) & "</td></tr>"
			if islooping then %>
				  <tr><td width="100%" colspan="7"><br /><span style="color:#FF0000;font-weight:bold">** </span><%=yyCatEx3%></td></tr>
<%			end if
			if hascatinprodsection then %>
				  <tr><td width="100%" colspan="7"><br /><ul><li><%=yyCPErr%></li></ul></td></tr>
<%			end if %>
				  <tr><td width="100%" colspan="7"><br /><ul><li><%=yyCatEx4%></li></ul></td></tr>
<%		else %>
				  <tr><td width="100%" colspan="7" align="center"><br /><strong><%=yyCatEx5%><br />&nbsp;</td></tr>
<%		end if
		rsCats.Close
		rs.open "SELECT COUNT(*) AS catcnt FROM sections WHERE topSection="&currgroup,cnn,0,1
		if NOT rs.EOF then ordingroup=rs("catcnt")+1
		rs.close
		rowsingrp=rowsingrp&"rowsingrp["&currgroup&"]="&(ordingroup-1)&";"
	else
		if trim(detlinkspacechar)<>"" then
			rowcounter=0
			jscript=""
			sSQL="SELECT sectionID,sectionWorkingName,sectionDescription,sectionURL,rootSection,sectionDisabled,0 AS sectionOrder FROM sections WHERE sectionURL LIKE '%"&IIfVr(mysqlserver,"\"&escape_string(detlinkspacechar),"["&escape_string(detlinkspacechar)&"]")&"%'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				print "<tr><td colspan=""7"" style=""color:#FF0000"">You have the detlinkspacechar parameter set as &quot;" & detlinkspacechar & "&quot; but have categories where the Static URL uses this character and these will not display properly. Consider removing the detlinkspacechar parameter, or replacing it with a space in the Static URL for these products.</td></tr>"
				displayheaderrow()
				do while NOT rs.EOF
					jscript=jscript&"pa[" & rowcounter & "]=[" & rs("sectionID") & ",0];" & vbCrLf %>
<tr id="tr<%=rowcounter%>"><td class="minicell">-</td><td><%=rs("sectionURL")%></td><td><%=rs("sectionWorkingName")%></td><td>-</td><td>-</td></tr>
<%					rowcounter=rowcounter+1
					rs.MoveNext
				loop
			end if
			rs.close
		end if
		numitems=0
		sSQL="SELECT COUNT(*) as totcount FROM sections"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			numitems=rs("totcount")
		end if
		rs.close
		print "<tr><td colspan=""7""><div class=""itemsdefine"">You have " & numitems & " categories defined.</div></td></tr>"
	end if %>
			  <tr style="height:50px">
				<td align="center" style="white-space:nowrap"><% if rowcounter>0 AND cract<>"" AND cract<>"none" AND cract<>"cor" then print "<input type=""hidden"" name=""resultcounter"" id=""resultcounter"" value=""" & rowcounter & """ /><input type=""button"" value=""" & yyUpdate & """ onclick=""quickupdate()"" /> <input type=""reset"" value=""" & yyReset & """ />" else print "&nbsp;"%></td>
                <td width="100%" colspan="5" align="center"><a href="admin.asp"><strong><%=yyAdmHom%></strong></a></td>
				<td>&nbsp;</td>
			  </tr>
            </table>
		  </form>
<script>
/* <![CDATA[ */
var pa=[];
<%	if qetype="image" then %>
function imageisvisible(img){
    var rect=img.getBoundingClientRect();
    return(rect.top<=(window.innerHeight||document.documentElement.clientHeight));
}
function checkvisibleimages(){
	var lzimgs=document.getElementById('catstable').getElementsByClassName('lazyload');
	var tarray=[];
	if(lzimgs.length==0){
		removeEventListener('scroll',setcheckvisibletimeout);
	}else{
		for(var lzi=0; lzi<lzimgs.length; lzi++){
			var telem=lzimgs[lzi];
			if(imageisvisible(telem)){
				telem.src=telem.getAttribute('data-src');
				tarray.push(telem.id);
			}
		}
		for(var lzi=0; lzi<tarray.length; lzi++){
			document.getElementById(tarray[lzi]).className='lazydoneload';
		}
	}
}
var checkvisibletimeout='';
function setcheckvisibletimeout(){
	if(checkvisibletimeout!='') clearTimeout(checkvisibletimeout);
	checkvisibletimeout=setTimeout(checkvisibleimages,300)
}
addEventListener('scroll',setcheckvisibletimeout);
addEventListener('load',checkvisibleimages);
<%	end if %>
<%=jscript%>
function createloginlevel(pid,sid){
	var mup='<select size="1" id="sec'+pid+'" onchange="this.name=\'pra_'+patch_pid(pidind)+'\'"><option value="0">No Restriction</option>';
	for(var cli=1; cli<=<%=maxloginlevels%>; cli++){
		mup+='<option value="'+cli+'"' + (cli==sid?' selected="selected"':'') + '>Login Level '+cli+'</option>';
	}
	mup+='<option value="127"' + (sid==127?' selected="selected"':'') + '>Disable Category</option>';
	return(mup+'</select>');
}
function patch_pid(pid){
	document.getElementById('pid'+pid).name='pid'+pid;
	document.getElementById('pid'+pid).value=pa[pid][1];
	return pid;
}
for(var pidind in pa){
	var ttr=document.getElementById('tr'+pidind);
	var stockcell=<% print IIfVr(useStockManagement,1,0) %>;
	ttr.cells[1].innerHTML+='<input type="hidden" id="pid'+pidind+'" value="" />';
	ttr.cells[4+stockcell].style.textAlign='center';
	ttr.cells[5+stockcell].style.textAlign='center';
	ttr.cells[5+stockcell].style.whiteSpace='nowrap';
	ttr.cells[4+stockcell].innerHTML='<input type="button" '+(pa[pidind][2]?' style="color:#F4E64B"':'')+' value="<% print jsescape(htmlspecials(yyAssign))%>" onclick="dsc(\''+pa[pidind][1]+'\')" />';
	ttr.cells[5+stockcell].innerHTML='<input type="button" value="M" style="width:30px;margin-right:4px" onclick="mr(\''+pa[pidind][1]+'\')" title="<% print jsescape(htmlspecials(yyModify))%>" />' +
		'<input type="button" value="C" style="width:30px;margin-right:4px" onclick="cr(\''+pa[pidind][1]+'\')" title="<% print jsescape(htmlspecials(yyClone))%>" />' +
		'<input type="button" value="X" style="width:30px" onclick="dr(\''+pa[pidind][1]+'\')" title="<% print jsescape(htmlspecials(yyDelete))%>" />';
<%	if cract<>"cor" AND cract<>"" then %>
	ttr.cells[0].innerHTML=
<%		if qetype="text" then %>
	pa[pidind][0]===false?'-':'<input type="text" id="chkbx'+pidind+'" size="<% print qesize%>" onchange="this.name=\'pra_'+patch_pid(pidind)+'\'" value="'+pa[pidind][0].replace('"','&quot;')+'" tabindex="'+(pidind+1)+'" />';
<%		elseif qetype="delbox" then %>
	'<input type="checkbox" id="chkbx'+pidind+'" onchange="this.name=\'pra_'+patch_pid(pidind)+'\'" value="del" tabindex="'+(pidind+1)+'" />';
<%		elseif qetype="checkbox" then %>
	'<input type="hidden" id="pra_'+pa[pidind][1]+'" value="1" /><input type="checkbox" id="chkbx'+pidind+'" onchange="this.name=\'prb_'+patch_pid(pidind)+'\';document.getElementById(\'pra_'+pa[pidind][1]+'\').name=\'pra_'+patch_pid(pidind)+'\'" value="1" '+(pa[pidind][0]==1?'checked="checked" ':'')+'tabindex="'+(pidind+1)+'" />';
<%		elseif qetype="image" then %>
	(pa[pidind][0]==''?'-':'<img class="lazyload" id="lazyimg'+pidind+'" src="adminimages/imageload.png" data-src="'+pa[pidind][0]+'" style="max-width:80px;cursor:pointer" alt="" onclick="mr(\''+pa[pidind][1]+'\')" />');
<%		elseif qetype="loginlevel" then %>
	createloginlevel(pidind,pa[pidind][0]);
<%		else %>
	'&nbsp;';
<%		end if
	end if %>
}
<%=rowsingrp%>
/* ]]> */
</script>
<% end if
cnn.Close
set rs=nothing
set rs2=nothing
set rsCats=nothing
set cnn=nothing
%>