<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protect under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
Dim sSQL,rs,alldata,success,cnn,rowcounter,allsections,alloptions,errmsg,prodoptions,aFields(6),dorefresh,thecat
success=true
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs2=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
themask=cStr(DateSerial(2003,12,11))
themask=replace(themask,"2003","yyyy")
themask=replace(themask,"12","mm")
themask=replace(themask,"11","dd")
dorefresh=FALSE
rtprodid=""
sub updateratings()
	sSQL="UPDATE products SET pNumRatings=0,pTotRating=0"
	ect_query(sSQL)
	sSQL="SELECT rtProdID,COUNT(*) AS numratings,SUM(rtRating) AS totrating FROM ratings WHERE rtApproved<>0 GROUP BY rtProdID"
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		numratings=rs("numratings")
		totrating=rs("totrating")
		if isnull(numratings) then numratings=0
		if isnull(totrating) then totrating=0
		sSQL="UPDATE products SET pNumRatings="&numratings&",pTotRating="&totrating&" WHERE pID='"& replace(rs("rtProdID"), "'", "''") & "'"
		ect_query(sSQL)
		rs.movenext
	loop
	rs.close
end sub
if getpost("posted")="1" then
	if getpost("act")="delete" then
		sSQL="SELECT rtProdID FROM ratings WHERE rtID=" & getpost("id")
		rs.open sSQL,cnn,0,1
			rtprodid=trim(rs("rtProdID")&"")
		rs.close
		dorefresh=TRUE
		sSQL="DELETE FROM ratings WHERE rtID=" & getpost("id")
		ect_query(sSQL)
	elseif getpost("act")="quickupdate" then
		for each objItem in request.form
			if left(objItem, 4)="pra_" then
				theid=right(objItem, len(objItem)-4)
				theval=getpost(objItem)
				pract=getpost("pract")
				sSQL=""
				if pract="del" then
					sSQL="DELETE FROM ratings"
				elseif pract="app" then
					sSQL="UPDATE ratings SET rtApproved=" & IIfVr(getpost("prb_"&theid)="1","1","0")
				elseif pract="pby" then
					sSQL="UPDATE ratings SET rtPosterName='" & escape_string(theval) & "'"
				elseif pract="pid" then
					sSQL="UPDATE ratings SET rtProdID='" & escape_string(theval) & "'"
				elseif pract="hed" then
					sSQL="UPDATE ratings SET rtHeader='" & escape_string(theval) & "'"
				elseif pract="rat" then
					sSQL="UPDATE ratings SET rtRating=" & theval
				end if
				if sSQL<>"" then
					sSQL=sSQL & " WHERE rtID="&theid
					ect_query(sSQL)
				end if
			end if
		next
		if getpost("pract")="del" OR getpost("pract")="app" OR getpost("pract")="rat" then
			call updateratings()
		end if
		if success then dorefresh=TRUE else errmsg=yyPOErr & "<br />" & errmsg
	elseif getpost("act")="domodify" then
		sSQL="UPDATE ratings SET " & _
			"rtProdID='"&replace(getpost("rtprodid"),"'","")&"'," & _
			"rtRating="&getpost("rtrating")&"," & _
			"rtApproved="&IIfVr(getpost("rtapproved")="yes", 1, 0)&"," & _
			"rtLanguage="&IIfVr(getpost("rtlanguage")<>"", getpost("rtlanguage"), 0)&"," & _
			"rtIPAddress='"&getpost("rtipaddress")&"'," & _
			"rtPosterName='"&escape_string(getpost("rtpostername"))&"'," & _
			"rtPosterEmail='"&escape_string(getpost("rtposteremail"))&"'," & _
			"rtDate=" & vsusdate(datevalue(getpost("rtdate"))) & "," & _
			"rtHeader='"&escape_string(getpost("rtheader"))&"'," & _
			"rtComments='"&escape_string(getpost("rtcomments"))&"' " & _
			"WHERE rtID="&getpost("id")
		' print sSQL
		ect_query(sSQL)
		rtprodid=getpost("rtprodid")
		dorefresh=TRUE
	elseif getpost("act")="doaddnew" then
		thedate=getpost("rtdate")
		if thedate<>"" then
			err.number=0
			on error resume next
			thedate=DateValue(thedate)
			if err.number <> 0 then
				thedate=Date()
			end if
			on error goto 0
		end if
		thedate=vsusdate(thedate)
		sSQL="INSERT INTO ratings (rtProdID,rtRating,rtDate,rtApproved,rtIPAddress,rtPosterName,rtPosterEmail,rtHeader,rtComments) VALUES (" & _
			"'"&replace(getpost("rtprodid"),"'","")&"'," & _
			getpost("rtrating")&"," & _
			thedate&"," & _
			IIfVr(getpost("rtapproved")="yes", 1, 0)&"," & _
			"'"&getpost("rtipaddress")&"'," & _
			"'"&escape_string(getpost("rtpostername"))&"'," & _
			"'"&escape_string(getpost("rtposteremail"))&"'," & _
			"'"&escape_string(getpost("rtheader"))&"'," & _
			"'"&escape_string(getpost("rtcomments"))&"')"
		ect_query(sSQL)
		rtprodid=getpost("rtprodid")
		dorefresh=TRUE
	elseif getpost("act")="updateratings" then
		print "<p align=""center"">" & yyUpdat & "...</p>"
		response.flush()
		call updateratings()
		dorefresh=TRUE
	end if
elseif getget("approve")="yes" then
	sSQL="UPDATE ratings SET rtApproved=1 WHERE rtID=" & getpost("id")
	ect_query(sSQL)
	sSQL="SELECT rtProdID FROM ratings WHERE rtID=" & getpost("id")
	rs.open sSQL,cnn,0,1
		rtprodid=trim(rs("rtProdID")&"")
	rs.close
elseif getget("unapprove")="yes" then
	sSQL="UPDATE ratings SET rtApproved=0 WHERE rtID=" & getpost("id")
	ect_query(sSQL)
	sSQL="SELECT rtProdID FROM ratings WHERE rtID=" & getpost("id")
	rs.open sSQL,cnn,0,1
		rtprodid=trim(rs("rtProdID")&"")
	rs.close
end if
if rtprodid<>"" then
	numratings=0
	totrating=0
	sSQL="SELECT COUNT(*) AS numratings, SUM(rtRating) AS totrating FROM ratings WHERE rtApproved<>0 AND rtProdID='" & rtprodid & "'"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		numratings=rs("numratings")
		totrating=rs("totrating")
		if isnull(numratings) OR isnull(totrating) then
			numratings=0
			totrating=0
		end if
	end if
	rs.close
	sSQL="UPDATE products SET pNumRatings="&numratings&",pTotRating="&totrating&" WHERE pID='"& replace(rtprodid, "'", "''") &"'"
	ect_query(sSQL)
end if
if dorefresh then
	print "<meta http-equiv=""refresh"" content=""1; url=adminratings.asp"
	print "?stext=" & urlencode(getpost("stext")) & "&mindate=" & request("mindate") & "&maxdate=" & request("maxdate") & "&stype=" & getpost("stype") & "&approved=" & getpost("approved") & "&pg=" & getpost("pg")
	print """>"
end if
if getpost("act")="modify" OR getpost("act")="addnew" OR getpost("act")="clone" then %>
<script src="popcalendar.js"></script>
<script>
try{languagetext('<%=adminlang%>');}catch(err){}
function formvalidator(theForm){
  return (true);
}
</script>
<%		if getpost("act")="modify" OR getpost("act")="clone" then
			rtID=getpost("id")
			sSQL="SELECT rtProdID,rtRating,rtDate,rtApproved,rtLanguage,rtIPAddress,rtPosterName,rtPosterEmail,rtHeader,rtComments FROM ratings WHERE rtID=" & rtID
			rs.open sSQL,cnn,0,1
				rtProdID=rs("rtProdID")
				rtRating=rs("rtRating")
				rtDate=rs("rtDate")
				rtApproved=rs("rtApproved")
				rtLanguage=rs("rtLanguage")
				rtIPAddress=rs("rtIPAddress")
				rtPosterName=rs("rtPosterName")
				rtPosterEmail=rs("rtPosterEmail")
				rtHeader=rs("rtHeader")
				rtComments=rs("rtComments")
			rs.close
			sSQL="SELECT pName FROM products WHERE pID='" & replace(rtProdID,"'","") & "'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then pName=rs("pName") else pName="Rating Not Found"
			rs.close
		else
			rtID=""
			rtProdID=""
			rtRating=0
			rtDate=Date()
			rtApproved=0
			rtLanguage=0
			rtIPAddress=left(request.servervariables("REMOTE_ADDR"), 32)
			rtPosterName=""
			rtPosterEmail=""
			rtHeader=""
			rtComments=""
			pName=""
		end if
%>
	<form name="mainform" method="post" action="adminratings.asp" onsubmit="return formvalidator(this)">
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
		<tr>
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
			<%	if getpost("act")="modify" then %>
			<input type="hidden" name="act" value="domodify" />
			<input type="hidden" name="id" value="<%=rtID%>" />
			<%	else %>
			<input type="hidden" name="act" value="doaddnew" />
			<%	end if
				call writehiddenvar("stock", getpost("stock"))
				call writehiddenvar("stext", getpost("stext"))
				call writehiddenvar("mindate", getpost("mindate"))
				call writehiddenvar("maxdate", getpost("maxdate"))
				call writehiddenvar("approved", getpost("approved"))
				call writehiddenvar("stype", getpost("stype"))
				call writehiddenvar("pg", getpost("pg")) %>
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="4" align="center"><strong><%=IIfVr(getpost("act")="clone",yyClone&": ",IIfVs(getpost("act")="modify",yyModify&": ")) & yyMPRRev%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
			    <td align="right"><%=redasterix&yyPrId%>:</td><td><input type="text" name="rtprodid" size="15" value="<%=rtProdID%>" /></td>
			    <td align="right"><%=redasterix&yyRatn%>:</td><td><select size="1" name="rtrating"><option value=""><%=yySelect%></option><%
						for rowcounter=0 to 10
							print "<option value='"&rowcounter&"'"
							if rowcounter=rtRating then print " selected=""selected"""
							print ">"&(rowcounter/2)& " "&yyStars&"</option>" &vbCrLf
						next %></select></td>
			  </tr>
			  <tr>
			    <td align="right"><%=redasterix&yyRatDat%>:</td><td><div style="position:relative;display:inline"><input type="text" name="rtdate" size="20" value="<%=rtDate%>" style="vertical-align:middle" /> <input type="button" onclick="popUpCalendar(this, document.forms.mainform.rtdate, '<%=themask%>', -205)" value="DP" /></div></td>
			    <td align="right"><%=redasterix&yyAppd%>:</td><td><select size="1" name="rtapproved"><option value="no"><%=yyNo%></option>"
				  <option value="yes"<% if rtApproved<>0 then print " selected=""selected"""%>><%=yyYes%></option></select></td>
			  </tr>
			  <tr>
			    <td align="right"><%=redasterix&yyPostBy%>:</td><td><input type="text" name="rtpostername" size="25" value="<%=htmlspecials(rtPosterName)%>" /></td>
			    <td align="right"><%=redasterix&yyIPAdd%>:</td><td><input type="text" name="rtipaddress" size="25" value="<%=htmlspecials(rtIPAddress)%>" /></td>
			  </tr>
			  <tr>
			    <td align="right"><%=redasterix&yyHeadi%>:</td><td<%=IIfVr(adminlanguages>0, "", " colspan=""3""")%>><input type="text" name="rtheader" size="35" value="<%=htmlspecials(rtHeader)%>" /></td>
<%			if adminlanguages>0 then %>
			    <td align="right"><%=redasterix&yyLanID%>:</td><td><select name="rtlanguage" size="1">
					<option value="0">1</option>
					<option value="1" <% if rtLanguage=1 then print "selected=""selected"""%>>2</option>
<%				if adminlanguages>1 then %>
					<option value="2" <% if rtLanguage=2 then print "selected=""selected"""%>>3</option>
<%				end if %>
					</select></td>
<%			end if %>
			  </tr>
			  <tr>
			    <td align="right"><%=redasterix&yyComme%>:</td><td colspan="3"><textarea name="rtcomments" cols="65" rows="8" wrap=virtual><%=htmlspecials(rtComments)%></textarea></td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="4">
				  <p>&nbsp;</p>
                  <p align="center"><input type="submit" value="<%=yySubmit%>" />&nbsp;&nbsp;<input type="reset" value="<%=yyReset%>" /></p>
                </td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
	</form>
<% elseif getpost("posted")="1" AND success then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <a href="adminratings.asp<%
							print "?rid="&getpost("rid")&"&stock="&getpost("stock")&"&stext=" & urlencode(getpost("stext")) & "&mindate=" & getpost("mindate") & "&maxdate=" & getpost("maxdate") & "&stype=" & getpost("stype") & "&approved=" & getpost("approved") & "&pg=" & getpost("pg")
						%>"><strong><%=yyClkHer%></strong></a>.<br />
                        <br />&nbsp;
                </td>
			  </tr>
			</table></td>
        </tr>
      </table>
<% elseif getpost("posted")="1" then %>
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
<%	else
		jscript=""
		pract=request.cookies("practrat")
		modclone=request.cookies("modclone") %>
<script src="popcalendar.js"></script>
<script>
<!--
try{languagetext('<%=adminlang%>');}catch(err){}
function setCookie(c_name,value,expiredays){
	var exdate=new Date();
	exdate.setDate(exdate.getDate()+expiredays);
	document.cookie=c_name+ "=" +escape(value)+((expiredays==null) ? "" : ";expires="+exdate.toGMTString());
}
function mr(id){
	document.mainform.action="adminratings.asp";
	document.mainform.id.value=id;
	document.mainform.act.value="modify";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function cr(id){
	document.mainform.action="adminratings.asp";
	document.mainform.id.value=id;
	document.mainform.act.value="clone";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function aprec(id){
	document.mainform.action="adminratings.asp?approve=yes";
	document.mainform.id.value=id;
	document.mainform.act.value="search";
	document.mainform.posted.value="";
	document.mainform.submit();
}
function newrec(id){
	document.mainform.action="adminratings.asp";
	document.mainform.id.value=id;
	document.mainform.act.value="addnew";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function dscnts(id){
	document.mainform.action="adminratings.asp";
	document.mainform.id.value=id;
	document.mainform.act.value="discounts";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function startsearch(){
	document.mainform.action="adminratings.asp";
	document.mainform.act.value="search";
	document.mainform.stock.value="";
	document.mainform.posted.value="";
	document.mainform.submit();
}
function quickupdate(){
	if(document.mainform.pract.value=="del"){
		if(!confirm("<%=jscheck(yyConDel)%>\n"))
			return;
	}
	document.mainform.action="adminratings.asp";
	document.mainform.act.value="quickupdate";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function dr(id){
if(confirm("<%=jscheck(yyConDel)%>\n")){
	document.mainform.action="adminratings.asp";
	document.mainform.id.value=id;
	document.mainform.act.value="delete";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
}
function updateratings(){
if(confirm("<%=jscheck(yySureCa)%>\n")){
	document.mainform.action="adminratings.asp";
	document.mainform.act.value="updateratings";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
}
function changepract(obj){
	setCookie('practrat',obj[obj.selectedIndex].value,600);
	startsearch();
}
function checkboxes(docheck){
	if(document.getElementById("resultcounter")){
		maxitems=document.getElementById("resultcounter").value;
		for(index=0;index<maxitems;index++){
			document.getElementById("chkbx"+index).checked=docheck;
		}
	}
}
function changemodclone(modclone){
	setCookie('modclone',modclone[modclone.selectedIndex].value,600);
	startsearch();
}
// -->
</script>
<h2><%=yyAdmRat%></h2>
		<form name="mainform" method="post" action="adminratings.asp">
		  <td width="100%">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="xxxxx" />
			<input type="hidden" name="stock" value="" />
			<input type="hidden" name="id" value="xxxxx" />
			<input type="hidden" name="pg" value="<%=IIfVr(getpost("act")="search", "1", getget("pg"))%>" />
			<table class="cobtbl" width="100%" border="0" cellspacing="1" cellpadding="3">
			  <tr> 
				<td class="cobhl" width="20%" align="right"><%=yySrchFr%>:</td>
				<td class="cobll" width="30%"><input type="text" name="stext" size="20" value="<%=request("stext")%>" /></td>
				<td class="cobhl" width="20%" align="right"><%=yyDatRan%>:</td>
				<td class="cobll" width="30%"><div style="position:relative;display:inline"><input type="text" name="mindate" size="10" value="<%=request("mindate")%>" style="vertical-align:middle" />&nbsp;<input type="button" onclick="popUpCalendar(this, document.forms.mainform.mindate, '<%=themask%>', -205)" value="DP" />&nbsp;<%=yyTo%>:&nbsp;<input type="text" name="maxdate" size="10" value="<%=request("maxdate")%>" style="vertical-align:middle" />&nbsp;<input type="button" onclick="popUpCalendar(this, document.forms.mainform.maxdate, '<%=themask%>', -205)" value="DP" /></div></td>
			  </tr>
			  <tr>
				<td class="cobhl"align="right"><%=yySrchTp%>:</td>
				<td class="cobll"><select name="stype" size="1">
					<option value=""><%=yySrchAl%></option>
					<option value="any"<% if request("stype")="any" then print " selected=""selected"""%>><%=yySrchAn%></option>
					<option value="exact"<% if request("stype")="exact" then print " selected=""selected"""%>><%=yySrchEx%></option>
					</select>
				</td>
				<td class="cobhl"align="right"><%=yyAppd%>:</td>
				<td class="cobll">
				  <select name="approved" size="1">
				  <option value="2"<% if request("approved")="2" then print " selected=""selected"""%>><%=yyAll%></option>
				  <option value=""<% if request("approved")="" then print " selected=""selected"""%>><%=yyNotApp%></option>
				  <option value="1"<% if request("approved")="1" then print " selected=""selected"""%>><%=yyAppd%></option>
				  </select>
				</td>
			  </tr>
			  <tr>
				<td class="cobhl"><%
					if pract="del" OR pract="app" then %>
						<input type="button" value="<%=yyCheckA%>" onclick="checkboxes(true);" /> <input type="button" value="<%=yyUCheck%>" onclick="checkboxes(false);" />
<%					else
						print "&nbsp;"
					end if %></td>
				<td class="cobll" colspan="3"><table width="100%" cellspacing="0" cellpadding="0" border="0">
					<tr>
					  <td class="cobll" align="center"><input type="button" value="<%=yyLiRat%>" onclick="startsearch();" /> 
						<input type="button" value="<%=yyNewRat%>" onclick="newrec();" />
						<input type="button" value="<%=yyUpdPrR%>" onclick="updateratings();" />
					  </td>
					  <td class="cobll" height="26" width="20%" align="right">&nbsp;</td>
					</tr>
				  </table></td>
			  </tr>
			</table>
<br />
            <table width="100%" class="stackable admin-table-a sta-white">
<%
		sub displayprodrow(xrs)
			%><tr id="tr<%=resultcounter%>"<% if cint(xrs("rtApproved"))=0 then print " style=""color:#FF0000"""%>><td class="minicell"><%
				if pract="pby" then
					print "<input type=""text"" id=""chkbx"&resultcounter&""" size=""18"" name=""pra_"&xrs("rtID")&""" value=""" & xrs("rtPosterName") & """ tabindex="""&(resultcounter+1)&"""/>"
				elseif pract="pid" then
					print "<input type=""text"" id=""chkbx"&resultcounter&""" size=""18"" name=""pra_"&xrs("rtID")&""" value=""" & xrs("rtProdID") & """ tabindex="""&(resultcounter+1)&"""/>"
				elseif pract="hed" then
					print "<input type=""text"" id=""chkbx"&resultcounter&""" size=""18"" name=""pra_"&xrs("rtID")&""" value=""" & xrs("rtHeader") & """ tabindex="""&(resultcounter+1)&"""/>"
				elseif pract="rat" then
					print "<select size=""1"" name=""pra_"&xrs("rtID")&""">"
					for rowcounter=0 to 10
						print "<option value='"&rowcounter&"'"
						if rowcounter=xrs("rtRating") then print " selected=""selected"""
						print ">"&(rowcounter/2)& " "&yyStars&"</option>" &vbCrLf
					next
					print "</select>"
				elseif pract="app" then
					print "<input type=""hidden"" name=""pra_"&xrs("rtID")&""" value="""" /><input type=""checkbox"" id=""chkbx"&resultcounter&""" name=""prb_"&xrs("rtID")&""" value=""1"" tabindex="""&(resultcounter+1)&"""" & IIfVr(xrs("rtApproved")," checked=""checked""","") & "/>"
				elseif pract="del" then
					print "<input type=""checkbox"" id=""chkbx"&resultcounter&""" name=""pra_"&xrs("rtID")&""" value=""del"" tabindex="""&(resultcounter+1)&"""/>"
				else
					print "&nbsp;"
				end if
			%></td><td><%
					thelink=""
					if NOT isnull(xrs("pName")) then thelink="../" & getdetailsurl(xrs("rtProdID"),xrs("pStaticPage"),xrs("pName"),trim(xrs("pStaticURL")&""),"","")
					print IIfVs(thelink<>"","<a href=""" & thelink & """ target=""viewdetails"">") & xrs("rtProdID") & IIfVs(thelink<>"","</a>")
			%></td><td><%="<a href=""javascript:mr(" & xrs("rtID") & ")"">" & htmlspecials(xrs("rtPosterName")) & "</a>"
			%></td><td><%=htmlspecials(xrs("rtIPAddress"))
			%></td><td><%=xrs("rtDate")
			%></td><td><%
					thecomments=xrs("rtHeader")
					print htmlspecials(left(thecomments, 180))
					if len(thecomments)>180 then print "..."
			%></td><td class="minicell"><%=cint(xrs("rtRating"))/2
			%></td><td class="minicell"><%
					if cint(xrs("rtApproved"))=0 then
			%><input type="button" value="<%=yyAppro%>" onclick="aprec('<%=replace(replace(xrs("rtID"),"\","\\"),"'","\'")%>')" /><%
					else
						print "&nbsp;"
					end if
			%></td><td>-</td></tr><%
			print vbCrLf
		end sub
		sub displayheaderrow() %>
			<tr>
				<th class="minicell">
					<select name="pract" id="pract" size="1" onchange="changepract(this)">
					<option value="none">Quick Entry...</option>
					<option value="pby"<% if pract="pby" then print " selected=""selected"""%>><%=yyPostBy%></option>
					<option value="pid"<% if pract="pid" then print " selected=""selected"""%>><%=yyPrId%></option>
					<option value="hed"<% if pract="hed" then print " selected=""selected"""%>><%=yyHeadi%></option>
					<option value="rat"<% if pract="rat" then print " selected=""selected"""%>><%=yyRatn%></option>
					<option value="app"<% if pract="app" then print " selected=""selected"""%>><%=yyAppd%></option>
					<option value="" disabled="disabled">------------------</option>
					<option value="del"<% if pract="del" then print " selected=""selected"""%>><%=yyDelete%></option>
					</select></th>
				<th><%=replace(yyPrId," ","&nbsp;")%></th>
				<th><%=yyPostBy%></th>
				<th><%=replace(yyIPAdd," ","&nbsp;")%></th>
				<th><%=replace(yyDateAd," ","&nbsp;")%></th>
				<th><%=yyHeadi%></th>
				<th class="minicell"><%=yyRatn%></th>
				<th class="minicell"><%=yyAppro%></th>
				<th class="minicell"><%=yyModify%></th>
			</tr>
<%		end sub
		whereand=" WHERE "
		sSQL="SELECT rtID,rtProdID,rtRating,rtDate,rtApproved,rtIPAddress,rtPosterName,rtPosterEmail,rtHeader,pStaticPage,pName,pStaticURL FROM ratings LEFT JOIN products on ratings.rtProdID=products.pID"
		if trim(request("approved"))<>"2" then
			if trim(request("approved"))="" then sSQL=sSQL & whereand & "rtApproved=0" else sSQL=sSQL & whereand & "rtApproved<>0"
			whereand=" AND "
		end if
		mindate=trim(request("mindate"))
		maxdate=trim(request("maxdate"))
		if mindate<>"" OR maxdate<>"" then
			if mindate<>"" then
				err.number=0
				on error resume next
				themindate=DateValue(mindate)
				if err.number <> 0 then
					themindate=""
				end if
				on error goto 0
			end if
			if maxdate<>"" then
				err.number=0
				on error resume next
				themaxdate=DateValue(maxdate)
				if err.number <> 0 then
					themaxdate=""
				end if
				on error goto 0
			end if
			if themindate<>"" AND themaxdate<>"" then
				sSQL=sSQL & whereand & "rtDate BETWEEN " & vsusdate(themindate) & " AND " & vsusdate(themaxdate+1)
				whereand=" AND "
			elseif themindate<>"" then
				sSQL=sSQL & whereand & "rtDate >= " & vsusdate(themindate)
				whereand=" AND "
			elseif themaxdate<>"" then
				sSQL=sSQL & whereand & "rtDate <= " & vsusdate(themaxdate)
				whereand=" AND "
			end if
		end if
		if trim(request("stext"))<>"" then
			sText=escape_string(request("stext"))
			aText=Split(sText)
			aFields(0)="rtID"
			aFields(1)="rtProdID"
			aFields(2)="rtIPAddress"
			aFields(3)="rtPosterName"
			aFields(4)="rtPosterEmail"
			aFields(5)="rtHeader"
			aFields(6)="rtComments"
			if request("stype")="exact" then
				sSQL=sSQL & whereand & "(rtID LIKE '%"&sText&"%' OR rtProdID LIKE '%"&sText&"%' OR rtIPAddress LIKE '%"&sText&"%' OR rtPosterName LIKE '%"&sText&"%' OR rtPosterEmail LIKE '%"&sText&"%' OR rtHeader LIKE '%"&sText&"%' OR rtComments LIKE '%"&sText&"%') "
				whereand=" AND "
			else
				sJoin="AND "
				if request("stype")="any" then sJoin="OR "
				sSQL=sSQL & whereand&"("
				whereand=" AND "
				for index=0 to 6
					sSQL=sSQL & "("
					for rowcounter=0 to UBOUND(aText)
						sSQL=sSQL & aFields(index) & " LIKE '%"&aText(rowcounter)&"%' "
						if rowcounter<UBOUND(aText) then sSQL=sSQL & sJoin
					next
					sSQL=sSQL & ") "
					if index < 6 then sSQL=sSQL & "OR "
				next
				sSQL=sSQL & ") "
			end if
		end if
		sSQL=sSQL & " ORDER BY rtDate DESC"
		if adminproductsperpage="" then adminproductsperpage=200
		rs.CursorLocation=3 ' adUseClient
		rs.CacheSize=adminproductsperpage
		rs.open sSQL, cnn
		if rs.eof or rs.bof then
			success=false
			iNumOfPages=0
		else
			success=true
			rs.MoveFirst
			rs.PageSize=adminproductsperpage
			CurPage=1
			if is_numeric(getget("pg")) then CurPage=int(getget("pg"))
			iNumOfPages=Int((rs.RecordCount + (adminproductsperpage-1)) / adminproductsperpage)
			rs.AbsolutePage=CurPage
		end if
		resultcounter=0
		if NOT rs.EOF then
			pblink="<a href=""adminratings.asp?stock="&request("stock")&"&approved="&request("approved")&"&stext="&urlencode(request("stext"))&"&stype="&request("stype")&"&mindate="&request("mindate")&IIfVr(request("maxdate")<>"","&maxdate="&request("maxdate"),"")&"&pg="
			if iNumOfPages > 1 then print "<tr><td colspan=""9"" align=""center"">" & writepagebar(CurPage,iNumOfPages,yyPrev,yyNext,pblink,FALSE) & "</td></tr>"
			displayheaderrow()
			addcomma=""
			do while NOT rs.EOF AND resultcounter<rs.PageSize
				jscript=jscript&"pa["&resultcounter&"]=["
				displayprodrow(rs)
				addcomma=","
				jscript=jscript&rs("rtID")&"];"&vbCrLf
				resultcounter=resultcounter + 1
				rs.MoveNext
			loop
			if iNumOfPages > 1 then print "<tr><td colspan=""9"" align=""center"">" & writepagebar(CurPage,iNumOfPages,yyPrev,yyNext,pblink,FALSE) & "</td></tr>"
		else
			print "<tr><td width=""100%"" colspan=""9"" align=""center""><br />"&yyItNone&"<br />&nbsp;</td></tr>"
		end if
		rs.close
%>			  <tr>
				<td align="center" style="white-space:nowrap"><% if resultcounter>0 AND pract<>"" AND pract<>"none" then print "<input type=""hidden"" name=""resultcounter"" id=""resultcounter"" value="""&resultcounter&""" /><input type=""button"" value="""&yyUpdate&""" onclick=""quickupdate()"" /> <input type=""reset"" value="""&yyReset&""" />" else print "&nbsp;"%></td>
                <td width="100%" colspan="7" align="center"><br /><a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br /></td>
				<td>&nbsp;</td>
			  </tr>
            </table>
		  </form>
<script>
/* <![CDATA[ */
var pa=[];
<%=jscript%>
for(var pidind in pa){
	var ttr=document.getElementById('tr'+pidind);
	ttr.cells[8].style.textAlign='center';
	ttr.cells[8].style.whiteSpace='nowrap';
	ttr.cells[8].innerHTML='<input type="button" value="M" style="width:30px;margin-right:4px" onclick="mr(\''+pa[pidind][0]+'\')" title="<%=jsescape(htmlspecials(yyModify))%>" />' +
		'<input type="button" value="C" style="width:30px;margin-right:4px" onclick="cr(\''+pa[pidind][0]+'\')" title="<%=jsescape(htmlspecials(yyClone))%>" />' +
		'<input type="button" value="X" style="width:30px" onclick="dr(\''+pa[pidind][0]+'\')" title="<%=jsescape(htmlspecials(yyDelete))%>" />';
}
/* ]]> */
</script>
<% end if
cnn.Close
set rs=nothing
set cnn=nothing
%>
