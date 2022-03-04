<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
success=TRUE
if mailinglistpurgedays="" then mailinglistpurgedays=32
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
haserror=FALSE
iscanadapost=TRUE ' This just tells the function callxmlfunction() to return a response even if the response code is not 200 ok
if lcase(adminencoding)="iso-8859-1" then raquo="»" else raquo=">"
resultcounter=0
sub writemenulevel(id,itlevel)
	dim wmlindex
	if itlevel<10 then
		for wmlindex=0 TO ubound(alldata,2)
			if alldata(2,wmlindex)=id then
				print "<option value='"&alldata(0,wmlindex)&"'"
				if IsArray(selstatus) then
					for ii=0 to UBOUND(selstatus)
						if int(selstatus(ii))=alldata(0,wmlindex) then print " selected=""selected"""
					next
				end if
				print ">"
				for index=0 to itlevel-2
					print raquo & " "
				next
				print alldata(1,wmlindex)&"</option>" & vbCrLf
				if alldata(3,wmlindex)=0 then call writemenulevel(alldata(0,wmlindex),itlevel+1)
			end if
		next
	end if
end sub
sub hiddenparams()
	call writehiddenvar("stext", getpost("stext"))
	call writehiddenvar("listem", getpost("listem"))
	call writehiddenvar("mindate", getpost("mindate"))
	call writehiddenvar("maxdate", getpost("maxdate"))
	call writehiddenvar("stype", getpost("stype"))
	call writehiddenvar("pg", getpost("pg"))
	call writehiddenvar("id", getpost("id"))
	call writehiddenvar("ordstate", request("ordstate"))
	call writehiddenvar("ordcountry", request("ordcountry"))
	call writehiddenvar("smanufacturer", request("smanufacturer"))
	call writehiddenvar("scat", request("scat"))
	call writehiddenvar("stsearch", request("stsearch"))
	call writehiddenvar("swholesale", request("swholesale"))
	call writehiddenvar("sortorder", request("sort"))
end sub
function checkmailchimperror(res)
	checkmailchimperror=TRUE
	if res<>"" then
		mctype="" : mctitle="" : mcdetail=""
		set cmejsonobj=new JSONobject
		set outputobj=cmejsonobj.parse(res)
		mctype=cmejsonobj.value("type")
		mctitle=cmejsonobj.value("title")
		mcdetail=cmejsonobj.value("detail")
		if mctype<>"" AND mctitle<>"" AND mcdetail<>"" then
			checkmailchimperror=FALSE
			print "<div style=""padding:30px;color:#f00;text-align:center"">The following error occurred: " & mcdetail & "</div>"
			print "<div style=""padding:30px;text-align:center""><input type=""button"" value=""Go Back And Try Again"" onclick=""history.go(-1)""</div>"
		end if
	end if
end function
if getpost("posted")="1" then
	if getpost("act")="confirm" then
		sSQL="UPDATE mailinglist SET isconfirmed=1 WHERE email='" & escape_string(getpost("id")) & "'"
		ect_query(sSQL)
	elseif getpost("act")="delete" then
		sSQL="DELETE FROM mailinglist WHERE email='" & escape_string(getpost("id")) & "'"
		ect_query(sSQL)
		dorefresh=TRUE
	elseif getpost("act")="doaddnew" then
		sSQL="INSERT INTO mailinglist (email,mlname,isconfirmed,mlConfirmDate,mlIPAddress) VALUES ('" & lcase(escape_string(getpost("email"))) & "','" & escape_string(getpost("mlname")) & "',1," & vsusdate(date())&",'"&left(request.servervariables("REMOTE_ADDR"), 48)&"')"
		on error resume next
		ect_query(sSQL)
		on error goto 0
		dorefresh=TRUE
	elseif getpost("act")="domodify" then
		if lcase(getpost("email"))<>lcase(getpost("id")) then
			sSQL="SELECT email FROM mailinglist WHERE email='" & lcase(escape_string(getpost("email"))) & "'"
			rs.open sSQL,cnn,0,1
			haserror=NOT rs.EOF
			rs.close
		end if
		if haserror then
			errormessage="Cannot rename email from &quot;" & htmlspecials(getpost("id")) & "&quot; to &quot;" & htmlspecials(getpost("email")) & "&quot; as that address is already in use."
		else
			sSQL="UPDATE mailinglist SET email='" & lcase(escape_string(getpost("email"))) & "',mlname='" & escape_string(getpost("mlname")) & "' WHERE email='" & lcase(escape_string(getpost("id"))) & "'"
			ect_query(sSQL)
			dorefresh=TRUE
		end if
	elseif getpost("act")="purgeunconfirmed" then
		sSQL="DELETE FROM mailinglist WHERE isconfirmed=0 AND mlConfirmDate<" & vsusdate(date()-mailinglistpurgedays)
		ect_query(sSQL)
		dorefresh=TRUE
	elseif getpost("act")="clearsent" then
		sSQL="UPDATE mailinglist SET emailsent=0"
		ect_query(sSQL)
		dorefresh=TRUE
	elseif getpost("act")="quickupdate" then
		for each objItem in request.form
			if left(objItem, 4)="pra_" then
				origid=right(objItem, len(objItem)-4)
				theid=getpost("pid"&origid)
				theval=getpost(objItem)
				mlact=getpost("mlact")
				sSQL=""
				if mlact="del" then
					if theval="del" then sSQL="DELETE FROM mailinglist"
				end if
				if sSQL<>"" then
					sSQL=sSQL & " WHERE email='"&escape_string(theid)&"'"
					ect_query(sSQL)
				end if
			end if
		next
		dorefresh=TRUE
	end if
end if
if dorefresh then
	print "<meta http-equiv=""refresh"" content=""1; url=adminmailinglist.asp"
	print "?stext=" & urlencode(getpost("stext"))
	print "&ordstate=" & urlencode(getpost("ordstate"))
	print "&ordcountry=" & urlencode(getpost("ordcountry"))
	print "&smanufacturer=" & urlencode(getpost("smanufacturer"))
	print "&scat=" & urlencode(getpost("scat"))
	print "&stsearch=" & getpost("stsearch")
	print "&swholesale=" & getpost("swholesale")
	print "&sortorder=" & getpost("sortorder")
	print "&stype=" & getpost("stype")
	print "&listem=" & urlencode(getpost("listem"))
	print "&mindate=" & getpost("mindate")
	print "&maxdate=" & getpost("maxdate")
	print "&mlact=" & getpost("mlact")
	print "&pg=" & getpost("pg")
	print """>"
end if
if getpost("posted")="1" AND haserror then
	print "<div style=""padding:50px;text-align:center"">" & yySorErr & "</div>"
	print "<div style=""padding:50px;text-align:center;color:#FF1010"">" & errormessage & "</div>"
	print "<div style=""padding:50px;text-align:center""><input type=""button"" onclick=""history.go(-1)"" value=""" & yyClkBac & """ /></div>"
elseif getpost("posted")="1" AND getpost("act")="dosendem" then
	server.scripttimeout=1800
	breatherseconds=IIfVr(debugmode,10,300)
%>
<script>
/* <![CDATA[ */
function breatherfunction(){
	var breathersecs=document.getElementById('breathersecs');
	document.getElementById('emerrordiv').innerHTML='<%="Taking breather. Sending next batch in <span id=""breathersecs"">"&breatherseconds&"</span> seconds."%><br /><br />'+document.getElementById('emerrordiv').innerHTML;
	setTimeout('document.getElementById(\'breatherform\').submit();',<%=(breatherseconds*1000)%>);
	setInterval('document.getElementById(\'breathersecs\').innerHTML=Math.max(parseInt(document.getElementById(\'breathersecs\').innerHTML)-1,0)',1000);
}
function nextbatchfunction(){
	document.getElementById('emerrordiv').innerHTML+='<div style="text-align:center"><input type="button" value="Send Next Batch Now" onclick="document.getElementById(\'breatherform\').submit();" /></div>';
	document.getElementById('totalsentsofar').value='';
}
/* ]]> */
</script>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
		  <td width="100%" align="center">
		  <form name="mainform" method="post" action="adminmailinglist.asp" onsubmit="return formvalidator(this)">
<%			call writehiddenvar("posted", "1")
			call writehiddenvar("act", "dosendem")
			call hiddenparams() %>
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
			  <tr>
                <td align="center"><strong><%=yyMaLiMa%> - Sending Emails</strong></td>
			  </tr>
			  <tr> 
                <td align="center"><br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;<br />
				<strong>Please do not refresh this page</strong>
				<br />&nbsp;<br />Sending email: <span name="sendspan" id="sendspan">1</span>
				<br />&nbsp;<br />&nbsp;<br />&nbsp;<br /><div id="emerrordiv"></div><br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;
				</td>
			  </tr>
			  <tr>
                <td align="center"><br />
                          <a href="adminmailinglist.asp"><strong><%=yyAdmHom%></strong></a><br />
                          &nbsp;</td>
			  </tr>
            </table>
		  </form>
		  </td>
        </tr>
      </table>
<%	if getpost("emformat")="1" OR getpost("emformat")="2" then htmlemails=TRUE else htmlemails=FALSE
	batchesof=getpost("batchesof")
	takebreather=getpost("takebreather")="ON"
	totalsentsofar=getpost("totalsentsofar")
	takingbreather=FALSE
	if NOT is_numeric(totalsentsofar) then totalsentsofar=0
	response.cookies("EMAILBATCHNUM")=batchesof
	response.cookies("EMAILBATCHNUM").Expires=Date()+1000
	if request.servervariables("HTTPS")="on" then response.cookies("EMAILBATCHNUM").secure=TRUE
	if NOT is_numeric(batchesof) then batchesof=0 else batchesof=clng(batchesof)
	if htmlemails=TRUE then emlNl="<br />" else emlNl=vbCrLf
	theemail=getpost("theemail")
	fromemail=getpost("fromemail")
	unsubscribe=(getpost("unsubscribe")="ON")
	unsublink=""
	print "</div>" ' to match the div that encloses this include file.
	index=0
	sSQL="SELECT email,mlName FROM mailinglist WHERE 1=1 "
	if getpost("sendto")="0" then
		sSQL="SELECT adminEmail AS email,'Admin' AS mlName FROM admin"
		batchesof=0 : takebreather=FALSE
	elseif getpost("sendto")="1" then
		sSQL=sSQL & "AND selected<>0 "
	elseif getpost("sendto")="2" then
		' Nothing - entire DB
	elseif getpost("sendto")="3" then ' Affiliates
		sSQL="SELECT affilEmail AS email,affilName AS mlName FROM affiliates WHERE 1=1 "
		unsubscribe=FALSE
	end if
	if (batchesof<>0 OR takebreather) AND getpost("sendto")<>"3" then
		sSQL=sSQL & "AND emailsent=0 "
	end if
	if getpost("sendto")<>"0" AND getpost("sendto")<>"3" then
		if NOT noconfirmationemail=TRUE then sSQL=sSQL & "AND isconfirmed<>0 "
		sSQL=sSQL & "ORDER BY email"
	end if
	hasunsubscribe=(instr(theemail, "%unsubscribe%")>0)
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		if unsubscribe then
			if NOT hasunsubscribe then unsublink=emlNl & emlNl & yyToUnsu & emlNl
			thelink=storeurl & "cart" & extension & "?unsubscribe=" & rs("email")
			if hasunsubscribe then
				unsublink = thelink
			else
				if htmlemails=TRUE then thelink="<a class=""unsubscribe"" href=""" & thelink & """>" & thelink & "</a>"
				unsublink=unsublink & thelink
			end if
		end if
		temailtxt=replaceemailtxt(theemail, "%name%", trim(rs("mlName")&""), replaceone)
		if hasunsubscribe then temailtxt=replace(temailtxt,"%unsubscribe%",unsublink)
		if NOT debugmode then call DoSendEmailEO(rs("email"),fromemail,"",getpost("emailsubject"),temailtxt & IIfVs(NOT hasunsubscribe,unsublink),emailObject,themailhost,theuser,thepass)
		if sendemailerrnum<>0 then
			print "<script>document.getElementById('emerrordiv').innerHTML+='Could not send: " & jsspecials(rs("email")) & " : " & jsspecials(sendemailerrdesc) & "<br />';</script>" & vbCrLf
		else
			ect_query("UPDATE mailinglist SET emailsent=1 WHERE email='"&escape_string(rs("email"))&"'")
		end if
		if batchesof<>0 then
			if totalsentsofar+index>=batchesof then exit do
		end if
		if index MOD 50=0 OR index=1 OR index=10 then
			print "<script>document.getElementById('sendspan').innerHTML=" & (totalsentsofar+index) & ";</script>" & vbCrLf
			response.flush
		end if
		if index=50 AND takebreather then
			print "<script>breatherfunction();</script>" & vbCrLf
			takingbreather=TRUE
			exit do
		end if
		index=index+1
		rs.movenext
	loop
	hassentall=rs.EOF
	if takebreather OR batchesof<>0 then %>
<form method="post" id="breatherform" action="adminmailinglist.asp">
<%		for each objItem in request.form
			if objItem<>"totalsentsofar" then print whv(objItem,getpost(objItem))
		next
		if is_numeric(getpost("totalsentsofar")) then totalsentsofar=int(getpost("totalsentsofar")) else totalsentsofar=0
		call writehiddenidvar("totalsentsofar",totalsentsofar+index) %>
</form>
<%	end if
	if batchesof<>0 AND NOT hassentall AND NOT takingbreather then
		print "<script>nextbatchfunction();</script>" & vbCrLf
	end if
	print "<script>document.getElementById('sendspan').innerHTML='" & (totalsentsofar+index) & " - All Done!';</script>" & vbCrLf
	print "</body></html>"
	response.flush
	response.end
elseif getpost("posted")="1" AND getpost("act")="mailchimpdeletelist" AND mailchimpapikey<>"" then
	mcsuccess=TRUE
	mcpwarray=split(mailchimpapikey,"-")
	xmlfnheaders=array(array("Content-Type","application/json"),array("X-HTTP-Method-Override","DELETE"),array("Authorization", "Basic " & vrbase64_encrypt("anystr:"&mailchimpapikey)))
	if callxmlfunction("https://"&mcpwarray(1)&".api.mailchimp.com/3.0/lists/"&getpost("listid"),"DELETE",res,"","Msxml2.ServerXMLHTTP",errormsg,FALSE) then
		mcsuccess=checkmailchimperror(res)
	end if
	if mcsuccess then print  "<div style=""padding:30px;text-align:center"">" & yyOpSuc & "</div><meta http-equiv=""refresh"" content=""1; url=adminmailinglist.asp?act=mailchimp"">"
elseif getpost("posted")="1" AND getpost("act")="syncmailchimp" AND mailchimpapikey<>"" then
	server.scripttimeout=1800
	mcsuccess=TRUE
	numemails=0
	rowcounter=0
	total_created=0
	total_updated=0
	error_count=0
	set aljsonobj=new JSONobject
	sSQL="SELECT COUNT(*) as tcount FROM mailinglist WHERE isconfirmed=1"&IIfVs(getpost("subact")="selected"," AND selected<>0")
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then numemails=rs("tcount")
	rs.close
	mcpwarray=split(mailchimpapikey,"-")
	if numemails>0 then
		json_data="{""members"": [" : addcomma=""
		sSQL="SELECT email,mlName,mlIPAddress,mlConfirmDate FROM mailinglist WHERE isconfirmed=1"&IIfVs(getpost("subact")="selected"," AND selected<>0")
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF AND mcsuccess
			call splitname(rs("mlName"),firstname,lastname)
			json_data=json_data&addcomma&"{""email_address"":" & json_encode(rs("email"))
			if rs("mlIPAddress")<>"" AND rs("mlIPAddress")<>"::1" then json_data=json_data&", ""ip_signup"":" & json_encode(rs("mlIPAddress"))
			if rs("mlConfirmDate")<>"" then json_data=json_data&", ""timestamp_signup"":" & json_encode(iso8601date(rs("mlConfirmDate"))&" 12:00:00")
			json_data=json_data&",""status"":""subscribed"",""merge_fields"":{""FNAME"":" & json_encode(firstname) & ",""LNAME"":" & json_encode(lastname) & "}}"
			addcomma=","
			rowcounter=rowcounter+1
			if rowcounter MOD 500=0 then
				json_data=json_data&"], ""update_existing"": true}"
				xmlfnheaders=array(array("Content-Type","application/json"),array("Authorization", "Basic " & vrbase64_encrypt("anystr:"&mailchimpapikey)))
				if callxmlfunction("https://"&mcpwarray(1)&".api.mailchimp.com/3.0/lists/"&getpost("listid"),json_data,res,"","Msxml2.ServerXMLHTTP",errormsg,FALSE) then
					mcsuccess=checkmailchimperror(res)
					aljsonobj.parse(res)
					if is_numeric(aljsonobj.value("total_created")&"") then total_created=total_created+int(aljsonobj.value("total_created"))
					if is_numeric(aljsonobj.value("total_updated")&"") then total_updated=total_updated+int(aljsonobj.value("total_updated"))
					if is_Numeric(aljsonobj.value("error_count")&"") then error_count=error_count+int(aljsonobj.value("error_count"))
				end if
				json_data="{""members"": [" : addcomma=""
			end if
			rs.movenext
		loop
		rs.close
		if mcsuccess then
			if rowcounter MOD 500<>0 then
				json_data=json_data&"], ""update_existing"": true}"
				xmlfnheaders=array(array("Content-Type","application/json"),array("Authorization", "Basic " & vrbase64_encrypt("anystr:"&mailchimpapikey)))
				if callxmlfunction("https://"&mcpwarray(1)&".api.mailchimp.com/3.0/lists/"&getpost("listid"),json_data,res,"","Msxml2.ServerXMLHTTP",errormsg,FALSE) then
					mcsuccess=checkmailchimperror(res)
					aljsonobj.parse(res)
					if is_numeric(aljsonobj.value("total_created")&"") then total_created=total_created+int(aljsonobj.value("total_created"))
					if is_numeric(aljsonobj.value("total_updated")&"") then total_updated=total_updated+int(aljsonobj.value("total_updated"))
					if is_Numeric(aljsonobj.value("error_count")&"") then error_count=error_count+int(aljsonobj.value("error_count"))
				end if
			end if
			print "<div style=""display:table;margin:auto"">"
				print "<div class=""ecttablerow"">"
					print "<div style=""text-align:right"">Total Created : </div><div> " & total_created & "</div>"
				print "</div>"
				print "<div class=""ecttablerow"">"
					print "<div style=""text-align:right"">Total Updated : </div><div> " & total_updated & "</div>"
				print "</div>"
				print "<div class=""ecttablerow"">"
					print "<div style=""text-align:right"">Error Count : </div><div> " & error_count & "</div>"
				print "</div>"
			print "</div>"
		end if
	else
		print "<div style=""padding:30px;text-align:center"">No Emails To Sync.</div>"
	end if
	if mcsuccess then
		print "<meta http-equiv=""refresh"" content=""5; url=adminmailinglist.asp?act=mailchimp"">"
		call adminsuccessforward("adminmailinglist.asp","act=mailchimp")
	end if
elseif getpost("posted")="1" AND getpost("act")="mailchimpcreatelist" AND mailchimpapikey<>"" then
	mcsuccess=TRUE
	json_data="{""name"":" & json_encode(getpost("mclistname")) & "," & _
		"""contact"":{""company"":" & json_encode(getpost("mccompany")) & "," & _
			"""address1"":" & json_encode(getpost("mcaddr1")) & "," & _
			"""address2"":" & json_encode(getpost("mcaddr2")) & "," & _
			"""city"":" & json_encode(getpost("mccity")) & "," & _
			"""state"":" & json_encode(getpost("mcstate")) & "," & _
			"""zip"":" & json_encode(getpost("mczip")) & "," & _
			"""country"":" & json_encode(getpost("mccountry")) & "," & _
			"""phone"":""""}," & _
		"""visibility"":" & json_encode(getpost("listpublicity")) & "," & _
		"""permission_reminder"":" & json_encode(getpost("howsignedup")) & "," & _
		"""notify_on_subscribe"":" & json_encode(getpost("mcnotifysubscribe")) & "," & _
		"""notify_on_unsubscribe"":" & json_encode(getpost("mcnotifyunsubscribe")) & "," & _
		"""campaign_defaults"":{""from_name"":" & json_encode(getpost("mcdefaultfromname")) & "," & _
			"""from_email"":" & json_encode(getpost("mcdefaultfromemail")) & "," & _
			"""subject"":" & json_encode(getpost("mcdefaultsubject")) & "," & _
			"""language"":""" & IIfVr(adminlang="","en",adminlang) & """}," & _
		"""email_type_option"":" & IIfVr(getpost("mcemailtype")="ON","true","false") & "}"
	
	mcpwarray=split(mailchimpapikey,"-")
	xmlfnheaders=array(array("Content-Type","application/json"),array("Authorization", "Basic " & vrbase64_encrypt("anystr:"&mailchimpapikey)))
	if callxmlfunction("https://" & mcpwarray(1) & ".api.mailchimp.com/3.0/lists",json_data,res,"","Msxml2.ServerXMLHTTP",errormsg,FALSE) then
		mcsuccess=checkmailchimperror(res)
	end if
	if mcsuccess then print  "<div style=""padding:30px;text-align:center"">" & yyOpSuc & "</div><meta http-equiv=""refresh"" content=""1; url=adminmailinglist.asp?act=mailchimp"">"
elseif (getpost("posted")="1" AND getpost("act")="mailchimp") OR getget("act")="mailchimp" then
	isresetkey=FALSE
	errormsg=""
	if getget("subact")="resetkey" then
		isresetkey=TRUE
	elseif getpost("subact")="updateapikey" then
		sSQL="UPDATE admin SET mailchimpAPIKey='" & escape_string(getpost("apikey")) & "' WHERE adminID=1"
		ect_query(sSQL)
		sSQL="SELECT mailchimpAPIKey FROM admin WHERE adminID=1"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then mailchimpapikey=rs("mailchimpAPIKey")
		rs.close
	elseif getpost("subact")="usemailchimplist" then
		sSQL="UPDATE admin SET mailchimpList='" & escape_string(getpost("listid")) & "' WHERE adminID=1"
		ect_query(sSQL)
		sSQL="SELECT mailchimpList FROM admin WHERE adminID=1"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then mailchimplist=rs("mailchimpList")
		rs.close
	end if
%>
<script>
/* <![CDATA[ */
function confirmdelete(){
if(confirm("Are you sure you want to delete the MailChimp API Key?")){
	document.forms.apikeyform.apikey.value="";
	document.forms.apikeyform.submit();
}
}
function formvalidator2(theForm){
if (theForm.apikey.value==""){
alert("<%=jscheck(yyPlsEntr&" """&"API Key")%>\".");
theForm.apikey.focus();
return(false);
}
if (theForm.apikey.value.indexOf("-")===-1){
alert("You must include the full MailChimp API Key, including the server (for example \"-us12\")\".");
theForm.apikey.focus();
return(false);
}
return true;
}
function formvalidator(theForm){
if (theForm.mclistname.value==""){
alert("<%=jscheck(yyPlsEntr&" """&"List Name")%>\".");
theForm.mclistname.focus();
return(false);
}
if (theForm.mcdefaultfromemail.value==""){
alert("<%=jscheck(yyPlsEntr&" """&"Default From Email")%>\".");
theForm.mcdefaultfromemail.focus();
return(false);
}
if (theForm.mcdefaultfromname.value==""){
alert("<%=jscheck(yyPlsEntr&" """&"Default From Name")%>\".");
theForm.mcdefaultfromname.focus();
return(false);
}
if (theForm.howsignedup.value==""){
alert("<%=jscheck(yyPlsEntr&" """&"Remind people how they signed up to your list")%>\".");
theForm.howsignedup.focus();
return(false);
}
if (theForm.mccompany.value==""){
alert("<%=jscheck(yyPlsEntr&" """&"Company / organization")%>\".");
theForm.mccompany.focus();
return(false);
}
if (theForm.addr1.value==""){
alert("<%=jscheck(yyPlsEntr&" """&"Address")%>\".");
theForm.addr1.focus();
return(false);
}
if (theForm.mccity.value==""){
alert("<%=jscheck(yyPlsEntr&" """&"City")%>\".");
theForm.mccity.focus();
return(false);
}
if (theForm.mcstate.value==""){
alert("<%=jscheck(yyPlsEntr&" """&"State / Province / Region")%>\".");
theForm.mcstate.focus();
return(false);
}
if (theForm.mczip.value==""){
alert("<%=jscheck(yyPlsEntr&" """&"Zip / Postal code")%>\".");
theForm.mczip.focus();
return(false);
}
if (theForm.mccountry.selectedIndex==0){
	alert("<%=jscheck(yyPlsSel&" """&"Country")%>\".");
	return(false);
}
return(true);
}
function synclist(listid,listname){
	if(confirm("<%=jscheck(yySureCa)%>")){
		document.getElementById('syncact').value='syncmailchimp';
		document.getElementById('synclistid').value=listid;
		document.getElementById('syncsubact').value='all';
		document.getElementById('synclistname').value=listname;
		document.getElementById('synclistform').submit();
	}
}
function synclistselected(listid,listname){
	if(confirm("<%=jscheck(yySureCa)%>")){
		document.getElementById('syncact').value='syncmailchimp';
		document.getElementById('synclistid').value=listid;
		document.getElementById('syncsubact').value='selected';
		document.getElementById('synclistname').value=listname;
		document.getElementById('synclistform').submit();
	}
}
function deletelist(listid,listname){
	if(confirm("<%=jscheck(yyConDel)%>")){
		document.getElementById('syncact').value='mailchimpdeletelist';
		document.getElementById('synclistid').value=listid;
		document.getElementById('synclistname').value=listname;
		document.getElementById('synclistform').submit();
	}
}
function uselist(listid){
	if(confirm("<%=jscheck(yySureCa)%>")){
		document.getElementById('syncact').value='mailchimp';
		document.getElementById('syncsubact').value='usemailchimplist';
		document.getElementById('synclistid').value=listid;
		document.getElementById('synclistform').submit();
	}
}
function actionmenu(selmenid,listid){
	var selmenobj=document.getElementById('actmen'+selmenid);
	var taction=selmenobj[selmenobj.selectedIndex].value;
	if(taction=='')
		alert("Please select an action");
	else if(taction=='1')
		synclist(listid,'');
	else if(taction=='2')
		synclistselected(listid,'');
	else if(taction=='3')
		deletelist(listid,'');
	else if(taction=='4')
		uselist(listid);
}
/* ]]> */
</script>
<%
	if mailchimpapikey<>"" AND instr(mailchimpapikey,"-")>0 AND NOT isresetkey then
		numallemails=0
		numselectedemails=0
		sSQL="SELECT COUNT(*) AS tcount FROM mailinglist WHERE isconfirmed=1"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then numallemails=rs("tcount")
		rs.close
		sSQL="SELECT COUNT(*) AS tcount FROM mailinglist WHERE isconfirmed=1 AND selected<>0"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then numselectedemails=rs("tcount")
		rs.close
		print "<form id=""synclistform"" method=""post"" action=""adminmailinglist.asp"">"
		print "<input type=""hidden"" name=""act"" id=""syncact"" value=""syncmailchimp"" />"
		print "<input type=""hidden"" name=""subact"" id=""syncsubact"" value="""" />"
		print "<input type=""hidden"" name=""posted"" value=""1"" />"
		print "<input type=""hidden"" name=""listid"" id=""synclistid"" value="""" />"
		print "<input type=""hidden"" name=""listname"" id=""synclistname"" value="""" />"
		print "</form>"
		mcpwarray=split(mailchimpapikey,"-",2)
		xmlfnheaders=array(array("Content-Type","application/json"),array("Authorization", "Basic " & vrbase64_encrypt("anystr:"&mailchimpapikey)))
		if mcpwarray(1)="" then
			isresetkey=TRUE
			errormsg="The MailChimp Server address is invalid"
		elseif callxmlfunction("https://" & mcpwarray(1) & ".api.mailchimp.com/3.0/","",res,"","Msxml2.ServerXMLHTTP",errormsg,FALSE) then
			' print "<pre>" & replace(res,",",",<br>") & "</pre><br><br>"
			set jsonobj=new JSONobject
			set outputobj=jsonobj.parse(res)
			mccompany=jsonObj.value("contact").value("company")
			mcaddr1=jsonObj.value("contact").value("addr1")
			mcaddr2=jsonObj.value("contact").value("addr2")
			mccity=jsonObj.value("contact").value("city")
			mcstate=jsonObj.value("contact").value("state")
			mczip=jsonObj.value("contact").value("zip")
			mccountry=jsonObj.value("contact").value("country")
			mcemail=jsonObj.value("email")
			proenabled=jsonObj.value("pro_enabled")
			print "<div id=""accountdetails"" class=""ecttable"">"
			print "<div class=""ecttablerow""><div>API Key</div><div>" & left(mcpwarray(0),4) & "************************" & right(mcpwarray(0),4) & "-" & mcpwarray(1) & " <a href=""adminmailinglist.asp?act=mailchimp&amp;subact=resetkey"">(Reset / Delete API Key)</a></div></div>"
			print "<div class=""ecttablerow""><div>Account ID</div><div>" & jsonObj.value("account_id") & "</div></div>"
			print "<div class=""ecttablerow""><div>Login ID</div><div>" & jsonObj.value("login_id") & "</div></div>"
			print "<div class=""ecttablerow""><div>Account Name</div><div>" & jsonObj.value("account_name") & "</div></div>"
			print "<div class=""ecttablerow""><div>Email</div><div>" & mcemail & "</div></div>"
			print "<div class=""ecttablerow""><div>Account Name</div><div>" & trim(jsonObj.value("first_name") & " " & jsonObj.value("last_name")) & "</div></div>"
			print "<div class=""ecttablerow""><div>-</div><div>-</div></div>"
			print "<div class=""ecttablerow""><div>Contact Information</div><div>" & mccompany & "<br />" & mcaddr1 & "<br />" & IIfVs(mcaddr2<>"",mcaddr2&"<br />") & mccity & IIfVs(mcstate<>"",", " & mcstate) & "<br />" & mczip & "<br />" & mccountry & "<br />" & "</div></div>"

			xmlfnheaders=array(array("Content-Type","application/json"),array("Authorization", "Basic " & vrbase64_encrypt("anystr:"&mailchimpapikey)))
			if callxmlfunction("https://"&mcpwarray(1)&".api.mailchimp.com/3.0/lists","",res,"","Msxml2.ServerXMLHTTP",errormsg,FALSE) then
				set mcrarray=jsonobj.parse(res)
				numlists=int(jsonObj.value("total_items"))
				rowcounter=1
				print "<div class=""ecttablerow""><div>Lists Defined</div><div>" & numlists & "</div></div>"
				for index=0 to numlists-1
					set jsonobj=new JSONobject
					set outputobj=jsonobj.parse(res)
					listid=jsonObj.value("lists")(index).value("id")
					listname=jsonObj.value("lists")(index).value("name")
					listmembercount=jsonObj.value("lists")(index).value("stats").value("member_count")
					print "<div class=""ecttablerow""><div style=""text-align:right" & IIfVs(mailchimplist=listid,";border:1px solid red") & """>" & listname & " - " & listmembercount & " Members : </div><div>"
						print "<select id=""actmen"&rowcounter&""" size=""1"" style=""width:300px;vertical-align:middle"">"
						print "<option value="""">" & yySelect & "</option>"
						print "<option value=""1"">Sync ALL emails (" & numallemails & ") with list: " & htmlspecials(listname) & "</option>"
						if numselectedemails<>numallemails AND numselectedemails<>0 then print "<option value=""2"">Sync selected emails (" & numselectedemails & ") with list: " & htmlspecials(listname) & "</option>"
						if mailchimplist<>listid then
							print "<option value=""3"">Delete list: " & htmlspecials(listname) & "</option>"
							print "<option value=""4"">Add new signups to this list: " & htmlspecials(listname) & "</option>"
						end if
						print "</select>"
						print " <input type=""button"" value=""Go"" onclick=""actionmenu("&rowcounter&",'" & listid & "')"" />"
					print "</div></div>"
					rowcounter=rowcounter+1
				next
				print "<div class=""ecttablerow""><div style=""text-align:right" & IIfVs(mailchimplist="",";border:1px solid red") & """>Ecommerce Plus Database - " & numallemails & " Members : </div><div>"
				print "<select id=""actmen"&rowcounter&""" size=""1"" style=""width:300px;vertical-align:middle"">"
				print "<option value="""">" & yySelect & "</option>"
				print "<option value=""4"">Add new signups to this list: Ecommerce Plus Database</option>"
				print "</select> <input type=""button"" value=""Go"" onclick=""actionmenu("&rowcounter&",'')"" />"
				print "</div></div>"
			end if
			print "<div class=""ecttablerow""><div>&nbsp;</div><div><input type=""button"" id=""createlistbutton"" value=""Create Mailchimp List"" onclick=""document.getElementById('accountdetails').style.display='none';document.getElementById('newmclist').style.display=''"" /></div></div>"
			print "</div>"

			' print "<br><br>" & $res . "<br><br>";
			
			print "<form method=""post"" action=""adminmailinglist.asp"" onsubmit=""return formvalidator(this)""><input type=""hidden"" name=""posted"" value=""1"" /><input type=""hidden"" name=""act"" value=""mailchimpcreatelist"" />"

			print "<div id=""newmclist"" style=""display:none;padding:10px 0px;width:50%"">"
			print "<h2>Create List</h2>"
			print "<h4>List details</h4>"
			print "<p style=""padding-top:10px"">" & redasterix & " List Name</p>"
			print "<input type=""text"" name=""mclistname"" style=""width:100%"" />"
			print "<p style=""font-size:0.8em;color:#555"">Your subscribers will see this, so make it something appropriate.<br />Good example: ""Acme Company Newsletter""<br />Bad example: ""Cust_11_01_2007""</p>"

			print "<p style=""padding-top:10px"">" & redasterix & " Default From Email</p>"
			print "<input type=""text"" name=""mcdefaultfromemail"" style=""width:100%"" />"
			print "<p style=""font-size:0.8em;color:#555"">This is the address people will reply to.</p>"

			print "<p style=""padding-top:10px"">" & redasterix & " Default From Name</p>"
			print "<input type=""text"" name=""mcdefaultfromname"" style=""width:100%"" />"
			print "<p style=""font-size:0.8em;color:#555"">This is who your emails will come from. Use something they\'ll instantly recognize, like your company name.</p>"

			print "<p style=""padding-top:10px"">" & redasterix & " Default Subject</p>"
			print "<input type=""text"" name=""mcdefaultsubject"" style=""width:100%"" />"
			print "<p style=""font-size:0.8em;color:#555"">Keep it relevant and non-spammy.</p>"

			print "<p style=""padding-top:10px"">List publicity: "
			print "<select size=""1"" name=""listpublicity""><option value=""pub"">Public</option><option value=""prv"">Private</option></select></p>"
			print "<p style=""font-size:0.8em;color:#555"">By default, campaigns are marked as public, but you can grant or revoke access to your data at any time.</p>"

			print "<p style=""padding-top:10px"">" & redasterix & " Remind people how they signed up to your list</p>"
			print "<textarea name=""howsignedup"" value="""" style=""height:5em;width:100%""></textarea>"
			print "<p style=""font-size:0.8em;color:#555"">Example: ""You are receiving this email because you opted in at our website &hellip; "" or ""We send special offers to customers who opted in at &hellip; ""</p>"
			
			print "<div id=""contactinfo"">"
				print "<p style=""padding-top:10px"">Contact information for the list. (<a target=""_blank"" href=""http://kb.mailchimp.com/accounts/compliance-tips/terms-of-use-and-anti-spam-requirements?&_ga=2.156609917.778392790.1498039288-1254932490.1498039288"">Why is this necessary</a>).</p>"
				print "<div style=""background-color:#ccebf3;padding:10px;margin:10px 0px"">" & mccompany & "<br />" & mcaddr1 & "<br />" & IIfVs(mcaddr2<>"",mcaddr2&"<br />") & mccity & IIfVs(mcstate<>"",", " & mcstate) & "<br />" & mczip & "<br />" & mccountry & "<br />" & "</div>"
				print "<input type=""button"" value=""Edit"" onclick=""document.getElementById('contactinfo').style.display='none';document.getElementById('addressdiv').style.display=''"" />"
			print "</div>"

			print "<div id=""addressdiv"" style=""display:none"">"
				print "<p style=""padding-top:10px"">" & redasterix & " Company / organization</p>"
				print "<input type=""text"" name=""mccompany"" style=""width:100%"" value=""" & htmlspecials(mccompany) & """ />"
				
				print "<p style=""padding-top:10px"">" & redasterix & " Address</p>"
				print "<input type=""text"" name=""mcaddr1"" style=""width:100%;margin-bottom:5px"" value=""" & htmlspecials(mcaddr1) & """ />"
				print "<input type=""text"" name=""mcaddr2"" style=""width:100%"" value=""" & htmlspecials(mcaddr2) & """ />"

				print "<p style=""padding-top:10px"">" & redasterix & " City</p>"
				print "<input type=""text"" name=""mccity"" style=""width:100%"" value=""" & htmlspecials(mccity) & """ />"
				
				print "<p style=""padding-top:10px"">" & redasterix & " State / Province / Region</p>"
				print "<input type=""text"" name=""mcstate"" style=""width:100%"" value=""" & htmlspecials(mcstate) & """ />"
				
				print "<p style=""padding-top:10px"">" & redasterix & " Zip / Postal code</p>"
				print "<input type=""text"" name=""mczip"" style=""width:100%"" value=""" & htmlspecials(mczip) & """ />"
				
				print "<p style=""padding-top:10px"">" & redasterix & " Country</p>"				
				print "<select name=""mccountry"" size=""1"" style=""width:100%"">"
				sSQL="SELECT countryCode,countryName FROM countries WHERE countryEnabled=1 ORDER BY countryOrder DESC,countryName"
				rs.open sSQL,cnn,0,1
				do while NOT rs.EOF
					print "<option value=""" & rs("countryCode") & """"
					if mccountry=rs("countryCode") then print " selected=""selected"""
					print ">" & rs("countryName") & "</option>"
					rs.movenext
				loop
				rs.close
				print "</select>"
			print "</div>"
			
			print "<h4 style=""padding-top:20px"">New subscriber notifications</h4>"
			
			print "<p style=""padding-top:10px"">Email subscribe notifications to:</p>"
			print "<input type=""text"" name=""mcnotifysubscribe"" style=""width:100%"" value="""" />"
			print "<p style=""font-size:0.8em;color:#555"">Get quick email alerts when subscribers join or leave this list (not recommended for large lists). Additional email addresses must be separated by a comma.</p>"
			
			print "<p style=""padding-top:10px"">Email unsubscribe notifications to:</p>"
			print "<input type=""text"" name=""mcnotifyunsubscribe"" style=""width:100%"" value="""" />"
			print "<p style=""font-size:0.8em;color:#555"">Get quick email alerts when subscribers join or leave this list (not recommended for large lists). Additional email addresses must be separated by a comma.</p>"
			
			print "<p style=""padding-top:10px"">Email daily digest to:</p>"
			print "<input type=""text"" name=""mcdailydigestemail"" style=""width:100%"" value="""" />"
			print "<p style=""font-size:0.8em;color:#555"">Get an end-of-the-day summary of subscribe and unsubscribe activity.</p>"
			
			print "<div>&nbsp;</div>"

			print "<div style=""float:left;height:40px""><input style=""margin:3px 10px"" class=""ectcheckbox"" type=""checkbox"" name=""mcemailtype"" value=""ON"" /></div><div>Let users pick plain-text or HTML emails.</div>"
			print "<div style=""font-size:0.8em;color:#555"">When people sign up for your list, you can let them specify which email format they prefer to receive. If they choose ""Plain-text"", then they won't receive your fancy HTML version.</div>"
			
			print "<div>&nbsp;</div>"

			print "<input type=""submit"" value=""Create Mailchimp List"" /> <input type=""button"" value=""Cancel"" onclick=""document.getElementById('accountdetails').style.display='';document.getElementById('newmclist').style.display='none'"" "

			print "</div>"
			print "</form>"
		else
			isresetkey=TRUE
		end if
	else
		isresetkey=TRUE
	end if
	if isresetkey then %>
		<form method="post" action="adminmailinglist.asp" name="apikeyform" onsubmit="return formvalidator2(this)">
		<input type="hidden" name="posted" value="1" />
		<input type="hidden" name="act" value="mailchimp" />
		<input type="hidden" name="subact" value="updateapikey" />
		<h2>Please enter your MailChimp API Key</h2>
		<div style="padding-bottom:10px">This will be of the form abab01cd45984fbaab856a8ef31680a0-us12</div>
<%		if errormsg<>"" then %>
		<div class="ectred" style="padding-bottom:10px"><%=errormsg%></div>
<%		end if
		if (mailchimpapikey<>"" AND instr(mailchimpapikey,"-")=0) OR errormsg<>"" then %>
		<div class="ectred" style="padding-bottom:10px">You must include the full MailChimp API Key, including the server (for example &quot;-us12&quot;)</div>
<%		end if %>
		<div style="padding-bottom:10px">API Key: <input type="text" name="apikey" size="50" /></div>
		<div style="padding-bottom:10px"><input type="submit" value="<%=yySubmit%>" />
<%		if mailchimpapikey<>"" then %>
			&nbsp;&nbsp;<input type="button" value="Delete API Key" onclick="confirmdelete()" />
<%		end if %>
		</div>
		<div style="padding-bottom:10px;font-size:0.8em;color:#555">(you only need to do this once.)</div>
		</form>
<%
	end if
elseif getpost("posted")="1" AND getpost("act")="sendem" then %>
<script>
/* <![CDATA[ */
function formvalidator(theForm){
<%		if htmleditor="ckeditor" OR htmleditor="froala" then %>
	if(wasusingfck){
		var inst=theemailfck;
		var sValue=inst.getData();
		if(sValue=='<br />') sValue='';
		document.getElementById("theemail").value=sValue;
	}
<%		end if %>
if (theForm.fromemail.value == ""){
alert("<%=jscheck(yyPlsEntr&" """&yyFrmEm)%>\".");
theForm.fromemail.focus();
return(false);
}
if (theForm.emailsubject.value == ""){
alert("<%=jscheck(yyPlsEntr&" """&yySubjc)%>\".");
theForm.emailsubject.focus();
return(false);
}
if (theForm.theemail.value == ""){
alert("<%=jscheck(yyPlsEntr&" """&yyMessag)%>\".");
if(!wasusingfck) theForm.theemail.focus();
return(false);
}
if (theForm.sendto.selectedIndex != 0){
	if(!confirm("<%=jscheck(yyCanSpm)%>")){
		return(false);
	}
}
<%		if htmleditor="froala" OR htmleditor="ckeditor" then %>
	if(wasusingfck){
		document.getElementById("fckrow").style.display='none';
		document.getElementById("textarearow").style.display='';
	}
<%		end if %>
return(true);
}
var wasusingfck=false;
function changeemailformat(obj){
<%		if htmleditor="ckeditor" OR htmleditor="froala" then
			if htmleditor="ckeditor" then %>
	var inst=CKEDITOR.instances.theemailfck;
<%			end if %>
	if(obj.selectedIndex==2){
		if(!wasusingfck){
<%			if htmleditor="ckeditor" then %>
			inst.setData(document.getElementById("theemail").value);
<%			else %>
			dfe_theemailfck();
			$('#theemailfck').froalaEditor('html.set',document.getElementById("theemail").value);
<%			end if %>
			document.getElementById("fckrow").style.display='';
			document.getElementById("textarearow").style.display='none';
		}
		wasusingfck=true;
	}else{
		if(wasusingfck){
<%			if htmleditor="ckeditor" then %>
			var sValue=inst.getData();
			if(sValue=='<br />') sValue='';
<%			else %>
			var sValue=$('#theemailfck').froalaEditor('html.get');
<%			end if %>
			document.getElementById("theemail").value=sValue;
			document.getElementById("fckrow").style.display='none';
			document.getElementById("textarearow").style.display='';
		}
		wasusingfck=false;
	}
<%		end if %>
}
/* ]]> */
</script>
<%		if htmleditor="ckeditor" then %>
<script src="ckeditor/ckeditor.js"></script>
<%		end if %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
		  <td width="100%" align="center">
		  <form name="mainform" method="post" action="adminmailinglist.asp" onsubmit="return formvalidator(this)">
<%			batchsent=0 : numselected=0
			sSQL="SELECT COUNT(*) AS batchsent FROM mailinglist WHERE emailsent<>0"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if NOT isnull(rs("batchsent")) then batchsent=rs("batchsent")
			end if
			rs.close
			sSQL="SELECT COUNT(*) AS numselected FROM mailinglist WHERE selected<>0"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if NOT isnull(rs("numselected")) then numselected=rs("numselected")
			end if
			rs.close
			call writehiddenvar("posted", "1")
			call writehiddenvar("act", "dosendem")
			call hiddenparams() %>
            <table width="100%" border="0" cellspacing="2" cellpadding="1">
			  <tr>
                <td colspan="2" align="center" height="34"><strong><%=yyMaLiMa&" - "&yySeEma%></strong></td>
			  </tr>
			  <tr>
				<td align="right" height="34"><%=yySenTo%>:</td>
				<td align="left"><select name="sendto" size="1">
							<option value="0"><%=yyAdmEm%></option>
<%			if numselected>0 then print "<option value=""1"">"&yySelEm&" ("&numselected&")</option>" %>
							<option value="2"><%=yyEntML%></option>
							<option value="3"><%=yyEntAL%></option>
						</select>
				</td>
			  </tr>
			  <tr>
				<td align="right" height="34"><%=yyEmlFm%>:</td>
				<td align="left"><select name="emformat" size="1" onchange="changeemailformat(this)">
							<option value="0"><%=yyText%></option>
							<option value="1">HTML</option>
<%			if htmleditor="ckeditor" OR htmleditor="froala" then %>
							<option value="2">HTML Using HTML Editor</option>
<%			end if %>
					</select>
				</td>
			  </tr>
			  <tr>
				<td align="right" height="34">Send in batches of:</td>
				<td align="left"><select name="batchesof" size="1">
							<option value="0">Unlimited</option>
							<option value="2"<% if request.cookies("EMAILBATCHNUM")="2" then print " selected=""selected"""%>>2</option>
							<option value="50"<% if request.cookies("EMAILBATCHNUM")="50" then print " selected=""selected"""%>>50</option>
							<option value="100"<% if request.cookies("EMAILBATCHNUM")="100" then print " selected=""selected"""%>>100</option>
							<option value="150"<% if request.cookies("EMAILBATCHNUM")="150" then print " selected=""selected"""%>>150</option>
							<option value="200"<% if request.cookies("EMAILBATCHNUM")="200" then print " selected=""selected"""%>>200</option>
							<option value="300"<% if request.cookies("EMAILBATCHNUM")="300" then print " selected=""selected"""%>>300</option>
							<option value="400"<% if request.cookies("EMAILBATCHNUM")="400" then print " selected=""selected"""%>>400</option>
							<option value="500"<% if request.cookies("EMAILBATCHNUM")="500" then print " selected=""selected"""%>>500</option>
							<option value="750"<% if request.cookies("EMAILBATCHNUM")="750" then print " selected=""selected"""%>>750</option>
							<option value="1000"<% if request.cookies("EMAILBATCHNUM")="1000" then print " selected=""selected"""%>>1000</option>
							<option value="1500"<% if request.cookies("EMAILBATCHNUM")="1500" then print " selected=""selected"""%>>1500</option>
							<option value="2000"<% if request.cookies("EMAILBATCHNUM")="2000" then print " selected=""selected"""%>>2000</option>
							<option value="3000"<% if request.cookies("EMAILBATCHNUM")="3000" then print " selected=""selected"""%>>3000</option>
							<option value="4000"<% if request.cookies("EMAILBATCHNUM")="4000" then print " selected=""selected"""%>>4000</option>
							<option value="5000"<% if request.cookies("EMAILBATCHNUM")="5000" then print " selected=""selected"""%>>5000</option>
							<option value="10000"<% if request.cookies("EMAILBATCHNUM")="10000" then print " selected=""selected"""%>>10000</option>
					</select>
<%			if batchsent<>0 then print " (" & batchsent & " Sent)" %>
				</td>
			  </tr>
			  <tr>
				<td align="right" height="34">Take Breather every 50 emails:</td>
				<td align="left"><input type="checkbox" name="takebreather" value="ON" />
				</td>
			  </tr>
			  <tr>
				<td align="right" height="34"><%=yyFrmEm%>:</td>
				<td align="left"><input type="text" name="fromemail" size="40" value="<%=emailAddr%>" />
				</td>
			  </tr>
			  <tr>
				<td align="right" height="34"><%=yySubjc%>:</td>
				<td align="left"><input type="text" name="emailsubject" size="40" />
				</td>
			  </tr>
			  <tr>
				<td align="right" height="34"><%=yyUnsubL%>:</td>
				<td align="left"><input type="checkbox" name="unsubscribe" value="ON" checked="checked" />
				</td>
			  </tr>
<%	if htmleditor="froala" OR htmleditor="ckeditor" then %>
			  <tr id="fckrow" style="display:none">
				<td align="right" height="34">&nbsp;</td>
				<td align="left"><textarea name="theemailfck" id="theemailfck" cols="70" rows="35"></textarea></td>
			  </tr>
<%	end if %>
			  <tr id="textarearow">
				<td align="right" height="34">&nbsp;</td>
				<td align="left"><textarea name="theemail" id="theemail" cols="70" rows="35"></textarea></td>
			  </tr>
			  <tr>
                <td colspan="2" align="center" height="34"><br /><input type="submit" value="<%=yySubmit%>" />&nbsp;<input type="reset" value="<%=yyReset%>" /><br />&nbsp;</td>
			  </tr>
			  <tr>
                <td colspan="2" align="center" height="34"><br />
                          <a href="adminmailinglist.asp"><strong><%=yyAdmHom%></strong></a><br />
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
		print "<script>"
		print "var theemailfck=CKEDITOR.replace('theemailfck',{width: 660,height: 800,toolbarStartupExpanded : false, toolbar : 'Basic', filebrowserBrowseUrl :'ckeditor/filemanager/browser/default/browser.html?Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserImageBrowseUrl : 'ckeditor/filemanager/browser/default/browser.html?Type=Image&Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserFlashBrowseUrl :'ckeditor/filemanager/browser/default/browser.html?Type=Flash&Connector="&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/connector.asp',filebrowserUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=File',filebrowserImageUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=Image',filebrowserFlashUploadUrl:'"&pathtovsadmin&"/ckeditor/filemanager/connectors/asp/upload.asp?Type=Flash'});" & vbCrLf
		print "theemailfck.on('instanceReady',function(event){var myToolbar='Basic';event.editor.on( 'beforeMaximize', function(){if(event.editor.getCommand('maximize').state == CKEDITOR.TRISTATE_ON && myToolbar != 'Basic'){theemailfck.setToolbar('Basic');myToolbar='Basic';theemailfck.execCommand('toolbarCollapse');}else if(event.editor.getCommand('maximize').state == CKEDITOR.TRISTATE_OFF && myToolbar != 'Full'){theemailfck.setToolbar('Full');myToolbar='Full';theemailfck.execCommand('toolbarCollapse');}});event.editor.on('contentDom', function(e){event.editor.document.on('blur', function(){if(!theemailfck.isToolbarCollapsed){theemailfck.execCommand('toolbarCollapse');theemailfck.isToolbarCollapsed=true;}});event.editor.document.on('focus',function(){if(theemailfck.isToolbarCollapsed){theemailfck.execCommand('toolbarCollapse');theemailfck.isToolbarCollapsed=false;}});});theemailfck.fire('contentDom');theemailfck.isToolbarCollapsed=true;});"
		print "</script>"
	elseif htmleditor="froala" then
		call displayfroalaeditor("theemailfck",yyMessag,"",TRUE,FALSE,1,TRUE)
	end if
elseif getpost("posted")="1" AND (getpost("act")="modify" OR getpost("act")="addnew") then
%>
<script>
<!--
function formvalidator(theForm){
if (theForm.email.value == ""){
alert("<%=jscheck(yyPlsEntr&" """&yyEmail)%>\".");
theForm.email.focus();
return(false);
}
return(true);
}
//-->
</script>
<%
		if getpost("act")="modify" then
			email=getpost("id")
			sSQL="SELECT isconfirmed,mlConfirmDate,mlIPAddress,mlName FROM mailinglist WHERE email='" & escape_string(email) & "'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				dateadded=rs("mlConfirmDate")
				ipaddress=trim(rs("mlIPAddress")&"")
				mlname=trim(rs("mlName")&"")
			end if
			rs.close
		else
			email=""
			mlname=""
		end if
%>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
		  <td width="100%" align="center">
		  <form name="mainform" method="post" action="adminmailinglist.asp" onsubmit="return formvalidator(this)">
<%			call writehiddenvar("posted", "1")
			if getpost("act")="modify" then call writehiddenvar("act", "domodify") else call writehiddenvar("act", "doaddnew")
			call hiddenparams() %>
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
			  <tr>
                <td width="100%" colspan="2" align="center" height="34"><strong><%
					print yyMaLiMa
					if getpost("act")="modify" then print " - " & htmlspecials(getpost("id"))%></strong></td>
			  </tr>
			  <tr>
				<td align="right" height="34"><strong><%=yyName%>:</strong></td>
				<td align="left"><input type="text" name="mlname" size="34" value="<%=htmlspecials(mlname)%>" /></td>
			  </tr>
			  <tr>
				<td align="right" height="34"><strong><%=yyEmail%>:</strong></td>
				<td align="left"><input type="text" name="email" size="34" value="<%=htmlspecials(email)%>" /></td>
			  </tr>
<%		if getpost("act")="modify" then %>
			  <tr>
				<td align="right" height="34"><strong><%=yyDateAd%>:</strong></td>
				<td align="left"><%=dateadded%></td>
			  </tr>
			  <tr>
				<td align="right" height="34"><strong><%=yyIPAdd%>:</strong></td>
				<td align="left"><%=ipaddress%></td>
			  </tr>
<%		end if %>
			  <tr>
                <td width="100%" colspan="2" align="center" height="34"><br /><input type="submit" value="<%=yySubmit%>" />&nbsp;<input type="reset" value="<%=yyReset%>" /><br />&nbsp;</td>
			  </tr>
			  <tr>
                <td width="100%" colspan="2" align="center" height="34"><br />
                          <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />
                          &nbsp;</td>
			  </tr>
            </table>
		  </form>
		  </td>
        </tr>
      </table>
<%
elseif getpost("posted")="1" AND getpost("act")<>"confirm" AND success then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr> 
                <td width="100%" colspan="2" align="center" height="34"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <a href="adminmailinglist.asp"><strong><%=yyClkHer%></strong></a>.<br />
                        <br />&nbsp;</td>
			  </tr>
			</table></td>
        </tr>
      </table>
<%
elseif getpost("posted")="1" AND getpost("act")<>"confirm" then %>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><span style="color:#FF0000;font-weight:bold"><%=yyOpFai%></span><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
			  </tr>
			</table></td>
        </tr>
      </table>
<%
else
	sSQL="SELECT count(*) AS thecount FROM mailinglist"
	if noconfirmationemail<>TRUE then sSQL=sSQL & " WHERE isconfirmed<>0"
	rs.open sSQL,cnn,0,1
	if NOT rs.eof then numemails=rs("thecount") else numemails=0
	rs.close
	sSQL="SELECT count(*) AS thecount FROM mailinglist WHERE emailsent<>0"
	rs.open sSQL,cnn,0,1
	if NOT rs.eof then numsentemails=rs("thecount") else numsentemails=0
	rs.close
	
	ordstate=trim(request("ordstate"))
	ordcountry=trim(request("ordcountry"))
	smanufacturer=trim(request("smanufacturer"))
	thecat=trim(request("scat"))
	stext=trim(request("stext"))
	stype=trim(request("stype"))
	stsearch=trim(request("stsearch"))
	swholesale=trim(request("swholesale"))
	sortorder=request("sort")
	mlact=request("mlact")
%>
<script src="popcalendar.js"></script>
<script>
<!--
try{languagetext('<%=adminlang%>');}catch(err){}
function mrec(id) {
	document.mainform.action="adminmailinglist.asp";
	document.mainform.id.value=id;
	document.mainform.act.value="modify";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function crec(id) {
	document.mainform.action="adminmailinglist.asp";
	document.mainform.id.value=id;
	document.mainform.act.value="confirm";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function newrec(id) {
	document.mainform.action="adminmailinglist.asp";
	document.mainform.id.value=id;
	document.mainform.act.value="addnew";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function sendem(id) {
	document.mainform.action="adminmailinglist.asp";
	document.mainform.act.value="sendem";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function mailchimp(id) {
	document.mainform.action="adminmailinglist.asp";
	document.mainform.act.value="mailchimp";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function drec(id){
if(confirm("<%=jscheck(yyConDel)%>\n")) {
	document.mainform.action="adminmailinglist.asp";
	document.mainform.id.value=id;
	document.mainform.act.value="delete";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
}
function startsearch(){
	document.mainform.action="adminmailinglist.asp";
	document.mainform.act.value="search";
	document.mainform.listem.value="";
	document.mainform.posted.value="";
	document.mainform.submit();
}
function listem(thelet){
	document.mainform.action="adminmailinglist.asp";
	document.mainform.act.value="search";
	document.mainform.listem.value=thelet;
	document.mainform.posted.value="";
	document.mainform.submit();
}
function removeuncon(){
if(confirm("<%=jscheck(yyConDel)%>\n")){
	document.mainform.action="adminmailinglist.asp";
	document.mainform.act.value="purgeunconfirmed";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
}
function clearsent(){
if(confirm("<%=jscheck(yySureCa)%>")) {
	document.mainform.action="adminmailinglist.asp";
	document.mainform.act.value="clearsent";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
}
function checkact(tmen){
	tact=tmen[tmen.selectedIndex].value;
	if(tact=='CSL') clearsent();
	if(tact=='ROU') removeuncon();
	if(tact=='DUS'){
		document.mainform.action="dumporders.asp";
		document.mainform.act.value="dumpemails";
		document.mainform.submit();
	}
	if(tact=='DUE'){
		document.mainform.action="dumporders.asp?entirelist=1";
		document.mainform.act.value="dumpemails";
		document.mainform.submit();
	}
	tmen.selectedIndex=0;
}
function changesortorder(men){
	var thesort=men[men.selectedIndex].value;
	document.mainform.action="adminmailinglist.asp<% if getpost("act")="search" OR getget("pg")<>"" then print "?pg=1&" else print "?"%>sort="+thesort;
	document.mainform.act.value="search";
	document.mainform.listem.value="";
	document.mainform.posted.value="";
	document.mainform.submit();
}
function changepract(obj){
	startsearch("search");
}
function quickupdate(){
	if(document.mainform.mlact.value=="del"){
		if(!confirm("<%=jscheck(yyConDel)%>\n"))
			return;
	}
	document.mainform.action="adminmailinglist.asp";
	document.mainform.act.value="quickupdate";
	document.mainform.posted.value="1";
	document.mainform.listem.value="";
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
// -->
</script>
		  <form name="mainform" method="post" action="adminmailinglist.asp">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="xxxxx" />
			<input type="hidden" name="listem" value="<%=request("listem")%>" />
			<input type="hidden" name="id" value="xxxxx" />
			<input type="hidden" name="pg" value="<%=IIfVr(getpost("act")="search", "1", getget("pg"))%>" />
			<table class="cobtbl" width="100%" border="0" cellspacing="1" cellpadding="3">
			  <tr>
				<td class="cobhl" colspan="4" align="center"><strong><%
					print numemails & " " & "Emails - "
					print "<a href=""javascript:listem('#')"">#</a> "
					for index=0 to 25
						print "<a href=""javascript:listem('"&chr(65+index)&"')"">"&chr(65+index) & "</a> "
					next
				%></strong></td>
			  </tr>
			  <tr> 
				<td class="cobhl" align="right"><select name="stsearch" size="1">
					<option value="srchemail"><%=yySrchFr&": "&yyEmail%></option>
					<option value="srchprodid" <% if stsearch="srchprodid" then print "selected"%>><%=yySrchFr&": "&yyPrId%></option>
					<option value="srchprodname" <% if stsearch="srchprodname" then print "selected"%>><%=yySrchFr&": "&yyPrName%></option>
					</select></td>
				<td class="cobll"><input type="text" name="stext" size="20" value="<%=stext%>" />
					<select name="stype" size="1">
					<option value=""><%=yySrchAl%></option>
					<option value="any" <% if stype="any" then print "selected"%>><%=yySrchAn%></option>
					<option value="exact" <% if stype="exact" then print "selected"%>><%=yySrchEx%></option>
					</select>
				</td>
				<td class="cobhl" align="right"><select name="swholesale" size="1" style="float:left">
					<option value=""><%=yyAll%></option>
					<option value="wholesale" <% if swholesale="wholesale" then print "selected"%>><%=yyWholes%></option>
					<option value="nonwholesale" <% if swholesale="nonwholesale" then print "selected"%>><%=yyNoWhol%></option>
					</select><%=yyDatRan%>:</td>
				<td class="cobll"><div style="position:relative;display:inline"><input type="text" name="mindate" size="10" value="<%=request("mindate")%>" style="vertical-align:middle" />&nbsp;<input type="button" onclick="popUpCalendar(this, document.forms.mainform.mindate, '<%=themask%>', -205)" value="DP" />&nbsp;<%=yyTo%>:&nbsp;<input type="text" name="maxdate" size="10" value="<%=request("maxdate")%>" style="vertical-align:middle" />&nbsp;<input type="button" onclick="popUpCalendar(this, document.forms.mainform.maxdate, '<%=themask%>', -205)" value="DP" /></div></td>
			  </tr>
			  <tr>
				<td class="cobhl" width="25%" align="center"><strong><%=yySection%></strong>&nbsp;&nbsp;<input type="checkbox" name="notsection" value="ON" <% if getpost("notsection")="ON" then print "checked "%>/><strong>...<%=yyNot%></strong></td>
				<td class="cobhl" width="25%" align="center"><strong><%=yyManuf%></strong>&nbsp;&nbsp;<input type="checkbox" name="notmanufacturer" value="ON" <% if getpost("notmanufacturer")="ON" then print "checked "%>/><strong>...<%=yyNot%></strong></td>
				<td class="cobhl" width="25%" align="center"><strong><%=yyState%></strong>&nbsp;&nbsp;<input type="checkbox" name="notstate" value="ON" <% if getpost("notstate")="ON" then print "checked "%>/><strong>...<%=yyNot%></strong></td>
				<td class="cobhl" width="25%" align="center"><strong><%=yyCountry%></strong>&nbsp;&nbsp;<input type="checkbox" name="notcountry" value="ON" <% if getpost("notcountry")="ON" then print "checked "%>/><strong>...<%=yyNot%></strong></td>
			  </tr>
			  <tr>
				<td class="cobll" align="center"><select name="scat" size="5" multiple="multiple"><%
						sSQL="SELECT sectionID,sectionWorkingName,topSection,rootSection FROM sections " & IIfVr(adminonlysubcats=TRUE, "WHERE rootSection=1 ORDER BY sectionWorkingName", "ORDER BY sectionOrder")
						rs.open sSQL,cnn,0,1
						if rs.eof then
							success=FALSE
						else
							alldata=rs.getrows
							success=TRUE
						end if
						rs.close
						if thecat<>"" then selstatus=split(thecat, ",") else selstatus=""
						if IsArray(alldata) then
							if adminonlysubcats=TRUE then
								for rowcounter=0 to UBOUND(alldata,2)
									print "<option value='"&alldata(0,rowcounter)&"'"
									if IsArray(selstatus) then
										for ii=0 to UBOUND(selstatus)
											if int(selstatus(ii))=alldata(0,rowcounter) then print " selected=""selected"""
										next
									end if
									print ">"&alldata(1,rowcounter)&"</option>"&vbCrLf
								next
							else
								call writemenulevel(0,1)
							end if
						end if %>
					  </select></td>
				<td class="cobll" align="center"><select name="smanufacturer" size="5" multiple="multiple"><%
						sSQL="SELECT scID,scName FROM searchcriteria WHERE scGroup=0 ORDER BY scName"
						rs.open sSQL,cnn,0,1
						if smanufacturer<>"" then selstatus=split(smanufacturer, ",") else selstatus=""
						do while NOT rs.EOF
							print "<option value=""" & rs("scID") & """"
							if IsArray(selstatus) then
								for ii=0 to UBOUND(selstatus)
									if int(selstatus(ii))=rs("scID") then print " selected=""selected"""
								next
							end if
							print ">" & rs("scName") & "</option>"
							rs.MoveNext
						loop
						rs.close %></select></td>
				<td class="cobll" align="center"><select name="ordstate" size="5" multiple="multiple"><%
						sSQL="SELECT stateID,stateName,stateAbbrev FROM states WHERE stateEnabled=1 AND stateCountryID=" & origCountryID & " ORDER BY stateName"
						rs.open sSQL,cnn,0,1
						if ordstate<>"" then selstatus=split(ordstate, ",") else selstatus=""
						do while NOT rs.EOF
							print "<option value=""" & IIfVr(usestateabbrev=TRUE,rs("stateAbbrev"),rs("stateName")) & """"
							if IsArray(selstatus) then
								for ii=0 to UBOUND(selstatus)
									if trim(selstatus(ii))=IIfVr(usestateabbrev=TRUE,rs("stateAbbrev"),rs("stateName")) then print " selected=""selected"""
								next
							end if
							print ">" & rs("stateName") & "</option>"
							rs.MoveNext
						loop
						rs.close %></select></td>
				<td class="cobll" align="center"><select name="ordcountry" size="5" multiple="multiple"><%
						sSQL="SELECT countryID,countryName FROM countries WHERE countryEnabled=1 ORDER BY countryOrder DESC, countryName"
						rs.open sSQL,cnn,0,1
						if ordcountry<>"" then selstatus=split(ordcountry, ",") else selstatus=""
						do while NOT rs.EOF
							print "<option value=""" & rs("countryName") & """"
							if IsArray(selstatus) then
								for ii=0 to UBOUND(selstatus)
									if trim(selstatus(ii))=rs("countryName") then print " selected=""selected"""
								next
							end if
							print ">" & rs("countryName") & "</option>"
							rs.MoveNext
						loop
						rs.close %></select></td>
			  </tr>
			  <tr>
				<td class="cobhl" align="center"><select onchange="checkact(this)">
						<option value=""><%=yyAct%>...</option>
						<option value="CSL">Clear &quot;Sent&quot; List<% if numsentemails<>0 then print " ("&numsentemails&")"%></option>
<%	mlcount=0
	sSQL="SELECT COUNT(*) AS mlcount FROM mailinglist WHERE isConfirmed=0 AND mlConfirmDate<" & vsusdate(date()-mailinglistpurgedays)
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		if NOT isnull(rs("mlcount")) then mlcount=rs("mlcount")
	end if
	rs.close
	if mlcount>0 then print "<option value=""ROU"">Remove Old Unconfirmed ("&mlcount&")</option>" %>
						<option value="DUS">Dump <%=yySelEm%></option>
						<option value="DUE">Dump <%=yyEntML%></option>
						</select>
<%		if mlact="del" then %>
					<div style="margin-top:2px"><input type="button" value="<%=yyCheckA%>" onclick="checkboxes(true);" /> <input type="button" value="<%=yyUCheck%>" onclick="checkboxes(false);" /></div>
<%		end if %>
						</td>
				<td class="cobll" colspan="3" align="center">
						<select name="sort" size="1" onchange="changesortorder(this)">
						<option value="naa"<% if sortorder="naa" then print " selected=""selected"""%>>Sort - Name ASC</option>
						<option value="nad"<% if sortorder="nad" then print " selected=""selected"""%>>Sort - Name DESC</option>
						<option value=""<% if sortorder="" then print " selected=""selected"""%>>Sort - Email ASC</option>
						<option value="emd"<% if sortorder="emd" then print " selected=""selected"""%>>Sort - Email DESC</option>
						<option value="daa"<% if sortorder="daa" then print " selected=""selected"""%>>Sort - Date ASC</option>
						<option value="dad"<% if sortorder="dad" then print " selected=""selected"""%>>Sort - Date DESC</option>
						<option value="coa"<% if sortorder="coa" then print " selected=""selected"""%>>Sort - Confirmed ASC</option>
						<option value="cod"<% if sortorder="cod" then print " selected=""selected"""%>>Sort - Confirmed DESC</option>
						<option value="nsf"<% if sortorder="nsf" then print " selected=""selected"""%>>No Sort (Fastest)</option>
						</select>
						<input type="button" value="<%=yyListRe%>" onclick="startsearch()" /> &nbsp;
						<input type="button" value="Add Email" onclick="newrec()" />
						<input type="button" value="Send Emails To List" onclick="sendem()" />
						<input type="button" value="Mailchimp Integration" onclick="mailchimp()" />
				</td>
			  </tr>
			</table>
<br />
            <table width="100%" class="stackable admin-table-a sta-white">
<%	if getpost("act")="search" OR getget("pg")<>"" OR getpost("act")="confirm" then
		jscript="" : qetype=""
		sub displayprodrow(xrs)
			jscript=jscript&"pa["&resultcounter&"]=["
%><tr class="<%=bgcolor%>" id="tr<%=resultcounter%>">
<td><%		print "-"
			if mlact="del" then
				jscript=jscript&"'del'"
				qetype="delbox"
			else
				print "-"
			end if %></td>
<td><%=htmlspecials(xrs("mlName")&"")%>&nbsp;</td><td></td><td><%=htmlspecials(xrs("mlConfirmDate")&"")%></td>
<td class="minicell"><% if noconfirmationemail<>TRUE AND cint(xrs("isconfirmed"))=0 then print "<input type=""button"" value="""&yyConfrm&""" onclick=""crec('"& htmlspecials(replace(xrs("email"),"'","\'")) & "')"" />" else print "&nbsp;"%></td>
<td class="minicell"><input type="button" value="<%=yyModify%>" onclick="mrec('<%=jsspecials(xrs("email"))%>')" /></td>
<td class="minicell"><input type="button" value="<%=yyDelete%>" onclick="drec('<%=jsspecials(xrs("email"))%>')" /></td></tr>
<%			jscript=jscript&",'"&jsspecials(xrs("email"))&"'];"&vbLf
			resultcounter=resultcounter+1
		end sub
		sub displayheaderrow() %>
			<tr>
				<th><select name="mlact" id="mlact" size="1" onchange="changepract(this)" style="width:150px">
					<option value="none">Quick Entry...</option>
					<option value="" disabled="disabled">---------------------</option>
					<option value="del"<% if mlact="del" then print " selected=""selected"""%>><%=yyDelete%></option>
					</select></th>
				<th class="maincell"><%=yyName%></th>
				<th class="maincell"><%=yyEmail%></th>
				<th class="maincell">Date</th>
				<th class="minicell"><% if noconfirmationemail<>TRUE then print yyConfrm else print "&nbsp;"%></th>
				<th class="minicell"><%=yyModify%></th>
				<th class="minicell"><%=yyDelete%></th>
			</tr>
<%		end sub
		sSQL="SELECT DISTINCT email,mlName,isconfirmed,mlConfirmDate FROM mailinglist "
		if (stext<>"" AND (stsearch="srchprodid" OR stsearch="srchprodname")) OR thecat<>"" OR smanufacturer<>"" OR ordstate<>"" OR ordcountry<>"" then sSQL=sSQL & "INNER JOIN (orders INNER JOIN (cart INNER JOIN products ON cart.cartProdID=products.pId) ON orders.ordID=cart.cartOrderID) ON mailinglist.email=orders.ordEmail  "
		whereand="WHERE"
		if trim(request("listem"))<>"" then
			if request("listem")="#" then
				sSQL=sSQL & "WHERE (email < 'A') "
			else
				sSQL=sSQL & "WHERE (email LIKE '"&escape_string(request("listem"))&"%') "
			end if
			whereand="AND"
		elseif stext<>"" then
			sText=escape_string(stext)
			aText=split(sText)
			if stype="exact" then
				sSQL=sSQL & whereand & " (email LIKE '%"&sText&"%') "
				whereand="AND"
			else
				if stype="any" then sJoin="OR " else sJoin="AND "
				sSQL=sSQL & whereand & " ("
				whereand="AND"
				for rowcounter=0 to UBOUND(aText)
					if stsearch="srchemail" OR stsearch="" then sSQL=sSQL & "email "
					if stsearch="srchprodid" then sSQL=sSQL & "cartProdId "
					if stsearch="srchprodname" then sSQL=sSQL & "cartProdName "
					sSQL=sSQL & " LIKE '%"&aText(rowcounter)&"%' "
					if rowcounter<UBOUND(aText) then sSQL=sSQL & sJoin
				next
				sSQL=sSQL & ")"
			end if
		end if
		if thecat<>"" then
			sectionids=getsectionids(thecat, TRUE)
			if sectionids<>"" then
				sSQL=sSQL & whereand & " " & IIfVr(getpost("notsection")="ON","NOT ","") & "(products.pSection IN (" & sectionids & ")) "
				whereand="AND"
			end if
		end if
		if smanufacturer<>"" then
			sSQL=sSQL & whereand & " " & IIfVr(getpost("notmanufacturer")="ON","NOT ","") & "(products.pManufacturer IN (" & smanufacturer & ")) "
			whereand="AND"
		end if
		if ordstate<>"" then
			sSQL=sSQL & whereand & " " & IIfVr(getpost("notstate")="ON","NOT ","") & "(ordState IN ('" & replace(replace(ordstate,", ",","),",","','") & "')) "
			whereand="AND"
		end if
		if ordcountry<>"" then
			sSQL=sSQL & whereand & " " & IIfVr(getpost("notcountry")="ON","NOT ","") & "(ordCountry IN ('" & replace(replace(ordcountry,", ",","),",","','") & "')) "
			whereand="AND"
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
				sSQL=sSQL & whereand & " mlConfirmDate BETWEEN " & vsusdate(themindate) & " AND " & vsusdate(themaxdate+1)
				whereand=" AND "
			elseif themindate<>"" then
				sSQL=sSQL & whereand & " mlConfirmDate >= " & vsusdate(themindate)
				whereand=" AND "
			elseif themaxdate<>"" then
				sSQL=sSQL & whereand & " mlConfirmDate <= " & vsusdate(themaxdate)
				whereand=" AND "
			end if
		end if
		if whereand="WHERE" then
			ect_query("UPDATE mailinglist SET selected=1")
		else
			ect_query("UPDATE mailinglist SET selected=0")
			rs.open sSQL,cnn,0,1
			do while NOT rs.EOF
				ect_query("UPDATE mailinglist SET selected=1 WHERE email='"&escape_string(rs("email"))&"'")
				rs.movenext
			loop
			rs.close
		end if
		if swholesale="nonwholesale" then
			sSQL="SELECT DISTINCT email FROM mailinglist LEFT JOIN customerlogin ON mailinglist.email=customerlogin.clEmail WHERE selected<>0 AND "
			if sqlserver OR mysqlserver then sSQL=sSQL&"(clActions&8)=8" else sSQL=sSQL&"(clActions\8) MOD 2=1"
			rs.open sSQL,cnn,0,1
			do while NOT rs.EOF
				ect_query("UPDATE mailinglist SET selected=0 WHERE email='"&escape_string(rs("email"))&"'")
				rs.movenext
			loop
			rs.close
		elseif swholesale="wholesale" then
			sSQL="SELECT DISTINCT email FROM mailinglist LEFT JOIN customerlogin ON mailinglist.email=customerlogin.clEmail WHERE selected<>0 AND "
			if sqlserver OR mysqlserver then sSQL=sSQL&"((clActions&8)<>8 OR clActions IS NULL)" else sSQL=sSQL&"((clActions\8) MOD 2<>1 OR clActions IS NULL)"
			rs.open sSQL,cnn,0,1
			do while NOT rs.EOF
				ect_query("UPDATE mailinglist SET selected=0 WHERE email='"&escape_string(rs("email"))&"'")
				rs.movenext
			loop
			rs.close
		end if
		thesort=" ORDER BY email"
		if sortorder="emd" then thesort=" ORDER BY email DESC"
		if sortorder="naa" then thesort=" ORDER BY mlName"
		if sortorder="nad" then thesort=" ORDER BY mlName DESC"
		if sortorder="daa" then thesort=" ORDER BY mlConfirmDate"
		if sortorder="dad" then thesort=" ORDER BY mlConfirmDate DESC"
		if sortorder="coa" then thesort=" ORDER BY isconfirmed,email"
		if sortorder="cod" then thesort=" ORDER BY isconfirmed DESC,email"
		if sortorder="nsf" then thesort=""
		sSQL="SELECT DISTINCT email,mlName,isconfirmed,mlConfirmDate FROM mailinglist WHERE selected<>0" & thesort
		if adminemailsperpage="" then adminemailsperpage=200
		rs.CursorLocation=3 ' adUseClient
		rs.CacheSize=adminemailsperpage
		rs.open sSQL, cnn
		if rs.eof or rs.bof then
			success=FALSE
			iNumOfPages=0
		else
			success=TRUE
			rs.MoveFirst
			rs.PageSize=adminemailsperpage
			CurPage=1
			if is_numeric(getget("pg")) then CurPage=int(getget("pg"))
			iNumOfPages=int((rs.RecordCount + (adminemailsperpage-1)) / adminemailsperpage)
			rs.AbsolutePage=CurPage
		end if
		Count=0
		haveerrprods=FALSE
		if NOT rs.EOF then
			pblink="<a href=""adminmailinglist.asp?stext="&urlencode(stext)&"&stype="&urlencode(stype)&"&ordstate="&urlencode(ordstate)&"&ordcountry="&urlencode(ordcountry)&"&smanufacturer="&urlencode(smanufacturer)&"&scat="&urlencode(thecat)&"&stsearch="&urlencode(stsearch)&"&swholesale="&urlencode(swholesale)&"&sort="&sortorder&"&mindate="&request("mindate")&IIfVr(request("maxdate")<>"","&maxdate="&request("maxdate"),"")&"&pg="
			if iNumOfPages > 1 then print "<tr><td colspan=""6"" align=""center"">" & writepagebar(CurPage,iNumOfPages,yyPrev,yyNext,pblink,FALSE) & "</td></tr>"
			displayheaderrow()
			addcomma=""
			do while NOT rs.EOF And Count < rs.PageSize
				if bgcolor="altdark" then bgcolor="altlight" else bgcolor="altdark"
				displayprodrow(rs)
				rs.MoveNext
				Count=Count + 1
			loop
			if haveerrprods then print "<tr><td width=""100%"" colspan=""6""><br />"&redasterix&yySeePr&"</td></tr>"
			if iNumOfPages > 1 then print "<tr><td colspan=""6"" align=""center"">" & writepagebar(CurPage,iNumOfPages,yyPrev,yyNext,pblink,FALSE) & "</td></tr>"
		else
			print "<tr><td width=""100%"" colspan=""6"" align=""center""><br />"&yyItNone&"<br />&nbsp;</td></tr>"
		end if
		rs.close
	else
		selectedunsent=0
		sSQL="SELECT COUNT(*) AS selectedunsent FROM mailinglist WHERE selected<>0 AND emailsent=0"
		if noconfirmationemail<>TRUE then sSQL=sSQL & " AND isconfirmed<>0"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			if NOT isnull(rs("selectedunsent")) then selectedunsent=rs("selectedunsent")
		end if
		rs.close
		if selectedunsent<>0 then %>
			<tr><td width="100%" colspan="6" align="center"><br /><%=selectedunsent%> Unsent from previous search<br />&nbsp;</td></tr>
<%		end if
		numitems=0
		sSQL="SELECT COUNT(*) as totcount FROM mailinglist"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			numitems=rs("totcount")
		end if
		rs.close
		print "<tr><td colspan=""6""><div class=""itemsdefine"">You have " & numitems & " mailing list entries.</div></td></tr>"
	end if
	if resultcounter>0 AND mlact<>"" AND mlact<>"none" then %>
			  <tr>
				<td align="center" style="white-space:nowrap"><%="<input type=""hidden"" name=""resultcounter"" id=""resultcounter"" value="""&resultcounter&""" /><input type=""button"" value="""&yyUpdate&""" onclick=""quickupdate()"" /> <input type=""reset"" value="""&yyReset&""" />"%></td>
                <td colspan="6">&nbsp;</td>
			  </tr>
<%	end if %>
            </table>
			<div style="text-align:center;margin:20px"><a href="admin.asp"><strong><%=yyAdmHom%></strong></a></div>
		  </form>
<script>
var pa=[];
<%
	print jscript%>
	function patch_pid(pid){
		document.getElementById('pid'+pid).name='pid'+pid;
		document.getElementById('pid'+pid).value=pa[pid][1];
		return pid;
	}
	for(var pidind in pa){
		var ttr=document.getElementById('tr'+pidind);
		ttr.cells[0].className='minicell';
		ttr.cells[2].innerHTML='<input type="hidden" id="pid'+pidind+'" value="" />'+pa[pidind][1];
		ttr.cells[0].innerHTML=
<%		if qetype="delbox" then %>
	'<input type="checkbox" id="chkbx'+pidind+'" onchange="this.name=\'pra_'+patch_pid(pidind)+'\'" value="del" tabindex="'+(pidind+1)+'" />';
<%		else %>
	'&nbsp;';
<%		end if %>
	}
</script>
<%
end if
cnn.Close
set rs=nothing
set cnn=nothing
%>
