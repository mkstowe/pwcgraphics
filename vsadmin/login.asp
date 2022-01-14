<!--#include file="db_conn_open.asp"-->
<!--#include file="includes.asp"-->
<!--#include file="inc/languageadmin.asp"-->
<!--#include file="inc/incfunctions.asp"-->
<% donotlogin=TRUE %>
<!--#include file="inc/incloginfunctions.asp"-->
<!--#include file="inc/incemail.asp"-->
<!--#include file="inc/md5.asp"-->
<%
if padssfeatures=TRUE then
	response.AddHeader "pragma","no-cache"
	response.CacheControl = "no-store,no-cache"
end if
%>
<!doctype html>
<head>
<title>Control panel login</title>
<!-- Header assets -->
<% call adminassets() %>
</head>

<body>
<div class="login">

<%
Dim sSQL,rs,success,cnn,errmsg
vsadmindir="vsadmin"
thisurl=request.servervariables("URL")
pos1 = instrrev(thisurl,"/")
if pos1>1 then
	pos1=pos1-1
	pos2 = instrrev(thisurl,"/",pos1-1)
	if pos2>0 then
		vsadmindir=mid(thisurl,pos2+1,pos1-pos2)
	end if
end if
if forceloginonhttps AND request.servervariables("HTTPS")="off" AND instr(storeurlssl,"https")>0 then response.redirect storeurlssl & vsadmindir & "/login.asp" & IIfVs(getget("loginkey")<>"","?loginkey="&htmlspecials(getget("loginkey"))) : response.end
success=true
dorefresh=FALSE
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
on error resume next
cnn.open sDSN
if err.number <> 0 then
	success=false
	errmsg = "<p><strong>Your database connection needs to be set before you can proceed.</strong><br /><br /></p>"
	if NOT (sqlserver=TRUE OR mysqlserver=TRUE) then errmsg = errmsg & "<p>The current setting is:<br />"&sDSN&"</p>"
	errmsg = errmsg & "<p>The following information may be helpful</p>" &_
		"<p><strong>Path to this directory<br />"&server.mappath("../")&"</strong></p><p>&nbsp;</p>"
end if
on error goto 0
loginkeyerror=FALSE
floodcontrol=FALSE
if success then
	alreadygotadmin = getadminsettings()
	if request.form("posted")="1" OR getget("loginkey")<>loginkey then
		if application("lastkeycheck")<>"" then
			if datediff("s",application("lastkeycheck"),now())<5 then floodcontrol=TRUE
		end if
	end if
	if floodcontrol then
		success=FALSE
		floodcontrol=TRUE
		disallowlogin=TRUE
	elseif loginkey<>"" AND getget("loginkey")<>loginkey then
		success=FALSE
		loginkeyerror=TRUE
		disallowlogin=TRUE
		application("lastkeycheck")=now()
	elseif request.form("posted")="1" then
		application("lastkeycheck")=now()
		if recaptchaenabled(16) then success=checkrecaptcha(errmsg)
		thashedpw=dohashpw(request.form("pass"))
		adminuser=""
		adminpassword=""
		if success then
			sSQL = "SELECT adminEmail,adminUser,adminPassword,adminPWLastChange FROM admin WHERE adminID=1"
			rs.open sSQL,cnn,0,1
			datelastchanged=rs("adminPWLastChange")
			adminuser=rs("adminUser")
			adminpassword=rs("adminPassword")
			rs.close
		end if
		if storesessionvalue="" then storesessionvalue="virtualstore"
		if NOT success then
		elseif disallowlogin=TRUE then
			success = FALSE
			errmsg = yyLogSor
		elseif NOT (trim(request.form("user"))=adminuser AND thashedpw=adminpassword) then
			sSQL="SELECT adminloginid,adminloginname,adminloginpassword,adminloginpermissions FROM adminlogin WHERE adminloginname='" & escape_string(request.form("user")) & "'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if rs("adminloginname")=request.form("user") AND rs("adminloginpassword")=thashedpw then
					SESSION("loggedon") = storesessionvalue
					SESSION("loggedonpermissions") = rs("adminloginpermissions")
					SESSION("loginid")=rs("adminloginid")
					SESSION("loginuser")=rs("adminloginname")
					dorefresh=TRUE
				else
					success = FALSE
					errmsg = yyLogSor
				end if
			else
				success = FALSE
				errmsg = yyLogSor
			end if
			rs.close
		else
			SESSION("loggedon") = storesessionvalue
			SESSION("loggedonpermissions") = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
			SESSION("loginid")=0
			SESSION("loginuser")=adminuser
			if thashedpw="50481f28d0f9c62842ad64b8985ab91c" AND ectdemostore<>TRUE then SESSION("mustchangepw")="A"
			if date()-datelastchanged>90 AND padssfeatures=TRUE then SESSION("mustchangepw")="B"
			dorefresh=TRUE
		end if
		call logevent(request.form("user"),"LOGIN",success,"LOGIN","")
		if notifyloginattempt=TRUE AND disallowlogin<>TRUE then
			if htmlemails=true then emlNl = "<br />" else emlNl=vbCrLf
			sMessage = "This is notification of a login attempt at your store."  & emlNl
			sMessage = sMessage & storeurl & emlNL
			if success OR (trim(request.form("user"))=adminuser AND trim(request.form("pass"))=adminpassword) then
				sMessage = sMessage & "A correct login / password was used." & emlNl
			else
				sMessage = sMessage & "An incorrect login was used." & emlNl & _
					"Username: " & request.form("user") & emlNl & _
					"Password: " & request.form("pass") & emlNl
			end if
			sMessage = sMessage & "User Agent: " & Request.ServerVariables("HTTP_USER_AGENT") & emlNl & _
				"IP: " & Request.ServerVariables("REMOTE_HOST") & emlNl
			Call DoSendEmail(emailAddr,emailAddr,"Login attempt at your store",sMessage)
		end if
		if success AND request.form("cook")="ON" then
			response.cookies("WRITECKL")=trim(request.form("user"))
			response.cookies("WRITECKL").Expires = Date()+365
			if request.servervariables("HTTPS")="on" then response.cookies("WRITECKL").secure=TRUE
			response.cookies("WRITECKP")=thashedpw
			response.cookies("WRITECKP").Expires = Date()+365
			if request.servervariables("HTTPS")="on" then response.cookies("WRITECKP").secure=TRUE
			response.cookies("loginkey")=getget("loginkey")
			response.cookies("loginkey").Expires = Date()+365
			if request.servervariables("HTTPS")="on" then response.cookies("loginkey").secure=TRUE
		end if
		if dorefresh then
			response.write "<meta http-equiv=""refresh"" content=""1; url=admin.asp"">"
		end if
	end if
end if
if cnn.State=1 then cnn.Close
set rs = nothing
set cnn = nothing
	if request.form("posted")="1" AND success then %>
	<div class="row centerit">
      <div class="login_message">
            <h2 class="centerit"><%=yyLogCor%></h2>
            <p><%=yyNowFrd%></p>
            <p><%=yyNoAuto%><a href="admin.asp"><strong><%=yyClkHer%></strong></a>.</p>
      </div>
    </div>
<%	else
		if disallowlogin then
			success=FALSE
			errmsg="<div class=""login_message"">" & "Login Disabled:<br />"
			if loginkeyerror then errmsg=errmsg&"<br />(Admin Login Key Error)"
			if floodcontrol then errmsg=errmsg&"<br />(Flood Control. Please wait 5 seconds between login attempts)"
			if floodcontrol AND request.form("posted")="1" then errmsg=errmsg&"<br /><br /><input type=""button"" value=""Go Back"" onclick=""history.go(-1)"" />"
			errmsg=errmsg& "</div>"
		end if %>
	<form method="post" name="mainform" action="login.asp<%=IIfVs(getget("loginkey")<>"","?loginkey="&htmlspecials(getget("loginkey")))%>" id="loginform" onsubmit="return checkloginform()">
	<input type="hidden" name="posted" value="1">
	<div class="row centerit">
        <div class="login_form round_all">
            <div class="login_header" onclick="document.location='admin.asp'"></div>
<%		if not success then %>
			  <div class="ectred"><%=errmsg%></div>
<%		end if
		if NOT disallowlogin then %>
			<table>
              <tr>
                <td width="30%" align="right"><strong><%=yyUname%>: </strong></td>
				<td align="left"><input type="text" name="user" id="user" size="20" /></td>
			  </tr>
			  <tr>
                <td align="right"><strong><%=yyPass%>: </strong></td>
				<td align="left"><input type="password" name="pass" size="20" autocomplete="off" /></td>
			  </tr>
			  <tr>
                <td align="right"><input type="checkbox" name="cook" value="ON" /></td>
				<td align="left" class="small"><strong><%=yyWrCoo%></strong><br /><span style="font-size:10px"><%=yyNoRec%></span></td>
			  </tr>
<%			if recaptchaenabled(16) then %>
			  <tr>
				<td align="center" colspan="2"><%
				call displayrecaptchajs("adminlogincaptcha",TRUE,FALSE)
				%><div id="adminlogincaptcha" class="g-recaptcha recaptchaadminlogin"></div>
				</td>
			  </tr>
<%			end if %>
			</table>
			<p><input type="submit" value="<%=yySubmit%>"></p>
<%		end if %>
			  <p class="credit"><a href="https://www.ecommercetemplates.com/">Shopping Cart Software</a> by Ecommerce Templates</p>
        </div>
    </div>
	</form>
<script>
<!--
function checkloginform(){
	var frm=document.getElementById('loginform');
	if(frm.user.value==''){
		alert("<%=jscheck(yyPlsEntr&" """&yyUname&"""")%>");
		frm.user.focus();
		return(false);
	}
<%	if recaptchaenabled(16) then print "if(!adminlogincaptchaok){ alert(""Please show you are a real human by completing the reCAPTCHA test"");return(false); }" %>
	return true;
}
document.getElementById('user').focus();
// -->
</script>
<%	end if %>
</div>
</body>
</html>