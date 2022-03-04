<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if alreadygotadmin<>TRUE then
	set rs=server.createobject("ADODB.RecordSet")
	set cnn=server.createobject("ADODB.Connection")
	cnn.open sDSN
	alreadygotadmin=getadminsettings()
	cnn.Close
	set rs=nothing
	set cnn=nothing
end if
pageqs=""
for each objitem in request.querystring
	if NOT (objitem="mode" AND (getget(objitem)="login" OR getget(objitem)="logout")) then pageqs=pageqs&IIfVs(pageqs<>"","&")&objitem&"="&getget(objitem)
next
if forceloginonhttps AND request.servervariables("HTTPS")="off" AND (replace(storeurl,"http:","https:")<>storeurlssl) then pagename="" else pagename=request.servervariables("URL") & IIfVs(pageqs<>"","?"&pageqs)
if displaysoftlogindone="" then displaysoftlogindone=""
call displaysoftlogin()
%>
      <table class="mincart" width="130" bgcolor="#FFFFFF">
        <tr> 
          <td class="mincart" bgcolor="#F0F0F0" align="center"><img src="images/minipadlock.png" style="vertical-align:text-top;" alt="<%=xxMLLIS%>" /> 
		<% if SESSION("clientID")<>"" AND customeraccounturl<>"" then %>
			&nbsp;<a class="ectlink mincart" href="<%=customeraccounturl%>"><strong><%=xxYouAcc%></strong></a>
		<% else %>
            &nbsp;<strong><%=xxMLLIS%></strong>
		<% end if %></td>
        </tr>
	<% if NOT enableclientlogin then %>
		<tr>
		  <td class="mincart" bgcolor="#F0F0F0" align="center">
		  <p class="mincart">Client login not enabled</p>
		  </td>
		</tr>
	<% elseif SESSION("clientID")<>"" AND request.querystring("mode")<>"logout" then %>
		<tr>
		  <td class="mincart" bgcolor="#F0F0F0" align="center">
		  <p class="mincart"><%=xxMLLIA%><strong><br /><%=server.htmlencode(SESSION("clientUser"))%></strong></p>
		  </td>
		</tr>
		<tr> 
          <td class="mincart" bgcolor="#F0F0F0" align="center"><span style="font-family:Verdana">&raquo;</span> <%=imageorlink(imgminilogout,xxLogout,"ectlink mincart","return dologoutaccount()",TRUE)%></td>
        </tr>
	<% else %>
		<tr>
		  <td class="mincart" bgcolor="#F0F0F0" align="center">
		  <p class="mincart"><%=xxMLNLI%></p>
		  </td>
		</tr>
		<tr> 
          <td class="mincart" bgcolor="#F0F0F0" align="center"><span style="font-family:Verdana">&raquo;</span> <%=imageorlink(imgminilogin,xxLogin,"ectlink mincart","return displayloginaccount()",TRUE)%></td>
        </tr>
	<% end if %>
      </table>