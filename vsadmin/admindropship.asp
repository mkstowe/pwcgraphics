<%
Response.Buffer = True
'=========================================
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protect under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
%>
<!--#include file="db_conn_open.asp"-->
<!--#include file="includes.asp"-->
<!--#include file="inc/languageadmin.asp"-->
<!--#include file="inc/incfunctions.asp"-->
<!--#include file="inc/incloginfunctions.asp"-->
<!--#include file="inc/md5.asp"-->
<!doctype html>
<head>

<title>Admin Drop Shipper</title>

<!-- Header assets -->
<% call adminassets() %>

</head>
<body <% if isprinter then response.write "class=""printbody"""%>>
<% if NOT isprinter then %>

<!-- Header section -->
<% call adminheader() %>

<!-- Left menus -->
<% call adminnavigation() %>

<% end if %>
<!-- main content -->

<div id="main">
<%	if Mid(SESSION("loggedonpermissions"),12,1)<>"X" then
		response.write "<table width=""100%"" border=""0"" bgcolor=""""><tr><td width=""100%"" colspan=""4"" align=""center""><p>&nbsp;</p><p>&nbsp;</p><p><strong>"&yyOpFai&"</strong></p><p>&nbsp;</p><p>"&yyNoPer&" <br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;<br /><a href=""admin.asp""><strong>"&yyAdmHom&"</strong></a>.</p><p>&nbsp;</p></td></tr></table>"
	else %>
<!--#include file="inc/incdropship.asp"-->
<%	end if %>
</div>

<!-- Footer -->
<% call adminfooter() %>

</body>
</html>
