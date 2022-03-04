<!--#include file="db_conn_open.asp"-->
<!--#include file="includes.asp"-->
<!--#include file="inc/languageadmin.asp"-->
<!--#include file="inc/incfunctions.asp"-->
<% donotlogin=TRUE %>
<!--#include file="inc/incloginfunctions.asp"-->
<% if storesessionvalue="" then storesessionvalue="virtualstore" %>
<!doctype html>
<head>

<title>EcommerceTemplates.com Admin Logout</title>

<!-- Header assets -->
<% call adminassets() %>

</head>
<body>

<div class="login">
<%	SESSION("loggedon") = ""
	response.cookies("WRITECKL")=""
	response.cookies("WRITECKL").Expires = Date()-30
	response.cookies("WRITECKP")=""
	response.cookies("WRITECKP").Expires = Date()-30
	response.cookies("loginkey")=""
	response.cookies("loginkey").Expires = Date()-30
%>
<meta http-equiv="Refresh" content="3; URL=admin.asp" />
  <div class="row centerit">
	<div class="login_message">
		<h2 class="centerit"><%=yyLogOut%></h2>
		<p><%=yyLOMes%></p>
	</div>
  </div>
</div>

</body>
</html>
