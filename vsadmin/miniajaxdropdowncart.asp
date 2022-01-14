<%
Response.Buffer=True
Response.Expires=60
Response.Expiresabsolute=Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl="no-cache"
%>
<!--#include file="db_conn_open.asp"-->
<!--#include file="inc/languageadmin.asp"-->
<!--#include file="inc/languagefile.asp"-->
<!--#include file="includes.asp"-->
<%
on error resume next
if lcase(adminencoding)<>"utf-8" then response.codepage=65001
on error goto 0
response.charset="utf-8"
%>
<!--#include file="inc/incfunctions.asp"-->
<% response.clear %>
<!--#include file="inc/incminidropdowncart.asp"-->
