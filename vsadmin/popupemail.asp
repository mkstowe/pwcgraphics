<%
Response.Buffer = True
'=========================================
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
%>
<!--#include file="inc/md5.asp"-->
<!--#include file="db_conn_open.asp"-->
<!--#include file="inc/languageadmin.asp"-->
<!--#include file="inc/incfunctions.asp"-->
<%	set rs=server.createobject("ADODB.RecordSet")
	set rs2=server.createobject("ADODB.RecordSet")
	set cnn=server.createobject("ADODB.Connection")
	cnn.open sDSN
	languageid=1
	rs.open "SELECT storelang FROM admin WHERE adminid=1",cnn,0,1
	storelangarr=split(rs("storelang"),"|")
	storelang = trim(rs("storelang")&"")
	if storelang<>"" then storelang=split(storelang,"|")(0) else storelang=""
	rs.close
	if request.form("posted")="1" AND is_numeric(request.form("id")) then
		sSQL="SELECT ordID,ordLang FROM orders WHERE ordID="&request.form("id")
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then languageid=rs("ordLang")+1 else languageid=1
		rs.close
		if languageid>1 then
			if languageid=2 AND UBOUND(storelangarr)>0 then storelang=storelangarr(1)
			if languageid=3 AND UBOUND(storelangarr)>1 then storelang=storelangarr(2)
		end if
	end if
	if storelang="de" then %>
<!--#include file="inc/languagefile_de.asp"-->
<%	elseif storelang="dk" then %>
<!--#include file="inc/languagefile_dk.asp"-->
<%	elseif storelang="es" then %>
<!--#include file="inc/languagefile_es.asp"-->
<%	elseif storelang="fr" then %>
<!--#include file="inc/languagefile_fr.asp"-->
<%	elseif storelang="it" then %>
<!--#include file="inc/languagefile_it.asp"-->
<%	elseif storelang="nl" then %>
<!--#include file="inc/languagefile_nl.asp"-->
<%	elseif storelang="pt" then %>
<!--#include file="inc/languagefile_pt.asp"-->
<%	else %>
<!--#include file="inc/languagefile_en.asp"-->
<%	end if %>
<!--#include file="includes.asp"-->
<!--#include file="inc/incemail.asp"-->
<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if NOT disallowlogin then
<!--#include file="inc/incloginfunctions.asp"-->
end if
if SESSION("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.redirect "login.asp"
isprinter=false
%>
<!doctype html>
<html>
<head>
<title>Email Popup</title>
<link rel="stylesheet" type="text/css" href="adminstyle.css"/>
<meta http-equiv="Content-Type" content="text/html; charset=<%=adminencoding%>"/>
</head>
<body>
&nbsp;<br />
<div>
<form method="post" action="popupemail.asp">
<%
	ordGrandTotal=0 : ordTotal=0 : ordStateTax=0 : ordHSTTax=0 : ordCountryTax=0 : ordShipping=0 : ordHandling=0 : ordDiscount=0
	affilID="" : ordState="" : ordCountry="" : ordDiscountText=""
	isresendemail=TRUE
	if request.form("posted")="1" then
		alreadygotadmin = getadminsettings()
		call do_order_success(request.form("id"),emailAddr,request.form("store")="1",FALSE,request.form("customer")="1",request.form("affiliate")="1",IIfVr(request.form("manufacturer")="1",2,FALSE))
%>
<p>&nbsp;</p>
<p align="center"><%=yyOpSuc%></p>
<p align="center"><a href="javascript:window.close()"><strong><%=xxClsWin%></strong></a></p>
<%	elseif request.querystring("id")<>"" then %>
<input type="hidden" name="posted" value="1">
<input type="hidden" name="id" value="<%=request.querystring("id")%>">
<table width="100%" cellspacing="2" cellpadding="2">
<tr><td colspan="2" align="center"><strong><%=yySendFo%></strong></td></tr>
<tr><td align="right" width="60%"><%=yyCusto%>: </td><td><input type="checkbox" name="customer" value="1" checked></td></tr>
<tr><td align="right"><%=yyAffili%>: </td><td><input type="checkbox" name="affiliate" value="1"></td></tr>
<tr><td align="right"><%=yyManDes%>: </td><td><input type="checkbox" name="manufacturer" value="1"></td></tr>
<tr><td align="right"><%=xxOrdStr%>: </td><td><input type="checkbox" name="store" value="1"></td></tr>
<tr><td colspan="2" align="center"><input type="submit" value="<%=yySubmit%>" /></td></tr>
</table>
<%	end if
cnn.Close
%>
</form>
</div>
</body>
</html>
