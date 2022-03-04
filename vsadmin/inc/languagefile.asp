<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if SESSION("languageid")<>"" then languageid=SESSION("languageid") else if languageid="" then languageid=1
if orstorelang<>"" then
	storelang=orstorelang
else
	if sDSN="" then response.write "Database connection not set" : response.end
	set cnn=Server.CreateObject("ADODB.Connection")
	cnn.open sDSN
	set rslang=cnn.execute("SELECT storelang FROM admin WHERE adminid=1")
	storelang = trim(rslang("storelang")&"")
	if storelang<>"" then
		ectstorelangarr=split(storelang,"|")
		if UBOUND(ectstorelangarr)>=(languageid-1) then storelang=ectstorelangarr(languageid-1)
	end if
	set rslang=nothing
	cnn.close
	set cnn=nothing 
end if
   if storelang="de" then %>
<!--#include file="languagefile_de.asp"-->
<% elseif storelang="dk" then %>
<!--#include file="languagefile_dk.asp"-->
<% elseif storelang="es" then %>
<!--#include file="languagefile_es.asp"-->
<% elseif storelang="fr" then %>
<!--#include file="languagefile_fr.asp"-->
<% elseif storelang="it" then %>
<!--#include file="languagefile_it.asp"-->
<% elseif storelang="nl" then %>
<!--#include file="languagefile_nl.asp"-->
<% elseif storelang="pt" then %>
<!--#include file="languagefile_pt.asp"-->
<% else %>
<!--#include file="languagefile_en.asp"-->
<% end if %>