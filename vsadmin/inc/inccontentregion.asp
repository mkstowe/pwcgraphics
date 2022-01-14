<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if sDSN="" then response.write "Database connection not set" : response.end
catname=""
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
theid=getget("region")
if regionid<>"url" AND is_numeric(regionid) then theid=regionid
if NOT is_numeric(theid) then theid=0
sSQL = "SELECT "&getlangid("contentData",32768)&" FROM contentregions WHERE contentID="&theid
rs.open sSQL,cnn,0,1
if NOT rs.EOF then contentdata=rs(getlangid("contentData",32768)) else contentdata = "Content Region ID " & theid & " not defined."
rs.close
cnn.Close
set cnn=nothing
set rs=nothing
response.write contentdata
regionid=""
%>

