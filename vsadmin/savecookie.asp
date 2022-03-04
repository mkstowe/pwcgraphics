<%
Response.Buffer = True
'=========================================
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protect under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if Trim(request.querystring("id1"))<>"" AND Trim(request.querystring("id2"))<>"" then
	response.cookies("id1")=request.querystring("id1")
	response.cookies("id1").Expires = Date()+180
	if request.servervariables("HTTPS")="on" then response.cookies("id1").secure=TRUE
	response.cookies("id2")=request.querystring("id2")
	response.cookies("id2").Expires = Date()+180
	if request.servervariables("HTTPS")="on" then response.cookies("id2").secure=TRUE
elseif Trim(request.querystring("PARTNER"))<>"" AND isnumeric(request.querystring("EXPIRES")) AND trim(request.querystring("EXPIRES"))<>"" then
	response.cookies("PARTNER")=Trim(request.querystring("PARTNER"))
	response.cookies("PARTNER").Expires = Date()+Int(request.querystring("EXPIRES"))
	if request.servervariables("HTTPS")="on" then response.cookies("PARTNER").secure=TRUE
elseif Trim(request.querystring("DELCK")) = "yes" then
	response.cookies("WRITECKL")=""
	response.cookies("WRITECKL").Expires = Date()-30
	response.cookies("WRITECKP")=""
	response.cookies("WRITECKP").Expires = Date()-30
elseif Trim(request.querystring("WRITECLL")) <> "" then
	response.cookies("WRITECLL")=Trim(request.querystring("WRITECLL"))
	if Trim(request.querystring("permanent")) = "Y" then response.cookies("WRITECLL").Expires = Date()+365
	if request.servervariables("HTTPS")="on" then response.cookies("WRITECLL").secure=TRUE
	response.cookies("WRITECLP")=Trim(request.querystring("WRITECLP"))
	if Trim(request.querystring("permanent")) = "Y" then response.cookies("WRITECLP").Expires = Date()+365
	if request.servervariables("HTTPS")="on" then response.cookies("WRITECLP").secure=TRUE
elseif Trim(request.querystring("DELCLL")) <> "" then
	response.cookies("WRITECLL")=""
	response.cookies("WRITECLL").Expires = Date()-30
	response.cookies("WRITECLP")=""
	response.cookies("WRITECLP").Expires = Date()-30
end if
response.flush
%>