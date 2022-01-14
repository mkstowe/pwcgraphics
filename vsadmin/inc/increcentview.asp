<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if sDSN="" then response.write "Database connection not set" : response.end
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
function join2pathsrv(stourl,securl)
	if instr(securl,"://")>0 then
		join2pathsrv=securl
	elseif right(stourl,1)="/" AND left(securl,1)="/" then
		slashpos=instr(10,stourl,"/")
		if slashpos>0 then tmpstourl=left(stourl,slashpos-1) else tmpstourl=stourl
		join2pathsrv=tmpstourl&securl
	else
		join2pathsrv=stourl&securl
	end if
end function
if getpost("sessionid")<>"" then thesessionid = replace(getpost("sessionid"),"'","") else thesessionid = getsessionid()
if incfunctionsdefined=TRUE then
	alreadygotadmin = getadminsettings()
else
	sSQL = "SELECT countryLCID,countryCurrency,adminStoreURL FROM admin INNER JOIN countries ON admin.adminCountry=countries.countryID WHERE adminID=1"
	rs.open sSQL,cnn,0,1
	if orlocale<>"" then
		Session.LCID = orlocale
	elseif rs("countryLCID")<>0 then
		Session.LCID = rs("countryLCID")
	end if
	storeurl = rs("adminStoreURL")
	if (left(LCase(storeurl),7) <> "http://") AND (left(LCase(storeurl),8) <> "https://") then storeurl = "http://" & storeurl
	if Right(storeurl,1) <> "/" then storeurl = storeurl & "/"
	rs.close
end if
if getpost("mode")<>"checkout" then
	sSQL = "SELECT rvProdName,rvProdURL,sectionName FROM recentlyviewed INNER JOIN sections ON recentlyviewed.rvProdSection=sections.sectionID WHERE rvProdID<>'"&escape_string(prodid)&"' AND " & IIfVr(SESSION("clientID")<>"", "rvCustomerID="&replace(SESSION("clientID"),"'",""), "(rvCustomerID=0 AND rvSessionID='"&thesessionid&"')")&" ORDER BY rvDate DESC"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
%>
      <table class="mincart" width="130" bgcolor="#FFFFFF">
        <tr> 
          <td class="mincart" bgcolor="#F0F0F0" align="center"><img src="images/recentview.png" style="vertical-align:text-top;" width="16" height="15" alt="<%=xxRecVie%>" /> 
            &nbsp;<strong><a class="ectlink mincart" href="<%=storeurl%>cart<%=extension%>"><%=xxRecVie%></a></strong></td>
        </tr>
<%		do while NOT rs.EOF %>
        <tr><td class="mincart" bgcolor="#F0F0F0" align="center">
		<span style="font-family:Verdana">&raquo;</span> <%=rs("sectionName")%><br />
		<a class="ectlink mincart" href="<%=join2pathsrv(storeurl,rs("rvProdURL"))%>"><%=rs("rvProdName")%></a></td></tr>
<%			rs.MoveNext
		loop %>		
      </table>
<%	end if
	rs.close
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>