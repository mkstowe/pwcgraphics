<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if sDSN="" then response.write "Database connection not set" : response.end
set rs=Server.CreateObject("ADODB.RecordSet")
set rs2=Server.CreateObject("ADODB.RecordSet")
set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
if trim(request.form("sessionid"))<>"" then thesessionid = replace(trim(request.form("sessionid")),"'","") else thesessionid=getsessionid()
if isempty(extension) then extension=".asp"
function FormatMCCurrency(amount)
	FormatMCCurrency=""
	if currPostAmount=0 then FormatMCCurrency=currSymbolHTML
	FormatMCCurrency=FormatMCCurrency&FormatNumber(amount,currDecimals,-1,0,IIfVr(currThousandsSep<>"",-1,0))
	if currPostAmount<>0 then FormatMCCurrency=FormatMCCurrency&currSymbolHTML
end function
mcgndtot=0
mcpdtxt=""
totquant=0
shipping=0
discounts=0
if SESSION("xscountrytax")<>"" then xscountrytax = cdbl(SESSION("xscountrytax")) else xscountrytax=0
if incfunctionsdefined=TRUE then
	alreadygotadmin = getadminsettings()
	if cartpageonhttps then pageurl=storeurlssl else pageurl=storeurl
else
	sSQL = "SELECT countryLCID,countryCurrency,adminStoreURL,currDecimalSep,currThousandsSep,currPostAmount,currDecimals,currSymbolHTML FROM admin INNER JOIN countries ON admin.adminCountry=countries.countryID WHERE adminID=1"
	rs.open sSQL,cnn,0,1
	if orlocale<>"" then
		Session.LCID = orlocale
	elseif rs("countryLCID")<>0 then
		Session.LCID = rs("countryLCID")
	end if
	countryCurrency = rs("countryCurrency")
	useEuro = (countryCurrency="EUR")
	currDecimalSep=rs("currDecimalSep")
	currThousandsSep=rs("currThousandsSep")
	currPostAmount=rs("currPostAmount")
	currDecimals=rs("currDecimals")
	currSymbolHTML=rs("currSymbolHTML")
	pageurl = rs("adminStoreURL")
	if (left(lcase(pageurl),7) <> "http://") AND (left(LCase(pageurl),8) <> "https://") then pageurl = "http://" & pageurl
	if right(pageurl,1) <> "/" then pageurl = pageurl & "/"
	rs.close
end if
if forceloginonhttps then pageurl=""
if request.form("mode")="checkout" then
	if trim(request.form("checktmplogin"))="x" then
		SESSION("clientID")=empty
	elseif is_numeric(request.form("checktmplogin")) then
		sSQL = "SELECT tmploginname FROM tmplogin WHERE tmploginid='" & escape_string(request.form("sessionid")) & "' AND tmploginchk=" & replace(trim(request.form("checktmplogin")),"'","")
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then SESSION("clientID")=rs("tmploginname")
		rs.close
	end if
end if
sSQL = "SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity FROM cart WHERE cartCompleted=0 AND " & getsessionsql()
rs2.Open sSQL,cnn,0,1
do while NOT rs2.EOF
	optPriceDiff=0
	mcpdtxt = mcpdtxt & "<div class=""minicartcnt"">"&rs2("cartQuantity") &" " & rs2("cartProdName") & "</div>"
	sSQL = "SELECT SUM(coPriceDiff) AS sumDiff FROM cartoptions WHERE coCartID="&rs2("cartID")
	rs.open sSQL,cnn,0,1
	if NOT IsNull(rs("sumDiff")) then optPriceDiff=rs("sumDiff")
	rs.close
	subtot = ((rs2("cartProdPrice")+optPriceDiff)*int(rs2("cartQuantity")))
	totquant = totquant + int(rs2("cartQuantity"))
	mcgndtot=mcgndtot+subtot
	rs2.MoveNext
loop
rs2.Close
cnn.Close
set rs = nothing
set rs2 = nothing
set cnn = nothing
%>
      <table class="mincart" width="130" bgcolor="#FFFFFF">
        <tr class="mcrowtitle"> 
          <td class="mincart" bgcolor="#F0F0F0" align="center"><img src="images/littlecart1.png" style="vertical-align:text-top;" width="16" height="16" alt="<%=xxMCSC%>" /> 
            &nbsp;<strong><a class="ectlink mincart" href="<%=pageurl%>cart<%=extension%>"><%=xxMCSC%></a></strong></td>
        </tr>
<%		if request.form("mode")="movetocart" then %>
		<tr class="mcrowclcheck"><td class="mincart" bgcolor="#F0F0F0" align="center">Please click checkout to view your cart contents.</td></tr>
<%		elseif request.form("mode")="update" then %>
		<tr class="mcrowmainwin"><td class="mincart" bgcolor="#F0F0F0" align="center"><%=xxMainWn%></td></tr>
<%		else %>
        <tr class="mcrowtotquant"><td class="mincart" bgcolor="#F0F0F0" align="center"><%="<span class=""ectMCquant"">" & totquant & "</span> " & xxMCIIC %></td></tr>
		<tr class="mcrowlineitems"><td class="mincart" bgcolor="#F0F0F0"><div class="mcLNitems"><%=mcpdtxt%></div></td></tr>
<%			if mcpdtxt<>"" AND SESSION("discounts")<>"" AND NOT nopriceanywhere then discounts=cdbl(SESSION("discounts")) %>
        <tr class="ecHidDsc"<% if discounts=0 then print " style=""display:none"""%>><td class="mincart" bgcolor="#F0F0F0" align="center"><span style="color:#FF0000"><%=xxDscnts & " <span class=""mcMCdsct"">" & FormatMCCurrency(discounts)%></span></span></td></tr>
<%			if mcpdtxt<>"" AND SESSION("xsshipping")<>"" then
				shipping = cdbl(SESSION("xsshipping"))
				if shipping=0 then showshipping=" style=""color:#FF0000;font-weight:bold"">"&xxFree else showshipping=">"&FormatMCCurrency(shipping) %>
        <tr class="mcrowshipping"><td class="mincart" bgcolor="#F0F0F0" align="center"><%=xxMCShpE & " <span class=""ectMCship"""&showshipping&"</span>"%></td></tr>
<%			end if 
			if mcpdtxt="" then xscountrytax=0
			if NOT nopriceanywhere then %>
        <tr class="mcrowtotal"><td class="mincart" bgcolor="#F0F0F0" align="center"><%=xxTotal & " <span class=""ectMCtot"">" & FormatMCCurrency((mcgndtot+shipping+xscountrytax)-discounts)%></span></td></tr>
<%			end if
		end if %>
        <tr class="mcrowcheckout"><td class="mincart" bgcolor="#F0F0F0" align="center"><span style="font-family:Verdana">&raquo;</span> <a class="ectlink mincart" href="<%=pageurl%>cart<%=extension%>"><strong><%=xxMCCO%></strong></a></td></tr>
      </table>