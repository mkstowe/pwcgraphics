<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if sDSN="" then response.write "Database connection not set" : response.end
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs2=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
if isempty(extension) then extension=".asp"
if xxCkCoVC="" then xxCkCoVC="Please click checkout to view your cart contents."
if cartpageonhttps then pageurl=storeurlssl else pageurl=storeurl
if getpost("sessionid")<>"" then thesessionid=replace(getpost("sessionid"),"'","") else thesessionid=getsessionid()
if minicssaction="onelineminicart" OR minicssaction="minicart" OR minicssaction="" then
	mcgndtot=0
	mcpdtxt=""
	totquant=0
	shipping=0
	discounts=0
	if SESSION("xscountrytax")<>"" then xscountrytax=cdbl(SESSION("xscountrytax")) else xscountrytax=0
	if request.form("mode")="checkout" then
		if trim(request.form("checktmplogin"))="x" then
			SESSION("clientID")=empty
		elseif trim(request.form("checktmplogin"))<>"" AND is_numeric(request.form("checktmplogin")) then
			sSQL="SELECT tmploginname FROM tmplogin WHERE tmploginid='" & escape_string(request.form("sessionid")) & "' AND tmploginchk=" & replace(trim(request.form("checktmplogin")),"'","")
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then SESSION("clientID")=rs("tmploginname")
			rs.close
		end if
	end if
	sSQL="SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity FROM cart WHERE cartCompleted=0 AND " & getsessionsql()
	rs2.Open sSQL,cnn,0,1
	do while NOT rs2.EOF
		optPriceDiff=0
		mcpdtxt=mcpdtxt & "<div class=""minicartcnt"">"&rs2("cartQuantity") &" " & rs2("cartProdName") & "</div>"
		sSQL="SELECT SUM(coPriceDiff) AS sumDiff FROM cartoptions WHERE coCartID="&rs2("cartID")
		rs.open sSQL,cnn,0,1
		if NOT IsNull(rs("sumDiff")) then optPriceDiff=rs("sumDiff")
		rs.close
		subtot=((rs2("cartProdPrice")+optPriceDiff)*Int(rs2("cartQuantity")))
		totquant=totquant + Int(rs2("cartQuantity"))
		mcgndtot=mcgndtot+subtot
		rs2.MoveNext
	loop
	rs2.Close
elseif minicssaction="minilogin" OR minicssaction="onelineminilogin" then
	if displaysoftlogindone="" then displaysoftlogindone=""
	if enableclientlogin then call displaysoftlogin()
	pageqs=""
	for each objitem in request.querystring
		if NOT (objitem="mode" AND (getget(objitem)="login" OR getget(objitem)="logout")) then pageqs=pageqs&IIfVs(pageqs<>"","&")&objitem&"="&getget(objitem)
	next
	if forceloginonhttps AND request.servervariables("HTTPS")="off" AND (replace(storeurl,"http:","https:")<>storeurlssl) then pagename="" else pagename=request.servervariables("URL") & IIfVs(pageqs<>"","?"&pageqs)
end if
if minicssaction="recentview" then
	if getpost("mode")<>"checkout" then
		if recentviewlayout="" then recentviewlayout="productname,productimage,category"
		recentviewlayoutarray=split(lcase(replace(recentviewlayout," ","")),",")
		sSQL="SELECT rvProdID,"&getlangid("pName",1)&",pSection,rvProdURL,pStaticPage,pStaticURL,"&getlangid("sectionName",256)&","&getlangid("sectionurl",2048)&" FROM (recentlyviewed INNER JOIN products ON recentlyviewed.rvProdID=products.pID) INNER JOIN sections ON products.pSection=sections.sectionID WHERE rvProdID<>'"&escape_string(prodid)&"' AND " & IIfVr(SESSION("clientID")<>"", "rvCustomerID="&replace(SESSION("clientID"),"'",""), "(rvCustomerID=0 AND rvSessionID='"&thesessionid&"')")&" ORDER BY rvDate DESC"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			if recentviewlink="" then recentviewlink=IIfVr(seocategoryurls,replace(seoprodurlpattern,"%s",""),"products.asp")&"?recentview=true"%>
		  <div class="recentview">
			<div class="recentviewheader"><img src="images/recentview.png" style="vertical-align:text-top;" alt="<%=xxRecVie%>" />
				&nbsp;<a class="ectlink recentview" href="<%=recentviewlink%>"><%=xxRecVie%></a></div>
<%			do while NOT rs.EOF
				thedetailslink=getdetailsurl(rs("rvProdID"),rs("pStaticPage"),rs(getlangid("pName",1)),trim(rs("pStaticURL")&""),"",pathtohere)
				print "<div class=""recentviewline ectclearfix"">"
				for each layoutoption in recentviewlayoutarray
					if layoutoption="productimage" then
						sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 1 ")&"imageSrc FROM productimages WHERE imageType=0 AND imageProduct='"&escape_string(rs("rvProdID"))&"' ORDER BY imageNumber"&IIfVs(mysqlserver=TRUE," LIMIT 0,1")
						rs2.open sSQL,cnn,0,1
						if NOT rs2.EOF then recentimage=rs2("imageSrc") else recentimage=""
						rs2.close
						if recentimage<>"" then %>
				<div class="recentviewimage">
					<a class="ectlink recentview" href="<%=thedetailslink%>"><img class="recentviewimage" src="<%=recentimage%>" alt="<%=strip_tags2(rs(getlangid("pName",1)))%>" /></a>
				</div>
<%						end if
					elseif layoutoption="category" then
						if trim(rs("sectionurl")&"")<>"" then
							recentsection=getcatid(rs("sectionurl"),rs("sectionurl"),seoprodurlpattern)
						else
							recentsection=IIfVs(NOT seocategoryurls,"products.asp?cat=") & getcatid(rs("pSection"),rs(getlangid("sectionName",256)),seoprodurlpattern)
						end if %>
				<div class="recentviewcategory">
					<a class="ectlink recentview" href="<%=recentsection%>"><%=rs(getlangid("sectionName",256))%></a>
				</div>
<%					elseif layoutoption="productname" then %>
				<div class="recentviewname">
					<a class="ectlink recentview" href="<%=thedetailslink%>"><%=rs(getlangid("pName",1))%></a>
				</div>
<%					end if
				next
				print "</div>"
				rs.MoveNext
			loop %>		
		  </div>
	<%	end if
		rs.close
	end if
elseif minicssaction="minilogin" then
%>
<div class="minicart">
	<div class="minicartcnt">
	<img src="images/minipadlock.png" style="vertical-align:text-top;" alt="<%=xxMLLIS%>" /> <%
	if SESSION("clientID")<>"" AND customeraccounturl<>"" then %>
		&nbsp;<a class="ectlink mincart" href="<%=customeraccounturl%>"><%=xxYouAcc%></a>
<%	else
		response.write "&nbsp;"&xxMLLIS
	end if %>
	</div>
<%	if NOT enableclientlogin then %>  
	<div class="minicartcnt">Client login not enabled</div>
<%	elseif SESSION("clientID")<>"" AND request.querystring("mode")<>"logout" then %>
	<div class="minicartcnt"><%=xxMLLIA&"<br />"&server.htmlencode(SESSION("clientUser"))%></div>
	<div class="minicartcnt"><%=imageorbuttontag(imgminicsslogout,xxLogout,"ectlink mincart","return dologoutaccount()",TRUE)%></div>
<%	else %>
	<div class="minicartcnt"><%=xxMLNLI%></div>
	<div class="minicartcnt"><%=imageorbuttontag(imgminicsslogin,xxLogin,"ectlink mincart","return displayloginaccount()",TRUE)%></div>
<%	end if %>
</div>
<%
elseif minicssaction="onelineminilogin" then
	startlink=""
	endlink=""
	if SESSION("clientID")<>"" AND customeraccounturl<>"" then
		startlink="<a class=""ectlink mincart"" href=""" & customeraccounturl & """>"
		endlink="</a>"
	end if
%>
<div class="minicartoneline">
	<div class="minicartoneline1"><%=startlink%><img src="images/minipadlock.png" alt="<%=xxMLLIS%>" /><%=endlink%></div>
<%	if NOT enableclientlogin then %>  
		<div class="minicartoneline1">Client login not enabled</div>
<%	elseif SESSION("clientID")<>"" AND request.querystring("mode")<>"logout" then %>
	<div class="minicartoneline2"><%=xxMLLIA&" "&server.htmlencode(SESSION("clientUser"))%></div>
	<div class="minicartoneline3"><%=imageorbuttontag(imgminicsslogout,xxLogout,"ectlink mincart","return dologoutaccount()",TRUE)%></div>
<%	else %>
	<div class="minicartoneline2"><%=xxMLNLI%></div>
	<div class="minicartoneline3"><%=imageorbuttontag(imgminicsslogin,xxLogin,"ectlink mincart","return displayloginaccount()",TRUE)%></div>
<%	end if %>
</div>
<%
elseif minicssaction="minisignup" then
	if SESSION("MLSIGNEDUP")=TRUE OR request.form("mode")="mailinglistsignup" then
		response.write "<div class=""minimailsignup"">"&xxThkSub&"</div>"
	else
		therp=Request.ServerVariables("URL")&IIfVr(trim(Request.ServerVariables("QUERY_STRING"))<>"","?","")&Request.ServerVariables("QUERY_STRING")
%>
<script>/* <![CDATA[ */
function mlvalidator(frm){
	var mlsuemail=document.getElementById('mlsuname');
	if(mlsuemail.value==""){
		alert("<%=jscheck(xxPlsEntr&" """&xxName)%>\".");
		mlsuemail.focus();
		return(false);
	}
	var mlsuemail=document.getElementById('mlsuemail');
	if(mlsuemail.value==""){
		alert("<%=jscheck(xxPlsEntr&" """&xxEmail)%>\".");
		mlsuemail.focus();
		return(false);
	}
	var regex=/[^@]+@[^@]+\.[a-z]{2,}$/i;
	if(!regex.test(mlsuemail.value)){
		alert("<%=jscheck(xxValEm)%>");
		mlsuemail.focus();
		return(false);
	}
	document.getElementById('mlsectgrp1').value=(document.getElementById('mlsuemail').value.split('@')[0].length);
	document.getElementById('mlsectgrp2').value=(document.getElementById('mlsuemail').value.split('@')[1].length);
	return(true);
}
/* ]]> */
</script>
<div id="ectform" class="minimailsignup">
	<form action="cart<%=extension%>" method="post" onsubmit="return mlvalidator(this)">
		<input type="hidden" name="mode" value="mailinglistsignup" />
		<input type="hidden" name="mlsectgrp1" id="mlsectgrp1" value="7418" />
		<input type="hidden" name="mlsectgrp2" id="mlsectgrp2" value="6429" />
		<input type="hidden" name="rp" value="<%=replace(replace(replace(therp,"""",""),"<",""),"&","&amp;")%>" />
		<input type="hidden" name="posted" value="1" />
		<div class="minimailcontainer">
			<label class="minimailsignup"><%=xxName%></label>
			<input class="ecttextinput minimailsignup" type="text" name="mlsuname" id="mlsuname" value="" maxlength="50" />
		</div>
		<div class="minimailcontainer">
			<label class="minimailsignup"><%=xxEmail%></label>
			<input class="ecttextinput minimailsignup" type="text" name="mlsuemail" id="mlsuemail" value="" maxlength="50" />
		</div>
		<div class="minimailcontainer">
			<%=imageorsubmit(imgmailformsubmit, xxSubmt, "minimailsignup minimailsubmit")%>
		</div>
	</form>
</div>
<%	end if
elseif minicssaction="onelineminicart" then
	if mcpdtxt<>"" AND SESSION("discounts")<>"" AND NOT nopriceanywhere then discounts=cdbl(SESSION("discounts")) %>
<div class="minicartoneline">
<%		if request.form("mode")="movetocart" then %>  
	<div class="minicartoneline1"><%=xxCkCoVC%></div>
<%		elseif request.form("mode")="update" then %>
	<div class="minicartoneline1"><%=xxMainWn%></div>
<%		else %>
	<div class="minicartoneline1"><span class="ectMCquant"><%=totquant & "</span> " & xxMCIIC %> | </div>
	<div class="minicartoneline2"><%=xxTotal & " <span class=""ectMCtot"">" & FormatEuroCurrency(mcgndtot-discounts)%></span> | </div>
<%		end if %>
	<div class="minicartoneline3"><img src="images/littlecart1.png" style="vertical-align:text-top;" width="16" height="16" alt="<%=xxMCSC%>" /> &nbsp;<a class="ectlink mincart" href="<%=pageurl%>cart<%=extension%>"><%=xxMCSC%></a></div>
</div>
<%
else %>
<div class="minicart">
	<div class="minicartcnt">
	<img src="images/littlecart1.png" style="vertical-align:text-top;" width="16" height="16" alt="<%=xxMCSC%>" /> &nbsp;<a class="ectlink mincart" href="<%=pageurl%>cart<%=extension%>"><%=xxMCSC%></a>
	</div>
<%		if request.form("mode")="movetocart" then %>  
	<div class="minicartcnt"><%=xxCkCoVC%></div>
<%		elseif request.form("mode")="update" then %>
	<div class="minicartcnt"><%=xxMainWn%></div>
<%		else %>
	<div class="minicartcnt"><span class="ectMCquant"><%=totquant & "</span> " & xxMCIIC %></div>
	<div class="mcLNitems"><%=mcpdtxt%></div>
<%			if mcpdtxt<>"" AND SESSION("discounts")<>"" then discounts=cdbl(SESSION("discounts")) %>
	<div class="ecHidDsc minicartcnt"<% if discounts=0 then print " style=""display:none"""%>><span class="minicartdsc"><%=xxDscnts & " <span class=""mcMCdsct"">" & FormatEuroCurrency(discounts)%></span></span></div>
<%			if mcpdtxt<>"" AND SESSION("xsshipping")<>"" then
				shipping=cdbl(SESSION("xsshipping"))
				if shipping=0 then showshipping=" minicartdsc"">"&xxFree else showshipping=""">"&FormatEuroCurrency(shipping) %>
	<div class="minicartcnt"><%=xxMCShpE & " <span class=""ectMCship" & showshipping&"</span>"%></div>
<%			end if 
			if mcpdtxt="" then xscountrytax=0
			if NOT nopriceanywhere then %>
	<div class="minicartcnt"><%=xxTotal & " <span class=""ectMCtot"">" & FormatEuroCurrency((mcgndtot+shipping+xscountrytax)-discounts)%></span></div>
<%			end if
		end if %>
	<div class="minicartcnt"><%=imageorbuttontag(imgminicsslogin,xxMCCO,"ectlink mincart",pageurl&"cart"&extension,FALSE)%></div>
</div>
<%
end if
cnn.Close
set rs=nothing
set rs2=nothing
set cnn=nothing
%>