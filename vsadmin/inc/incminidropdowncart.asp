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

if xxCkCoVC="" then xxCkCoVC="Please click checkout to view your cart contents."
if xxCMCSC="" then xxCMCSC="View Cart"
if xxCMCEmp="" then xxCMCEmp="Your shopping cart is currently empty"
if xxCMCClo="" then xxCMCClo="Close"
alreadygotadmin = getadminsettings()
if trim(request.form("sessionid"))<>"" then thesessionid=replace(trim(request.form("sessionid")),"'","") else thesessionid=getsessionid()

if getget("action")="deleteitem" AND is_numeric(getget("cartid")) then
	sSQL="SELECT cartID FROM cart WHERE cartID=" & escape_string(getget("cartid")) & " AND cartCompleted=0 AND " & getsessionsql()
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		ect_query("DELETE FROM cartoptions WHERE coCartID=" & rs("cartID"))
		ect_query("DELETE FROM cart WHERE cartID=" & rs("cartID"))
		ect_query("DELETE FROM giftcertificate WHERE gcCartID=" & rs("cartID"))
	end if
	rs.close
end if
mcgndtot=0 : totquant=0 : shipping=0 : mcdiscounts=0
if SESSION("xscountrytax")<>"" then xscountrytax=cdbl(SESSION("xscountrytax")) else xscountrytax=0
mcpdtxt=""
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
	sSQL = "SELECT imageSrc FROM productimages WHERE imageType=0 AND imageProduct='" & escape_string(rs2("cartProdID")) & "' ORDER BY imageNumber"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then productimageurl=rs("imageSrc") else productimageurl=""
	rs.close

	optPriceDiff=0
	sSQL="SELECT SUM(coPriceDiff) AS sumDiff FROM cartoptions WHERE coCartID="&rs2("cartID")
	rs.open sSQL,cnn,0,1
	if NOT IsNull(rs("sumDiff")) then optPriceDiff=rs("sumDiff")
	rs.close

	productqty=rs2("cartQuantity")
	productprice=rs2("cartProdPrice")+optPriceDiff

	subtot=((rs2("cartProdPrice")+optPriceDiff)*int(rs2("cartQuantity")))
	totquant=totquant + int(rs2("cartQuantity"))
	mcgndtot=mcgndtot+subtot
	prod=""

	mcpdtxt=mcpdtxt&"<div class=""minicartcnt_dp""><div class=""ectdp_minicartitemImg"">"
	if productimageurl="" then
		mcpdtxt=mcpdtxt&"<div style=""min-width:40px;"">&nbsp;</div>"
	else
		mcpdtxt=mcpdtxt&"<img src="""&productimageurl&""" class=""ectdp_minicartImg"" alt="""">"
	end if
	mcpdtxt=mcpdtxt&"</div><div class=""ectdp_minicartitemName"">" & IIfVs(nopriceanywhere,productqty&" ") & rs2("cartProdName") & IIfVs(NOT nopriceanywhere,"<br /><span class=""ectMCtot"">" & productqty & " x " & FormatEuroCurrency(productprice) & "</span>") & "</div><div class=""ectdp_minicartitemDelete""><img class=""cartdelete"" src=""images/delete.png"" alt="""&xxDelete&""" onclick=""dodeleteitem(" & rs2("cartID") & ")"" style=""cursor:pointer""></div><div style=""clear:both;""></div></div>"
	rs2.MoveNext
loop
rs2.close
	if getget("action")="" then %>
<div class="ectdp_minicartmainwrapper_ct" id="ectdp_minicartmainwrapper_ct">
<%	end if %>
<div class="ectdp_minicartmainwrapper">
    
	<div class="minicartcnt_dp ectdp_minicartopen" onmouseover="domcopen()" onmouseout="startmcclosecount()">
	<img src="images/arrow-down.png" style="vertical-align:text-top;" width="16" height="16" alt="<%=xxMCSC%>" /> &nbsp;<a class="ectlink mincart" href="<%=storeurl%>cart<%=extension%>"><%=xxCMCSC & IIfVs(NOT (getpost("mode")="movetocart" OR getpost("mode")="update")," (<span class=""ectMCquant"">" & totquant & "</span>)")%></a>
	</div>
    <div class="ectdp_minicartcontainer" id="ectdp_minicartcontainer" style="display:none;" onmouseover="domcopen()" onmouseout="startmcclosecount()">
<%
		if totquant=0 then
			print "<div class=""minicartcnt_dp ectdp_empty"">" & xxCMCEmp & "</div>"
			print "<div class=""minicartcnt_dp"">"&imageorbuttontag(imgminicartclose,xxCMCClo,"dpminicartclose","domcclose()",TRUE)&"</div>"
		else
			if getpost("mode")="movetocart" then %>  
	<div class="minicartcnt_dp ectdp_pincart"><%=xxCkCoVC%></div>
<%			elseif getpost("mode")="update" then %> 
	<div class="minicartcnt_dp ectdp_pincart"><%=xxMainWn%></div>
<%			else %>
	<div class="minicartcnt_dp ectdp_pincart"><%=totquant & " " & xxMCIIC %></div>
	<div class="mcdpLNitems"><%=mcpdtxt%></div>
<%				if mcpdtxt<>"" AND SESSION("discounts")<>"" AND NOT nopriceanywhere then mcdiscounts=cdbl(SESSION("discounts")) %>
	<div class="ecHidDsc minicartcnt_dp"<% if mcdiscounts=0 then print " style=""display:none"""%>><span class="minicartdsc"><%=xxDscnts & " <span class=""mcMCdsct"">" & FormatEuroCurrency(mcdiscounts)%></span></span></div>
<%				if mcpdtxt<>"" AND SESSION("xsshipping")<>"" then
					shipping=cdbl(SESSION("xsshipping"))
					if shipping=0 then showshipping=" minicartdsc"">" & xxFree else showshipping=""">" & FormatEuroCurrency(shipping) %>
   	<div class="minicartcnt_dp"><%=xxMCShpE & " <span class=""ectMCship" & showshipping & "</span>"%></div>
<%				end if
				if mcpdtxt="" then xscountrytax=0
				if NOT nopriceanywhere then %>
	<div class="minicartcnt_dp ectdp_minicarttotal"><%=xxTotal & " <span class=""ectMCtot"">" & FormatEuroCurrency((mcgndtot+shipping+xscountrytax)-mcdiscounts)%></span></div>
<%				end if
			end if %>
        <div class="minicartcnt_dp"> 
			<%=imageorbuttontag(imgminicartclose,xxCMCClo,"dpminicartclose","domcclose()",TRUE)%>
			&nbsp;&nbsp;&nbsp;&nbsp;
			<%=imageorbuttontag(imgminicartcheckout,xxMCCO,"dpminicartcheckout",storeurl&"cart"&extension,FALSE)%>
		</div>
<%		end if %>
    </div>
</div>
<%	if getget("action")="" then %>
</div>
<script>
var mctmrid=0,ajaxobj,ajaxobjrf;
function domcopen(){
	clearTimeout(mctmrid);
	document.getElementById('ectdp_minicartcontainer').style.display='';
}
function domcclose(){
	document.getElementById('ectdp_minicartcontainer').style.display='none';
}
function startmcclosecount(){
	mctmrid=setTimeout("domcclose()",400);
}
function mcpagerefresh(){
	if(ajaxobj.readyState==4){
		//document.getElementById('ectdp_minicartmainwrapper_ct').innerHTML=ajaxobj.responseText;
		document.location.reload();
	}
}
function refreshmcwindow(){
	if(ajaxobjrf.readyState==4){
		//alert(ajaxobjrf.responseText);
		document.getElementById('ectdp_minicartmainwrapper_ct').innerHTML=ajaxobjrf.responseText;
	}
}
function dodeleteitem(cartid){
	ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
	ajaxobj.onreadystatechange=mcpagerefresh;
	ajaxobj.open("GET", "vsadmin/miniajaxdropdowncart.asp?action=deleteitem&cartid="+cartid,true);
	ajaxobj.setRequestHeader("Content-type","application/x-www-form-urlencoded");
	ajaxobj.send('');
}
function dorefreshmctimer(){
	setTimeout("dorefreshmc()",1000);
}
function dorefreshmc(){
	ajaxobjrf=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
	ajaxobjrf.onreadystatechange=refreshmcwindow;
	ajaxobjrf.open("GET", "vsadmin/miniajaxdropdowncart.asp?action=refresh",true);
	ajaxobjrf.setRequestHeader("Content-type","application/x-www-form-urlencoded");
	ajaxobjrf.send('');
}
function addOnclick(elem, func) {
    var old=elem.onclick;
    if(typeof elem.onclick!='function'){
        elem.onclick=func;
    }else{
        elem.onclick=function(){
            if(old) old();
            func();
        };        
    }
}
function addbuttonclickevent(){
	var buybuttons=document.getElementsByClassName('buybutton');
	for(var i = 0; i < buybuttons.length; i++) {
		var buybutton=buybuttons[i];
		addOnclick(buybutton, dorefreshmctimer);
	}
}
if(window.addEventListener){
	window.addEventListener("load",addbuttonclickevent);
}else if(window.attachEvent)
    window.attachEvent("load", addbuttonclickevent);
</script>
<%	end if
cnn.close
set rs=nothing
set rs2=nothing
set cnn=nothing
%>