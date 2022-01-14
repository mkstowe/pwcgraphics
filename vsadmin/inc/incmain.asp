<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
Dim sSQL,rs,alldata,success,cnn,errmsg,index,allcountries
success=true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
session.LCID=1033
numshipmethods=10
if getpost("posted")="1" then
	adminLangSettings=0
	for each objItem in request.form("adminlangsettings")
		adminLangSettings = adminLangSettings + Int(objItem)
	next
	DIM pfarr(2),pftext(2),pftext2(2),pftext3(2)
	for numfilters=0 to 1
		pfarr(numfilters)=0
		pftext(numfilters)=""
		pftext2(numfilters)=""
		pftext3(numfilters)=""
		for index=0 to 5
			if getpost("filtercb"&numfilters&"_"&index)="ON" then pfarr(numfilters)=pfarr(numfilters)+(2 ^ index)
			if index<>0 then
				pftext(numfilters)=pftext(numfilters)&"&"
				pftext2(numfilters)=pftext2(numfilters)&"&"
				pftext3(numfilters)=pftext3(numfilters)&"&"
			end if
			pftext(numfilters)=pftext(numfilters)&replace(getpost("filtertext"&numfilters&"_"&index),"&","%26")
			pftext2(numfilters)=pftext2(numfilters)&replace(getpost("filtertext"&numfilters&"_"&index&"x2"),"&","%26")
			pftext3(numfilters)=pftext3(numfilters)&replace(getpost("filtertext"&numfilters&"_"&index&"x3"),"&","%26")
		next
	next
	sortoptions=0
	for index=1 to 20
		if getpost("sortid"&index)="ON" then sortoptions=sortoptions+(2 ^ (index-1))
	next
	sSQL = "UPDATE admin SET adminEmail='"&getpost("xxemaddr")&"',emailfromname='"&escape_string(getpost("emailfromname"))&"',adminStoreURL='"&getpost("url")&"',adminStoreURLSSL='"&getpost("urlssl")&"',adminProdsPerPage='"&getpost("prodperpage")&"',adminShipping="&getpost("shipping")&",adminIntShipping="&getpost("intshipping")&",adminZipCode='"&getpost("zipcode")&"',adminCountry="&getpost("countrySetting")&",adminDelUncompleted="&getpost("deleteUncompleted")&",adminClearCart="&getpost("adminClearCart")&",adminStockManage="&getpost("stockManage")&",adminHandling="&IIfVr(NOT is_numeric(getpost("handling")), 0, getpost("handling"))&",adminHandlingPercent="&IIfVr(getpost("handlingpercent")="" OR NOT is_numeric(getpost("handlingpercent")), 0, getpost("handlingpercent"))&",emailObject="&getpost("emailObject")&","
	sSQL=sSQL&"smtpserver='"&getpost("smtpserver")&"',emailUser='"&getpost("emailuser")&"'"
	sSQL=sSQL&",smtpport='" & escape_string(getpost("smtpport")) & "'"
	sSQL=sSQL&",smtpsecure='" & escape_string(getpost("smtpsecure")) & "'"
	sSQL=sSQL&",htmlemails=" & escape_string(getpost("htmlemails"))
	if getpost("deleteemailpass")="delete" then
		sSQL=sSQL&",emailPass=''"
	elseif getpost("emailpass")<>"" then
		sSQL=sSQL&",emailPass='"&getpost("emailpass")&"'"
	end if
	if getpost("deletecurrconvpass")="delete" then
		sSQL=sSQL&",currConvPw=''"
	elseif getpost("currConvPw")<>"" then
		sSQL=sSQL&",currConvPw='"&getpost("currConvPw")&"'"
	end if
	if getpost("emailconfirm")="ON" then adminEmailConfirm=1 else adminEmailConfirm=0
	if getpost("affilconfirm")="ON" then adminEmailConfirm=adminEmailConfirm+2
	if getpost("customerconfirm")="ON" then adminEmailConfirm=adminEmailConfirm+4
	if getpost("reviewconfirm")="ON" then adminEmailConfirm=adminEmailConfirm+8
	sSQL = sSQL & ",adminEmailConfirm="&adminEmailConfirm&","
	sSQL = sSQL & "adminUnits=" & (int(getpost("adminUnits")) + int(getpost("adminDims")))
	for index=1 to 3
		sSQL = sSQL & ",currRate" & index & "=" & IIfVr(is_numeric(getpost("currRate" & index)),getpost("currRate" & index),0)
		sSQL = sSQL & ",currSymbol" & index & "='" & escape_string(getpost("currSymbol" & index)) & "'"
	next
	sSQL = sSQL & ",currLastUpdate=" & vsusdatetime(Now()-10)
	sSQL = sSQL & ",currConvUser='" & getpost("currConvUser") & "'"
	sSQL = sSQL & ",cardinalProcessor='" & escape_string(getpost("cardinalprocessor")) & "'"
	sSQL = sSQL & ",cardinalMerchant='" & escape_string(getpost("cardinalmerchant")) & "'"
	sSQL = sSQL & ",cardinalPwd='" & escape_string(getpost("cardinalpwd")) & "'"
	sSQL = sSQL & ",adminlanguages=" & getpost("adminlanguages")
	sSQL = sSQL & ",adminlang='" & getpost("adminlang") & "'"
	sSQL = sSQL & ",storelang='" & getpost("storelang1") & "|" & getpost("storelang2") & "|" & getpost("storelang3") & "'"
	sSQL = sSQL & ",adminAltRates=" & getpost("adminAltRates")
	sSQL = sSQL & ",prodFilter=" & pfarr(0)
	sSQL = sSQL & ",sideFilter=" & pfarr(1)
	sSQL = sSQL & ",prodFilterOrder='" & escape_string(getpost("prodfilterorder0")) & "'"
	sSQL = sSQL & ",sideFilterOrder='" & escape_string(getpost("prodfilterorder1")) & "'"
	sSQL = sSQL & ",prodFilterText='" & escape_string(pftext(0)) & "'"
	sSQL = sSQL & ",sideFilterText='" & escape_string(pftext(1)) & "'"
	if (adminlangsettings AND 262144)=262144 then
		if adminlanguages>=1 then sSQL = sSQL & ",prodFilterText2='" & escape_string(pftext2(0)) & "',sideFilterText2='" & escape_string(pftext2(1)) & "'"
		if adminlanguages>=2 then sSQL = sSQL & ",prodFilterText3='" & escape_string(pftext3(0)) & "',sideFilterText3='" & escape_string(pftext3(1)) & "'"
	end if
	sSQL = sSQL & ",sortOrder=" & getpost("sortorder")
	sSQL = sSQL & ",sortOptions=" & sortoptions
	sSQL = sSQL & ",reCAPTCHAsitekey='" & escape_string(getpost("reCAPTCHAsitekey")) & "'"
	sSQL = sSQL & ",reCAPTCHAsecret='" & escape_string(getpost("reCAPTCHAsecret")) & "'"
	sSQL = sSQL & ",uploadDir='" & escape_string(getpost("uploaddir")) & "'"
	sSQL = sSQL & ",shipInsuranceInt=" & IIfVr(is_numeric(getpost("shipinsuranceint")),getpost("shipinsuranceint"),0)
	sSQL = sSQL & ",shipInsuranceDom=" & IIfVr(is_numeric(getpost("shipinsurancedom")),getpost("shipinsurancedom"),0)
	sSQL = sSQL & ",insuranceIntMin=" & IIfVr(is_numeric(getpost("insuranceintmin")),getpost("insuranceintmin"),0)
	sSQL = sSQL & ",insuranceDomMin=" & IIfVr(is_numeric(getpost("insurancedommin")),getpost("insurancedommin"),0)
	sSQL = sSQL & ",insuranceIntPercent=" & IIfVr(is_numeric(getpost("insuranceintpercent")),getpost("insuranceintpercent"),0)
	sSQL = sSQL & ",insuranceDomPercent=" & IIfVr(is_numeric(getpost("insurancedompercent")),getpost("insurancedompercent"),0)
	sSQL = sSQL & ",noCarrierDomIns=" & IIfVr(is_numeric(getpost("nocarrierdomins")),getpost("nocarrierdomins"),0)
	sSQL = sSQL & ",noCarrierIntIns=" & IIfVr(is_numeric(getpost("nocarrierintins")),getpost("nocarrierintins"),0)
	
	sSQL = sSQL & ",onvacation=" & getpost("onvacation")
	sSQL = sSQL & ",vacationmessage='" & escape_string(getpost("vacationmessage")) & "'"

	recaptchauseon=0
	if getpost("recaptcha1")="ON" then recaptchauseon=recaptchauseon+1
	if getpost("recaptcha2")="ON" then recaptchauseon=recaptchauseon+2
	if getpost("recaptcha3")="ON" then recaptchauseon=recaptchauseon+4
	if getpost("recaptcha4")="ON" then recaptchauseon=recaptchauseon+8
	if getpost("recaptcha5")="ON" then recaptchauseon=recaptchauseon+16
	if getpost("recaptcha6")="ON" then recaptchauseon=recaptchauseon+32
	if getpost("recaptcha7")="ON" then recaptchauseon=recaptchauseon+64
	if getpost("recaptcha8")="ON" then recaptchauseon=recaptchauseon+128
	if getpost("recaptcha9")="ON" then recaptchauseon=recaptchauseon+256
	if getpost("recaptcha10")="ON" then recaptchauseon=recaptchauseon+512
	sSQL = sSQL & ",reCAPTCHAuseon=" & recaptchauseon
	
	sSQL = sSQL & ",adminlangsettings=" & adminLangSettings & " WHERE adminID=1"
	if NOT ectdemostore then ect_query(sSQL)
	
	sSQL = "UPDATE adminshipping SET adminPacking="&getpost("packing")&" WHERE adminShipID=1"
	if NOT ectdemostore then ect_query(sSQL)
	
	sSQL = "SELECT adminSecret FROM admin WHERE adminID=1"
	rs.open sSQL,cnn,0,1
	currsecret=trim(rs("adminSecret")&"")
	rs.close
	randomize	
	if currsecret="" AND NOT ectdemostore then ect_query("UPDATE admin SET adminSecret='itz a "&(int(100000000 * rnd) + 100000000)&" real deal secret "&(int(100000000 * rnd) + 100000000)&" no' WHERE adminID=1")
	
	altrateids = split(getpost("altrateids"),",")
	altrateidsintl = split(getpost("altrateidsintl"),",")
	altrateuse = split(getpost("altrateuse"),",")
	altrateuseintl = split(getpost("altrateuseintl"),",")
	altratetext = split(getpost("altratetext"),",")
	altratetext2 = split(getpost("altratetext2"),",")
	altratetext3 = split(getpost("altratetext3"),",")
	
	for index=1 to numshipmethods
		if index=1 AND trim(altratetext(index-1))="" then altratetext(index-1)=yyFlatShp
		if index=2 AND trim(altratetext(index-1))="" then altratetext(index-1)=yyWghtShp
		if index=3 AND trim(altratetext(index-1))="" then altratetext(index-1)=yyUSPS
		if index=4 AND trim(altratetext(index-1))="" then altratetext(index-1)=yyUPS
		if index=5 AND trim(altratetext(index-1))="" then altratetext(index-1)=yyPriShp
		if index=6 AND trim(altratetext(index-1))="" then altratetext(index-1)=yyCanPos
		if index=7 AND trim(altratetext(index-1))="" then altratetext(index-1)=yyFedex
		if index=8 AND trim(altratetext(index-1))="" then altratetext(index-1)="FedEx SmartPost"
		if index=9 AND trim(altratetext(index-1))="" then altratetext(index-1)=yyDHLShp
		
		if trim(altratetext2(index-1))="" then altratetext2(index-1)=altratetext(index-1)
		if trim(altratetext3(index-1))="" then altratetext3(index-1)=altratetext(index-1)
		
		altratetext(index-1)=left(urldecode(altratetext(index-1)),255)
		altratetext2(index-1)=left(urldecode(altratetext2(index-1)),255)
		altratetext3(index-1)=left(urldecode(altratetext3(index-1)),255)

		sSQL = "UPDATE alternaterates SET altratetext='"&escape_string(altratetext(index-1))&"',altratetext2='"&escape_string(altratetext2(index-1))&"',altratetext3='"&escape_string(altratetext3(index-1))&"', usealtmethod="&altrateuse(index-1)&",altrateorder="&index&" WHERE altrateid="&altrateids(index-1)
		ect_query(sSQL)
		sSQL = "UPDATE alternaterates SET usealtmethodintl="&altrateuseintl(index-1)&",altrateorderintl="&index&" WHERE altrateid="&altrateidsintl(index-1)
		ect_query(sSQL)
	next
	print "<meta http-equiv=""refresh"" content=""1; url=adminmain.asp"">"
else
	sSQL = "SELECT countryID,countryName,countryCurrency FROM countries WHERE countryLCID<>'0' AND countryLCID<>''  ORDER BY countryOrder DESC, countryName"
	rs.open sSQL,cnn,0,1
	allcountries=rs.getrows
	rs.close
	sSQL = "SELECT DISTINCT countryCurrency FROM countries ORDER BY countryCurrency"
	rs.open sSQL,cnn,0,1
	allcurrencies=rs.getrows
	rs.close
end if
if getpost("posted")="1" AND success then %>
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
				<td width="100%" colspan="2" align="center"><br /><%
				if ectdemostore then %>
				<strong>Updating of Admin Settings has been disabled for the Demo Store.</strong>
<%				else %>
				<strong><%=yyUpdSuc%></strong>
<%				end if %>
				<br /><br /><%=yyNowFrd%><br /><br />
						<%=yyNoAuto%><a href="adminmain.asp"><strong><%=yyClkHer%></strong></a>.<br /><br />&nbsp;</td>
			  </tr>
			</table>
<%
else
	sSQL="SELECT adminPacking FROM adminshipping WHERE adminShipID=1"
	rs.open sSQL,cnn,0,1
	adminPacking=rs("adminPacking")
	rs.close

	sSQL="SELECT adminEmail,emailfromname,adminStoreURL,adminStoreURLSSL,adminProdsPerPage,adminShipping,adminIntShipping,adminZipCode,adminEmailConfirm,adminCountry,adminUnits,adminDelUncompleted,adminClearCart,adminStockManage,adminHandling,adminHandlingPercent,currRate1,currSymbol1,currRate2,currSymbol2,currRate3,currSymbol3,currConvUser,currConvPw,cardinalProcessor,cardinalMerchant,cardinalPwd,emailObject,smtpserver,smtpport,smtpsecure,htmlemails,emailUser,emailPass,adminlanguages,adminlangsettings,adminAltRates,prodFilter,prodFilterOrder,prodFilterText,prodFilterText2,prodFilterText3,sideFilter,sideFilterOrder,sideFilterText,sideFilterText2,sideFilterText3,sortOrder,sortOptions,adminlang,storelang,reCAPTCHAsitekey,reCAPTCHAsecret,reCAPTCHAuseon,onvacation,uploadDir,shipInsuranceInt,shipInsuranceDom,insuranceIntMin,insuranceDomMin,insuranceIntPercent,insuranceDomPercent,noCarrierDomIns,noCarrierIntIns FROM admin WHERE adminID=1"
	rs.open sSQL,cnn,0,1
%>
<script>
<!--
function formvalidator(theForm){
  if(theForm.prodperpage.value==""){
	alert("<%=jscheck(yyPlsEntr&" """&yyPPP)%>\".");
	theForm.prodperpage.focus();
	return (false);
  }
  var checkOK = "0123456789";
  var checkStr = theForm.prodperpage.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++){
	ch = checkStr.charAt(i);
	for (j = 0;  j < checkOK.length;  j++)
		if (ch == checkOK.charAt(j))
			break;
	if (j == checkOK.length){
		allValid = false;
			break;
	}
  }
  if (!allValid){
	alert("<%=jscheck(yyOnlyNum&" """&yyPPP)%>\".");
	theForm.prodperpage.focus();
	return (false);
  }
for(index=1;index<=3;index++){
  var checkOK = "0123456789.";
  var thisRate = eval("theForm.currRate" + index);
  var checkStr = thisRate.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++){
	ch = checkStr.charAt(i);
	for (j = 0;  j < checkOK.length;  j++)
		if (ch == checkOK.charAt(j))
			break;
	if (j == checkOK.length){
		allValid = false;
			break;
	}
  }
  if(!allValid){
	alert("<%=jscheck(yyOnlyDec&" """&yyConRat)%> " + index + "\".");
	thisRate.focus();
	return (false);
  }
}

  if(theForm.handling.value==""){
	alert('<%=jscheck(yyPlsEntr)%> "<%=jscheck(yyHanChg)%>". <%=jscheck(yyNoHan)%>');
	theForm.handling.focus();
	return (false);
  }
  var checkOK = "0123456789.";
  var checkStr = theForm.handling.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++){
	ch = checkStr.charAt(i);
	for (j = 0;  j < checkOK.length;  j++)
		if (ch == checkOK.charAt(j))
			break;
	if (j == checkOK.length){
		allValid = false;
			break;
	}
  }
  if(!allValid){
	alert("<%=jscheck(yyOnlyDec&" """&yyHanChg)%>\".");
	theForm.handling.focus();
	return (false);
  }
  
	var altrateids="";
	var altrateidsintl="";
	var altrateuse="";
	var altrateuseintl="";
	var altratetext="";
	var altratetext2="";
	var altratetext3="";
	for(var index=1; index<=<%=numshipmethods%>; index++){
		altrateids+=(document.getElementById("altrateids"+index).value+',');
		altrateidsintl+=(document.getElementById("altrateidsintl"+index).value+',');
		altrateuse+=((document.getElementById("altrateuse"+index).checked?'1':'0')+',');
		altrateuseintl+=((document.getElementById("altrateuseintl"+index).checked?'1':'0')+',');
		altratetext+=(encodeURIComponent(document.getElementById("altratetext_"+index).value)+',');
		altratetext2+=(encodeURIComponent(document.getElementById("altratetext2_"+index).value)+',');
		altratetext3+=(encodeURIComponent(document.getElementById("altratetext3_"+index).value)+',');
	}
	altrateids=altrateids.substr(0,altrateids.length-1);
	altrateidsintl=altrateidsintl.substr(0,altrateidsintl.length-1);
	altrateuse=altrateuse.substr(0,altrateuse.length-1);
	altrateuseintl=altrateuseintl.substr(0,altrateuseintl.length-1);
	altratetext=altratetext.substr(0,altratetext.length-1);
	altratetext2=altratetext2.substr(0,altratetext2.length-1);
	altratetext3=altratetext3.substr(0,altratetext3.length-1);
	// alert(altrateids+"\n"+altrateuse+"\n"+altrateuseintl+"\n"+altratetext+"\n"+altratetext2+"\n"+altratetext3);
	document.getElementById("altrateids").value=altrateids;
	document.getElementById("altrateidsintl").value=altrateidsintl;
	document.getElementById("altrateuse").value=altrateuse;
	document.getElementById("altrateuseintl").value=altrateuseintl;
	document.getElementById("altratetext").value=altratetext;
	document.getElementById("altratetext2").value=altratetext2;
	document.getElementById("altratetext3").value=altratetext3;
	return (true);
}
<%	if trim(rs("prodFilterOrder")&"")="" then prodfilterorder="1,2,4,8,16,32" else prodfilterorder=trim(rs("prodFilterOrder"))
	if trim(rs("sideFilterOrder")&"")="" then sidefilterorder="1,2,4,8,16,32" else sidefilterorder=trim(rs("sideFilterOrder"))
	prodfilterorder=replace(prodfilterorder&",","1,","") ' Because Manufacturer (1) is now included with attributes
	sidefilterorder=replace(sidefilterorder&",","1,","")
	if right(prodfilterorder,1)="," then prodfilterorder=left(prodfilterorder,len(prodfilterorder)-1)
	if right(sidefilterorder,1)="," then sidefilterorder=left(sidefilterorder,len(sidefilterorder)-1)
%>
var currfilterorder=[];
currfilterorder[0]='<%=prodfilterorder%>'.split(',');
currfilterorder[1]='<%=sidefilterorder%>'.split(',');
function swapFilterTRows(whichfilter,fromrow,torow){
	var srtable=document.getElementById("filtertable"+whichfilter);
	if(srtable.moveRow){
		srtable.moveRow(fromrow+1,torow+1);
	}else{ // FF etc
		var firstRow=srtable.rows[fromrow+1];
		firstRow.parentNode.insertBefore(srtable.rows[torow+1],firstRow);
	}
}
function swapfilteritems(whichfilter,item){
	var thisrow=document.getElementById('filter'+whichfilter+'item'+item);
	var thistemphtml=thisrow.innerHTML;
	var thistempid=thisrow.id,currpos;
	for(var ii in currfilterorder[whichfilter]){
		if(currfilterorder[whichfilter][ii]==item) currpos=parseInt(ii);
	}
	if(currpos!=0){
		swapFilterTRows(whichfilter,currpos-1,currpos)
		var temporder=currfilterorder[whichfilter][currpos-1];
		currfilterorder[whichfilter][currpos-1]=currfilterorder[whichfilter][currpos]
		currfilterorder[whichfilter][currpos]=temporder;
		document.getElementById('prodfilterorder'+whichfilter).value=currfilterorder[whichfilter].join();
	}
	return false;
}
var savetabletext=[];
function showhiderow(ishide,tablenum,rownum){
	var thetable=document.getElementById('maintable'+tablenum);
	thetable.rows[rownum].className=ishide?'maintablehidden':'maintablevisible';
}
function hidetablerows(tablenum){
	var thetable=document.getElementById('maintable'+tablenum);
	var tbutton=document.getElementById('mainbutton'+tablenum);
	var tablerows=thetable.rows.length;
	for(var i=1;i < tablerows; i++){
		if(i==1)
			thetable.rows[i].className='maintablehidden';
		else
			setTimeout('showhiderow(true,'+tablenum+','+i+');',30*(i-1));
	}
	tbutton.innerHTML=savetabletext[tablenum]+' <div style="float:right">&#9660;</div>';
}
function showtablerows(tablenum,ttext){
	var thetable=document.getElementById('maintable'+tablenum);
	var tbutton=document.getElementById('mainbutton'+tablenum);
	savetabletext[tablenum]=ttext
	if(thetable.rows[1].className=='maintablevisible'){
		hidetablerows(tablenum);
	}else{
		var tablerows=thetable.rows.length;
		for(var i=1;i < tablerows; i++){
			if(i==1)
				thetable.rows[i].className='maintablevisible';
			else
				setTimeout('showhiderow(false,'+tablenum+','+i+');',30*(i-1));
		}
		tbutton.innerHTML=ttext+' <div style="float:right">&#9650;</div>';
	}
}
function showhidealtrates(obj){
	if(obj.options[obj.selectedIndex].value=="0"){
		document.getElementById('altraterowtitle').style.display='none';
		document.getElementById('altraterow').style.display='none';
		document.getElementById('domesticraterow0').style.display='';
		document.getElementById('domesticraterow1').style.display='';
		document.getElementById('domesticraterow2').style.display='';
		document.getElementById('domesticraterow3').style.display='';
		domesticraterow3
	}else{
		document.getElementById('altraterowtitle').style.display='';
		document.getElementById('altraterow').style.display='';
		document.getElementById('domesticraterow0').style.display='none';
		document.getElementById('domesticraterow1').style.display='none';
		document.getElementById('domesticraterow2').style.display='none';
		document.getElementById('domesticraterow3').style.display='none';
	}
}
function swaptbrows(rid){
	if(rid!=1){
		rid2=rid-1;
		var altrateids=document.getElementById("altrateids"+rid).value;
		var altrateuse=document.getElementById("altrateuse"+rid).checked;
		var methodname=document.getElementById("methodname"+rid).innerHTML;
		var methodnamedom=document.getElementById("methodnamedom"+rid).innerHTML;
		var altratetext=document.getElementById("altratetext_"+rid).value;
		var altratetext2=document.getElementById("altratetext2_"+rid).value;
		var altratetext3=document.getElementById("altratetext3_"+rid).value;
		
		document.getElementById("altrateids"+rid).value=document.getElementById("altrateids"+rid2).value;
		document.getElementById("altrateuse"+rid).checked=document.getElementById("altrateuse"+rid2).checked;
		document.getElementById("methodname"+rid).innerHTML=document.getElementById("methodname"+rid2).innerHTML;
		document.getElementById("methodnamedom"+rid).innerHTML=document.getElementById("methodnamedom"+rid2).innerHTML;
		document.getElementById("altratetext_"+rid).value=document.getElementById("altratetext_"+rid2).value;
		document.getElementById("altratetext2_"+rid).value=document.getElementById("altratetext2_"+rid2).value;
		document.getElementById("altratetext3_"+rid).value=document.getElementById("altratetext3_"+rid2).value;
		
		document.getElementById("altrateids"+rid2).value=altrateids;
		document.getElementById("altrateuse"+rid2).checked=altrateuse;
		document.getElementById("methodname"+rid2).innerHTML=methodname;
		document.getElementById("methodnamedom"+rid2).innerHTML=methodnamedom;
		document.getElementById("altratetext_"+rid2).value=altratetext;
		document.getElementById("altratetext2_"+rid2).value=altratetext2;
		document.getElementById("altratetext3_"+rid2).value=altratetext3;
	}
	return false;
}
function swaptbrowsintl(rid){
	if(rid!=1){
		rid2=rid-1;
		var altrateidsintl=document.getElementById("altrateidsintl"+rid).value;
		var altrateuseintl=document.getElementById("altrateuseintl"+rid).checked;
		var methodnameintl=document.getElementById("methodnameintl"+rid).innerHTML;

		document.getElementById("altrateidsintl"+rid).value=document.getElementById("altrateidsintl"+rid2).value;
		document.getElementById("altrateuseintl"+rid).checked=document.getElementById("altrateuseintl"+rid2).checked;
		document.getElementById("methodnameintl"+rid).innerHTML=document.getElementById("methodnameintl"+rid2).innerHTML;

		document.getElementById("altrateidsintl"+rid2).value=altrateidsintl;
		document.getElementById("altrateuseintl"+rid2).checked=altrateuseintl;
		document.getElementById("methodnameintl"+rid2).innerHTML=methodnameintl;
	}
	return false;
}
//-->
</script>
<form method="post" action="adminmain.asp" onsubmit="return formvalidator(this)">
<input type="hidden" name="posted" value="1" />
<%	if not success then %>
		<p style="text-align:center"><br /><span style="color:#FF0000"><%=errmsg%></span></p>
<%	end if %>
	<h3 class="round_top half_top"><button class="mainbutton" id="mainbutton1" type="button" onclick="showtablerows(1,'<%=jsescape(yyStoSet)%>')"><%=yyStoSet%> <div style="float:right">&#9660;</div></button></h3>
	<table class="admin-table-b keeptable mainsettings" id="maintable1">
		<tr>
			<th colspan="2" scope="col"><%=yyCouSet & ", " & yyStoreURL & ", " & yyPPP & ", " & yyDefSor%></th>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%=yyCouSet%>: </strong></td>
			<td><select name="countrySetting" size="1"><%
				for index=0 to UBOUND(allcountries,2)
					print "<option value='"&allcountries(0,index)&"'"
					if rs("adminCountry")=allcountries(0,index) then print " selected=""selected"""
					print ">"&allcountries(1,index)&"</option>"&vbCrLf
				next %></select></td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><%=yyURLEx & " " & yyExample%>:<br /><%
				guessURL ="http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO")
				wherevs = InStr(guessURL,"vsadmin")
				if wherevs > 0 then
					guessURL = Left(guessURL,wherevs-1)
				else
					guessURL = "http://www.myurl.com/mystore/"
				end if
				print guessURL
			%></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%=yyStoreURL%>:</strong></td>
			<td><input type="text" name="url" size="45" value="<%=rs("adminStoreURL")%>" /></td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><div><%=yyURLSEx & " " & yyExample%>:</div><%
				print "<div>" & replace(guessURL,"http:","https:") & "</div>"
				if pathtossl<>"" then print "<div style=""color:#FF1010"">You have the parameter &quot;pathtossl&quot; set in your includes.asp file. This parameter is no longer used as the parameter is set here in admin. If you wish to override the parameter, use &quot;orstoreurlssl&quot;.</div>"
			%></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong>HTTPS <%=yyStoreURL%>:</strong></td>
			<td><input type="text" name="urlssl" size="45" value="<%=rs("adminStoreURLSSL")%>" /></td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><%=yyHMPPP%></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%=yyPPP%>:</strong></td>
			<td><input type="text" name="prodperpage" class="smallinput" size="10" value="<%=rs("adminProdsPerPage")%>" /></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%=yyDefSor%>:</strong></td>
			<td><select name="sortorder" size="1">
				<option value="0"><%=yySelect%></option>
				<option value="1"<%if rs("sortOrder")=1 then print " selected=""selected"""%>>&#x25B2; <%=yySortAl%></option>
				<option value="11"<%if rs("sortOrder")=11 then print " selected=""selected"""%>>&#x25BC; <%=yySortAl%></option>
				<option value="2"<%if rs("sortOrder")=2 then print " selected=""selected"""%>>&#x25B2; <%=yySortID%></option>
				<option value="12"<%if rs("sortOrder")=12 then print " selected=""selected"""%>>&#x25BC; <%=yySortID%></option>
				<option value="14"<%if rs("sortOrder")=14 then print " selected=""selected"""%>>&#x25B2; Sort By SKU</option>
				<option value="15"<%if rs("sortOrder")=15 then print " selected=""selected"""%>>&#x25BC; Sort By SKU</option>
				<option value="3"<%if rs("sortOrder")=3 then print " selected=""selected"""%>>&#x25B2; <%=yySortPA%></option>
				<option value="4"<%if rs("sortOrder")=4 then print " selected=""selected"""%>>&#x25BC; <%=yySortPA%></option>
				<option value="6"<%if rs("sortOrder")=6 then print " selected=""selected"""%>>&#x25B2; <%=yySortOA%></option>
				<option value="7"<%if rs("sortOrder")=7 then print " selected=""selected"""%>>&#x25BC; <%=yySortOA%></option>
				<option value="8"<%if rs("sortOrder")=8 then print " selected=""selected"""%>>&#x25B2; <%=yySortDA%></option>
				<option value="9"<%if rs("sortOrder")=9 then print " selected=""selected"""%>>&#x25BC; <%=yySortDA%></option>
				<option value="10"<%if rs("sortOrder")=10 then print " selected=""selected"""%>><%=yySortMa%></option>
				<option value="5"<%if rs("sortOrder")=5 then print " selected=""selected"""%>><%=yySortNS%></option>
				<option value="16"<%if rs("sortOrder")=16 then print " selected=""selected"""%>>Number of Ratings</option>
				<option value="17"<%if rs("sortOrder")=17 then print " selected=""selected"""%>>Average Rating</option>
				<option value="18"<%if rs("sortOrder")=18 then print " selected=""selected"""%>>Number of Sales</option>
				<option value="19"<%if rs("sortOrder")=19 then print " selected=""selected"""%>>Popularity</option>
			</select></td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><div class="mainshow"><button class="mainbuttonbottom" type="button" onclick="hidetablerows(1)"><%=yyCliHid%> &#9650;</button> <button class="mainbuttonbottom" type="submit"><%=IIfVr(ectdemostore,"Disabled",yySubmit)%></button></div></td>
		</tr>
	</table>

	<h3 class="round_top half_top"><button class="mainbutton" id="mainbutton2" type="button" onclick="showtablerows(2,'<%=jsescape(yyEmlSet)%>')"><%=yyEmlSet%> <div style="float:right">&#9660;</div></button></h3>
	<table class="admin-table-b keeptable mainsettings" id="maintable2">
		<tr>
			<th colspan="2"><%="Confirmation Emails" & ", " & yyEmail & ", " & "Email &quot;From&quot; Name" & ", " & "Email SMTP Service" & ", " & "SMTP Username / Password"%></th>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2">
				<div style="display:inline-block;width:20%"><input type="checkbox" name="emailconfirm" value="ON" <%if (rs("adminEmailConfirm") AND 1)=1 then print "checked=""checked"""%> /> :<strong><%="New Order"%></strong></div>
				<div style="display:inline-block;width:20%"><input type="checkbox" name="affilconfirm" value="ON" <%if (rs("adminEmailConfirm") AND 2)=2 then print "checked=""checked"""%> /> :<strong><%="New Affiliate"%></strong></div>
				<div style="display:inline-block;width:20%"><input type="checkbox" name="customerconfirm" value="ON" <%if (rs("adminEmailConfirm") AND 4)=4 then print "checked=""checked"""%> /> :<strong><%="New Customer"%></strong></div>
				<div style="display:inline-block;width:20%"><input type="checkbox" name="reviewconfirm" value="ON" <%if (rs("adminEmailConfirm") AND 8)=8 then print "checked=""checked"""%> /> :<strong><%="New Review"%></strong></div>
			</td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><%=yyCEAddr%></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%=yyEmail%>:</strong></td>
			<td><input type="text" name="xxemaddr" size="30" value="<%=rs("adminEmail")%>" /></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong>Email &quot;From&quot; Name:</strong></td>
			<td><input type="text" name="emailfromname" size="30" value="<%=rs("emailfromname")%>" placeholder="Eg: Your Store Name" /></td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><%=yyEmObjs%><br />
				  <span style="font-size:10px"><%=yyEmCDO%> 
				  <%=yyEmMoInf%> <a href="https://www.ecommercetemplates.com/help/ecommplus/parameters.asp#mail" target="_blank"><strong><%=yyHere%></strong></a>. <%=yyEmGen%> <a href="https://www.ecommercetemplates.com/help/email-help.asp" target="_blank"><strong><%=yyHere%></strong></a>.</span></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%=yyEmailObj%>:</strong></td>
			<td><select name="emailObject" size="1"><option value="99"><%=yyNone%></option><%
					gotobject=false
					function checkemail(objnum)
						if objnum=rs("emailObject") then
							checkemail = " selected=""selected"""
							gotobject=true
						else
							checkemail=""
						end if
					end function
					on error resume next
					err.number=0
					Set EmailObj = Server.CreateObject("CDONTS.NewMail")
					if err.number = 0 then print "<option value=""0"""&checkemail(0)&">CDONTS</option>"
					Set EmailObj = nothing
					err.number=0
					Set EmailObj = Server.CreateObject("CDO.Message")
					if err.number = 0 then print "<option value=""1"""&checkemail(1)&">CDO</option>"
					Set EmailObj = nothing
					err.number=0
					Set EmailObj = Server.CreateObject("Persits.MailSender")
					if err.number = 0 then print "<option value=""2"""&checkemail(2)&">ASP Email (PERSITS)</option>"
					Set EmailObj = nothing
					err.number=0
					Set EmailObj = Server.CreateObject("SMTPsvg.Mailer")
					if err.number = 0 then print "<option value=""3"""&checkemail(3)&">ASP Mail (ServerObjects)</option>"
					Set EmailObj = nothing
					err.number=0
					Set EmailObj = Server.CreateObject("JMail.SMTPMail")
					if err.number = 0 then print "<option value=""4"""&checkemail(4)&">JMail SMTPMail (Dimac)</option>"
					Set EmailObj = nothing
					err.number=0
					Set EmailObj = Server.CreateObject("SoftArtisans.SMTPMail")
					if err.number = 0 then print "<option value=""5"""&checkemail(5)&">SMTPMail (SoftArtisans)</option>"
					Set EmailObj = nothing
					err.number=0
					Set EmailObj = Server.CreateObject("JMail.Message")
					if err.number = 0 then print "<option value=""6"""&checkemail(6)&">JMail (Dimac)</option>"
					Set EmailObj = nothing
					on error goto 0
					%></select></td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><%=yySMTPEn%></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%=yySMTPSe%>:</strong></td>
			<td><input type="text" name="smtpserver" size="30" value="<%=rs("smtpserver")%>" placeholder="Eg: smtp.yourhost.com" /></td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><%=yySMTPSt%></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%=yyUname%>: </strong><br /><br /><strong><%=yyPass%>: </strong></td>
			<td><input type="text" name="emailuser" size="15" value="<%=rs("emailUser")%>" /><br /><br /><input type="password" name="emailpass" id="emailpass" size="15" value="" /> <select size="1" name="deleteemailpass" onchange="document.getElementById('emailpass').disabled=this.selectedIndex==1"><option value="">Set: YES</option><option value="delete"<% if trim(rs("emailPass")&"")="" then print " selected=""selected"""%>>Set: NO (Or Delete)</option></select><div style="font-size:11px">(Enter new value to change)</div></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong>SMTP Port: </strong></td>
			<td><input type="text" size="15" name="smtpport" value="<%=rs("smtpport")%>" placeholder="Eg: 587" /></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong>SMTP Secure: </strong></td>
			<td><select name="smtpsecure" size="1">
				<option value="">None</option>
				<option value="ssl"<% if rs("smtpsecure")="ssl" then print " selected=""selected"""%>>SSL</option>
				<option value="tls"<% if rs("smtpsecure")="tls" then print " selected=""selected"""%>>TLS (Persits Only)</option>
				</select></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong>Email Format: </strong></td>
			<td><select name="htmlemails" size="1">
				<option value="0">Text Format Emails</option>
				<option value="1"<% if rs("htmlemails")<>0 then print " selected=""selected"""%>>HTML Format Emails</option>
				</select></td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><div class="mainshow"><button class="mainbuttonbottom" type="button" onclick="hidetablerows(2)"><%=yyCliHid%> &#9650;</button> <button class="mainbuttonbottom" type="submit"><%=IIfVr(ectdemostore,"Disabled",yySubmit)%></button></div></td>
		</tr>
	</table>
	 
<%	DIM prodfilterarr(2),textarray(2),textarray2(2),textarray3(2)
	prodfilterarr(0)=rs("prodFilter")
	prodfilterarr(1)=rs("sideFilter")
	textarray(0)=split(rs("prodFilterText")&"","&")
	textarray2(0)=split(rs("prodFilterText2")&"","&")
	textarray3(0)=split(rs("prodFilterText3")&"","&")
	textarray(1)=split(rs("sideFilterText")&"","&")
	textarray2(1)=split(rs("sideFilterText2")&"","&")
	textarray3(1)=split(rs("sideFilterText3")&"","&")
	DIM filtertext(2,10),filtertext2(2,10),filtertext3(2,10)
	for numfilters=0 to 1
		if isarray(textarray(numfilters)) then
			for index=0 to 9
				if UBOUND(textarray(numfilters))>=index then filtertext(numfilters,index)=replace(textarray(numfilters)(index),"%26","&")
			next
		end if
		if isarray(textarray2(numfilters)) then
			for index=0 to 9
				if UBOUND(textarray2(numfilters))>=index then filtertext2(numfilters,index)=replace(textarray2(numfilters)(index),"%26","&")
			next
		end if
		if isarray(textarray3(numfilters)) then
			for index=0 to 9
				if UBOUND(textarray3(numfilters))>=index then filtertext3(numfilters,index)=replace(textarray3(numfilters)(index),"%26","&")
			next
		end if
	next
	sortoptions=rs("sortOptions")
%>
	<input type="hidden" name="prodfilterorder0" id="prodfilterorder0" value="<%=prodfilterorder%>" />
	<input type="hidden" name="prodfilterorder1" id="prodfilterorder1" value="<%=sidefilterorder%>" />
	<h3 class="round_top half_top"><button class="mainbutton" id="mainbutton3" type="button" onclick="showtablerows(3,'<%=jsescape(yyPrFiBr)%>')"><%=yyPrFiBr%> <div style="float:right">&#9660;</div></button></h3>
	<table class="admin-table-b keeptable mainsettings" id="maintable3">
		<tr>
			<th colspan="2"><%="Filter products by attributes, keyword, price, keywords"%></th>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2">
<%	for numfilters=0 to 1
			%><div class="mainfilterdiv">
				<table class="admin-table-b keeptable" id="filtertable<%=numfilters%>">
					<tr>
						<th colspan="3"><%=IIfVr(numfilters=0,yyFilSec,"Configuration for side filter bar (<a href=""https://www.ecommercetemplates.com/proddetail.asp?prod=ECT-Side-Filter-Bar"" style=""color:#FF9966;font-weight:bold"">IF INSTALLED</a>).")%></th>
					</tr>
<%		if numfilters=0 then filterorderarray=split(prodfilterorder,",") else filterorderarray=split(sidefilterorder,",")
		for indexfilterorder=0 to UBOUND(filterorderarray)
			select case filterorderarray(indexfilterorder)
			case 2 %>
					<tr id="filter<%=numfilters%>item2">
						<td width="19"><a href="#"><img src="adminimages/uparrow.png" alt="Move Up" onclick="return swapfilteritems(<%=numfilters%>,2)" /></a></td>
						<td><strong><%=yyFilScr%>: </strong></td>
						<td><input type="checkbox" name="filtercb<%=numfilters%>_1" value="ON" <% if (prodfilterarr(numfilters) AND 2)=2 then print "checked=""checked"" "%>/> <%=yyLabOpt%> <input type="text" name="filtertext<%=numfilters%>_1" class="smallinput" size="20" maxlength="50" value="<%=htmlspecials(filtertext(numfilters,1))%>" />
<%				if (adminlangsettings AND 262144)=262144 then
							if adminlanguages>=1 then	print "<input type=""text"" name=""filtertext"&numfilters&"_1x2"" class=""smallinput"" size=""20"" maxlength=""50"" value=""" & htmlspecials(filtertext2(numfilters,1)) & """ /> "
							if adminlanguages>=2 then	print "<input type=""text"" name=""filtertext"&numfilters&"_1x3"" class=""smallinput"" size=""20"" maxlength=""50"" value=""" & htmlspecials(filtertext3(numfilters,1)) & """ /> "
				end if %></td>
					</tr>
<%			case 4 %>
					<tr id="filter<%=numfilters%>item4">
						<td width="19"><a href="#"><img src="adminimages/uparrow.png" alt="Move Up" onclick="return swapfilteritems(<%=numfilters%>,4)" /></a></td>
						<td><strong><%=yyFilPri%>: </strong></td>
						<td><input type="checkbox" name="filtercb<%=numfilters%>_2" value="ON" <% if (prodfilterarr(numfilters) AND 4)=4 then print "checked=""checked"" "%>/> <%=yyLabOpt%> <input type="text" name="filtertext<%=numfilters%>_2" class="smallinput" size="20" maxlength="50" value="<%=htmlspecials(filtertext(numfilters,2))%>" />
<%				if (adminlangsettings AND 262144)=262144 then
					if adminlanguages>=1 then	print "<input type=""text"" name=""filtertext"&numfilters&"_2x2"" class=""smallinput"" size=""20"" maxlength=""50"" value=""" & htmlspecials(filtertext2(numfilters,2)) & """ /> "
					if adminlanguages>=2 then	print "<input type=""text"" name=""filtertext"&numfilters&"_2x3"" class=""smallinput"" size=""20"" maxlength=""50"" value=""" & htmlspecials(filtertext3(numfilters,2)) & """ /> "
				end if %></td>
					</tr>
<%			case 8 %>
					<tr id="filter<%=numfilters%>item8">
						<td width="19"><a href="#"><img src="adminimages/uparrow.png" alt="Move Up" onclick="return swapfilteritems(<%=numfilters%>,8)" /></a></td>
						<td><strong><%=yyCusSor%>: </strong></td>
						<td><input type="checkbox" name="filtercb<%=numfilters%>_3" value="ON" <% if (prodfilterarr(numfilters) AND 8)=8 then print "checked=""checked"" "%>/> <%=yyLabOpt%> <input type="text" name="filtertext<%=numfilters%>_3" class="smallinput" size="20" maxlength="50" value="<%=htmlspecials(filtertext(numfilters,3))%>" />
<%				if (adminlangsettings AND 262144)=262144 then
					if adminlanguages>=1 then	print "<input type=""text"" name=""filtertext"&numfilters&"_3x2"" class=""smallinput"" size=""20"" maxlength=""50"" value=""" & htmlspecials(filtertext2(numfilters,3)) & """ /> "
					if adminlanguages>=2 then	print "<input type=""text"" name=""filtertext"&numfilters&"_3x3"" class=""smallinput"" size=""20"" maxlength=""50"" value=""" & htmlspecials(filtertext3(numfilters,3)) & """ /> "
				end if %></td>
					</tr>
<%			case 16 %>
					<tr id="filter<%=numfilters%>item16">
						<td width="19"><a href="#"><img src="adminimages/uparrow.png" alt="Move Up" onclick="return swapfilteritems(<%=numfilters%>,16)" /></a></td>
						<td><strong><%=yyProPag%>: </strong></td>
						<td><input type="checkbox" name="filtercb<%=numfilters%>_4" value="ON" <% if (prodfilterarr(numfilters) AND 16)=16 then print "checked=""checked"" "%>/> <%=yyLabOpt%> <input type="text" name="filtertext<%=numfilters%>_4" class="smallinput" size="20" maxlength="50" value="<%=htmlspecials(filtertext(numfilters,4))%>" />
<%				if (adminlangsettings AND 262144)=262144 then
					if adminlanguages>=1 then	print "<input type=""text"" name=""filtertext"&numfilters&"_4x2"" class=""smallinput"" size=""20"" maxlength=""50"" value=""" & htmlspecials(filtertext2(numfilters,4)) & """ /> "
					if adminlanguages>=2 then	print "<input type=""text"" name=""filtertext"&numfilters&"_4x3"" class=""smallinput"" size=""20"" maxlength=""50"" value=""" & htmlspecials(filtertext3(numfilters,4)) & """ /> "
				end if %></td>
					</tr>
<%			case 32 %>
					<tr id="filter<%=numfilters%>item32">
						<td width="19"><a href="#"><img src="adminimages/uparrow.png" alt="Move Up" onclick="return swapfilteritems(<%=numfilters%>,32)" /></a></td>
						<td><strong><%=yyFilKey%>: </strong></td>
						<td><input type="checkbox" name="filtercb<%=numfilters%>_5" value="ON" <% if (prodfilterarr(numfilters) AND 32)=32 then print "checked=""checked"" "%>/> <%=yyLabOpt%> <input type="text" name="filtertext<%=numfilters%>_5" class="smallinput" size="20" maxlength="50" value="<%=htmlspecials(filtertext(numfilters,5))%>" />
<%				if (adminlangsettings AND 262144)=262144 then
					if adminlanguages>=1 then	print "<input type=""text"" name=""filtertext"&numfilters&"_5x2"" class=""smallinput"" size=""20"" maxlength=""50"" value=""" & htmlspecials(filtertext2(numfilters,5)) & """ /> "
					if adminlanguages>=2 then	print "<input type=""text"" name=""filtertext"&numfilters&"_5x3"" class=""smallinput"" size=""20"" maxlength=""50"" value=""" & htmlspecials(filtertext3(numfilters,5)) & """ /> "
				end if %></td>
					</tr>
<%			end select
		next %>
				</table>
			</div><%
	next %>
			</td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2" style="padding-left:20px">Options for Customer Defined Sort: (<a href="https://www.ecommercetemplates.com/help/ecommplus/parameters.asp#filterbar" target="_blank">The text for these can be changed using these parameters...</a>)</td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2">
			<table width="100%"><tr>
			<td style="font-size:10px;border:0;width:16.66%"><label><input type="checkbox" name="sortid1" value="ON" <% if (sortoptions AND (2 ^ 0))<>0 then print "checked=""checked"" "%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> &#x25B2; <%=yySortAl%></label></td>
			<td style="font-size:10px;border:0;width:16.66%"><label><input type="checkbox" name="sortid2" value="ON" <% if (sortoptions AND (2 ^ 1))<>0 then print "checked=""checked"" "%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> &#x25B2; <%=yySortID%></label></td>
			<td style="font-size:10px;border:0;width:16.66%"><label><input type="checkbox" name="sortid14" value="ON" <% if (sortoptions AND (2 ^ 13))<>0 then print "checked=""checked"" "%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> &#x25B2; Sort By SKU</label></td>
			<td style="font-size:10px;border:0;width:16.66%"><label><input type="checkbox" name="sortid3" value="ON" <% if (sortoptions AND (2 ^ 2))<>0 then print "checked=""checked"" "%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> &#x25B2; <%=yySortPA%></label></td>
			<td style="font-size:10px;border:0;width:16.66%"><label><input type="checkbox" name="sortid6" value="ON" <% if (sortoptions AND (2 ^ 5))<>0 then print "checked=""checked"" "%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> &#x25B2; <%=yySortOA%></label></td>
			<td style="font-size:10px;border:0;width:16.66%"><label><input type="checkbox" name="sortid8" value="ON" <% if (sortoptions AND (2 ^ 7))<>0 then print "checked=""checked"" "%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> &#x25B2; <%=yySortDA%></label></td>
			</tr><tr>
			<td style="font-size:10px;border:0"><label><input type="checkbox" name="sortid11" value="ON" <% if (sortoptions AND (2 ^ 10))<>0 then print "checked=""checked"" "%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> &#x25BC; <%=yySortAl%></label></td>
			<td style="font-size:10px;border:0"><label><input type="checkbox" name="sortid12" value="ON" <% if (sortoptions AND (2 ^ 11))<>0 then print "checked=""checked"" "%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> &#x25BC; <%=yySortID%></label></td>
			<td style="font-size:10px;border:0"><label><input type="checkbox" name="sortid15" value="ON" <% if (sortoptions AND (2 ^ 14))<>0 then print "checked=""checked"" "%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> &#x25BC; Sort By SKU</label></td>
			<td style="font-size:10px;border:0"><label><input type="checkbox" name="sortid4" value="ON" <% if (sortoptions AND (2 ^ 3))<>0 then print "checked=""checked"" "%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> &#x25BC; <%=yySortPA%></label></td>
			<td style="font-size:10px;border:0"><label><input type="checkbox" name="sortid7" value="ON" <% if (sortoptions AND (2 ^ 6))<>0 then print "checked=""checked"" "%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> &#x25BC; <%=yySortOA%></label></td>
			<td style="font-size:10px;border:0"><label><input type="checkbox" name="sortid9" value="ON" <% if (sortoptions AND (2 ^ 8))<>0 then print "checked=""checked"" "%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> &#x25BC; <%=yySortDA%></label></td>
			</tr><tr>
			<td style="font-size:10px;border:0"><label><input type="checkbox" name="sortid5" value="ON" <% if (sortoptions AND (2 ^ 4))<>0 then print "checked=""checked"" "%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> <%=yySortNS%></label></td>
			<td style="font-size:10px;border:0"><label><input type="checkbox" name="sortid10" value="ON" <% if (sortoptions AND (2 ^ 9))<>0 then print "checked=""checked"" "%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> <%=yySortMa%></label></td>
			<td style="font-size:10px;border:0"><label><input type="checkbox" name="sortid16" value="ON" <% if (sortoptions AND (2 ^ 15))<>0 then print "checked=""checked"" "%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> Number of Ratings</label></td>
			<td style="font-size:10px;border:0"><label><input type="checkbox" name="sortid17" value="ON" <% if (sortoptions AND (2 ^ 16))<>0 then print "checked=""checked"" "%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> Average Rating</label></td>
			<td style="font-size:10px;border:0"><label><input type="checkbox" name="sortid18" value="ON" <% if (sortoptions AND (2 ^ 17))<>0 then print "checked=""checked"" "%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> Number of Sales</label></td>
			<td style="font-size:10px;border:0"><label><input type="checkbox" name="sortid19" value="ON" <% if (sortoptions AND (2 ^ 18))<>0 then print "checked=""checked"" "%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> Popularity</label></td>
			</tr>
			</table>
			</td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><div class="mainshow"><button class="mainbuttonbottom" type="button" onclick="hidetablerows(3)"><%=yyCliHid%> &#9650;</button> <button class="mainbuttonbottom" type="submit"><%=IIfVr(ectdemostore,"Disabled",yySubmit)%></button></div></td>
		</tr>
	</table>
		
	<h3 class="round_top half_top"><button class="mainbutton" id="mainbutton4" type="button" onclick="showtablerows(4,'<%=jsescape(yyShHaSe)%>')"><%=yyShHaSe%> <div style="float:right">&#9660;</div></button></h3>
	<table class="admin-table-b keeptable mainsettings" id="maintable4">
		<tr><th colspan="2"><%=yyShpMet & ", " & "Alternate Rates" & ", " & "Product Packing" & ", " & "Origin Zip" & ", " & yyHanChg & ", " & "Shipping Insurance"%></th></tr>
		<tr class="maintablehidden">
			<td colspan="2"><%=yyWAltRa%></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%=yyUseAlt%>: </strong></td>
			<td><select name="adminAltRates" size="1" onchange="showhidealtrates(this)">
				<option value="0"><%=yyNoAltR%></option>
				<option value="1"<%if rs("adminAltRates")=1 then print " selected=""selected"""%>><%=yyAlRaMe%></option>
				<option value="2"<%if rs("adminAltRates")=2 then print " selected=""selected"""%>><%=yyAlRaTo%></option>
				</select>
			</td>
		</tr>
		<tr class="maintablehidden" id="domesticraterow0"<%if rs("adminAltRates")<>0 then print " style=""display:none"""%>>
			<td colspan="2"><%=yySelShD%></td>
		</tr>
		<tr class="maintablehidden" id="domesticraterow1"<%if rs("adminAltRates")<>0 then print " style=""display:none"""%>>
			<td><strong><%=yyShpTyp%>: </strong></td>
			<td><select name="shipping" size="1">
					<option value="0"><%=yyNoShp%></option>
					<option value="1" <%if int(rs("adminShipping"))=1 then print "selected=""selected"""%>><%=yyFlatShp%></option>
					<option value="2" <%if int(rs("adminShipping"))=2 then print "selected=""selected"""%>><%=yyWghtShp%></option>
					<option value="5" <%if int(rs("adminShipping"))=5 then print "selected=""selected"""%>><%=yyPriShp%></option>
					<option value="3" <%if int(rs("adminShipping"))=3 then print "selected=""selected"""%>><%=yyUSPS%></option>
					<option value="4" <%if int(rs("adminShipping"))=4 then print "selected=""selected"""%>><%=yyUPS%></option>
					<option value="6" <%if int(rs("adminShipping"))=6 then print "selected=""selected"""%>><%=yyCanPos%></option>
					<option value="7" <%if int(rs("adminShipping"))=7 then print "selected=""selected"""%>><%=yyFedex%></option>
					<option value="8" <%if int(rs("adminShipping"))=8 then print "selected=""selected"""%>>FedEx SmartPost</option>
					<option value="9" <%if int(rs("adminShipping"))=9 then print "selected=""selected"""%>><%=yyDHLShp%></option>
					<option value="10" <%if int(rs("adminShipping"))=10 then print "selected=""selected"""%>>Australia Post</option>
					</select></td>
		</tr>
		<tr class="maintablehidden" id="domesticraterow2"<%if rs("adminAltRates")<>0 then print " style=""display:none"""%>>
			<td colspan="2"><%=yySelShI%></td>
		</tr>
		<tr class="maintablehidden" id="domesticraterow3"<%if rs("adminAltRates")<>0 then print " style=""display:none"""%>>
			<td><strong><%=yyShpTyp%>: </strong></td>
			<td><select name="intshipping" size="1">
					<option value="0"><%=yySamDom%></option>
					<option value="1" <%if int(rs("adminIntShipping"))=1 then print "selected=""selected"""%>><%=yyFlatShp%></option>
					<option value="2" <%if int(rs("adminIntShipping"))=2 then print "selected=""selected"""%>><%=yyWghtShp%></option>
					<option value="5" <%if int(rs("adminIntShipping"))=5 then print "selected=""selected"""%>><%=yyPriShp%></option>
					<option value="3" <%if int(rs("adminIntShipping"))=3 then print "selected=""selected"""%>><%=yyUSPS%></option>
					<option value="4" <%if int(rs("adminIntShipping"))=4 then print "selected=""selected"""%>><%=yyUPS%></option>
					<option value="6" <%if int(rs("adminIntShipping"))=6 then print "selected=""selected"""%>><%=yyCanPos%></option>
					<option value="7" <%if int(rs("adminIntShipping"))=7 then print "selected=""selected"""%>><%=yyFedex%></option>
					<option value="9" <%if int(rs("adminIntShipping"))=9 then print "selected=""selected"""%>><%=yyDHLShp%></option>
					<option value="10" <%if int(rs("adminIntShipping"))=10 then print "selected=""selected"""%>>Australia Post</option>
					</select></td>
		</tr>
		<tr class="maintablehidden" id="altraterowtitle"<%if rs("adminAltRates")=0 then print " style=""display:none"""%>>
			<td colspan="2"><%=yyAltSel%></td>
		</tr>
		<tr class="maintablehidden" id="altraterow"<%if rs("adminAltRates")=0 then print " style=""display:none"""%>>
			<td colspan="2" align="center">
				<input type="hidden" name="altrateids" id="altrateids" value="" />
				<input type="hidden" name="altrateidsintl" id="altrateidsintl" value="" />
				<input type="hidden" name="altrateuse" id="altrateuse" value="" />
				<input type="hidden" name="altrateuseintl" id="altrateuseintl" value="" />
				<input type="hidden" name="altratetext" id="altratetext" value="" />
				<input type="hidden" name="altratetext2" id="altratetext2" value="" />
				<input type="hidden" name="altratetext3" id="altratetext3" value="" />
				<table>
				<tr>
				<td valign="top" width="40%">
					<div style="font-weight:bold;text-align:center">Shipping Method Label</div>
					<table>
<%				index=1
				sSQL = "SELECT altrateid,altratename,altratetext,altratetext2,altratetext3,usealtmethod,usealtmethodintl FROM alternaterates ORDER BY ABS(usealtmethod "&IIfVr(sqlserver,"|","OR")&" usealtmethodintl) DESC,altrateorder,altrateid"
				rs2.open sSQL,cnn,0,1
				do while NOT rs2.EOF %>
				<tr height="40">
					<td><span id="methodname<%=index%>"><%=rs2("altratename") %></span></td>
					<td><input type="text" id="altratetext_<%=index%>" size="30" value="<%=htmlspecials(rs2("altratetext")) %>" /><br />
<%						for index2=2 to 3
							if index2<=(adminlanguages+1) AND (adminlangsettings AND 65536)=65536 then %>
					<input type="text" id="altratetext<%=index2%>_<%=index%>" size="30" value="<%=htmlspecials(rs2("altratetext"&index2)) %>" /><br />
<%							else %>
					<input type="hidden" id="altratetext<%=index2%>_<%=index%>" value="<%=htmlspecials(rs2("altratetext"&index2)) %>" />
<%							end if
						next %></td>
				</tr>
<%					index=index+1
					rs2.movenext
				loop
				rs2.close %>
					</table>
				</td>
				<td valign="top" width="30%">
					<div style="font-weight:bold;text-align:center">Domestic Shipping Order</div>
					<table>
<%				index=1
				sSQL = "SELECT altrateid,altratename,usealtmethod FROM alternaterates ORDER BY ABS(usealtmethod "&IIfVr(sqlserver,"|","OR")&" usealtmethodintl) DESC,altrateorder,altrateid"
				rs2.open sSQL,cnn,0,1
				do while NOT rs2.EOF %>
				  <tr height="40">
					<td>
					<input type="hidden" id="altrateids<%=index%>" value="<%=rs2("altrateid")%>" />
					<%	if index=0 then
							print "&nbsp;"
						else %>
					<a href="#"><img src="adminimages/uparrow.png" alt="Move Up" onclick="return swaptbrows(<%=index%>)" /></a>
<%						end if %></td>
					<td><input type="checkbox" id="altrateuse<%=index%>" value="ON" <%=IIfVr(rs2("usealtmethod"),"checked=""checked"" ","")%>/></td>
					<td><span id="methodnamedom<%=index%>"><%=rs2("altratename") %></span></td>
				  </tr>
<%					index=index+1
					rs2.movenext
				loop
				rs2.close %>
					</table>
				</td>
				<td valign="top" width="30%">
					<div style="font-weight:bold;text-align:center">International Shipping Order</div>
					<table>
<%				index=1
				sSQL = "SELECT altrateid,altratename,usealtmethodintl FROM alternaterates ORDER BY ABS(usealtmethodintl "&IIfVr(sqlserver,"|","OR")&" usealtmethod) DESC,altrateorderintl,altrateid"
				rs2.open sSQL,cnn,0,1
				do while NOT rs2.EOF %>
				  <tr height="40">
					<td>
					<input type="hidden" id="altrateidsintl<%=index%>" value="<%=rs2("altrateid")%>" />
					<%	if index=0 then
							print "&nbsp;"
						else %>
					<a href="#"><img src="adminimages/uparrow.png" alt="Move Up" onclick="return swaptbrowsintl(<%=index%>)" /></a>
<%						end if %></td>
					<td><input type="checkbox" id="altrateuseintl<%=index%>" value="ON" <%=IIfVr(rs2("usealtmethodintl"),"checked=""checked"" ","")%>/></td>
					<td><span id="methodnameintl<%=index%>"><%=rs2("altratename") %></span></td>
				  </tr>
<%					index=index+1
					rs2.movenext
				loop
				rs2.close %>
					</table>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><%=yyHowPck%><br /><span style="font-size:10px"><%=yyOnlyAf%></span></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%=yyPackPr%>: </strong></td>
			<td><select name="packing" size="1">
					<option value="0"><%=yyPckSep%></option>
					<option value="1" <%if int(adminPacking)=1 then print "selected=""selected"""%>><%=yyPckTog%></option>
					</select></td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><%=yyEntZip%></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%=yyZip%>: </strong></td>
			<td><input type="text" name="zipcode" class="smallinput" size="10" value="<%=rs("adminZipCode")%>" /></td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><%=yyUPSUnt%></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%=yyShpUnt%>: </strong><br /><br /><strong><%=yyDims%>: </strong></td>
			<td><select name="adminUnits" size="1">
					<option value="1" <%if (int(rs("adminUnits")) AND 3)=1 then print "selected=""selected"""%>>LBS</option>
					<option value="0" <%if (int(rs("adminUnits")) AND 3)=0 then print "selected=""selected"""%>>KGS</option>
				</select><br /><br />
				<select name="adminDims" size="1">
					<option value="0"><%=yyNotSpe%></option>
					<option value="4" <%if (int(rs("adminUnits")) AND 12)=4 then print "selected=""selected"""%>>IN</option>
					<option value="8" <%if (int(rs("adminUnits")) AND 12)=8 then print "selected=""selected"""%>>CM</option>
				</select></td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><ul>
				  <li><span style="font-size:10px"><%=redasterix&yyUntNote%></span></li>
				</ul></td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><%=yyHandEx%></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%=yyHanChg%>: </strong><br /><br /><strong><%=yyHanChg & " (" & yyPercen & ")"%>: </strong></td>
			<td><input type="text" name="handling" class="smallinput" size="10" value="<%=rs("adminHandling")%>" /><br /><br /><input type="text" name="handlingpercent" class="smallinput" size="10" style="text-align:right" value="<%=rs("adminHandlingPercent")%>" />%</td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2">If you wish to add shipping insurance, please set it below. Please note that the online carriers will calculate insurance automatically. Rates set here will be used in the case of Free Shipping however, and for Weight / Price based shipping.</td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%="Domestic Shipments Insurance"%>: </strong></td>
			<td><select name="shipinsurancedom" size="1">
					<option value="0">Don't add insurance</option>
					<option value="1"<% if rs("shipInsuranceDom")=1 then print " selected=""selected"""%>>Automatically add insurance</option>
					<option value="2"<% if rs("shipInsuranceDom")=2 then print " selected=""selected"""%>>Ask the customer if they would like insurance</option>
					<option value="3"<% if rs("shipInsuranceDom")=3 then print " selected=""selected"""%>>Force the customer to choose if they want insurance or not</option>
					</select>
					</td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%="International Shipments Insurance"%>: </strong></td>
			<td><select name="shipinsuranceint" size="1">
					<option value="0">Don't add insurance</option>
					<option value="1"<% if rs("shipInsuranceInt")=1 then print " selected=""selected"""%>>Automatically add insurance</option>
					<option value="2"<% if rs("shipInsuranceInt")=2 then print " selected=""selected"""%>>Ask the customer if they would like insurance</option>
					<option value="3"<% if rs("shipInsuranceInt")=3 then print " selected=""selected"""%>>Force the customer to choose if they want insurance or not</option>
					</select>
					</td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2">If not using an online shipping rates service (or in some Free Shipping scenarios) please enter the default for insurance</td>
		</tr>
		<tr class="maintablehidden">
			<td>
				<strong><%="Domestic Insurance Cost"%>: </strong>
				<div style="padding-top:10px"><input type="checkbox" value="1" name="nocarrierdomins"<%=IIfVs(rs("noCarrierDomIns")<>0," checked=""checked""")%> /> Also use for online carriers.</div>
			</td>
			<td>
				<div style="display:table">
					<div class="ecttablerow"><div style="text-align:right">Percentage:</div><div><input type="text" name="insurancedompercent" class="smallinput" size="10" value="<%=rs("insuranceDomPercent")%>" /></div></div>
					<div class="ecttablerow"><div style="text-align:right">Minimum:</div><div><input type="text" name="insurancedommin" class="smallinput" size="10" value="<%=rs("insuranceDomMin")%>" /></div></div>
				</div>
			</td>
		</tr>
		<tr class="maintablehidden">
			<td>
				<strong><%="International Insurance Cost"%>: </strong>
				<div style="padding-top:10px"><input type="checkbox" value="1" name="nocarrierintins"<%=IIfVs(rs("noCarrierIntIns")<>0," checked=""checked""")%> /> Also use for online carriers.</div>
			</td>
			<td>
				<div style="display:table">
					<div class="ecttablerow"><div style="text-align:right">Percentage:</div><div><input type="text" name="insuranceintpercent" class="smallinput" size="10" value="<%=rs("insuranceIntPercent")%>" /></div></div>
					<div class="ecttablerow"><div style="text-align:right">Minimum:</div><div><input type="text" name="insuranceintmin" class="smallinput" size="10" value="<%=rs("insuranceIntMin")%>" /></div></div>
				</div>
			</td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><div class="mainshow"><button class="mainbuttonbottom" type="button" onclick="hidetablerows(4)"><%=yyCliHid%> &#9650;</button> <button class="mainbuttonbottom" type="submit"><%=IIfVr(ectdemostore,"Disabled",yySubmit)%></button></div></td>
		</tr>
	</table>

	<h3 class="round_top half_top"><button class="mainbutton" id="mainbutton5" type="button" onclick="showtablerows(5,'<%=jsescape(yyOrdMan)%>')"><%=yyOrdMan%> <div style="float:right">&#9660;</div></button></h3>
	<table class="admin-table-b keeptable mainsettings" id="maintable5">
		<tr>
			<th colspan="2"><%=yyStock & ", " & "Incomplete Order Deletion" & ", " & "Vacation Message"%></th>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%=yyStock%>: </strong></td>
			<td><select name="stockManage" size="1">
					<option value="0"><%=yyNoStk%></option>
					<option value="1" <% if Int(rs("adminStockManage"))<>0 then print "selected=""selected"""%>> &nbsp;&nbsp; <%=yyOn%></option>
			</select></td>
		</tr>			  
		<tr class="maintablehidden">
			<td colspan="2"><%=yyDelUnc%></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%=yyDelAft%>:</strong></td>
			<td><select name="deleteUncompleted" size="1">
					<option value="0"><%=yyNever%></option>
					<option value="1" <% if int(rs("adminDelUncompleted"))=1 then print "selected=""selected"""%>>1 <%=yyDay%></option>
					<option value="2" <% if int(rs("adminDelUncompleted"))=2 then print "selected=""selected"""%>>2 <%=yyDays%></option>
					<option value="3" <% if int(rs("adminDelUncompleted"))=3 then print "selected=""selected"""%>>3 <%=yyDays%></option>
					<option value="4" <% if int(rs("adminDelUncompleted"))=4 then print "selected=""selected"""%>>4 <%=yyDays%></option>
					<option value="7" <% if int(rs("adminDelUncompleted"))=7 then print "selected=""selected"""%>>1 <%=yyWeek%></option>
					<option value="14" <% if int(rs("adminDelUncompleted"))=14 then print "selected=""selected"""%>>2 <%=yyWeeks%></option>
					<option value="28" <% if int(rs("adminDelUncompleted"))=28 then print "selected=""selected"""%>>4 <%=yyWeeks%></option>
					<option value="70" <% if int(rs("adminDelUncompleted"))=70 then print "selected=""selected"""%>>10 <%=yyWeeks%></option>
					<option value="140" <% if int(rs("adminDelUncompleted"))=140 then print "selected=""selected"""%>>20 <%=yyWeeks%></option>
					<option value="210" <% if int(rs("adminDelUncompleted"))=210 then print "selected=""selected"""%>>30 <%=yyWeeks%></option>
					<option value="364" <% if int(rs("adminDelUncompleted"))=364 then print "selected=""selected"""%>>52 <%=yyWeeks%></option>
					<option value="525" <% if int(rs("adminDelUncompleted"))=525 then print "selected=""selected"""%>>75 <%=yyWeeks%></option>
					<option value="728" <% if int(rs("adminDelUncompleted"))=728 then print "selected=""selected"""%>>104 <%=yyWeeks%></option>
					</select><%
			if NOT enableclientlogin then call writehiddenvar("adminClearCart",rs("adminClearCart")) %></td>
		</tr>
<%			if enableclientlogin then %>
		<tr class="maintablehidden">
			<td colspan="2"><%=yyRemLII%></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%=yyDelAft%>:</strong></td>
			<td><select name="adminClearCart" size="1">
					<option value="0"><%=yyNever%></option>
					<option value="14" <% if int(rs("adminClearCart"))=14 then print "selected=""selected"""%>>2 <%=yyWeeks%></option>
					<option value="28" <% if int(rs("adminClearCart"))=28 then print "selected=""selected"""%>>4 <%=yyWeeks%></option>
					<option value="70" <% if int(rs("adminClearCart"))=70 then print "selected=""selected"""%>>10 <%=yyWeeks%></option>
					<option value="140" <% if int(rs("adminClearCart"))=140 then print "selected=""selected"""%>>20 <%=yyWeeks%></option>
					<option value="210" <% if int(rs("adminClearCart"))=210 then print "selected=""selected"""%>>30 <%=yyWeeks%></option>
					<option value="364" <% if int(rs("adminClearCart"))=364 then print "selected=""selected"""%>>52 <%=yyWeeks%></option>
					<option value="525" <% if int(rs("adminClearCart"))=525 then print "selected=""selected"""%>>75 <%=yyWeeks%></option>
					<option value="728" <% if int(rs("adminClearCart"))=728 then print "selected=""selected"""%>>104 <%=yyWeeks%></option>
					</select></td>
		</tr>
<%			end if %>
		<tr class="maintablehidden">
			<td colspan="2">Temporarily close your store for a vacation, for instance</td>
		</tr>
		<tr class="maintablehidden">
			<td><strong>On vacation:</strong></td>
			<td><select name="onvacation" size="1" onchange="document.getElementById('vacationmessage').style.display=this.selectedIndex==0?'none':''"><option value="0">No</option><option value="1"<% if rs("onvacation")<>0 then print " selected=""selected"" "%>>Yes, close my store</option></select></td>
		</tr>
		<tr id="vacationmessage"<% if rs("onvacation")=0 then print " style=""display:none"""%>>
			<td><strong><%="Vacation Message"%>: </strong></td>
			<td><textarea name="vacationmessage" cols="58" rows="5"><%=htmlspecials(vacationmessage)%></textarea></td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><div class="mainshow"><button class="mainbuttonbottom" type="button" onclick="hidetablerows(5)"><%=yyCliHid%> &#9650;</button> <button class="mainbuttonbottom" type="submit"><%=IIfVr(ectdemostore,"Disabled",yySubmit)%></button></div></td>
		</tr>
	</table>

	<h3 class="round_top half_top"><button class="mainbutton" id="mainbutton6" type="button" onclick="showtablerows(6,'<%=jsescape(yyLaSet)%>')"><%=yyLaSet%> <div style="float:right">&#9660;</div></button></h3>
	<table class="admin-table-b keeptable mainsettings" id="maintable6">
		<tr><th colspan="2">Admin Language and Store Languages</th></tr>
		<tr class="maintablehidden">
			<td><strong>Admin Language: </strong></td>
			<td><select name="adminlang" size="1">
					<option value="">English</option>
					<option value="fr" <% if rs("adminlang")="fr" then print "selected=""selected"""%>>French / Fran&ccedil;ais</option>
					<option value="de" <% if rs("adminlang")="de" then print "selected=""selected"""%>>German / Deutsch</option>
					<option value="it" <% if rs("adminlang")="it" then print "selected=""selected"""%>>Italian / Italiano</option>
					<option value="nl" <% if rs("adminlang")="nl" then print "selected=""selected"""%>>Nederlands / Dutch</option>
					<option value="es" <% if rs("adminlang")="es" then print "selected=""selected"""%>>Spanish / Espa&ntilde;ol</option>
					</select></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong>Store Language: </strong></td>
			<td><%	storelang1="" : storelang2="" : storelang3=""
				if trim(rs("storelang")&"")="" then
					storelang1="en"
				else
					storelangarr=split(rs("storelang"),"|")
					storelang1=storelangarr(0)
					if UBOUND(storelangarr)>0 then storelang2=storelangarr(1)
					if UBOUND(storelangarr)>1 then storelang3=storelangarr(2)
				end if %><div><select name="storelang1" size="1"><option value="">English</option>
				<option value="dk" <% if storelang1="dk" then print "selected=""selected"""%>>Danish / Dansk</option>
				<option value="fr" <% if storelang1="fr" then print "selected=""selected"""%>>French / Fran&ccedil;ais</option>
				<option value="de" <% if storelang1="de" then print "selected=""selected"""%>>German / Deutsch</option>
				<option value="it" <% if storelang1="it" then print "selected=""selected"""%>>Italian / Italiano</option>
				<option value="nl" <% if storelang1="nl" then print "selected=""selected"""%>>Nederlands / Dutch</option>
				<option value="pt" <% if storelang1="pt" then print "selected=""selected"""%>>Portugese / Portugu&ecirc;s</option>
				<option value="es" <% if storelang1="es" then print "selected=""selected"""%>>Spanish / Espa&ntilde;ol</option>
				</select></div>
				<div id="storelang2"<% if int(rs("adminlanguages"))<1 then print " style=""display:none""" %>><select name="storelang2" size="1">
				<option value="">English</option>
				<option value="dk" <% if storelang2="dk" then print "selected=""selected"""%>>Danish / Dansk</option>
				<option value="fr" <% if storelang2="fr" then print "selected=""selected"""%>>French / Fran&ccedil;ais</option>
				<option value="de" <% if storelang2="de" then print "selected=""selected"""%>>German / Deutsch</option>
				<option value="it" <% if storelang2="it" then print "selected=""selected"""%>>Italian / Italiano</option>
				<option value="nl" <% if storelang2="nl" then print "selected=""selected"""%>>Nederlands / Dutch</option>
				<option value="pt" <% if storelang2="pt" then print "selected=""selected"""%>>Portugese / Portugu&ecirc;s</option>
				<option value="es" <% if storelang2="es" then print "selected=""selected"""%>>Spanish / Espa&ntilde;ol</option>
				</select></div>
				<div id="storelang3"<% if int(rs("adminlanguages"))<2 then print " style=""display:none""" %>><select name="storelang3" size="1">
				<option value="">English</option>
				<option value="dk" <% if storelang3="dk" then print "selected=""selected"""%>>Danish / Dansk</option>
				<option value="fr" <% if storelang3="fr" then print "selected=""selected"""%>>French / Fran&ccedil;ais</option>
				<option value="de" <% if storelang3="de" then print "selected=""selected"""%>>German / Deutsch</option>
				<option value="it" <% if storelang3="it" then print "selected=""selected"""%>>Italian / Italiano</option>
				<option value="nl" <% if storelang3="nl" then print "selected=""selected"""%>>Nederlands / Dutch</option>
				<option value="pt" <% if storelang3="pt" then print "selected=""selected"""%>>Portugese / Portugu&ecirc;s</option>
				<option value="es" <% if storelang3="es" then print "selected=""selected"""%>>Spanish / Espa&ntilde;ol</option>
				</select></div></td>
		</tr>
		<tr class="maintablehidden"><td colspan="2"><%=yyHowLan%></td></tr>
		<tr class="maintablehidden">
			<td><strong><%=yyNumLan%>: </strong></td>
			<td><select name="adminlanguages" size="1" onchange="document.getElementById('storelang3').style.display=this.selectedIndex<2?'none':'';document.getElementById('storelang2').style.display=this.selectedIndex<1?'none':'';">
				<option value="0">1</option>
				<option value="1" <% if Int(rs("adminlanguages"))=1 then print "selected=""selected"""%>>2</option>
				<option value="2" <% if Int(rs("adminlanguages"))=2 then print "selected=""selected"""%>>3</option>
				</select></td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><%=yyWhMull%><br />
			<span style="font-size:10px"><%=yyLonrel%></span></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%=yyLaSet%>: </strong></td>
			<td><select name="adminlangsettings" size="5" multiple="multiple">
				<option value="1" <% if (Int(rs("adminlangsettings")) AND 1)=1 then print "selected=""selected"""%>><%=yyPrName%></option>
				<option value="2" <% if (Int(rs("adminlangsettings")) AND 2)=2 then print "selected=""selected"""%>><%=yyDesc%></option>
				<option value="4" <% if (Int(rs("adminlangsettings")) AND 4)=4 then print "selected=""selected"""%>><%=yyLnDesc%></option>
				<option value="1048576" <% if (Int(rs("adminlangsettings")) AND 1048576)=1048576 then print "selected=""selected"""%>><%=yyStaNam%></option>
				<option value="8" <% if (Int(rs("adminlangsettings")) AND 8)=8 then print "selected=""selected"""%>><%=yyCntNam%></option>
				<option value="16" <% if (Int(rs("adminlangsettings")) AND 16)=16 then print "selected=""selected"""%>><%=yyPOName%></option>
				<option value="32" <% if (Int(rs("adminlangsettings")) AND 32)=32 then print "selected=""selected"""%>><%=yyPOChoi%></option>
				<option value="64" <% if (Int(rs("adminlangsettings")) AND 64)=64 then print "selected=""selected"""%>><%=yyOrdSta%></option>
				<option value="128" <% if (Int(rs("adminlangsettings")) AND 128)=128 then print "selected=""selected"""%>><%=yyPayMet%></option>
				<option value="256" <% if (Int(rs("adminlangsettings")) AND 256)=256 then print "selected=""selected"""%>><%=yyCatNam%></option>
				<option value="512" <% if (Int(rs("adminlangsettings")) AND 512)=512 then print "selected=""selected"""%>><%=yyCatDes%></option>
				<option value="524288" <% if (Int(rs("adminlangsettings")) AND 524288)=524288 then print "selected=""selected"""%>>Category Header</option>
				<option value="1024" <% if (Int(rs("adminlangsettings")) AND 1024)=1024 then print "selected=""selected"""%>><%=yyDisTxt%></option>
				<option value="2048" <% if (Int(rs("adminlangsettings")) AND 2048)=2048 then print "selected=""selected"""%>><%=yyCatURL%></option>
				<option value="4096" <% if (Int(rs("adminlangsettings")) AND 4096)=4096 then print "selected=""selected"""%>><%=yyEmlHdr%></option>
				<option value="8192" <% if (Int(rs("adminlangsettings")) AND 8192)=8192 then print "selected=""selected"""%>><%=yyManURL%></option>
				<option value="16384" <% if (Int(rs("adminlangsettings")) AND 16384)=16384 then print "selected=""selected"""%>><%=yyManDsc%></option>
				<option value="32768" <% if (Int(rs("adminlangsettings")) AND 32768)=32768 then print "selected=""selected"""%>><%=yyContReg%></option>
				<option value="65536" <% if (Int(rs("adminlangsettings")) AND 65536)=65536 then print "selected=""selected"""%>><%=yyAltShM%></option>
				<option value="131072" <% if (Int(rs("adminlangsettings")) AND 131072)=131072 then print "selected=""selected"""%>><%=yySeaCri%></option>
				<option value="262144" <% if (Int(rs("adminlangsettings")) AND 262144)=262144 then print "selected=""selected"""%>>Filter Bar</option>
				<option value="2097152" <% if (Int(rs("adminlangsettings")) AND 2097152)=2097152 then print "selected=""selected"""%>>Page Title / Meta Description</option>
				<option value="4194304" <% if (Int(rs("adminlangsettings")) AND 4194304)=4194304 then print "selected=""selected"""%>><%=yyAddSrP%></option>
			</select></td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><div class="mainshow"><button class="mainbuttonbottom" type="button" onclick="hidetablerows(6)"><%=yyCliHid%> &#9650;</button> <button class="mainbuttonbottom" type="submit"><%=IIfVr(ectdemostore,"Disabled",yySubmit)%></button></div></td>
		</tr>
	</table>

	<h3 class="round_top half_top"><button class="mainbutton" id="mainbutton7" type="button" onclick="showtablerows(7,'<%=jsescape(yyCurenc)%>')"><%=yyCurenc%> <div style="float:right">&#9660;</div></button></h3>
	<table class="admin-table-b keeptable mainsettings" id="maintable7">
		<tr><th colspan="2"><%=yy3CurCon%><br /><span style="font-size:10px"><%=yyNo3Con%></span></th></tr>
		<tr class="maintablehidden">
			<td><strong><%=yyConv%> 1: </strong></td>
			<td>&nbsp;<%=yyRate%> <input type="text" name="currRate1" class="smallinput" size="10" value="<% if rs("currRate1")<>0 then print rs("currRate1")%>" />&nbsp;&nbsp;&nbsp;Symbol <select name="currSymbol1" size="1"><option value="">None</option>
<%			for index=0 to UBOUND(allcurrencies,2)
				print "<option value='"&allcurrencies(0,index)&"'"
				if rs("currSymbol1")=allcurrencies(0,index) then print " selected=""selected"""
				print ">"&allcurrencies(0,index)&"</option>"&vbCrLf
			next
%></select></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%=yyConv%> 2: </strong></td>
			<td>&nbsp;<%=yyRate%> <input type="text" name="currRate2" class="smallinput" size="10" value="<% if rs("currRate2")<>0 then print rs("currRate2")%>" />&nbsp;&nbsp;&nbsp;Symbol <select name="currSymbol2" size="1"><option value="">None</option>
<%			for index=0 to UBOUND(allcurrencies,2)
				print "<option value='"&allcurrencies(0,index)&"'"
				if rs("currSymbol2")=allcurrencies(0,index) then print " selected=""selected"""
				print ">"&allcurrencies(0,index)&"</option>"&vbCrLf
			next
%></select></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%=yyConv%> 3: </strong></td>
			<td>&nbsp;<%=yyRate%> <input type="text" name="currRate3" class="smallinput" size="10" value="<% if rs("currRate3")<>0 then print rs("currRate3")%>" />&nbsp;&nbsp;&nbsp;Symbol <select name="currSymbol3" size="1"><option value="">None</option>
<%			for index=0 to UBOUND(allcurrencies,2)
				print "<option value='"&allcurrencies(0,index)&"'"
				if rs("currSymbol3")=allcurrencies(0,index) then print " selected=""selected"""
				print ">"&allcurrencies(0,index)&"</option>"&vbCrLf
			next
%></select></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong>Update rates automatically: </strong></td>
			<td>&nbsp;<select name="currConvUser" size="1"><option value="">Do not update automatically</option>
							<option value="y"<%=IIfVs(rs("currConvUser")<>""," selected=""selected""") %>>Yes, update rates automatically</option></select></td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><div class="mainshow"><button class="mainbuttonbottom" type="button" onclick="hidetablerows(7)"><%=yyCliHid%> &#9650;</button> <button class="mainbuttonbottom" type="submit"><%=IIfVr(ectdemostore,"Disabled",yySubmit)%></button></div></td>
		</tr>
	</table>

	<h3 class="round_top half_top"><button class="mainbutton" id="mainbutton9" type="button" onclick="showtablerows(9,'<%=jsescape("Google reCaptcha")%>')"><%="Google reCaptcha"%> <div style="float:right">&#9660;</div></button></h3>
	<table class="admin-table-b keeptable mainsettings" id="maintable9">
		<tr><th colspan="2">Automatically integrate Google's reCAPTCHA security system on your website.</th></tr>
		<tr class="maintablehidden">
			<td><strong><%="reCaptcha Site key"%>: </strong></td>
			<td><input type="text" name="reCAPTCHAsitekey" size="30" value="<%=htmlspecials(rs("reCAPTCHAsitekey"))%>" /></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%="reCaptcha Secret key"%>: </strong></td>
			<td><input type="text" name="reCAPTCHAsecret" size="30" value="<%=htmlspecials(rs("reCAPTCHAsecret"))%>" /></td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2" style="padding-left:40px"><%="Use recaptcha on"%>:</td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2" style="padding-left:40px">
				<table><tr>
				<td width="125" style="font-size:10px;border:0"><label><input type="checkbox" name="recaptcha1" value="ON" <%=IIfVs((rs("reCAPTCHAuseon") AND 1)=1,"checked=""checked"" ")%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> <%="Card Entry"%></label></td>
				<td width="125" style="font-size:10px;border:0"><label><input type="checkbox" name="recaptcha2" value="ON" <%=IIfVs((rs("reCAPTCHAuseon") AND 2)=2,"checked=""checked"" ")%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> <%="Ask A Question"%></label></td>
				<td width="125" style="font-size:10px;border:0"><label><input type="checkbox" name="recaptcha3" value="ON" <%=IIfVs((rs("reCAPTCHAuseon") AND 4)=4,"checked=""checked"" ")%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> <%="New Affiliate"%></label></td>
				<td width="125" style="font-size:10px;border:0"><label><input type="checkbox" name="recaptcha4" value="ON" <%=IIfVs((rs("reCAPTCHAuseon") AND 8)=8,"checked=""checked"" ")%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> <%="New Account"%></label></td>
				</tr><tr>
				<td width="125" style="font-size:10px;border:0"><label><input type="checkbox" name="recaptcha5" value="ON" <%=IIfVs((rs("reCAPTCHAuseon") AND 16)=16,"checked=""checked"" ")%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> <%="Admin Login"%></label></td>
				<td width="125" style="font-size:10px;border:0"><label><input type="checkbox" name="recaptcha6" value="ON" <%=IIfVs((rs("reCAPTCHAuseon") AND 32)=32,"checked=""checked"" ")%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> <%="Post Review"%></label></td>
				<td width="125" style="font-size:10px;border:0"><label><input type="checkbox" name="recaptcha7" value="ON" <%=IIfVs((rs("reCAPTCHAuseon") AND 64)=64,"checked=""checked"" ")%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> <%="Gift Certificate"%></label></td>
				<td width="125" style="font-size:10px;border:0"><label><input type="checkbox" name="recaptcha9" value="ON" <%=IIfVs((rs("reCAPTCHAuseon") AND 256)=256,"checked=""checked"" ")%>style="padding:0;margin:0;vertical-align:bottom;top:-1px;" /> <%="Email Orders"%></label></td>
				</tr></table>
			</td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><div class="mainshow"><button class="mainbuttonbottom" type="button" onclick="hidetablerows(9)"><%=yyCliHid%> &#9650;</button> <button class="mainbuttonbottom" type="submit"><%=IIfVr(ectdemostore,"Disabled",yySubmit)%></button></div></td>
		</tr>
	</table>
	
	<h3 class="round_top half_top"><button class="mainbutton" id="mainbutton10" type="button" onclick="showtablerows(10,'<%=jsescape("Image File Uploads")%>')"><%="Image File Uploads"%> <div style="float:right">&#9660;</div></button></h3>
	<table class="admin-table-b keeptable mainsettings" id="maintable10">
		<tr>
			<th colspan="2"><%
				print "If you wish to allow file or image uploads, you need to set the upload directory location. You are recommended to make this location outside of the web root so it cannot be accessed via HTTP."
				print "To help set this value, this is the location of the web root: " & Server.MapPath("/")
			%></th>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%="File Upload Directory"%>: </strong></td>
			<td><input type="text" name="uploaddir" size="50" value="<%=htmlspecials(rs("uploadDir"))%>" /></td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><div class="mainshow"><button class="mainbuttonbottom" type="button" onclick="hidetablerows(10)"><%=yyCliHid%> &#9650;</button> <button class="mainbuttonbottom" type="submit"><%=IIfVr(ectdemostore,"Disabled",yySubmit)%></button></div></td>
		</tr>
	</table>

	<h3 class="round_top half_top"><button class="mainbutton" id="mainbutton11" type="button" onclick="showtablerows(11,'<%=jsescape("Cardinal Commerce")%>')"><%="Cardinal Commerce"%> <div style="float:right">&#9660;</div></button></h3>
	<table class="admin-table-b keeptable mainsettings" id="maintable11">
		<tr><th colspan="2"><%=yyCaCoAc%></th></tr>
		<tr class="maintablehidden">
			<td><strong><%="Cardinal Processor ID"%>: </strong></td>
			<td><input type="text" name="cardinalprocessor" size="30" value="<%=htmlspecials(rs("cardinalProcessor"))%>" /></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%="Cardinal Merchant ID"%>: </strong></td>
			<td><input type="text" name="cardinalmerchant" size="30" value="<%=htmlspecials(rs("cardinalMerchant"))%>" /></td>
		</tr>
		<tr class="maintablehidden">
			<td><strong><%="Cardinal Transaction Password"%>: </strong></td>
			<td><input type="text" name="cardinalpwd" size="30" value="<%=htmlspecials(rs("cardinalPwd"))%>" /></td>
		</tr>
		<tr class="maintablehidden">
			<td colspan="2"><div class="mainshow"><button class="mainbuttonbottom" type="button" onclick="hidetablerows(11)"><%=yyCliHid%> &#9650;</button> <button class="mainbuttonbottom" type="submit"><%=IIfVr(ectdemostore,"Disabled",yySubmit)%></button></div></td>
		</tr>
	</table>

</form>
<script>
<!--
<% if trim(rs("emailPass")&"")="" then print "document.getElementById('emailpass').disabled=true;" %>
//-->
</script>
<%	rs.close
end if
cnn.Close
set rs = nothing
set cnn = nothing%>