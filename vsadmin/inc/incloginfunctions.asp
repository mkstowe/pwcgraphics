<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if incfunctionsdefined<>TRUE AND isadmincsv<>TRUE then
	response.write "Illegal Call"
	response.end
end if
set cnn=Server.CreateObject("ADODB.Connection")
set rs=Server.CreateObject("ADODB.RecordSet")
cnn.open sDSN
alreadygotadmin=getadminsettings()
rs.open "SELECT adminlang FROM admin WHERE adminid=1",cnn,0,1
adminlang=trim(rs("adminlang")&"")
rs.close
   if adminlang="de" then %>
<!--#include file="languageadmin_de.asp"-->
<!--#include file="languagefile_de.asp"-->
<% elseif adminlang="es" then %>
<!--#include file="languageadmin_es.asp"-->
<!--#include file="languagefile_es.asp"-->
<% elseif adminlang="fr" then %>
<!--#include file="languageadmin_fr.asp"-->
<!--#include file="languagefile_fr.asp"-->
<% elseif adminlang="it" then %>
<!--#include file="languageadmin_it.asp"-->
<!--#include file="languagefile_it.asp"-->
<% elseif adminlang="nl" then %>
<!--#include file="languageadmin_nl.asp"-->
<!--#include file="languagefile_nl.asp"-->
<% else %>
<!--#include file="languageadmin.asp"-->
<!--#include file="languagefile_en.asp"-->
<% end if
if forceloginonhttps then
	thisurl=request.servervariables("URL")
	if left(thisurl,1)="/" then thisurl=right(thisurl,len(thisurl)-1)
	if request.servervariables("HTTPS")="off" AND instr(storeurlssl,"https:")>0 then response.redirect storeurlssl & thisurl : response.end
end if
if storesessionvalue="" then storesessionvalue="virtualstore"
mustchangefordate=FALSE
if padssfeatures=TRUE then
	response.AddHeader "pragma","no-cache"
	response.CacheControl = "no-store,no-cache"
end if
sub adminsuccessforward(turl,tquery)
	response.write "<div style=""text-align:center;padding:50px 0"">"
		response.write "<div style=""font-weight:bold"">" & yyUpdSuc & "</div>"
		response.write "<div style=""padding:50px 0"">" & yyNowFrd & "</div>"
        response.write "<div style=""padding:0 0 50px 0"">" & yyNoAuto & " <a href=""" & turl & IIfVs(tquery<>"","?"&tquery) & """>" & yyClkHer & "</a>.</div>"
	response.write "</div>" & vbLf
end sub
sub adminfailback(errmsg)
	response.write "<div style=""text-align:center;padding:50px 0"">"
		response.write "<div style=""color:#FF0000;font-weight:bold"">" & yyOpFai & "</div>"
		response.write "<div style=""padding:50px 0"">" & errmsg & "</div>"
		response.write "<div style=""padding:0 0 50px 0""><a href=""javascript:history.go(-1)"">" & yyClkBac & "</div>"
	response.write "</div>" & vbLf
end sub
sub updaterchecker()
	set cnnu=server.createobject("ADODB.Connection")
	cnnu.open sDSN
	set rsu=server.createobject("ADODB.RecordSet")
	sSQL="SELECT adminVersion,updLastCheck,updRecommended,updSecurity,updShouldUpd,adminStoreURL FROM admin WHERE adminID=1"
	rsu.open sSQL,cnnu,0,1
	storeVersion=rsu("adminVersion")
	updLastCheck=rsu("updLastCheck")
	if isdate(updLastCheck) then updLastCheck=datevalue(updLastCheck)
	recommendedversion=rsu("updRecommended")
	securityrelease=rsu("updSecurity")
	shouldupdate=rsu("updShouldUpd")
	storeURL=rsu("adminStoreURL")
	rsu.close
	cnnu.close
	set cnnu=nothing
	set rsu=nothing
	checkupdates=(date()-updLastCheck>=3) OR NOT isdate(updLastCheck)
	if disableupdatechecker then
		checkupdates=FALSE
	else
%>
<script>
/* <![CDATA[ */
function ajaxcallback(){
	if(ajaxobj.readyState==4){
		var newtxt='';
		var xmlDoc=ajaxobj.responseXML.documentElement;
		var recver=xmlDoc.getElementsByTagName("recommendedversion")[0].childNodes[0].nodeValue;
		var shouldupdate=(xmlDoc.getElementsByTagName("shouldupdate")[0].childNodes[0].nodeValue=='true');
		var securityupdate=(xmlDoc.getElementsByTagName("securityupdate")[0].childNodes[0].nodeValue=='true');
		var haserror=(xmlDoc.getElementsByTagName("haserror")[0].childNodes[0].nodeValue=='true');
		if(haserror){
			newtxt='<span style="color:#FF0000;font-weight:bold">' + recver + '!</span><br /><%=replace(yyChkMan,"'","\'")%> <a href="https://www.ecommercetemplates.com/updaters.asp" target="_blank"><%=replace(yyClkHer,"'","\'")%></a><br />';
			newtxt += 'To disable this function please <a href="https://www.ecommercetemplates.com/help/ecommplus/parameters.asp#dissupcheck" target="_blank"><%=replace(yyClkHer,"'","\'")%></a><br />';
		}else{
			if(shouldupdate) newtxt='<a href="https://www.ecommercetemplates.com/updaters.asp" target="_blank"><%=replace(yyNewRec,"'","\'")%>: v' + recver + '</a><br />';
			if(securityupdate) newtxt += '<span style="color:#FF0000;font-weight:bold"><%=replace(yyRUSec,"'","\'")%></span><br />';
		}
		document.getElementById("checkupdates").innerHTML=(shouldupdate?'<div class="should_update">'+newtxt+'</div>':'<div class="updates_okay"><%=replace(yyNoNew,"'","\'")%></div>');
	}
}
function checkforupdates(){
	if(window.XMLHttpRequest)
		ajaxobj=new XMLHttpRequest();
	else
		ajaxobj=new ActiveXObject("MSXML2.XMLHTTP");
	ajaxobj.onreadystatechange=ajaxcallback;
	ajaxobj.open("GET", "ajaxservice.asp?action=checkupdates&storever=<%=urlencode(storeVersion)%>", true);
	ajaxobj.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
	ajaxobj.send(null);
}
<% if checkupdates then response.write "checkforupdates();" & vbCrLf
%>/* ]]> */
</script>
<%	end if ' disableupdatechecker
%><div class="updatecheck">
<div class="current_version<% if shouldupdate then response.write " old_version"%>"><%=replace(storeVersion, " v", "ASP v")%></div>
<span id="checkupdates"><%
	if disableupdatechecker then
		response.write "<div class=""updates_okay"">Auto update feature disabled! " & yyChkMan & " <a href=""https://www.ecommercetemplates.com/updaters.asp"" target=""_blank"">" & yyClkHer & "</a></div>"
	elseif checkupdates then
		response.write "<div class=""updates_okay"">" & yyChkNew & "...</div>"
	else
		if shouldupdate then
			response.write "<div class=""should_update"&IIfVs(securityrelease," security_update")&""">"
			response.write "<a href=""https://www.ecommercetemplates.com/updaters.asp"" target=""_blank"">" & yyNewRec & ": v" & recommendedversion & "</a>"
			if securityrelease then response.write "<br><span>" & yyRUSec & "</span>"
			response.write "</div>"
		else
			response.write "<div class=""updates_okay"">" & yyNoNew & "<div class=""last_update"">" & yyLasChk & ": <span class=""update_date""><a href=""javascript:checkforupdates()"">" & updLastCheck & "</a></span></div></div>"
		end if
	end if %></span>
</div>
<%
	if padssfeatures=TRUE then %>
<script>
/* <![CDATA[ */
var ecttimo=0;
function dokeepalive(){
	clearTimeout(ecttimo);
	if(window.XMLHttpRequest)
		ajaxobj=new XMLHttpRequest();
	else
		ajaxobj=new ActiveXObject("MSXML2.XMLHTTP");
	ajaxobj.open("GET", "ajaxservice.asp", true);
	ajaxobj.send('');
	document.getElementById('logindiv').style.display='none';
	setlotimos();
	return(false);
}
function setlotimos(){
	setTimeout("document.getElementById('logindiv').style.display='block';document.getElementById('contbutton').focus();",870000);
	ecttimo=setTimeout("document.location='logout.asp';",900000);
}
if(ecttimo==0)setlotimos();
/* ]]> */
</script>
<div id="logindiv" style="display:none;position:absolute;width:100%;height:2000px;background-image:url(adminimages/opaquepixel.png);top:0px;left:0px;text-align:center;z-index:10000;">
<br /><br /><br /><br /><br /><br /><br /><br />
<table width="100%"><tr><td align="center">
<form method="post" action="admin.asp" onsubmit="return dokeepalive()">
<table width="350" cellspacing="2" cellpadding="2" bgcolor="#FFFFFF"><tr><td align="center"><br /><br /><%=yyWilLog%><br /><br /><%=yyCliCon%><br /><br />
<%=yyFinMor%> <a href="https://www.ecommercetemplates.com/pa-dss-compliance.asp#padss5" target="_blank"><strong><%=yyClkHer%></strong></a>.<br /><br />
</td></tr>
<tr><td align="center"><br /><input type="submit" id="contbutton" value="<%=yyContin%>" /> &nbsp; <input type="button" value="<%=yyCancel%>" onclick="document.getElementById('logindiv').style.display='none'" /><br /><br /></td></tr></table>
</form>
</td></tr></table>
</div>
<%	end if
end sub
sub logeventlif(userid,eventtype,eventsuccess,eventorigin,areaaffected)
	if padssfeatures=TRUE then 
		sSQL = "SELECT logID FROM auditlog WHERE eventType='STARTLOG'"
		rs.open sSQL,cnn,0,1
		if rs.EOF then cnn.execute("INSERT INTO auditlog (userID,eventType,eventDate,eventSuccess,eventOrigin,areaAffected) VALUES (" & _
			"'" & escape_string(left(userid,48)) & "','STARTLOG'," & vsusdatetime(now) & ",1," & _
			"'" & escape_string(left(eventorigin,48)) & "','" & escape_string(left(areaaffected,48)) & "')")
		rs.close
		sSQL = "INSERT INTO auditlog (userID,eventType,eventDate,eventSuccess,eventOrigin,areaAffected) VALUES (" & _
			"'" & escape_string(left(userid,48)) & "','" & escape_string(left(eventtype,48)) & "'," & _
			vsusdatetime(now) & "," & IIfVr(eventsuccess,1,0) & "," & _
			"'" & escape_string(left(eventorigin,48)) & "','" & escape_string(left(areaaffected,48)) & "')"
		cnn.execute(sSQL)
		cnn.execute("DELETE FROM auditlog WHERE eventDate<" & vsusdatetime(date()-365))
	end if
end sub
sub adminassets() %>
<meta http-equiv="Content-Type" content="text/html; charset=<%=adminencoding%>"/>
<meta http-equiv="X-UA-Compatible" content="IE=9" />
<!-- Mobile Specific Meta
================================================== -->
<meta name="robots" content="noindex,nofollow">
<meta name="viewport" content="width=device-width, initial-scale=1" />
<link rel="stylesheet" type="text/css" href="adminstyle.css?ver=1" />
<link rel="stylesheet" type="text/css" href="ectadmincustom.css" />
<%	if trim(SESSION("loggedon"))<>"" then %>
<script src="//ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js"></script>
<% 		if htmleditor="froala" then %>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.4.0/css/font-awesome.min.css">
<link rel="stylesheet" href="froala/css/froala_editor.min.css">
<link rel="stylesheet" href="froala/css/froala_style.min.css">
<link rel="stylesheet" href="froala/css/plugins/code_view.min.css">
<link rel="stylesheet" href="froala/css/plugins/colors.min.css">
<link rel="stylesheet" href="froala/css/plugins/fullscreen.min.css">
<link rel="stylesheet" href="froala/css/plugins/image_manager.min.css">
<link rel="stylesheet" href="froala/css/plugins/image.min.css">
<link rel="stylesheet" href="froala/css/plugins/table.min.css">
<link rel="stylesheet" href="froala/css/plugins/video.min.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.3.0/codemirror.min.css">

<script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.3.0/codemirror.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.3.0/mode/xml/xml.min.js"></script>
<script src="froala/js/froala_editor.min.js"></script>
<script src="froala/js/plugins/align.min.js"></script>
<script src="froala/js/plugins/code_beautifier.min.js"></script>
<script src="froala/js/plugins/code_view.min.js"></script>
<script src="froala/js/plugins/colors.min.js"></script>
<script src="froala/js/plugins/draggable.min.js"></script>
<script src="froala/js/plugins/font_family.min.js"></script>
<script src="froala/js/plugins/font_size.min.js"></script>
<script src="froala/js/plugins/fullscreen.min.js"></script>
<script src="froala/js/plugins/image.min.js"></script>
<script src="froala/js/plugins/image_manager.min.js"></script>
<script src="froala/js/plugins/link.min.js"></script>
<script src="froala/js/plugins/lists.min.js"></script>
<script src="froala/js/plugins/paragraph_format.min.js"></script>
<script src="froala/js/plugins/paragraph_style.min.js"></script>
<script src="froala/js/plugins/table.min.js"></script>
<script src="froala/js/plugins/video.min.js"></script>
<script src="froala/js/plugins/url.min.js"></script>
<%		end if %>
<script src="assets/ectadmin.js"></script>
<%	end if
end sub
function displayfroalaeditor(regionid,regionname,extras,initimmediate,displayinline,uploadloc,wrapinfunction)
' initimmediate=FALSE,displayinline=FALSE,uploadloc=1,wrapinfunction=FALSE
toobarbuttons="toolbarButtons:['fullscreen','html','undo','redo','selectAll','clearFormatting','|','bold','italic','underline','strikeThrough','superscript','insertHR','|','fontFamily', 'fontSize', 'color','paragraphFormat','paragraphStyle','align','formatOL','formatUL','indent','outdent','|','insertLink','insertTable'],"
htmlallowedtags="htmlAllowedTags:['a','abbr','address','area','article','aside','audio','b','base','bdi','bdo','blockquote','br','button','canvas','caption','cite','code','col','colgroup','datalist','dd','del','details','dfn','dialog','div','dl','dt','ecttab','em','embed','fieldset','figcaption','figure','font','footer','form','h1','h2','h3','h4','h5','h6','header','hgroup','hr','i','iframe','img','input','ins','kbd','keygen','label','legend','li','link','main','map','mark','menu','menuitem','meter','nav','noscript','object','ol','optgroup','option','output','p','param','pre','progress','queue','rp','rt','ruby','s','samp','script','style','section','select','small','source','span','strike','strong','sub','summary','sup','table','tbody','td','textarea','tfoot','th','thead','time','title','tr','track','u','ul','var','video','wbr'],"
htmlallowedattrs="htmlAllowedAttrs:['accept','accept-charset','accesskey','action','align','allowfullscreen','allowtransparency','alt','async','autocomplete','autofocus','autoplay','autosave','background','bgcolor','border','charset','cellpadding','cellspacing','checked','cite','class','color','cols','colspan','content','contenteditable','contextmenu','controls','coords','data','data-.*','datetime','default','defer','dir','dirname','disabled','download','draggable','dropzone','enctype','face','for','form','formaction','frameborder','headers','height','hidden','high','href','hreflang','http-equiv','icon','id','ismap','itemprop','keytype','kind','label','lang','language','list','loop','low','max','maxlength','media','method','min','mozallowfullscreen','multiple','muted','name','novalidate','onblur','onchange','onclick','ondblclick','onerror','onfocus','onload','onselect','onsubmit','onreset','onkeydown','onkeypress','onkeyup','onmouseover','onmouseout','onmousedown','onmouseup','onmousemove','onresize','onunload','open','optimum','pattern','ping','placeholder','playsinline','poster','preload','pubdate','radiogroup','readonly','rel','required','reversed','rows','rowspan','sandbox','scope','scoped','scrolling','seamless','selected','shape','size','sizes','span','src','srcdoc','srclang','srcset','start','step','summary','spellcheck','style','tabindex','target','title','type','translate','usemap','value','valign','webkitallowfullscreen','width','wrap'],"
htmlallowedemptytags="htmlAllowedEmptyTags:['.fa','a','ecttab','iframe','label','object','style','script','textarea','video'],"
htmlremovetags="htmlRemoveTags:[''],"
htmldonotwraptags="htmlDoNotWrapTags:['ecttab','script','style'],"
ectstorecss=""
if customcsslocation<>"" then ectstorecss=ectstorecss&","""& replace(customcsslocation,",",""",""") & """"
response.write "<script>" & IIfVs(wrapinfunction,"function dfe_"&regionid&"(){") & "$('#"&regionid&"').froalaEditor({iframeStyle:'body{height:auto !important;background:none !important;display:block !important;border:10px !important;margin:8px !important;}',iframeStyleFiles:['../css/ectcart.css','../css/ectstyle.css','../css/style.css'" & ectstorecss & "],iframe:true,key:""IG1A3A3A1pD1D1E1A3E1J4A14B3A7C7kWd1WDPTa1ZNRGe1OC1c1==""," & toobarbuttons & htmlallowedtags & htmlallowedattrs & htmlallowedemptytags & htmlremovetags & htmldonotwraptags & "placeholderText:'Click to Edit "&jsescape(regionname)&"',language: '"&adminlang&"'"&IIfVs(NOT initimmediate,",initOnClick:true")&IIfVs(displayinline,",toolbarInline:true")&",imageUploadURL:'froala/upload_image.php?loc="&uploadloc&"'})"&extras& IIfVs(wrapinfunction,"}") & "</script>" & vbLf
end function
sub adminheader() %>
<div id="header1">
<div class="inner_half"><p class="viewstore"><a class="topbar" href="../" target="_blank"><%=yyLmVwSt%></a></p></div>
<div class="inner_half aright last">
<p class="log-out"><a class="topbar" href="logout.asp"><%=yyLLLogO%></a></p>
</div>
</div>
<div id="header">
<div class="logotop"><a href="admin.asp"><img src="adminimages/logo.png" alt="Ecommerce Templates" class="ectlogo" /></a></div>
<div class="toplinks">
<a class="topbar" href="https://www.ecommercetemplates.com/help.asp" target="_blank"><img src="adminimages/icon-help.png" title="<%=yyLLHelp%>" alt="<%=yyLLHelp%>"></a> 
<a href="https://www.ecommercetemplates.com/support/default.asp" target="_blank" class="topbar"><img src="adminimages/icon-forum.png" title="<%=yyLLForu%>" alt="<%=yyLLForu%>"></a>
<a href="https://www.ecommercetemplates.com/support/search.asp" target="_blank" class="topbar"><img src="adminimages/icon-search.png" title="<%=yyLLForS%>" alt="<%=yyLLForS%>" ></a>
<a href="https://www.ecommercetemplates.com/updaters.asp" target="_blank" class="topbar"><img src="adminimages/icon-update.png" alt="<%=yyLLUpda%>" title="<%=yyLLUpda%>"></a>
</div>
</div>
<%
	set cnn=Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.RecordSet")
	cnn.open sDSN
	if trim(request.querystring("cleardnalert"))="true" then
		sSQL="UPDATE admin SET adminDeviceNotifAlert='' WHERE adminID=1"
		cnn.execute(sSQL)
	else
		sSQL="SELECT updRecommended,updSecurity,updShouldUpd,adminDeviceNotifAlert FROM admin WHERE adminID=1"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			recommendedversion=rs("updRecommended")
			securityrelease=rs("updSecurity")
			shouldupdate=rs("updShouldUpd")
		end if
		if trim(rs("adminDeviceNotifAlert")&"")<>"" OR shouldupdate then
			response.write "<div class=""headeradvis"">"
			if trim(rs("adminDeviceNotifAlert")&"")<>"" then response.write "<div><input type=""button"" value=""Clear"" onclick=""document.location='admin.asp?cleardnalert=true'"" /> " & rs("adminDeviceNotifAlert") & "</div>"
			if shouldupdate then
				response.write "<div class=""should_update" & IIfVs(securityrelease," security_update") & """>"
				response.write "<a href=""https://www.ecommercetemplates.com/updaters.asp"" target=""_blank"">" & yyNewRec & ": v" & recommendedversion & "</a>"
				if securityrelease then response.write "<br><span>" & yyRUSec & "</span>"
				response.write "</div>"
			end if
			response.write "</div>"
		end if
		rs.close
	end if
	cnn.close
end sub
sub adminnavigation() %>
<div id="admin_menu">
<div class='menu-button'>Menu</div>
<nav>
<ul id="nav" role="navigation">
 <li class="top-level"><a href="admin.asp"><%=yyLMStAd%></a>
<ul class="sub-menu">
	<li><a href="admin.asp"><%=lcase(yyDashbd)%></a></li>
    <li><a href="adminmain.asp"><%=yyLLMain%></a></li>
    <li><a href="adminlogin.asp"><%=yyLLPass%></a></li>
    <li><a href="adminaffil.asp"><%=yyLLAffl%></a></li>
	<li><a href="adminemailmsgs.asp"><%=yyLMEmla%></a></li>
    <li><a href="adminmailinglist.asp"><%=yyLMMaLi%></a></li>
    <li><a href="admincontent.asp"><%=lcase(yyContReg)%></a></li>
	<li><a href="adminipblock.asp"><%=lcase(yyIPBlock)%></a></li>
	<li><a href="admindbutility.asp">database utility</a></li>
</ul>
 <li class="top-level"><a href="adminorders.asp"><%=yyOrdAdm%></a>
<ul class="sub-menu">
    <li><a href="adminorders.asp"><%=yyLLOrds%></a></li>
<%	if FALSE then
		Set cnndn=Server.CreateObject("ADODB.Connection")
		cnndn.open sDSN
		set rsdn = Server.CreateObject("ADODB.RecordSet")
		sSQL="SELECT dnID FROM devicenotifications ORDER BY dnID"
		rsdn.open sSQL,cnndn,0,1
		if NOT rsdn.EOF then %>
	<li><a href="adminorders.asp?act=dnotif" style="color:#FF7711">device notfications</a></li>
<%		end if
		rsdn.close
		set rsdn=nothing
	end if %>
    <li><a href="adminpayprov.asp"><%=yyLLPayP%></a></li>
    <li><a href="adminclientlog.asp"><%=yyLLClLo%></a></li>
    <li><a href="adminordstatus.asp"><%=yyLLOrSt%></a></li>
    <li><a href="admingiftcert.asp"><%=yyLLGftC%></a></li>
	<li><a href="adminstats.asp">order stats</a></li>
</ul>
 <li class="top-level"><a href="adminprods.asp"><%=yyLMPrAd%></a>
<ul class="sub-menu">
	<li><a href="adminprods.asp"><%=yyLLProA%></a></li>
    <li><a href="adminprodopts.asp"><%=yyLLProO%></a></li>
    <li><a href="admincats.asp"><%=yyLLCats%></a></li>
    <li><a href="admindiscounts.asp"><%=yyLLDisc%></a></li>
	<li><a href="adminsearchcriteria.asp"><%=lcase(yySeaCri)%></a></li>
    <li><a href="adminpricebreak.asp"><%=yyLLQuan%></a></li>
    <li><a href="adminratings.asp"><%=xxLMRaRv%></a></li>
	<li><a href="admincsv.asp"><%=lcase(yyCSVUpl)%></a></li>
</ul>
 <li class="top-level"><a href="adminuspsmeths.asp"><%=yyLMShAd%></a>
<ul class="sub-menu">
	<li><a href="adminstate.asp"><%=yyLLStat%></a></li>
    <li><a href="admincountry.asp"><%=yyLLCoun%></a></li>
    <li><a href="adminzones.asp"><%=yyLLZone%></a></li>
    <li><a href="adminuspsmeths.asp"><%=yyLLShpM%></a></li>
    <li><a href="admindropship.asp"><%=yyDrShpr%></a></li>
</ul>
</ul>
</nav>
</div>
<%	if onvacation<>0 then %>
<div style="color:#ED2803;text-align:center;font-size:1.3em;padding-top:16px">
The Vacation Setting is Currently Active. <a href="adminmain.asp#vacationmessage">To Turn This off, click here.</a>
</div>
<%	end if
end sub
sub adminfooter() %>
<div id="adminfooter">
<div class="row footer_block">
<div class="one_third footer_half">
<h4>Store Help</h4>
<ul>
<li><a href="https://www.ecommercetemplates.com/help/ecommplus/about.asp">ASP Help Files</a></li>
<li><a href="https://www.ecommercetemplates.com/help/admin-help.asp">Admin Help Files</a></li>
<li><a href="https://www.ecommercetemplates.com/free_downloads.asp#usermanual">User manual</a></li>
<li><a href="https://www.ecommercetemplates.com/help/ecommplus/parameters.asp">Store settings</a></li>
<li><a href="https://www.ecommercetemplates.com/tutorials/">Tutorials</a></li>
<li><a href="https://www.ecommercetemplates.com/support/">Support Forum</a></li>
</ul>
</div>
<div class="one_third footer_half">
<h4>Resources</h4>
<ul>
<li><a href="https://www.ecommercetemplates.com/affiliateinfo.asp" target="_blank"><%=yyLLAffP%></a></li>
<li><a href="https://www.ecommercetemplates.com/addsite.asp" target="_blank"><%=yyLLSubm%></a></li>
<li><a href="https://www.ecommercetemplates.com/payment_processors.asp">Payment providers</a></li>
<li><a href="https://www.ecommercetemplates.com/free_downloads.asp">Store downloads</a></li>
<li><a href="https://www.ecommercetemplates.com/ecommercetools.asp">Store tools &amp; add-ons</a></li>
<li><a href="https://www.ecommercetemplates.com/newsletter/default.asp">Ecommerce Templates News</a></li>
</ul>
</div>
<div class="one_third last">
<h4>Social media</h4>
<p>
<a href="https://www.facebook.com/EcommerceTemplates" target="_blank"><img src="adminimages/fb.gif" alt="Facebook" width="32" height="32" /></a>
<a href="https://twitter.com/etemplates/" target="_blank"><img src="adminimages/tw.gif" alt="Twitter" width="32" height="32" /></a>
<a href="https://www.linkedin.com/in/ecommercetemplates" target="_blank"><img src="adminimages/li.gif" alt="Linkedin" width="32" height="32" /></a>
<a href="https://www.youtube.com/user/EcommerceTemplates" target="_blank"><img src="adminimages/yt.gif" alt="YouTube" width="32" height="32" border="0" /></a>
</p>
</div>
</div>
<%	on error resume next
	updaterchecker()
	on error goto 0 %>
</div>
<script> 
$("[role='navigation']").mainnav(); 
$('#responsive-table').stacktable({myClass:'admin-table-a-small'});
</script>
<% end sub
if NOT donotlogin then
	if SESSION("loggedon") <> storesessionvalue AND trim(request.cookies("WRITECKL"))<>"" AND disallowlogin<>TRUE AND NOT (loginkey<>"" AND trim(request.cookies("loginkey"))<>loginkey) then
		if padssfeatures=TRUE then session.timeout=15
		sSQL="SELECT adminID,adminUser,adminPWLastChange FROM admin WHERE adminPassword='" & replace(request.cookies("WRITECKP"),"'","''") & "' AND adminID=1"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			if rs("adminUSER")=request.cookies("WRITECKL") then
				SESSION("loggedon") = storesessionvalue
				SESSION("loggedonpermissions") = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
				SESSION("loginid")=0
				SESSION("loginuser")=rs("adminUser")
				if date()-rs("adminPWLastChange")>90 AND padssfeatures=TRUE then SESSION("mustchangepw")="B" : mustchangefordate=TRUE
			end if
		end if
		rs.close
		if SESSION("loggedon") <> storesessionvalue then
			sSQL="SELECT adminloginid,adminloginname,adminloginpermissions,adminLoginLastChange FROM adminlogin WHERE adminloginpassword='" & replace(request.cookies("WRITECKP"),"'","''") & "'"
			rs.open sSQL,cnn,0,1
			do while NOT rs.EOF
				if rs("adminloginname")=request.cookies("WRITECKL") then
					SESSION("loggedon") = storesessionvalue
					SESSION("loggedonpermissions") = rs("adminloginpermissions")
					SESSION("loginid")=rs("adminloginid")
					SESSION("loginuser")=rs("adminloginname")
					if date()-rs("adminLoginLastChange")>90 AND padssfeatures=TRUE then SESSION("mustchangepw")="B" : mustchangefordate=TRUE
				end if
				rs.movenext
			loop
			rs.close
		end if
		call logeventlif(request.cookies("WRITECKL"),"LOGIN",SESSION("loggedon")=storesessionvalue,"LOGIN","")
	end if
	if SESSION("loggedon") <> storesessionvalue OR disallowlogin=TRUE then response.redirect "login.asp" & IIfVs(getget("loginkey")<>"","?loginkey="&htmlspecials(getget("loginkey")))
	if (SESSION("mustchangepw")<>"" OR mustchangefordate) AND NOT (thispagename="adminlogin") then response.redirect "adminlogin.asp"
end if
isprinter=false
cnn.close
set cnn=nothing
%>