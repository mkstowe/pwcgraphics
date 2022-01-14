<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
Dim sSQL,rs,alladmin,success,cnn,rowcounter,errmsg,aFields(1)
success=true
Set rs = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin = getadminsettings()
Session.LCID = 1033
sSQL = ""
if maxloginlevels="" then maxloginlevels=5
dorefresh=FALSE
sub dodeletediscount(did)
	sSQL="DELETE FROM cpnassign WHERE cpaCpnID=" & did
	ect_query(sSQL)
	sSQL="DELETE FROM coupons WHERE cpnID=" & did
	ect_query(sSQL)
end sub
if getpost("posted")="1" then
	if getpost("act")="delete" then
		call dodeletediscount(getpost("id"))
		dorefresh=TRUE
	elseif getpost("act")="quickupdate" then
		for each objItem in request.form
			if left(objItem, 4)="pra_" then
				origid=right(objItem, len(objItem)-4)
				theid=getpost("pid"&origid)
				theval=getpost(objItem)
				cract=getpost("cract")
				sSQL=""
				if cract="del" then
					if theval="del" then dodeletediscount(theid)
					sSQL=""
				end if
				if sSQL<>"" then
					sSQL=sSQL & " WHERE cpnID="&int(theid)
					ect_query(sSQL)
				end if
			end if
		next
		dorefresh=TRUE
	elseif getpost("act")="domodify" then
		cpnName = getpost("cpnName")
		sSQL = "UPDATE coupons SET cpnName='" & escape_string(getpost("cpnName")) & "'"
			for index=2 to adminlanguages+1
				if (adminlangsettings AND 1024)=1024 then sSQL=sSQL & ",cpnName"&index&"='" & escape_string(getpost("cpnName"&index)) & "'"
			next
			if getpost("cpnWorkingName")<>"" then
				sSQL=sSQL & ",cpnWorkingName='" & escape_string(getpost("cpnWorkingName"))&"'"
			else
				sSQL=sSQL & ",cpnWorkingName='" & escape_string(cpnName)&"'"
			end if
			if getpost("cpnIsCoupon")="0" then
				sSQL=sSQL & ",cpnNumber='',"
			else
				sSQL=sSQL & ",cpnNumber='" & escape_string(getpost("cpnNumber")) & "',"
			end if
			sSQL=sSQL & "cpnType=" & getpost("cpnType") & ","
			if isdate(getpost("cpnStartDate")) then
				sSQL=sSQL & "cpnStartDate=" & vsusdate(cdate(getpost("cpnStartDate"))) & ","
			else
				sSQL=sSQL & "cpnStartDate=" & vsusdate(DateSerial(2000,1,1)) & ","
			end if
			if getpost("cpnEndDate")="Expired" then
				sSQL=sSQL & "cpnEndDate=" & vsusdate(date()-30) & ","
			elseif isdate(getpost("cpnEndDate")) then
				sSQL=sSQL & "cpnEndDate=" & vsusdate(cdate(getpost("cpnEndDate"))) & ","
			else
				sSQL=sSQL & "cpnEndDate=" & vsusdate(DateSerial(3000,1,1)) & ","
			end if
			if is_numeric(getpost("cpnDiscount")) AND getpost("cpnType")<>"0" then
				sSQL=sSQL & "cpnDiscount=" & getpost("cpnDiscount") & ","
			else
				sSQL=sSQL & "cpnDiscount=0,"
			end if
			if is_numeric(getpost("cpnThreshold")) then
				sSQL=sSQL & "cpnThreshold=" & getpost("cpnThreshold") & ","
			else
				sSQL=sSQL & "cpnThreshold=0,"
			end if
			if is_numeric(getpost("cpnThresholdMax")) then
				sSQL=sSQL & "cpnThresholdMax=" & getpost("cpnThresholdMax") & ","
			else
				sSQL=sSQL & "cpnThresholdMax=0,"
			end if
			if is_numeric(getpost("cpnThresholdRepeat")) then
				sSQL=sSQL & "cpnThresholdRepeat=" & getpost("cpnThresholdRepeat") & ","
			else
				sSQL=sSQL & "cpnThresholdRepeat=0,"
			end if
			if is_numeric(getpost("cpnQuantity")) then
				sSQL=sSQL & "cpnQuantity=" & getpost("cpnQuantity") & ","
			else
				sSQL=sSQL & "cpnQuantity=0,"
			end if
			if is_numeric(getpost("cpnQuantityMax")) then
				sSQL=sSQL & "cpnQuantityMax=" & getpost("cpnQuantityMax") & ","
			else
				sSQL=sSQL & "cpnQuantityMax=0,"
			end if
			if is_numeric(getpost("cpnQuantityRepeat")) then
				sSQL=sSQL & "cpnQuantityRepeat=" & getpost("cpnQuantityRepeat") & ","
			else
				sSQL=sSQL & "cpnQuantityRepeat=0,"
			end if
			if is_numeric(getpost("cpnNumAvail")) then
				sSQL=sSQL & "cpnNumAvail=" & getpost("cpnNumAvail") & ","
			else
				sSQL=sSQL & "cpnNumAvail=30000000,"
			end if
			if getpost("cpnType")="0" then
				sSQL=sSQL & "cpnCntry=" & getpost("cpnCntry") &","
			else
				sSQL=sSQL & "cpnCntry=0,"
			end if
			cpnLoginLevel=int(getpost("cpnLoginLevel"))
			if getpost("cpnLoginLt")="1" then cpnLoginLevel=-1-cpnLoginLevel
			sSQL=sSQL & "cpnLoginLevel="&cpnLoginLevel&","
			if is_numeric(getpost("cpnHandling")) then sSQL=sSQL & "cpnHandling=" & getpost("cpnHandling") &","
			if is_numeric(getpost("cpnInsurance")) then sSQL=sSQL & "cpnInsurance=" & getpost("cpnInsurance") &","
			sSQL=sSQL & "cpnIsCoupon=" & getpost("cpnIsCoupon") &","
			if getpost("cpnType")="0" then
				sSQL=sSQL & "cpnSitewide=1"
			else
				sSQL=sSQL & "cpnSitewide=" & getpost("cpnSitewide")
			end if
			sSQL=sSQL & " WHERE cpnID="&getpost("id")
		ect_query(sSQL)
		dorefresh=TRUE
	elseif getpost("act")="doaddnew" then
		cpnName = getpost("cpnName")
		sSQL = "INSERT INTO coupons (cpnName,"
			for index=2 to adminlanguages+1
				if (adminlangsettings AND 1024)=1024 then sSQL=sSQL & "cpnName"&index&","
			next
			sSQL=sSQL & "cpnWorkingName,cpnNumber,cpnType,cpnStartDate,cpnEndDate,cpnDiscount,cpnThreshold,cpnThresholdMax,cpnThresholdRepeat,cpnQuantity,cpnQuantityMax,cpnQuantityRepeat,cpnNumAvail,cpnCntry,cpnLoginLevel,cpnHandling,cpnInsurance,cpnIsCoupon,cpnSitewide) VALUES (" & _
			"'"&escape_string(cpnName)&"',"
			for index=2 to adminlanguages+1
				if (adminlangsettings AND 1024)=1024 then sSQL=sSQL & "'"&escape_string(getpost("cpnName"&index))&"',"
			next
			if getpost("cpnWorkingName")<>"" then
				sSQL=sSQL & "'"&escape_string(getpost("cpnWorkingName"))&"',"
			else
				sSQL=sSQL & "'"&escape_string(cpnName)&"',"
			end if
			if getpost("cpnIsCoupon")="0" then
				sSQL=sSQL & "'',"
			else
				sSQL=sSQL & "'"&escape_string(getpost("cpnNumber"))&"',"
			end if
			sSQL=sSQL & getpost("cpnType") & ","
			numdays=0
			if isdate(getpost("cpnStartDate")) then
				sSQL=sSQL & vsusdate(cdate(getpost("cpnStartDate"))) & ","
			else
				sSQL=sSQL & vsusdate(DateSerial(2000,1,1)) & ","
			end if
			if getpost("cpnEndDate")="Expired" then
				sSQL=sSQL & vsusdate(date()-30) & ","
			elseif getpost("cpnEndDate")<>"" then
				sSQL=sSQL & vsusdate(cdate(getpost("cpnEndDate"))) & ","
			else
				sSQL=sSQL & vsusdate(DateSerial(3000,1,1)) & ","
			end if
			if is_numeric(getpost("cpnDiscount")) AND getpost("cpnType")<>"0" then
				sSQL=sSQL & getpost("cpnDiscount") & ","
			else
				sSQL=sSQL & "0,"
			end if
			if is_numeric(getpost("cpnThreshold")) then
				sSQL=sSQL & getpost("cpnThreshold") & ","
			else
				sSQL=sSQL & "0,"
			end if
			if is_numeric(getpost("cpnThresholdMax")) then
				sSQL=sSQL & getpost("cpnThresholdMax") & ","
			else
				sSQL=sSQL & "0,"
			end if
			if is_numeric(getpost("cpnThresholdRepeat")) then
				sSQL=sSQL & getpost("cpnThresholdRepeat") & ","
			else
				sSQL=sSQL & "0,"
			end if
			if is_numeric(getpost("cpnQuantity")) then
				sSQL=sSQL & getpost("cpnQuantity") & ","
			else
				sSQL=sSQL & "0,"
			end if
			if is_numeric(getpost("cpnQuantityMax")) then
				sSQL=sSQL & getpost("cpnQuantityMax") & ","
			else
				sSQL=sSQL & "0,"
			end if
			if is_numeric(getpost("cpnQuantityRepeat")) then
				sSQL=sSQL & getpost("cpnQuantityRepeat") & ","
			else
				sSQL=sSQL & "0,"
			end if
			if is_numeric(getpost("cpnNumAvail")) then
				sSQL=sSQL & getpost("cpnNumAvail") & ","
			else
				sSQL=sSQL & "30000000,"
			end if
			if getpost("cpnType")="0" then
				sSQL=sSQL & getpost("cpnCntry") &","
			else
				sSQL=sSQL & "0,"
			end if
			cpnLoginLevel=int(getpost("cpnLoginLevel"))
			if getpost("cpnLoginLt")="1" then cpnLoginLevel=-1-cpnLoginLevel
			sSQL=sSQL & cpnLoginLevel&","
			if is_numeric(getpost("cpnHandling")) then sSQL=sSQL & getpost("cpnHandling") &"," else sSQL=sSQL & "0,"
			if is_numeric(getpost("cpnInsurance")) then sSQL=sSQL & getpost("cpnInsurance") &"," else sSQL=sSQL & "0,"
			sSQL=sSQL & getpost("cpnIsCoupon") &","
			if getpost("cpnType")="0" then
				sSQL=sSQL & "1)"
			else
				sSQL=sSQL & getpost("cpnSitewide") & ")"
			end if
		ect_query(sSQL)
		dorefresh=TRUE
	end if
	forwardquery="qe="&getpost("cract")&"&stext="&urlencode(request("stext"))&"&stype="&request("stype")&"&scpds="&request("scpds")&"&sefct="&request("sefct")&"&sort="&request("sort")&"&pg=" & request("pg")
	if dorefresh then
		print "<meta http-equiv=""refresh"" content=""1; url=admindiscounts.asp?" & forwardquery & """ />"
	end if
end if
%>
<script>
<!--
var savebg, savebc, savecol;
function formvalidator(theForm)
{
  if(theForm.cpnName.value == ""){
    alert("<%=jscheck(yyPlsEntr&" """&yyDisTxt)%>\".");
    theForm.cpnName.focus();
    return (false);
  }
  if(theForm.cpnName.value.length > 255){
    alert("<%=jscheck(yyMax255&" """&yyDisTxt)%>\".");
    theForm.cpnName.focus();
    return (false);
  }
  if(theForm.cpnType.selectedIndex!=0){
	if(theForm.cpnDiscount.value == ""){
	  alert("<%=jscheck(yyPlsEntr&" """&yyDscAmt)%>\".");
	  theForm.cpnDiscount.focus();
	  return (false);
	}
	if(theForm.cpnType.selectedIndex==2){
	  if(theForm.cpnDiscount.value < 0 || theForm.cpnDiscount.value > 100){
		alert("<%=jscheck(yyNum100&" """&yyDscAmt)%>\".");
		theForm.cpnDiscount.focus();
		return (false);
	  }
	}
  }
  if(theForm.cpnIsCoupon.selectedIndex==1){
	if(theForm.cpnNumber.value == ""){
	  alert("<%=jscheck(yyPlsEntr&" """&yyCpnCod)%>\".");
	  theForm.cpnNumber.focus();
	  return (false);
	}
	var regex=/^[0-9A-Za-z\_\-]+$/;
	if (!regex.test(theForm.cpnNumber.value)){
		alert("<%=jscheck(yyAlpha2&" """&yyCpnCod)%>\".");
		theForm.cpnNumber.focus();
		return (false);
	}
  }
  var regex=/^[0-9]*$/;
  if (!regex.test(theForm.cpnNumAvail.value)){
	alert("<%=jscheck(yyOnlyNum&" """&yyNumAvl)%>\".");
	theForm.cpnNumAvail.focus();
	return (false);
  }
  if(theForm.cpnNumAvail.value != "" && theForm.cpnNumAvail.value > 1000000){
    alert("<%=jscheck(yyNumMil&" """&yyNumAvl)%>\"<%=jscheck(yyOrBlank)%>");
    theForm.cpnNumAvail.focus();
    return (false);
  }
  var regex=/^[0-9\-\/]*$/;
  if(!regex.test(theForm.cpnStartDate.value)){
	alert("<%=jscheck(yyValEnt&" """&"Start Date")%>\".");
	theForm.cpnStartDate.focus();
	return (false);
  }
  if(!regex.test(theForm.cpnEndDate.value)&&theForm.cpnEndDate.value!="Expired"){
	alert("<%=jscheck(yyValEnt&" """&"End Date")%>\".");
	theForm.cpnEndDate.focus();
	return (false);
  }
  var regex=/^[0-9\.]*$/;
  if (!regex.test(theForm.cpnThreshold.value)){
	alert("<%=jscheck(yyOnlyDec&" """&yyMinPur)%>\".");
	theForm.cpnThreshold.focus();
	return (false);
  }
  var regex=/^[0-9\.]*$/;
  if (!regex.test(theForm.cpnThresholdRepeat.value)){
	alert("<%=jscheck(yyOnlyDec&" """&yyRepEvy)%>\".");
	theForm.cpnThresholdRepeat.focus();
	return (false);
  }
  var regex=/^[0-9\.]*$/;
  if (!regex.test(theForm.cpnThresholdMax.value)){
	alert("<%=jscheck(yyOnlyDec&" """&yyMaxPur)%>\".");
	theForm.cpnThresholdMax.focus();
	return (false);
  }
  var regex=/^[0-9]*$/;
  if (!regex.test(theForm.cpnQuantity.value)){
	alert("<%=jscheck(yyOnlyNum&" """&yyMinQua)%>\".");
	theForm.cpnQuantity.focus();
	return (false);
  }
  var regex=/^[0-9]*$/;
  if (!regex.test(theForm.cpnQuantityRepeat.value)){
	alert("<%=jscheck(yyOnlyNum&" """&yyRepEvy)%>\".");
	theForm.cpnQuantityRepeat.focus();
	return (false);
  }
  var regex=/^[0-9]*$/;
  if (!regex.test(theForm.cpnQuantityMax.value)){
	alert("<%=jscheck(yyOnlyNum&" """&yyMaxQua)%>\".");
	theForm.cpnQuantityMax.focus();
	return (false);
  }
  var regex=/^[0-9\.]*$/;
  if (!regex.test(theForm.cpnDiscount.value)){
	alert("<%=jscheck(yyOnlyDec&" """&yyDscAmt)%>\".");
	theForm.cpnDiscount.focus();
	return (false);
  }
  document.mainform.cpnNumber.disabled=false;
  document.mainform.cpnDiscount.disabled=false;
  document.mainform.cpnCntry.disabled=false;
  document.mainform.cpnHandling.disabled=false;
  document.mainform.cpnInsurance.disabled=false;
  document.mainform.cpnSitewide.disabled=false;
  document.mainform.cpnThresholdRepeat.disabled=false;
  document.mainform.cpnQuantityRepeat.disabled=false;
  return (true);
}
function couponcodeactive(forceactive){
	if(document.mainform.cpnIsCoupon.selectedIndex==0){
		document.mainform.cpnNumber.style.backgroundColor="#DDDDDD";
		document.mainform.cpnNumber.disabled=true;
	}
	else if(document.mainform.cpnIsCoupon.selectedIndex==1){
		document.mainform.cpnNumber.style.backgroundColor=savebg;
		document.mainform.cpnNumber.disabled=false;
	}
}
function changecouponeffect(forceactive){
	if(document.mainform.cpnType.selectedIndex==0){
		document.mainform.cpnDiscount.style.backgroundColor="#DDDDDD";
		document.mainform.cpnDiscount.disabled=true;

		document.mainform.cpnCntry.style.backgroundColor=savebg;
		document.mainform.cpnCntry.disabled=false;
		
		document.mainform.cpnHandling.style.backgroundColor=savebg;
		document.mainform.cpnHandling.disabled=false;
		
		document.mainform.cpnInsurance.style.backgroundColor=savebg;
		document.mainform.cpnInsurance.disabled=false;

		document.mainform.cpnSitewide.style.backgroundColor="#DDDDDD";
		document.mainform.cpnSitewide.disabled=true;
	}else{
		document.mainform.cpnDiscount.style.backgroundColor=savebg;
		document.mainform.cpnDiscount.disabled=false;

		document.mainform.cpnCntry.style.backgroundColor="#DDDDDD";
		document.mainform.cpnCntry.disabled=true;
		
		document.mainform.cpnHandling.style.backgroundColor="#DDDDDD";
		document.mainform.cpnHandling.disabled=true;
		
		document.mainform.cpnInsurance.style.backgroundColor="#DDDDDD";
		document.mainform.cpnInsurance.disabled=true;

		document.mainform.cpnSitewide.style.backgroundColor=savebg;
		document.mainform.cpnSitewide.disabled=false;
	}
	if(document.mainform.cpnType.selectedIndex==1){
		document.mainform.cpnThresholdRepeat.style.backgroundColor=savebg;
		document.mainform.cpnThresholdRepeat.disabled=false;

		document.mainform.cpnQuantityRepeat.style.backgroundColor=savebg;
		document.mainform.cpnQuantityRepeat.disabled=false;
	}else{
		document.mainform.cpnThresholdRepeat.style.backgroundColor="#DDDDDD";
		document.mainform.cpnThresholdRepeat.disabled=true;

		document.mainform.cpnQuantityRepeat.style.backgroundColor="#DDDDDD";
		document.mainform.cpnQuantityRepeat.disabled=true;
	}
}
function setloglev(isequal){
var tobj=document.getElementById('cpnLoginLevel');
if(isequal.selectedIndex==0)
	tobj[0].text="<%=yyNoRes%>";
else
	tobj[0].text="<%=yyLiLev & " 0"%>";
}
//-->
</script>
<%
if getpost("posted")="1" AND (getpost("act")="modify" OR getpost("act")="clone" OR getpost("act")="addnew") then %>
<script src="popcalendar.js"></script>
<%
	themask=cStr(DateSerial(2003,12,11))
	themask=replace(themask,"2003","yyyy")
	themask=replace(themask,"12","mm")
	themask=replace(themask,"11","dd")
	isexpired=FALSE
	if (getpost("act")="modify" OR getpost("act")="clone") AND is_numeric(getpost("id")) then
		sSQL = "SELECT cpnName,cpnName2,cpnName3,cpnWorkingName,cpnNumber,cpnType,cpnStartDate,cpnEndDate,cpnDiscount,cpnThreshold,cpnThresholdMax,cpnThresholdRepeat,cpnQuantity,cpnQuantityMax,cpnQuantityRepeat,cpnNumAvail,cpnCntry,cpnIsCoupon,cpnSitewide,cpnHandling,cpnInsurance,cpnLoginLevel FROM coupons WHERE cpnID="&getpost("id")
		rs.open sSQL,cnn,0,1
		cpnName = rs("cpnName")
		cpnName2 = rs("cpnName2")&""
		cpnName3 = rs("cpnName3")&""
		cpnWorkingName = rs("cpnWorkingName")
		cpnNumber = rs("cpnNumber")
		cpnType = rs("cpnType")
		Session.LCID = saveLCID
		cpnStartDate = IIfVr(rs("cpnStartDate")=dateserial(2000,1,1),"",cstr(rs("cpnStartDate")&""))
		isexpired=rs("cpnEndDate")-Date()<0
		cpnEndDate = IIfVr(rs("cpnEndDate")=dateserial(3000,1,1),"",cstr(rs("cpnEndDate")&""))
		session.LCID=1033
		cpnDiscount = rs("cpnDiscount")
		cpnThreshold = rs("cpnThreshold")
		cpnThresholdMax = rs("cpnThresholdMax")
		cpnThresholdRepeat = rs("cpnThresholdRepeat")
		cpnQuantity = rs("cpnQuantity")
		cpnQuantityMax = rs("cpnQuantityMax")
		cpnQuantityRepeat = rs("cpnQuantityRepeat")
		cpnNumAvail = rs("cpnNumAvail")
		cpnCntry = rs("cpnCntry")
		cpnIsCoupon = rs("cpnIsCoupon")
		cpnSitewide = rs("cpnSitewide")
		cpnHandling = rs("cpnHandling")
		cpnInsurance = rs("cpnInsurance")
		cpnLoginLevel = rs("cpnLoginLevel")
		cpnLoginLt = (cpnLoginLevel<0)
		cpnLoginLevel = abs(cpnLoginLevel)
		rs.close
	else
		cpnName = ""
		cpnName2 = ""
		cpnName3 = ""
		cpnWorkingName = ""
		cpnNumber = ""
		cpnType = 0
		cpnStartDate=""
		cpnEndDate=""
		cpnDiscount = ""
		cpnThreshold = 0
		cpnThresholdMax = 0
		cpnThresholdRepeat = 0
		cpnQuantity = 0
		cpnQuantityMax = 0
		cpnQuantityRepeat = 0
		cpnNumAvail = 30000000
		cpnCntry = 0
		cpnIsCoupon = 0
		cpnSitewide = 0
		cpnHandling = 0
		cpnInsurance = 0
		cpnLoginLevel = 0
		cpnLoginLt = FALSE
	end if
%>
		  <form name="mainform" method="post" action="admindiscounts.asp" onsubmit="return formvalidator(this)">
			<input type="hidden" name="posted" value="1" />
		<% if getpost("act")="modify" AND is_numeric(getpost("id")) then %>
			<input type="hidden" name="act" value="domodify" />
			<input type="hidden" name="id" value="<%=getpost("id")%>" />
		<% else %>
			<input type="hidden" name="act" value="doaddnew" />
		<% end if
			call writehiddenvar("scpds", getpost("scpds"))
			call writehiddenvar("sefct", getpost("sefct"))
			call writehiddenvar("stext", getpost("stext"))
			call writehiddenvar("sort", getpost("sort"))
			call writehiddenvar("stype", getpost("stype"))
			call writehiddenvar("pg", getpost("pg")) %>
            <table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><strong><%=yyDscNew%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><%=yyCpnDsc%>:</td>
				<td width="60%"><select name="cpnIsCoupon" size="1" onchange="couponcodeactive(false);">
					<option value="0"><%=yyDisco%></option>
					<option value="1" <% if Int(cpnIsCoupon)=1 then print "selected=""selected""" %>><%=yyCoupon%></option>
					</select></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><%=yyDscEff%>:</td>
				<td width="60%"><select name="cpnType" size="1" onchange="changecouponeffect(false);">
					<option value="0"><%=yyFrSShp%></option>
					<option value="1" <% if Int(cpnType)=1 then print "selected=""selected""" %>><%=yyFlatDs%></option>
					<option value="2" <% if Int(cpnType)=2 then print "selected=""selected""" %>><%=yyPerDis%></option>
					</select></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><%=yyDisTxt%>:</td>
				<td width="60%"><input type="text" name="cpnName" size="30" value="<%=htmlspecials(cpnName)%>" /></td>
			  </tr>
<%				for index=2 to adminlanguages+1
					cpnName=""
					if (adminlangsettings AND 1024)=1024 then
						if getpost("act")="modify" then
							if index=2 then cpnName=cpnName2
							if index=3 then cpnName=cpnName3
						end if
			%><tr>
				<td width="40%" align="right"><%=yyDisTxt & " " & index%>:</td>
				<td width="60%"><input type="text" name="cpnName<%=index%>" size="30" value="<%=htmlspecials(cpnName)%>" /></td>
			  </tr><%
					end if
				next %>
			  <tr>
				<td width="40%" align="right"><%=yyWrkNam%>:</td>
				<td width="60%"><input type="text" name="cpnWorkingName" size="30" value="<%=htmlspecials(cpnWorkingName)%>" /></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><%=yyCpnCod%>:</td>
				<td width="60%"><input type="text" name="cpnNumber" size="30" value="<%=cpnNumber%>" /></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><%=yyNumAvl%>:</td>
				<td width="60%"><input type="text" name="cpnNumAvail" size="10" value="<% if Int(cpnNumAvail)<>30000000 then print cpnNumAvail%>" /></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><input type="button" onclick="popUpCalendar(this,document.getElementById('cpnStartDate'),'<%=themask%>',0)" value="Start Date >=" /></td>
				<td width="60%"><div style="position:relative;display:inline"><input type="text" style="vertical-align:middle" id="cpnStartDate" name="cpnStartDate" size="10" value="<%
				print cpnStartDate %>" /> <input type="button" value="None" onclick="document.getElementById('cpnStartDate').value=''" /></div></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><input type="button" onclick="popUpCalendar(this,document.getElementById('cpnEndDate'),'<%=themask%>',0)" value="End Date <=" /></td>
				<td width="60%"><div style="position:relative;display:inline"><input type="text" style="vertical-align:middle" id="cpnEndDate" name="cpnEndDate" size="10" value="<%
				if cpnEndDate<>"" then
					if isexpired then print "Expired" else print cpnEndDate
				end if %>" /> <input type="button" value="None" onclick="document.getElementById('cpnEndDate').value=''" /> <input type="button" value="Expired" onclick="document.getElementById('cpnEndDate').value='Expired'" /></div></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><%=yyMinPur%>:</td>
				<td width="60%"><input type="text" name="cpnThreshold" size="10" value="<% if Int(cpnThreshold)>0 then print cpnThreshold%>" /> <%=yyRepEvy%>: <input type="text" name="cpnThresholdRepeat" size="10" value="<% if Int(cpnThresholdRepeat)>0 then print cpnThresholdRepeat%>" /></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><%=yyMaxPur%>:</td>
				<td width="60%"><input type="text" name="cpnThresholdMax" size="10" value="<% if Int(cpnThresholdMax)>0 then print cpnThresholdMax%>" /></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><%=yyMinQua%>:</td>
				<td width="60%"><input type="text" name="cpnQuantity" size="10" value="<% if Int(cpnQuantity)>0 then print cpnQuantity%>" /> <%=yyRepEvy%>: <input type="text" name="cpnQuantityRepeat" size="10" value="<% if Int(cpnQuantityRepeat)>0 then print cpnQuantityRepeat%>" /></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><%=yyMaxQua%>:</td>
				<td width="60%"><input type="text" name="cpnQuantityMax" size="10" value="<% if Int(cpnQuantityMax)>0 then print cpnQuantityMax%>" /></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><%=yyDscAmt%>:</td>
				<td width="60%"><input type="text" name="cpnDiscount" size="10" value="<%=cpnDiscount%>" /></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><%=yyScope%>:</td>
				<td width="60%"><select name="cpnSitewide" size="1">
					<option value="0"><%=yyIndCat%></option>
					<option value="3" <% if Int(cpnSitewide)=3 then print "selected=""selected""" %>><%=yyDsCaTo%></option>
					<option value="2" <% if Int(cpnSitewide)=2 then print "selected=""selected""" %>><%=yyGlInPr%></option>
					<option value="1" <% if Int(cpnSitewide)=1 then print "selected=""selected""" %>><%=yyGlPrTo%></option>
					</select></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><%=yyAplHan%>:</td>
				<td width="60%"><select name="cpnHandling" size="1">
					<option value="0"><%=yyNo%></option>
					<option value="1"<% if int(cpnHandling)<>0 then print " selected=""selected""" %>><%=yyYes%></option>
					</select></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><%="Also Applies To Services"%>:</td>
				<td width="60%"><select name="cpnInsurance" size="1">
					<option value="0"><%=yyNo%></option>
					<option value="1"<% if int(cpnInsurance)<>0 then print " selected=""selected""" %>><%=yyYes%></option>
					</select></td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><%=yyLiLev%>:</td>
				<td width="60%">
					<select name="cpnLoginLt" size="1" onchange="setloglev(this)">
					<option value="0">&gt;=</option>
					<option value="1" <% if cpnLoginLt then print "selected=""selected""" %>>=</option>
					</select>
					<select name="cpnLoginLevel" id="cpnLoginLevel" size="1">
						<option value="0"><%=IIfVr(cpnLoginLt,yyLiLev & " 0",yyNoRes)%></option>
						<%	for index=1 to maxloginlevels
								print "<option value="""&index&""""
								if (cpnLoginLt AND cpnLoginLevel-1=index) OR (NOT cpnLoginLt AND cpnLoginLevel=index) then print " selected=""selected"""
								print ">" & yyLiLev & " " & index & "</option>"
							next%>
						<option value="127"<% if cpnLoginLevel=127 then print " selected=""selected"""%>><%=yyDisabl%></option>
					</select>
				</td>
			  </tr>
			  <tr>
				<td width="40%" align="right"><%=yyRestr%>:</td>
				<td width="60%"><select name="cpnCntry" size="1">
					<option value="0"><%=yyAppAll%></option>
					<option value="1" <% if Int(cpnCntry)=1 then print "selected=""selected""" %>><%=yyYesRes%></option>
					</select></td>
			  </tr>
			  <tr>
                <td width="100%" colspan="2" align="center"><br /><input type="submit" value="<%=yySubmit%>" /><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="2" align="center"><br />
                          <a href="admin.asp"><%=yyAdmHom%></a><br />
                          &nbsp;</td>
			  </tr>
            </table>
		  </form>
<script>
<!--
savebg=document.mainform.cpnNumber.style.backgroundColor;
couponcodeactive(false);
changecouponeffect(false);
//-->
</script>
<%
elseif getpost("posted")="1" AND success then
	call adminsuccessforward("admindiscounts.asp",forwardquery)
elseif getpost("posted")="1" then
	call adminfailback(errmsg)
else
	sortorder=request("sort")
	cract=getget("qe")
	modclone=request.cookies("modclone") %>
<script>
<!--
function setCookie(c_name,value,expiredays){
	var exdate=new Date();
	exdate.setDate(exdate.getDate()+expiredays);
	document.cookie=c_name+ "=" +escape(value)+((expiredays==null) ? "" : ";expires="+exdate.toGMTString());
}
function mr(id){
	document.mainform.id.value = id;
	document.mainform.act.value = "modify";
	document.mainform.submit();
}
function cr(id){
	document.mainform.id.value = id;
	document.mainform.act.value = "clone";
	document.mainform.submit();
}
function newrec(id){
	document.mainform.id.value = id;
	document.mainform.act.value = "addnew";
	document.mainform.submit();
}
function dr(id){
if(confirm("<%=jscheck(yyConDel)%>\n")) {
	document.mainform.id.value = id;
	document.mainform.act.value = "delete";
	document.mainform.submit();
}
}
function startsearch(){
	document.mainform.action="admindiscounts.asp";
	document.mainform.act.value="search";
	document.mainform.posted.value="";
	document.mainform.submit();
}
function changemodclone(modclone){
	setCookie('modclone',modclone[modclone.selectedIndex].value,600);
	startsearch();
}
function quickupdate(){
	if(document.mainform.cract.value=="del"){
		if(!confirm("<%=jscheck(yyConDel)%>\n"))
			return;
	}
	document.mainform.action='admindiscounts.asp';
	document.mainform.act.value="quickupdate";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function changecract(obj){
	document.mainform.action='admindiscounts.asp?qe='+obj[obj.selectedIndex].value;
	document.mainform.act.value="search";
	document.mainform.posted.value="";
	document.mainform.submit();
}
function checkboxes(docheck){
	maxitems=document.getElementById("resultcounter").value;
	for(index=0;index<maxitems;index++){
		var thiselem=document.getElementById("chkbx"+index);
		if(thiselem.checked!=docheck){
			thiselem.checked=docheck;
			thiselem.onchange();
		}
	}
}
// -->
</script>
<h2><%=yyAdmCoD%></h2>
	<form name="mainform" method="post" action="admindiscounts.asp">
	<input type="hidden" name="posted" value="1" />
	<input type="hidden" name="act" value="xxxxx" />
	<input type="hidden" name="id" value="xxxxx" />
	<input type="hidden" name="pg" value="<%=IIfVr(getpost("act")="search", "1", getget("pg"))%>" />
	<input type="hidden" name="selectedq" value="1" />
	<input type="hidden" name="newval" value="1" />
	<table class="cobtbl" width="100%" border="0" cellspacing="1" cellpadding="3">
	  <tr height="30"> 
		<td class="cobhl" width="25%" align="right"><%=yySrchFr%>:</td>
		<td class="cobll" width="25%"><input type="text" name="stext" size="20" value="<%=request("stext")%>" /></td>
		<td class="cobhl" width="25%" align="right"><%=yySrchTp%>:</td>
		<td class="cobll" width="25%"><select name="stype" size="1">
			<option value=""><%=yySrchAl%></option>
			<option value="any"<% if request("stype")="any" then print " selected=""selected"""%>><%=yySrchAn%></option>
			<option value="exact"<% if request("stype")="exact" then print " selected=""selected"""%>><%=yySrchEx%></option>
			</select>
		</td>
	  </tr>
	  <tr height="30"> 
		<td class="cobhl" width="25%" align="right"><%=yyCpnDsc%>:</td>
		<td class="cobll" width="25%"><select name="scpds" size="1">
			<option value=""><%=yyAll%></option>
			<option value="cpn"<% if request("scpds")="cpn" then print " selected=""selected"""%>><%=yyCoupon%></option>
			<option value="dsc"<% if request("scpds")="dsc" then print " selected=""selected"""%>><%=yyDisco%></option>
			</select>
		</td>
		<td class="cobhl" width="25%" align="right"><%=yyDscEff%>:</td>
		<td class="cobll" width="25%"><select name="sefct" size="1">
			<option value=""><%=yyAll%></option>
			<option value="frshp"<% if request("sefct")="frshp" then print " selected=""selected"""%>><%=yyFrSShp%></option>
			<option value="fltra"<% if request("sefct")="fltra" then print " selected=""selected"""%>><%=yyFlatDs%></option>
			<option value="percd"<% if request("sefct")="percd" then print " selected=""selected"""%>><%=yyPerDis%></option>
			</select>
		</td>
	  </tr>
	  <tr height="30">
		<td class="cobhl" align="center"><%
			if cract="del" then %>
					<input type="button" value="<%=yyCheckA%>" onclick="checkboxes(true)" /> <input type="button" value="<%=yyUCheck%>" onclick="checkboxes(false)" />
<%			end if
		%></td>
		<td class="cobll" colspan="3" align="center">
				<select name="sort" size="1" style="vertical-align: middle">
				<option value="">Sort - <%=yyDisTxt%></option>
				<option value="dam"<% if sortorder="dam" then print " selected=""selected"""%>>Sort - <%=yyDscAmt%></option>
				<option value="dex"<% if sortorder="dex" then print " selected=""selected"""%>>Sort - <%=yyExpDat%></option>
				</select>
				<input type="submit" value="List Discounts" onclick="startsearch();" />
				<input type="button" value="<%=yyNewDsc%>" onclick="newrec()" />
	  </tr>
	</table>
<br />
            <table width="100%" class="stackable admin-table-a sta-white">
<%	sub displayheaderrow() %>
			  <tr>
			  	<th class="small minicell">
					<select name="cract" id="cract" size="1" onchange="changecract(this)" style="width:150px">
					<option value="none">Quick Entry...</option>
					<option value="" disabled="disabled">---------------------</option>
					<option value="del"<% if cract="del" then print " selected=""selected"""%>><%=yyDelete%></option>
					</select>
				</th>
				<th class="maincell"><%=yyWrkNam%></th>
				<th class="minicell"><%=yyType%></th>
				<th class="minicell"><%=yyExpDat%></th>
				<th class="minicell"><%=yyGlobal%></th>
				<th class="minicell"><%=yyModify%></th>
			  </tr>
<%	end sub
	if getpost("act")="search" OR getget("pg")<>"" then
		Session.LCID = saveLCID
		sSQL = "SELECT cpnID,cpnWorkingName,cpnSitewide,cpnIsCoupon,cpnEndDate,cpnNumAvail FROM coupons"
		whereand=" WHERE "
		if trim(request("scpds"))<>"" then
			if request("scpds")="cpn" then sSQL=sSQL & whereand & "cpnIsCoupon<>0" else sSQL=sSQL & whereand & "cpnIsCoupon=0"
			whereand=" AND "
		end if
		if trim(request("sefct"))<>"" then
			if request("sefct")="frshp" then
				sSQL=sSQL & whereand & "cpnType=0"
			elseif request("sefct")="fltra" then
				sSQL=sSQL & whereand & "cpnType=1"
			else
				sSQL=sSQL & whereand & "cpnType=2"
			end if
			whereand=" AND "
		end if
		if trim(request("stext"))<>"" then
			hassearch=TRUE
			Xstext = escape_string(request("stext"))
			aText = Split(Xstext)
			maxsearchindex=1
			aFields(0)="cpnName"
			aFields(1)="cpnWorkingName"
			if request("stype")="exact" then
				sSQL=sSQL & whereand & "(cpnName LIKE '%"&Xstext&"%' OR cpnWorkingName LIKE '%"&Xstext&"%') "
				whereand=" AND "
			else
				sJoin="AND "
				if request("stype")="any" then sJoin="OR "
				sSQL=sSQL & whereand&"("
				whereand=" AND "
				for index=0 to maxsearchindex
					sSQL=sSQL & "("
					for rowcounter=0 to UBOUND(aText)
						sSQL=sSQL & aFields(index) & " LIKE '%"&aText(rowcounter)&"%' "
						if rowcounter<UBOUND(aText) then sSQL=sSQL & sJoin
					next
					sSQL=sSQL & ") "
					if index < maxsearchindex then sSQL=sSQL & "OR "
				next
				sSQL=sSQL & ") "
			end if
		end if
		if sortorder="dam" then
			sSQL=sSQL & " ORDER BY cpnDiscount,cpnWorkingName"
		elseif sortorder="dex" then
			sSQL=sSQL & " ORDER BY cpnEndDate,cpnWorkingName"
		else
			sSQL=sSQL & " ORDER BY cpnWorkingName"
		end if
		if admindiscountsperpage="" then admindiscountsperpage=600
		rs.CursorLocation = 3 ' adUseClient
		rs.CacheSize = admindiscountsperpage
		rs.open sSQL, cnn
		if NOT rs.eof then
			Count=0
			rs.MoveFirst
			rs.PageSize = admindiscountsperpage
			CurPage = 1
			if is_numeric(getget("pg")) then CurPage=int(getget("pg"))
			iNumOfPages = int((rs.RecordCount + (admindiscountsperpage-1)) / admindiscountsperpage)
			rs.AbsolutePage = CurPage
			pblink = "<a href=""admindiscounts.asp?stext="&urlencode(request("stext"))&"&stype="&request("stype")&"&scpds="&request("scpds")&"&sefct="&request("sefct")&"&sort="&sortorder&"&pg="
			if iNumOfPages > 1 then print "<tr><td colspan=""5"" align=""center"">" & writepagebar(CurPage,iNumOfPages,yyPrev,yyNext,pblink,FALSE) & "<br />&nbsp;</td></tr>"
			call displayheaderrow()
			do while NOT rs.EOF AND Count < rs.PageSize
				jscript=jscript&"pa["&Count&"]=[" %>
				  <tr id="tr<%=Count%>">
					<td class="minicell"><%
				qetype="text"
				qesize="18"
				if cract="del" then
					jscript=jscript&"'del'"
					qetype="delbox"
				else
					qetype=""
				end if %></td>
					<td class="maincell"><%=rs("cpnWorkingName")%></td>
					<td class="minicell"><%	if cint(rs("cpnIsCoupon"))=1 then print yyCoupon else print yyDisco%></td>
					<td class="minicell"><%	if rs("cpnEndDate")=DateSerial(3000,1,1) then
												print yyNever
											elseif rs("cpnEndDate")-Date() < 0 then
												print "<span style=""color:#FF0000"">"&yyExpird&"</span>"
											else
												print rs("cpnEndDate")
											end if
											if rs("cpnNumAvail")<=0 then print " / <span style=""color:#FF0000""> 0 Available</span>" %></td>
<td class="minicell"><% if cint(rs("cpnSitewide"))=1 OR cint(rs("cpnSitewide"))=2 then print yyYes else print yyNo %></td><td>-</td></tr>
	<%			jscript=jscript&","&rs("cpnID")&"];"&vbCrLf
				rs.MoveNext
				Count=Count+1
			loop
			if iNumOfPages > 1 then print "<tr><td colspan=""5"" align=""center""><br />" & writepagebar(CurPage,iNumOfPages,yyPrev,yyNext,pblink,FALSE) & "<br />&nbsp;</td></tr>"
		else %>
			  <tr> 
				<td width="100%" colspan="5" align="center"><br /><%=yyNoDsc%><br />&nbsp;</td>
			  </tr>
<%		end if
		rs.close
	else
		numitems=0
		sSQL="SELECT COUNT(*) as totcount FROM coupons"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			numitems=rs("totcount")
		end if
		rs.close
		print "<tr><td colspan=""5""><div class=""itemsdefine"">You have " & numitems & " discounts defined.</div></td></tr>"
	end if
%>
			  <tr>
				<td align="center" style="white-space:nowrap"><% if Count>0 AND cract<>"" AND cract<>"none" then print "<input type=""hidden"" name=""resultcounter"" id=""resultcounter"" value=""" & Count & """ /><input type=""button"" value=""" & yyUpdate & """ onclick=""quickupdate()"" /> <input type=""reset"" value=""" & yyReset & """ />" else print "&nbsp;"%></td>
                <td width="100%" colspan="5" align="center"><br /><a href="admin.asp"><%=yyAdmHom%></a><br />&nbsp;</td>
			  </tr>
            </table>
	</form>
<script>
/* <![CDATA[ */
var pa=[];
<%=jscript%>
function patch_pid(pid){
	document.getElementById('pid'+pid).name='pid'+pid;
	document.getElementById('pid'+pid).value=pa[pid][1];
	return pid;
}
for(var pidind in pa){
	var ttr=document.getElementById('tr'+pidind);
	ttr.cells[1].innerHTML+='<input type="hidden" id="pid'+pidind+'" value="" />';
	ttr.cells[5].style.textAlign='center';
	ttr.cells[5].style.whiteSpace='nowrap';
	ttr.cells[5].innerHTML='<input type="button" value="M" style="width:30px;margin-right:4px" onclick="mr(\''+pa[pidind][1]+'\')" title="<%=jsescape(htmlspecials(yyModify))%>" />' +
		'<input type="button" value="C" style="width:30px;margin-right:4px" onclick="cr(\''+pa[pidind][1]+'\')" title="<%=jsescape(htmlspecials(yyClone))%>" />' +
		'<input type="button" value="X" style="width:30px" onclick="dr(\''+pa[pidind][1]+'\')" title="<%=jsescape(htmlspecials(yyDelete))%>" />';
<%		if qetype="text" then %>
	ttr.cells[0].innerHTML=pa[pidind][0]===false?'-':'<input type="text" id="chkbx'+pidind+'" size="<% print qesize%>" onchange="this.name=\'pra_'+patch_pid(pidind)+'\'" value="'+pa[pidind][0].replace('"','&quot;')+'" tabindex="'+(pidind+1)+'" />';
<%		elseif qetype="delbox" then %>
	ttr.cells[0].innerHTML='<input type="checkbox" id="chkbx'+pidind+'" onchange="this.name=\'pra_'+patch_pid(pidind)+'\'" value="del" tabindex="'+(pidind+1)+'" />';
<%		elseif qetype="checkbox" then %>
	ttr.cells[0].innerHTML='<input type="hidden" id="pra_'+pa[pidind][1]+'" value="1" /><input type="checkbox" id="chkbx'+pidind+'" onchange="this.name=\'prb_'+patch_pid(pidind)+'\';document.getElementById(\'pra_'+pa[pidind][1]+'\').name=\'pra_'+patch_pid(pidind)+'\'" value="1" '+(pa[pidind][0]==1?'checked="checked" ':'')+'tabindex="'+(pidind+1)+'" />';
<%		elseif qetype="image" then %>
	ttr.cells[0].innerHTML=(pa[pidind][0]==''?'-':'<img class="lazyload" id="lazyimg'+pidind+'" src="adminimages/imageload.png" data-src="'+pa[pidind][0]+'" style="max-width:80px;cursor:pointer" alt="" onclick="mr(\''+pa[pidind][1]+'\')" />');
<%		end if %>
}
/* ]]> */
</script>
<%
end if
cnn.Close
set rs = nothing
set cnn = nothing
%>