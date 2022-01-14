<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
Dim sSQL,rs,alldata,success,cnn,rowcounter,netnav,errmsg,aOption,index,iID,bOption,fieldDims,aFields(3)
success=true
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs2=Server.CreateObject("ADODB.RecordSet")
Set rs3=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
Session.LCID=1033
alldata=""
errmsg=""
resultcounter=0
dorefresh=FALSE
if htmlemails=TRUE then emlNl="<br />"&vbCrLf else emlNl=vbCrLf
function dodeleteoption(oid)
	index=0
	sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 100 ")&"poID,poProdID FROM prodoptions INNER JOIN products ON prodoptions.poProdID=products.pID WHERE poOptionGroup=" & oid & IIfVs(mysqlserver=TRUE," LIMIT 0,100")
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		success=false
		errmsg=errmsg & yyPOUse & "<br /><br />"
		do while NOT rs.EOF
			errmsg=errmsg & "<form method=""post"" action=""adminprods.asp"" style=""display:inline"">" & whv("posted",1) & whv("act","modify") & "<input type=""submit"" name=""id"" value=""" & rs("poProdID") & """></form>"
			index=index+1
			if index>=10 then print "<br />" : index=0
			rs.movenext
		loop
	end if
	rs.close
	index=0
	showmessage=TRUE
	sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 100 ")&"optGroup,optName,optDependants FROM options WHERE optDependants LIKE '%" & oid & "%'" & IIfVs(mysqlserver=TRUE," LIMIT 0,100")
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		if instr(","&rs("optDependants")&",",","&oid&",")>0 then
			if showmessage then errmsg=errmsg & "<br /><br />This option is a dependent option of the following options:<br /><br />"
			showmessage=FALSE
			errmsg=errmsg & "<form method=""post"" action=""adminprodopts.asp"" style=""display:inline"">" & whv("posted",1) & whv("act","modify") & whv("id",rs("optGroup")) & "<input type=""submit"" value=""" & htmlspecials(rs("optName")) & """></form>"
			index=index+1
			if index>=10 then print "<br />" : index=0
			success=FALSE
		end if
		rs.movenext
	loop
	rs.close
	if success then
		ect_query("DELETE FROM options WHERE optGroup=" & oid)
		ect_query("DELETE FROM optiongroup WHERE optGrpID=" & oid)
		ect_query("DELETE FROM prodoptions WHERE poOptionGroup=" & oid)
	end if
	dodeleteoption=success
end function
sub checknotifystock(theoid)
	if useStockManagement AND notifybackinstock then
		sSQL="SELECT "&getlangid("notifystocksubject",4096)&","&getlangid("notifystockemail",4096)&" FROM emailmessages WHERE emailID=1"
		rs.open sSQL,cnn,0,1
		oemailsubject=trim(rs(getlangid("notifystocksubject",4096))&"")
		oemailmessage=rs(getlangid("notifystockemail",4096))&""
		rs.close
		
		idlist=""
		if mysqlserver then
			sSQL="SELECT DISTINCT nsProdID FROM notifyinstock INNER JOIN prodoptions ON notifyinstock.nsProdID=prodoptions.poProdID INNER JOIN options ON prodoptions.poOptionGroup=options.optGroup WHERE nsOptID=-1 AND optID="&theoid
		else
			sSQL="SELECT DISTINCT nsProdID FROM notifyinstock INNER JOIN (prodoptions INNER JOIN options ON prodoptions.poOptionGroup=options.optGroup) ON notifyinstock.nsProdID=prodoptions.poProdID WHERE nsOptID=-1 AND optID="&theoid
		end if
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			gotall=TRUE
			sSQL="SELECT poOptionGroup FROM prodoptions INNER JOIN optiongroup ON prodoptions.poOptionGroup=optiongroup.optGrpID WHERE poProdID='"&escape_string(rs("nsProdID"))&"'"
			rs2.Open sSQL,cnn,0,1
			do while NOT rs2.EOF
				sSQL="SELECT optID FROM options WHERE optStock>0 AND optGroup="&rs2("poOptionGroup")
				rs3.Open sSQL,cnn,0,1
				if rs3.EOF then gotall=FALSE
				rs3.Close
				rs2.movenext
			loop
			rs2.Close
			if gotall then idlist=idlist&"'"&escape_string(rs("nsProdID"))&"',"
			rs.movenext
		loop
		rs.close
		if idlist<>"" then idlist=left(idlist,len(idlist)-1)

		pStockByOpts=0
		sSQL="SELECT pId,pName,pStockByOpts,pStaticPage,pStaticURL,pInStock,nsEmail FROM products INNER JOIN notifyinstock ON products.pID=notifyinstock.nsProdID WHERE nsOptId="&theoid
		if idlist<>"" then sSQL=sSQL & " OR (nsOptID=-1 AND nsProdID IN ("&idlist&"))"
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			nspid=rs("pId")
			pName=trim(rs("pName"))
			pStockByOpts=rs("pStockByOpts")
			pStaticPage=rs("pStaticPage")
			pStaticURL=rs("pStaticURL")
			pInStock=rs("pInStock")
			theemail=rs("nsEmail")
			thelink=storeurl & getdetailsurl(nspid,pStaticPage,pName,trim(pStaticURL&""),"","")
			if htmlemails=TRUE AND thelink<>"" then thelink="<a href=""" & thelink & """>" & thelink & "</a>"
			emailsubject=replace(oemailsubject,"%pid%",trim(nspid))
			emailsubject=replace(emailsubject,"%pname%",pName)
			emailmessage=replace(oemailmessage,"%pid%",trim(nspid))
			emailmessage=replace(emailmessage,"%pname%",pName)
			emailmessage=replace(emailmessage,"%link%",thelink)
			emailmessage=replace(emailmessage,"%storeurl%",storeurl)
			emailmessage=replace(emailmessage, "<br />", emlNl)
			emailmessage=replace(emailmessage, "%nl%", emlNl)
			call DoSendEmailEO(theemail,emailAddr,"",emailsubject,emailmessage,emailObject,themailhost,theuser,thepass)
			rs.movenext
		loop
		rs.close
		sSQL="DELETE FROM notifyinstock WHERE nsOptId="&theoid
		if idlist<>"" then sSQL=sSQL & " OR (nsOptID=-1 AND nsProdID IN ("&idlist&"))"
		ect_query(sSQL)
	end if
end sub
if getpost("posted")="1" then
	if getpost("act")="delete" then
		if dodeleteoption(getpost("id")) then
			dorefresh=TRUE
		else
			errmsg=yyPOErr & "<br />" & errmsg
		end if
	elseif getpost("act")="quickupdate" then
		for each objItem in request.form
			if left(objItem, 4)="pra_" then
				theid=right(objItem, len(objItem)-4)
				theval=getpost(objItem)
				pract=getpost("pract")
				sSQL=""
				if pract="del" then
					if theval="del" then dodeleteoption(theid)
					sSQL=""
				elseif pract="own" then
					sSQL="UPDATE optiongroup SET optGrpWorkingName='" & escape_string(theval) & "'"
				elseif pract="oty" then
					ect_query("UPDATE optiongroup SET optType=" & escape_string(theval) & " WHERE optType>0 AND optGrpID="&theid)
					ect_query("UPDATE optiongroup SET optType=-" & escape_string(theval) & " WHERE optType<0 AND optGrpID="&theid)
				elseif pract="opn" then
					sSQL="UPDATE optiongroup SET optGrpName='" & escape_string(theval) & "'"
				elseif pract="opn2" then
					sSQL="UPDATE optiongroup SET optGrpName2='" & escape_string(theval) & "'"
				elseif pract="opn3" then
					sSQL="UPDATE optiongroup SET optGrpName3='" & escape_string(theval) & "'"
				end if
				if sSQL<>"" then
					sSQL=sSQL & " WHERE optGrpID="&theid
					ect_query(sSQL)
				end if
			end if
		next
		if success then dorefresh=TRUE else errmsg=yyPOErr & "<br />" & errmsg
	elseif getpost("act")="domodify" OR getpost("act")="doaddnew" then
		sSQL=""
		maxoptnumber=getpost("maxoptnumber")
		Redim aOption(12,maxoptnumber)
		bOption=false
		optFlags=0
		if getpost("pricepercent")="1" then optFlags=1
		if getpost("weightpercent")="1" then optFlags=optFlags + 2
		if getpost("singleline")="1" then optFlags=optFlags + 4
		if getpost("optdefault")<>"" then optDefault=int(getpost("optdefault")) else optDefault=-1
		for rowcounter=0 to maxoptnumber-1
			if getpost("opt"&rowcounter)<>"" then bOption=true
			aOption(0,rowcounter)=escape_string(getpost("opt"&rowcounter))
			for index=2 to adminlanguages+1
				if (adminlangsettings AND 32)=32 then aOption(9+index,rowcounter)=escape_string(getpost("opl"&index&"x"&rowcounter))
			next
			if is_numeric(getpost("pri"&rowcounter)) then
				aOption(1,rowcounter)=getpost("pri"&rowcounter)
			else
				aOption(1,rowcounter)=0
			end if
			if is_numeric(getpost("wsp"&rowcounter)) then
				aOption(4,rowcounter)=getpost("wsp"&rowcounter)
			else
				aOption(4,rowcounter)=0
			end if
			if is_numeric(getpost("wei"&rowcounter)) then
				aOption(2,rowcounter)=getpost("wei"&rowcounter)
			else
				aOption(2,rowcounter)=0
			end if
			if is_numeric(getpost("optStock"&rowcounter)) then
				aOption(3,rowcounter)=getpost("optStock"&rowcounter)
			else
				aOption(3,rowcounter)=0
			end if
			aOption(5,rowcounter)=escape_string(getpost("regexp"&rowcounter))
			aOption(6,rowcounter)=getpost("orig"&rowcounter)
			aOption(7,rowcounter)=escape_string(getpost("altimg"&rowcounter))
			aOption(8,rowcounter)=escape_string(getpost("altlimg"&rowcounter))
			aOption(9,rowcounter)=""
			depotpnum=1
			do while getpost("depopts"&rowcounter&"_"&depotpnum)<>""
				if is_numeric(getpost("depopts"&rowcounter&"_"&depotpnum)) then aOption(9,rowcounter)=aOption(9,rowcounter)&getpost("depopts"&rowcounter&"_"&depotpnum)&","
				depotpnum=depotpnum+1
			loop
			if aOption(9,rowcounter)<>"" then aOption(9,rowcounter)=left(aOption(9,rowcounter),len(aOption(9,rowcounter))-1)
			aOption(10,rowcounter)=escape_string(getpost("cls"&rowcounter))
		next
		if (getpost("secname")="" OR NOT bOption) AND getpost("optType")<>"3" AND getpost("optType")<>"5" then
			success=false
			errmsg=yyPOErr & "<br />"
			errmsg=errmsg & yyPOOne
		else
			if getpost("optType")="3" OR getpost("optType")="5" then ' Text / Date Picker option
				fieldDims=getpost("pri0")&"."
				if int(getpost("fieldheight")) < 10 then fieldDims=fieldDims & "0"
				fieldDims=fieldDims & getpost("fieldheight")
				optTxtCharge=getpost("optTxtCharge")
				if NOT is_numeric(optTxtCharge) OR trim(optTxtCharge)="" then optTxtCharge=0
				if getpost("act")="doaddnew" then
					rs.open "optiongroup",cnn,1,3,&H0002
					rs.AddNew
					rs.Fields("optGrpName")	= getpost("secname")
					for index=2 to adminlanguages+1
						if (adminlangsettings AND 16)=16 then rs.Fields("optGrpName" & index)=getpost("secname" & index)
					next
					rs.Fields("optType")=IIfVr(getpost("forceselec")="ON",getpost("optType"),0-int(getpost("optType")))
					rs.Fields("optTxtMaxLen")=getpost("optTxtMaxLen")
					rs.Fields("optTxtCharge")=IIfVr(getpost("iscostperentry")="1", 0-optTxtCharge, optTxtCharge)
					rs.Fields("optMultiply")=IIfVr(getpost("optMultiply")="ON", 1, 0)
					rs.Fields("optAcceptChars")=getpost("optAcceptChars")
					rs.Fields("optGrpWorkingName")=IIfVr(getpost("workingname")="",getpost("secname"),getpost("workingname"))
					rs.Fields("optFlags")=optFlags
					rs.Fields("optTooltip")=getpost("opttooltip")
					rs.Update
					if mysqlserver=true then
						rs.close
						rs.open "SELECT LAST_INSERT_ID() AS lstIns",cnn,0,1
						iID=rs("lstIns")
					else
						iID =rs.Fields("optGrpID")
					end if
					rs.close
					sSQL="INSERT INTO options (optGroup,optName,optPlaceholder,optPriceDiff"
					for index=2 to adminlanguages+1
						if (adminlangsettings AND 16)=16 then
							sSQL=sSQL & ",optName" & index
							sSQL=sSQL & ",optPlaceholder" & index
						end if
					next
					sSQL=sSQL & ",optWeightDiff,optClass) VALUES ("&iID&",'"&escape_string(getpost("opt0"))&"','"&escape_string(getpost("oph0"))&"',"&fieldDims
					for index=2 to adminlanguages+1
						if (adminlangsettings AND 16)=16 then
							sSQL=sSQL & ",'" & escape_string(getpost("opl" & index & "x0"))&"'"
							sSQL=sSQL & ",'" & escape_string(getpost("oph" & index & ""))&"'"
						end if
					next
					sSQL=sSQL & ",0,'"&escape_string(getpost("cls"))&"')"
					ect_query(sSQL)
				else
					iID=getpost("id")
					sSQL="UPDATE optiongroup SET optGrpName='"&escape_string(getpost("secname"))&"'"
					for index=2 to adminlanguages+1
						if (adminlangsettings AND 16)=16 then sSQL=sSQL & ",optGrpName" & index & "='"& escape_string(getpost("secname" & index))&"'"
					next
					sSQL=sSQL & ",optType=" & IIfVr(getpost("forceselec")="ON",getpost("optType"),0-int(getpost("optType")))
					sSQL=sSQL & ",optTxtMaxLen=" & getpost("optTxtMaxLen")
					sSQL=sSQL & ",optTxtCharge=" & IIfVr(getpost("iscostperentry")="1", 0-optTxtCharge, optTxtCharge)
					sSQL=sSQL & ",optMultiply=" & IIfVr(getpost("optMultiply")="ON", 1, 0)
					sSQL=sSQL & ",optAcceptChars='" & escape_string(getpost("optAcceptChars")) & "'"
					sSQL=sSQL & ",optFlags=" & optFlags
					sSQL=sSQL & ",optGrpWorkingName='"& escape_string(IIfVr(getpost("workingname")="",getpost("secname"),getpost("workingname")))&"'"
					sSQL=sSQL & ",optTooltip='" & escape_string(getpost("opttooltip")) & "' "
					sSQL=sSQL & "WHERE optGrpID="&iID
					ect_query(sSQL)
					sSQL="UPDATE options SET optName='"&escape_string(getpost("opt0"))&"',optPlaceholder='"&escape_string(getpost("oph0"))&"',optPriceDiff="&fieldDims
					for index=2 to adminlanguages+1
						if (adminlangsettings AND 16)=16 then sSQL=sSQL & ",optName" & index & "='"& escape_string(getpost("opl" & index & "x0"))&"',optPlaceholder" & index & "='"& escape_string(getpost("oph" & index))&"'"
					next
					sSQL=sSQL & ",optClass='" & escape_string(getpost("cls")) & "' "
					sSQL=sSQL & " WHERE optGroup="&iID
					ect_query(sSQL)
				end if
			else ' Non-text Option
				if getpost("act")="doaddnew" then
					rs.open "optiongroup",cnn,1,3,&H0002
					rs.AddNew
					rs.Fields("optGrpName")	= getpost("secname")
					for index=2 to adminlanguages+1
						if (adminlangsettings AND 16)=16 then rs.Fields("optGrpName" & index)=getpost("secname" & index)
					next
					rs.Fields("optType")=IIfVr(getpost("forceselec")="ON",getpost("optType"),0-int(getpost("optType")))
					rs.Fields("optGrpWorkingName")=IIfVr(getpost("workingname")="",getpost("secname"),getpost("workingname"))
					rs.Fields("optFlags")=optFlags
					rs.Fields("optGrpSelect")=IIfVr(getpost("optgrpselect")="1",1,0)
					rs.Fields("optTooltip")=getpost("opttooltip")
					rs.Update
					if mysqlserver=true then
						rs.close
						rs.open "SELECT LAST_INSERT_ID() AS lstIns",cnn,0,1
						iID=rs("lstIns")
					else
						iID =rs.Fields("optGrpID")
					end if
					rs.close
				else
					iID=getpost("id")
					sSQL="UPDATE optiongroup SET optGrpName='"& escape_string(getpost("secname"))&"'"
					for index=2 to adminlanguages+1
						if (adminlangsettings AND 16)=16 then sSQL=sSQL & ",optGrpName" & index & "='"& escape_string(getpost("secname" & index))&"'"
					next
					sSQL=sSQL & ",optType=" & IIfVr(getpost("forceselec")="ON",getpost("optType"),0-int(getpost("optType")))
					sSQL=sSQL & ",optFlags=" & optFlags
					sSQL=sSQL & ",optGrpSelect=" & IIfVr(getpost("optgrpselect")="1",1,0)
					sSQL=sSQL & ",optGrpWorkingName='" & escape_string(IIfVr(getpost("workingname")="",getpost("secname"),getpost("workingname"))) & "'"
					sSQL=sSQL & ",optTooltip='" & escape_string(getpost("opttooltip")) & "' "
					sSQL=sSQL & "WHERE optGrpID="&iID
					ect_query(sSQL)
				end if
				for rowcounter=0 to UBOUND(aOption,2)
					if trim(aOption(0,rowcounter))<>"" then
						if aOption(6,rowcounter)<>"" then
							sSQL="UPDATE options SET optClass='"&aOption(10,rowcounter)&"',optName='"&aOption(0,rowcounter)&"',optRegExp='"&aOption(5,rowcounter)&"',optAltImage='"&aOption(7,rowcounter)&"',optAltLargeImage='"&aOption(8,rowcounter)&"',optPriceDiff="&aOption(1,rowcounter)&",optWeightDiff="&aOption(2,rowcounter)&",optStock="&aOption(3,rowcounter)
							if wholesaleoptionpricediff=TRUE then sSQL=sSQL & ",optWholesalePriceDiff="&aOption(4,rowcounter)
							for index=2 to adminlanguages+1
								if (adminlangsettings AND 32)=32 then sSQL=sSQL & ",optName" & index & "='" & aOption(9+index,rowcounter) & "'"
							next
							sSQL=sSQL & ",optDefault=" & IIfVr(rowcounter=optDefault,"1","0")
							sSQL=sSQL & ",optDependants='" & aOption(9,rowcounter) & "'"
							sSQL=sSQL & " WHERE optID=" & aOption(6,rowcounter)
							ect_query(sSQL)
							if aOption(3,rowcounter)>0 then
								call checknotifystock(aOption(6,rowcounter))
							end if
						else
							sSQL="INSERT INTO options (optGroup,optClass,optName,optRegExp,optAltImage,optAltLargeImage,optPriceDiff,optWeightDiff,optStock,optDefault,optDependants"
							if wholesaleoptionpricediff=TRUE then sSQL=sSQL & ",optWholesalePriceDiff"
							for index=2 to adminlanguages+1
								if (adminlangsettings AND 32)=32 then sSQL=sSQL & ",optName" & index
							next
							sSQL=sSQL & ") VALUES ("&iID&",'"&aOption(10,rowcounter)&"','"&aOption(0,rowcounter)&"','"&aOption(5,rowcounter)&"','"&aOption(7,rowcounter)&"','"&aOption(8,rowcounter)&"',"&aOption(1,rowcounter)&","&aOption(2,rowcounter)&","&aOption(3,rowcounter)&","&IIfVr(rowcounter=optDefault,"1","0")&",'"&aOption(9,rowcounter)&"'"
							if wholesaleoptionpricediff=TRUE then sSQL=sSQL & "," & aOption(4,rowcounter)
							for index=2 to adminlanguages+1
								if (adminlangsettings AND 32)=32 then sSQL=sSQL & ",'" & aOption(9+index,rowcounter) & "'"
							next
							sSQL=sSQL & ")"
							ect_query(sSQL)
						end if
					else
						if aOption(6,rowcounter)<>"" then
							ect_query("DELETE FROM options WHERE optID=" & aOption(6,rowcounter))
						end if
					end if
				next
			end if
		end if
		dorefresh=TRUE
	end if
end if
if dorefresh then
	print "<meta http-equiv=""refresh"" content=""1; url=adminprodopts.asp"
	print "?disp=" & getpost("disp") & "&stext=" & urlencode(getpost("stext")) & "&stype=" & getpost("stype") & "&pg=1" ' & getpost("pg")
	print """>"
end if
%>
<script>
/* <![CDATA[ */
var oAR=new Array();
<%	sSQL="SELECT optGrpID,optGrpWorkingName,optType FROM optiongroup ORDER BY optGrpWorkingName"
	rs.open sSQL,cnn,0,1
	rowcounter=0
	do while NOT rs.EOF
		print "oAR["&rowcounter&"]=["&rs("optGrpID")&",'"&jsescape(rs("optGrpWorkingName"))&"',"&rs("optType")&"];"&vbCrLf
		rowcounter=rowcounter+1
		rs.movenext
	loop
	rs.close
%>
function addoptionselect(oSelect){
	oSelect.name=oSelect.id;
	var spanid=oSelect.id.split('_')[0];
	var optnum=parseInt(oSelect.id.split('_')[1]);	

	var select=document.createElement("select");
	select.setAttribute("id",spanid+'_'+(optnum+1));
	select.style.width="140px";
	select.onchange=function(){addoptionselect(this)};
	select.onmouseover=function(){populateoptionsselect(this)};
	var option;
	option=document.createElement("option");
	option.setAttribute("value","x");
	option.innerHTML="<%=jscheck(yySelect)%>";
	select.appendChild(option);

	document.getElementById(spanid).appendChild(select);
	oSelect.onchange='';
}
function populateoptionsselect(oSelect){
	var insbefore=oSelect.selectedIndex!=0;
	var existingitem=oSelect.options[oSelect.selectedIndex];
	var osarray;
	osarray=oAR;
	for(var i=0;i<osarray.length;i++){
		if(existingitem.value==osarray[i][0]){
			insbefore=false;
		}else{
			var y=document.createElement('option');
			y.innerHTML=osarray[i][1];
			y.value=osarray[i][0];
			if(insbefore){
				try{oSelect.add(y,existingitem);} // FF etc
				catch(ex){oSelect.add(y,oSelect.selectedIndex);} // IE
			}else{
				try{oSelect.add(y,null);} // FF etc
				catch(ex){oSelect.add(y);} // IE
			}
		}
	}
	oSelect.onmouseover='';
}
function formvalidator(theForm){
	var maxrow=document.getElementById("maxoptnumber").value;
	if(theForm.secname.value == ""){
		alert("<%=jscheck(yyPlsEntr&" """&yyPOName)%>\".");
		theForm.secname.focus();
		return (false);
	}
	for(index=0;index<maxrow;index++){
		document.getElementById("altimg" + index).disabled=(document.getElementById("altimg" + index).name=='xxx'?true:false);
		document.getElementById("altlimg" + index).disabled=(document.getElementById("altlimg" + index).name=='xxx'?true:false);
<%		if useStockManagement then print "document.getElementById('optStock' + index).disabled=(document.getElementById('optStock' + index).name=='xxx'?true:false);"
		if wholesaleoptionpricediff=TRUE then print "document.getElementById('wsp' + index).disabled=(document.getElementById('wsp' + index).name=='xxx'?true:false);" %>
		document.getElementById("pri" + index).disabled=(document.getElementById("pri" + index).name=='xxx'?true:false);
		document.getElementById("regexp" + index).disabled=(document.getElementById("regexp" + index).name=='xxx'?true:false);
		document.getElementById("wei" + index).disabled=(document.getElementById("wei" + index).name=='xxx'?true:false);
	}
	return (true);
}
function changeunits(){
	var maxrow=document.getElementById("maxoptnumber").value;
	for(index=0;index<maxrow;index++){
		wel=document.getElementById("wunitspan" + index);
		pel=document.getElementById("punitspan" + index);
		if(document.forms.mainform.weightpercent.checked)
			wel.style.display='';
		else
			wel.style.display='none';
		if(document.forms.mainform.pricepercent.checked)
			pel.style.display='';
		else
			pel.style.display='none';
	}
}
function doswitcher(){
	var maxrow=document.getElementById("maxoptnumber").value;
	var switcher=document.getElementById("switcher");
	var hideraquo;
	var hidestock=false;
	if(switcher.selectedIndex==0){
		doswon='block';
		doswoff='none';
		depopts='none';
		hideraquo=false;
<% if NOT useStockManagement then print "hidestock=true;" %>
	}else if(switcher.options[1].disabled){
		switcher.selectedIndex=0;
		return;
	}else if(switcher.selectedIndex==2){ // Dependent Options
		doswon='none';
		doswoff='none';
		depopts='block';
		hideraquo=true;
	}else{
		doswon='none';
		doswoff='block';
		depopts='none';
		hideraquo=false;
	}
	for(index=-1;index<maxrow;index++){
		if(index==-1)theindex='';else theindex=index;
		document.getElementById("swprdiff" + theindex).style.display=doswon;
		document.getElementById("swaltid" + theindex).style.display=doswoff;
		document.getElementById("swwtdiff" + theindex).style.display=doswon;
		document.getElementById("swaltimg" + theindex).style.display=doswoff;
		document.getElementById("swstk" + theindex).style.display=doswon;
		document.getElementById("swaltlgim" + theindex).style.display=doswoff;
		document.getElementById("optclass" + theindex).style.display=switcher.selectedIndex==0?'block':'none';
		document.getElementById("depopts" + theindex).style.display=depopts;
		document.getElementById("depcell" + theindex).style.textAlign=hideraquo?"left":"center";
		if(index>=0){
			hasaltid=(document.getElementById("regexp" + theindex).value.replace(/ /,'')!='');
<% if useStockManagement then print "document.getElementById('optStock' + theindex).disabled=hasaltid;" %>
		}
	}
	if(raquo=document.getElementById('raquo1a'))
		raquo.style.visibility=hideraquo?"collapse":"";
	document.getElementById('stkcol').style.visibility=hidestock?"collapse":"";
	document.getElementById('raquo15').style.visibility=hideraquo?"collapse":"";
	document.getElementById('raquo2').style.visibility=hideraquo?"collapse":"";
	document.getElementById('raquo25').style.visibility=hideraquo?"collapse":"";
	document.getElementById('raquo3').style.visibility=hideraquo||hidestock?"collapse":"";
	document.getElementById('raquo4').style.visibility=switcher.selectedIndex==0?'':'collapse';
}
<%	if adminlanguages>1 AND ((adminlangsettings AND 32)=32) then %>
function doswitchlang(){
var langid=document.getElementById("langid");
var theid=langid[langid.selectedIndex].value;
var maxrow=document.getElementById("maxoptnumber").value;
for(index=0;index<maxrow;index++){
<%		for index=2 to adminlanguages+1 %>
document.getElementById("lang<%=index%>x" + index).style.display='none';
<%		next %>
}
for(index=0;index<maxrow;index++){
document.getElementById("lang" + theid + "x" + index).style.display='block';
}
}
<%	end if %>
function doaddrow(){
var rownumber=document.getElementById("maxoptnumber").value;
opttable=document.getElementById('optiontable');
newrow=opttable.insertRow(opttable.rows.length);
newcell=newrow.insertCell(0);
newcell.align='center';
newcell.innerHTML='<input type="radio" name="optdefault" value="'+rownumber+'" />';
newcell=newrow.insertCell(1);
newcell.innerHTML='<input type="button" id="insertopt'+rownumber+'" value="+" onclick="insertoption(this)" />';
newcell=newrow.insertCell(2);
newcell.align='center';
newcell.innerHTML='<input type="text" name="opt'+rownumber+'" id="opt'+rownumber+'" style="width:142px" value="" />';
newcell=newrow.insertCell(3);
newcell.innerHTML='&raquo;';
<%	extracells=0
	if adminlanguages>=1 AND ((adminlangsettings AND 32)=32) then
		extracells=2
		langtext=""
		for index=2 to adminlanguages+1
			langtext=langtext & "<span id=""lang"&index&"x'+rownumber+'"""
			if index>2 then langtext=langtext & " style=""display:none"">" else langtext=langtext & ">"
			langtext=langtext & "<input type=""text"" name=""opl"&index&"x'+rownumber+'"" id=""opl"&index&"x'+rownumber+'"" size=""20"" /></span>"
		next %>
newcell=newrow.insertCell(4);
newcell.align='center';
newcell.innerHTML='<%=langtext%>';

newcell=newrow.insertCell(5);
newcell.innerHTML='&raquo;';
<%	end if

langtext="<span id=""swprdiff'+rownumber+'"">"
langtext=langtext & "&nbsp;&nbsp;&nbsp;&nbsp;<input type=""text"" name=""pri'+rownumber+'"" id=""pri'+rownumber+'"" size=""5"" />"
if wholesaleoptionpricediff=TRUE then
	langtext=langtext & " / <input type=""text"" name=""wsp'+rownumber+'"" id=""wsp'+rownumber+'"" size=""5"" />"
end if
langtext=langtext & "<span id=""punitspan'+rownumber+'"" style=""padding:2px;'+(document.forms.mainform.pricepercent.checked?'':'display:none')+'"">%</span>"
langtext=langtext & "</span><span id=""swaltid'+rownumber+'"" style=""display:none""><input type=""text"" name=""regexp'+rownumber+'"" id=""regexp'+rownumber+'"" size=""12"" /></span>"
%>
newcell=newrow.insertCell(<%=(4+extracells)%>);
newcell.align='center';
newcell.innerHTML='<%=langtext%>';

newcell=newrow.insertCell(<%=(5+extracells)%>);
newcell.innerHTML='&raquo;';

<%
langtext="<span id=""swwtdiff'+rownumber+'"">"
langtext=langtext & "&nbsp;&nbsp;&nbsp;&nbsp;<input type=""text"" name=""wei'+rownumber+'"" id=""wei'+rownumber+'"" size=""5"" /><span id=""wunitspan'+rownumber+'"" style=""padding:2px;'+(document.forms.mainform.weightpercent.checked?'':'display:none')+'"">%</span>"
langtext=langtext & "</span><span id=""swaltimg'+rownumber+'"" style=""display:none""><input type=""text"" name=""altimg'+rownumber+'"" id=""altimg'+rownumber+'"" size=""20"" /></span>"
%>
newcell=newrow.insertCell(<%=(6+extracells)%>);
newcell.align='center';
newcell.innerHTML='<%=langtext%>';

newcell=newrow.insertCell(<%=(7+extracells)%>);
newcell.whiteSpace='nowrap';
newcell.innerHTML='&raquo;';

<%
langtext="<span id=""swstk'+rownumber+'"">"
if useStockManagement then langtext=langtext & "<input type=""text"" name=""optStock'+rownumber+'"" id=""optStock'+rownumber+'"" size=""4"" />"
langtext=langtext & "</span><span id=""swaltlgim'+rownumber+'"" style=""display:none""><input type=""text"" name=""altlimg'+rownumber+'"" id=""altlimg'+rownumber+'"" size=""20"" /></span>" & _
	"<span id=""depopts'+rownumber+'"" style=""display:none"">" & _
	"<select id=""depopts'+rownumber+'_1"" onmouseover=""populateoptionsselect(this)"" onchange=""addoptionselect(this)"" style=""width:140px""><option value=""x"">"&jscheck(yySelect)&"</option></select>" & _
	"</span>"
%>
newcell=newrow.insertCell(<%=(8+extracells)%>);
newcell.align='center';
newcell.id='depcell'+rownumber;
newcell.innerHTML='<%=langtext%>';

newcell=newrow.insertCell(<%=(9+extracells)%>);
newcell.whiteSpace='nowrap';
newcell.innerHTML='&raquo;';

newcell=newrow.insertCell(<%=(10+extracells)%>);
newcell.align='center';
newcell.id='clscell'+rownumber;
newcell.innerHTML='<span id="optclass'+rownumber+'"><input type="text" name="cls'+rownumber+'" id="cls'+rownumber+'" size="10" /></span>';

document.getElementById("maxoptnumber").value=parseInt(rownumber)+1;
}
function addmorerows(){
	numextrarows=document.getElementById("numextrarows").value;
	numextrarows=parseInt(numextrarows);
	if(isNaN(numextrarows))numextrarows=1;
	if(numextrarows==0)numextrarows=1;
	if(numextrarows>100)numextrarows=100;
	for(index=0;index<numextrarows;index++){
		doaddrow();
	}
	doswitcher();
<%	if adminlanguages>1 AND ((adminlangsettings AND 32)=32) then %>
	doswitchlang();
<% 	end if %>
}
function moveitemup(tid,tindex){
	if(document.getElementById('opt' + tindex).value!=''&&document.getElementById(tid + tindex).name=='xxx')
		document.getElementById(tid + tindex).name=document.getElementById(tid + tindex).id;
	document.getElementById(tid + tindex).value=document.getElementById(tid + (tindex-1)).value;
}
function insertoption(theval){
	var maxoptnumber=parseInt(document.getElementById("maxoptnumber").value);
	var theid=theval.id;
	theid=parseInt(theid.replace(/insertopt/, ''));
	if(document.getElementById('opt' + (maxoptnumber-1)).value!=''){
		doaddrow();
		doswitcher();
<%	if adminlanguages>1 AND ((adminlangsettings AND 32)=32) then %>
		doswitchlang();
<% 	end if %>
		maxoptnumber++;
	}
	for(index=maxoptnumber-1;index>theid;index--){
		document.getElementById('opt' + index).value=document.getElementById('opt' + (index-1)).value;
<%	if adminlanguages>1 AND ((adminlangsettings AND 32)=32) then
		for index=2 to adminlanguages+1
			print "moveitemup('opl"&index&"x',index);" & vbCrLf
		next
	end if
	if wholesaleoptionpricediff=TRUE then print "moveitemup('wsp',index);" & vbCrLf
	if useStockManagement then print "moveitemup('optStock',index);" & vbCrLf
%>		moveitemup('pri',index);
		moveitemup('regexp',index);
		moveitemup('wei',index);
		moveitemup('altimg',index);
		moveitemup('altlimg',index);
		moveitemup('cls',index);
	}
	document.getElementById('opt' + theid).value='';
<%	if adminlanguages>1 AND ((adminlangsettings AND 32)=32) then
		for index=2 to adminlanguages+1
			print "document.getElementById('opl"&index&"x' + theid).value='';" & vbCrLf
		next
	end if
	if wholesaleoptionpricediff=TRUE then print "document.getElementById('wsp' + theid).value='';" & vbCrLf %>
	document.getElementById('pri' + theid).value='';
	document.getElementById('regexp' + index).value='';
	document.getElementById('wei' + index).value='';
	document.getElementById('altimg' + index).value='';
<%	if useStockManagement then print "document.getElementById('optStock' + index).value='';" %>
	document.getElementById('altlimg' + index).value='';
	document.getElementById('cls' + index).value='';
}
function checkmultipurchase(opttype){
	var theopttype=opttype[opttype.selectedIndex];
	var maxrow=document.getElementById("maxoptnumber").value;
	var switcher=document.getElementById('switcher');
	document.getElementById('plsselspan').innerHTML=(theopttype.value==4?'<%=replace(yyDtPgOn," ","&nbsp;")%>':'<%=replace(yyPlsSLi," ","&nbsp;")%>');
	if(switcher.selectedIndex==2){
		switcher.selectedIndex=0;
		doswitcher();
	}
	switcher.options[2].disabled=(theopttype.value==4);
}
function switchtextinput(numrows){
	if(numrows>5) numrows=5;
	document.getElementById("opt0").rows=numrows;
	document.getElementById("opt0").style.whiteSpace=(numrows==1?"nowrap":"");
	document.getElementById("oph0").rows=numrows;
	document.getElementById("oph0").style.whiteSpace=(numrows==1?"nowrap":"");
<%	for index=2 to adminlanguages+1
		if (adminlangsettings AND 16)=16 then
			print "document.getElementById(""opl"&index&"x0"").rows=numrows;" & vbCrLf
			print "document.getElementById(""opl"&index&"x0"").style.whiteSpace=(numrows==1?'nowrap':'');" & vbCrLf
			
			print "document.getElementById(""oph"&index&""").rows=numrows;" & vbCrLf
			print "document.getElementById(""oph"&index&""").style.whiteSpace=(numrows==1?'nowrap':'');" & vbCrLf
		end if
	next %>
}
function disableelem(theelemtxt,isdis){
	var theelem=document.getElementById(theelemtxt);
	if(isdis){
		theelem.disabled=true;
		theelem.style.backgroundColor="#DDDDDD";
	}else{
		theelem.disabled=false;
		theelem.style.backgroundColor="#FFFFFF";
	}
}
function checkre(theval){
if(document.getElementById('regexp'+theval).value!=''){
	disableelem('pri'+theval,true);
	disableelem('wei'+theval,true);
	if(document.getElementById('wsp'+theval)) disableelem('wsp'+theval,true);
<% if useStockManagement then print "disableelem('optStock'+theval,true);" %>
}else{
	disableelem('pri'+theval,false);
	disableelem('wei'+theval,false);
	if(document.getElementById('wsp'+theval)) disableelem('wsp'+theval,false);
<% if useStockManagement then print "disableelem('optStock'+theval,false);" %>
}
}
var curropttype=0;
/* ]]> */
</script>
<%	if getpost("posted")="1" AND (getpost("act")="modify" OR getpost("act")="clone" OR getpost("act")="addnew") then
		iscloning=(getpost("act")="clone")
		if (getpost("act")="modify" OR iscloning) AND is_numeric(getpost("id")) then
			doaddnew=false
			sSQL="SELECT optID,optName,optGrpName,optGrpWorkingName,optPriceDiff,optType,optWeightDiff,optFlags,optStock,optWholesalePriceDiff,optRegExp,optName2,optName3,optGrpName2,optGrpName3,optDefault,optGrpSelect,optAltImage,optAltLargeImage,optTxtMaxLen,optTxtCharge,optMultiply,optAcceptChars,optDependants,optPlaceholder,optPlaceholder2,optPlaceholder3,optToolTip,optClass FROM options INNER JOIN optiongroup ON optiongroup.optGrpID=options.optGroup WHERE optGroup="&getpost("id")&" ORDER BY optID"
			rs.open sSQL,cnn,0,1
			alldata=rs.getrows
			rs.close
			optName=alldata(1,0)
			optGrpName=alldata(2,0)
			optGrpWorkingName=alldata(3,0)
			optPriceDiff=alldata(4,0)
			optType=alldata(5,0)
			optWeightDiff=alldata(6,0)
			optFlags=alldata(7,0)
			optStock=alldata(8,0)
			optWholesalePriceDiff=alldata(9,0)
			optName2=alldata(11,0)
			optName3=alldata(12,0)
			optGrpName2=alldata(13,0)
			optGrpName3=alldata(14,0)
			optDefault=alldata(15,0)
			optGrpSelect=alldata(16,0)
			optAltImage=alldata(17,0)
			optAltLargeImage=alldata(18,0)
			optTxtMaxLen=alldata(19,0)
			optTxtCharge=alldata(20,0)
			optMultiply=alldata(21,0)
			optAcceptChars=alldata(22,0)
			optplaceholder=alldata(24,0)
			optplaceholder2=alldata(25,0)
			optplaceholder3=alldata(26,0)
			opttooltip=alldata(27,0)
			optclass=alldata(28,0)
			maxoptnumber=UBOUND(alldata,2)
		else
			doaddnew=true
			optName=""
			optGrpName=""
			optGrpWorkingName=""
			optPriceDiff=15
			optType=int(getpost("optType"))
			optWeightDiff=""
			optFlags=0
			optStock=""
			optWholesalePriceDiff=""
			optName2=""
			optName3=""
			optGrpName2=""
			optGrpName3=""
			optDefault=""
			optGrpSelect=1
			optAltImage=""
			optAltLargeImage=""
			optTxtMaxLen=0
			optTxtCharge=0
			optMultiply=0
			optAcceptChars=""
			optplaceholder="" : optplaceholder2="" : optplaceholder3=""
			opttooltip=""
			optclass=""
			maxoptnumber=-1
		end if
		iscostperentry=optTxtCharge<0
		optTxtCharge=abs(optTxtCharge)
%>
	<form name="mainform" method="post" action="adminprodopts.asp" onsubmit="return formvalidator(this)">
	<input type="hidden" name="posted" value="1" />
	<%	if iscloning OR getpost("act")="addnew" then %>
	<input type="hidden" name="act" value="doaddnew" />
	<%	else %>
	<input type="hidden" name="act" value="domodify" />
	<input type="hidden" name="id" value="<%=getpost("id")%>" />
	<%	end if
		call writehiddenvar("disp", getpost("disp"))
		call writehiddenvar("stext", getpost("stext"))
		call writehiddenvar("stype", getpost("stype"))
		call writehiddenvar("pg", getpost("pg"))
		if abs(optType)=3 OR abs(optType)=5 then print "<input type=""hidden"" id=""optType"" name=""optType"" value="""&IIfVr(abs(optType)=3,3,5)&""" />" %>
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
		  <td align="center">
			<table border="0" cellspacing="0" cellpadding="3">
<%		if abs(optType)=3 OR abs(optType)=5 then ' Text option
			fieldHeight=cInt((cdbl(optPriceDiff)-Int(optPriceDiff))*100.0)
%>			  <tr>
                <td width="100%" colspan="4" align="center"><strong><%=IIfVr(getpost("act")="clone",yyClone,yyModify)&": "&yyPOAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td align="right" height="30"><%=yyPOName%>:</td><td align="left"><input type="text" name="secname" size="30" value="<%=htmldisplay(optGrpName)%>" /></td>
				<td align="right"><%=yyDefTxt%>:</td><td align="left"><textarea name="opt0" id="opt0" cols="30" rows="<%=fieldHeight%>"><%=htmldisplay(optName)%></textarea></td>
			  </tr>
			  <tr>
				<td align="right" height="30">&nbsp;</td><td align="left">&nbsp;</td>
				<td align="right">Placeholder:</td><td align="left"><textarea name="oph0" id="oph0" cols="30" rows="<%=fieldHeight%>"><%=htmldisplay(optPlaceholder)%></textarea></td>
			  </tr>
<%				for index=2 to adminlanguages+1
					if index=2 then optGrpName=optGrpName2 : optName=optName2 : optPlaceholder=optPlaceholder2
					if index=3 then optGrpName=optGrpName3 : optName=optName3 : optPlaceholder=optPlaceholder3
					if (adminlangsettings AND 16)=16 then %>
			  <tr>
				<td align="right" height="30"><%=yyPOName & " " & index%>:</td><td align="left"><input type="text" name="secname<%=index%>" size="30" value="<%=htmldisplay(optGrpName)%>" /></td>
				<td align="right"><%=yyDefTxt & " " & index%>:</td><td align="left"><textarea name="opl<%=index%>x0" id="opl<%=index%>x0" cols="30" rows="<%=fieldHeight%>"><%=htmldisplay(optName)%></textarea></td>
			  </tr>
			  <tr>
				<td align="right" height="30">&nbsp;</td><td align="left">&nbsp;</td>
				<td align="right"><%="Placeholder" & " " & index%>:</td><td align="left"><textarea name="oph<%=index%>" id="oph<%=index%>" cols="30" rows="<%=fieldHeight%>"><%=htmldisplay(optPlaceholder)%></textarea></td>
			  </tr>
<%					end if
				next %>
			  <tr>
				<td align="right" rowspan="3" height="30"><%=yyWrkNam%>:</td>
				<td align="left" rowspan="3"><input type="text" name="workingname" size="30" value="<%=htmlspecials(optGrpWorkingName)%>" /></td>
				<td align="right" height="30"><%=yyFldWdt%>:</td>
				<td align="left"><select name="pri0" size="1"><%
					for rowcounter=1 to 35
						print "<option value='"&rowcounter&"'"
						if rowcounter=Int(optPriceDiff) then print " selected"
						print ">&nbsp; "&rowcounter&" </option>"&vbCrLf
					next
				%>
				</select></td>
			  </tr>
			  <tr>
				<td align="right" height="30"><%=yyFldHgt%>:</td>
				<td align="left"><select name="fieldheight" size="1" onchange="switchtextinput(this.selectedIndex+1)"><%
					for rowcounter=1 to 15
						print "<option value='"&rowcounter&"'"
						if rowcounter=fieldHeight then print " selected"
						print ">&nbsp; "&rowcounter&" </option>"&vbCrLf
					next
				%>
				</select></td>
			  </tr>
			  <tr>
				<td align="right" height="30"><%=yyMaxEnt%>:</td>
				<td align="left"><select name="optTxtMaxLen" size="1">
				<option value="0">MAX</option>
				<%	for rowcounter=1 to 255
						print "<option value='"&rowcounter&"'"
						if rowcounter=optTxtMaxLen then print " selected=""selected"""
						print ">&nbsp; "&rowcounter&" </option>"&vbCrLf
					next
				%>
				</select></td>
			  </tr>
			  <tr>
				<td align="right" height="30"><%=yyForSel%>:</td><td align="left"><input type="checkbox" name="forceselec" value="ON"<% if optType > 0 then print " checked=""checked"""%> /></td>
				<td align="right" height="30">Class:</td><td align="left"><input type="text" size="20" name="cls" value="<%=optclass%>" /></td>
			  </tr>			  
			  <tr>
				<td align="right" height="30"><%=yyCosPer%>:</td><td align="left"><select name="iscostperentry" size="1"><option value=""><%=yyCosCha%></option><option value="1"<% if iscostperentry then print " selected=""selected"""%>><%=yyCosEnt%></option></select> <input type="text" name="optTxtCharge" value="<%=htmlspecials(optTxtCharge)%>" size="5" /></td>
				<td align="right"><%=yyAccCha%>:</td>
				<td align="left"><input type="text" name="optAcceptChars" value="<%=htmlspecials(optAcceptChars)%>" size="20" /></td>
			  </tr>
			  <tr>
				<td align="right" height="30"><%=yyIsMult%>:</td>
				<td align="left"><input type="checkbox" name="optMultiply"<% if optMultiply<>0 then print " checked=""checked"""%> value="ON" /></td>
				<td align="right">Is Date Picker:</td>
				<td align="left"><input type="checkbox" name="isdatepicker"<% if abs(optType)=5 then print " checked=""checked"""%> value="ON" onchange="document.getElementById('optType').value=this.checked?5:3" /></td>
			  </tr>
			  <tr>
				<td align="right">Tooltip:</td>
				<td colspan="3"><textarea name="opttooltip" cols="50" rows="10"><%=opttooltip%></textarea></td>
			  </tr>
			  <tr>
				<td colspan="4" align="left">
				  <ul>
				  <li><span style="font-size:10px"><%=yyPOEx1%></span></li>
				  <li><span style="font-size:10px"><%=yyPOEx2%></span></li>
				  <li><span style="font-size:10px"><%=yyPOEx3%></span></li>
				  </ul>
				  <input type="hidden" name="maxoptnumber" id="maxoptnumber" value="0" />
                </td>
			  </tr>
<%		else %>
			  <tr>
				<td width="30%" align="center">
				  <table border="0" cellspacing="0" cellpadding="3">
				  <tr><td align="right"><%=replace(yyPOName," ","&nbsp;")%></td><td align="left" colspan="3">
				  <input type="text" name="secname" size="30" value="<%=htmldisplay(optGrpName)%>" /></td></tr>
<%				for index=2 to adminlanguages+1
					if index=2 then optGrpName=optGrpName2
					if index=3 then optGrpName=optGrpName3
					if (adminlangsettings AND 16)=16 then
						%><tr><td align="right"><%=replace(yyPOName & " " & index," ","&nbsp;")%></td><td align="left" colspan="3">
				  <input type="text" name="secname<%=index%>" size="30" value="<%=htmldisplay(optGrpName)%>" /></td></tr><%
					end if
				next %>
				  <tr><td align="right"><%=replace(yyWrkNam," ","&nbsp;")%></td><td align="left" colspan="3"><input type="text" name="workingname" size="30" value="<%=htmldisplay(optGrpWorkingName)%>" /></td></tr>
				  <tr><td align="right"><%=replace(yyOptSty," ","&nbsp;")%></td><td align="left" colspan="3"><select name="optType" id="optType" size="1" onclick="curropttype=this.selectedIndex" onchange="checkmultipurchase(this)"><option value="2"><%=yyDDMen%></option><option value="1"<% if abs(optType)=1 then print " selected=""selected"""%>><%=yyRadBut%></option><option value="4"<% if abs(optType)=4 then print " selected=""selected"""%>><%=yyMulPur%></option></select></td></tr>
				  <tr><td align="right" style="white-space:nowrap"><%=replace(yyForSel," ","&nbsp;")%></td><td align="left" style="white-space:nowrap"><input type="checkbox" name="forceselec" value="ON"<% if optType > 0 then print " checked=""checked"""%> />&nbsp;</td><td align="right" style="white-space:nowrap">&nbsp;<input type="radio" name="optdefault" value="" /></td><td align="left" style="white-space:nowrap"><%=replace(yyNoDefa," ","&nbsp;")%></td></tr>
				  <tr><td align="right" style="white-space:nowrap"><%=replace(yySinLin," ","&nbsp;")%></td><td align="left" style="white-space:nowrap"><input type="checkbox" name="singleline" value="1"<% if (optFlags AND 4)=4 then print " checked=""checked"""%> /></td><td align="right" style="white-space:nowrap"><input type="checkbox" name="optgrpselect" value="1"<% if cint(optGrpSelect)<>0 then print " checked=""checked"""%> /></td><td align="left" style="white-space:nowrap"><span id="plsselspan"><%=replace(IIfVr(abs(optType)=4,yyDtPgOn,yyPlsSLi)," ","&nbsp;")%></span></td></tr>
				  </table>
                </td>
				<td colspan="2" align="left">
				  <p align="center"><%=IIfVr(getpost("act")="clone",yyClone,yyModify)&": "&yyPOAdm%></p>
				  <ul>
				  <li><span style="font-size:10px"><%=yyPOEx1%></span></li>
				  <li><span style="font-size:10px"><%=yyPOEx4%></span></li>
				  <li><span style="font-size:10px"><%=yyPOEx5%></span></li>
				  <% if useStockManagement then %>
				  <li><span style="font-size:10px"><%=yyPOEx6%></span></li>
				  <% end if %>
				  </ul>
				  <div style="text-align:center"><table style="margin:0 auto"><tr><td align="right">Tooltip</td><td><textarea name="opttooltip" cols="50" rows="3" onfocus="this.rows=10" onblur="this.rows=3"><%=opttooltip%></textarea></td></tr></table></div>
                </td>
			  </tr>
			</table>
			<table id="optiontable" width="500" border="0" cellspacing="0" cellpadding="3">
			<col /><col /><col /><col id="raquo1" /><% if adminlanguages>=1 AND ((adminlangsettings AND 32)=32) then print "<col /><col id=""raquo1a"" />"%><col id="raquo15" /><col id="raquo2" /><col id="raquo25" /><col id="raquo3" <% if NOT useStockManagement then print "style=""visibility:collapse"" "%>/><col id="stkcol" <% if NOT useStockManagement then print "style=""visibility:collapse"" "%>/><col id="raquo4" /><col id="classcol" />
			  <tr>
				<td><%=yyDefaul%></td>
				<td width="3%" align="center">&nbsp;</td>
				<td align="center"><select name="switcher" id="switcher" size="1" onchange="doswitcher()"><option value="1"><%=yyPOOpts&" / "&yyVals%></option><option value="2"><%=yyPOOpts&" / "&yyAlts%></option><option value="3"<% if abs(optType)=4 then print " disabled=""disabled"""%>>Dependent Options</option></select></td>
				<td width="3%" align="center">&nbsp;</td>
<%			if adminlanguages>=1 AND ((adminlangsettings AND 32)=32) then
				print "<td align=""center""><select name=""langid"" id=""langid"" size=""1"" onchange=""doswitchlang()"">"
				for index=2 to adminlanguages+1
					print "<option value="""&index&""">" & yyPOOpts & " Language " & index & "</option>"
				next
				print "</select></td><td align=""center"">&nbsp;</td>"
			end if	%>
				<td align="center" style="white-space:nowrap;"><span id="swprdiff"><% if wholesaleoptionpricediff=TRUE then print yyPrWsa else print yyPOPrDf%>&nbsp;%<input class="noborder" type="checkbox" name="pricepercent" value="1" onclick="changeunits();"<% if (optFlags AND 1)=1 then print " checked=""checked"""%> /></span><span id="swaltid" style="display:none"><%=yyAltPId%></span></td>
				<td width="3%" align="center">&nbsp;</td>
				<td align="center" style="white-space:nowrap;"><span id="swwtdiff"><%=yyPOWtDf%>&nbsp;%<input class="noborder" type="checkbox" name="weightpercent" value="1" onclick="changeunits();"<% if (optFlags AND 2)=2 then print " checked=""checked"""%> /></span><span id="swaltimg" style="display:none"><%=yyAltIm%></span></td>
				<td width="3%" align="center">&nbsp;</td>
				<td align="center" style="white-space:nowrap;" id="depcell"><span id="swstk"><%=yyStkLvl%></span><span id="swaltlgim" style="display:none"><%=yyAltLIm%></span><span id="depopts" style="display:none">Dependent Options</span></td>
				<td width="3%" align="center">&nbsp;</td>
				<td align="center" style="white-space:nowrap;" id="depcell"><span id="optclass">Class</span></td>
			  </tr>
<%			for rowcounter=0 to vrmax(14, maxoptnumber+5)
				if rowcounter<=maxoptnumber then optclass=alldata(28,rowcounter) else optclass="" %>
			  <tr>
				<td align="center"><input type="radio" name="optdefault" value="<%=rowcounter%>"<% if rowcounter<=maxoptnumber then if cint(alldata(15,rowcounter))<>0 then print " checked=""checked"""%> /></td>
				<td align="center"><input type="button" id="insertopt<%=rowcounter%>" value="+" onclick="insertoption(this)" /></td>
				<td align="center"><%
					if rowcounter<=maxoptnumber AND NOT iscloning then print "<input type=""hidden"" name=""orig" & rowcounter & """ value=""" & alldata(0,rowcounter) & """ />"
					print "<input type=""text"" name=""opt"&rowcounter&""" id=""opt"&rowcounter&""" style=""width:142px"" value="""
					if rowcounter<=maxoptnumber then print replace(alldata(1,rowcounter)&"","""", "&quot;")
					print """ /><br />"&vbCrLf
				%></td><td>&raquo;</td>
<%				if adminlanguages>=1 AND ((adminlangsettings AND 32)=32) then
					print "<td align=""center"">"
					for index=2 to adminlanguages+1
						print "<span id=""lang"&index&"x"&rowcounter&""""
						if index>2 then print " style=""display:none"">" else print ">"
						print "<input type=""text"" name=""opl"&index&"x"&rowcounter&""" id=""opl"&index&"x"&rowcounter&""" size=""20"" value="""
						if rowcounter<=maxoptnumber then print replace(alldata(9+index,rowcounter)&"","""", "&quot;")
						print """ onchange=""this.name=this.id"" /></span>"
					next
					print "</td><td>&raquo;</td>"
				end if %>
				<td align="center"><span id="swprdiff<%=rowcounter%>" style="white-space:nowrap"><%
					if rowcounter<=maxoptnumber then optvalue=alldata(4,rowcounter) else optvalue=0
					print "&nbsp;&nbsp;&nbsp;&nbsp;<input type=""text"" name=""" & IIfVr(optvalue<>0,"pri" & rowcounter,"xxx") & """ id=""pri"&rowcounter&""" size=""5"" value="""
					if rowcounter<=maxoptnumber then print optvalue
					print """ onchange=""this.name=this.id"" />"
					if wholesaleoptionpricediff=TRUE then
						if rowcounter<=maxoptnumber then optvalue=alldata(9,rowcounter) else optvalue=0
						print " / <input type=""text"" name=""" & IIfVr(optvalue<>0,"wsp" & rowcounter,"xxx") & """ id=""wsp"&rowcounter&""" size=""5"" value='"
						if rowcounter<=maxoptnumber then print optvalue
						print "' onchange=""this.name=this.id"" />"
					end if
					print "<span id=""punitspan"&rowcounter&""" style=""padding:2px"&IIfVs((optFlags AND 1)<>1,";display:none")&""">%</span>"
					if rowcounter<=maxoptnumber then optvalue=alldata(10,rowcounter) else optvalue=""
				%></span><span id="swaltid<%=rowcounter%>" style="display:none"><input type="text" name="<%=IIfVr(optvalue<>"","regexp" & rowcounter,"xxx")%>" id="regexp<%=rowcounter%>" onchange="this.name=this.id;checkre(<%=rowcounter%>)" size="12" value="<% print optvalue %>" /></span></td>
				<td>&raquo;</td>
				<td align="center" style="white-space:nowrap;"><span id="swwtdiff<%=rowcounter%>"><%
					if rowcounter<=maxoptnumber then optvalue=alldata(6,rowcounter) else optvalue=0
					print "&nbsp;&nbsp;&nbsp;&nbsp;<input type=""text"" name=""" & IIfVr(optvalue<>0,"wei" & rowcounter,"xxx") & """ id=""wei"&rowcounter&""" size=""5"" value='"
					if rowcounter<=maxoptnumber then print optvalue
					print "' onchange=""this.name=this.id"" /><span id=""wunitspan"&rowcounter&""" style=""padding:2px"&IIfVs((optFlags AND 2)<>2,";display:none")&""">%</span>"
					if rowcounter<=maxoptnumber then optvalue=alldata(17,rowcounter) else optvalue=""
				%></span><span id="swaltimg<%=rowcounter%>" style="display:none"><input type="text" name="<%=IIfVr(optvalue<>"","altimg" & rowcounter,"xxx")%>" id="altimg<%=rowcounter%>" size="20" value="<% print optvalue %>" onchange="this.name=this.id" /></span>
				</td>
				<td>&raquo;</td>
				<td align="center" style="white-space:nowrap" id="depcell<%=rowcounter%>"><span id="swstk<%=rowcounter%>"><%
					if rowcounter<=maxoptnumber then optvalue=alldata(8,rowcounter) else optvalue=0
					if useStockManagement then
						print "<input type=""text"" name=""" & IIfVr(optvalue<>0,"optStock" & rowcounter,"xxx") & """ id=""optStock"&rowcounter&""" size=""4"" value="""
						if rowcounter<=maxoptnumber then
							print optvalue
							if trim(alldata(10,rowcounter))<>"" then print """ disabled=""disabled"
						end if
						print """ onchange=""this.name=this.id"" />"
					' elseif rowcounter<=maxoptnumber then
					'	print "<input type=""hidden"" name=""optStock"&rowcounter&""" id=""optStock"&rowcounter&""" value=""" & optvalue & """ />n/a"
					end if
					if rowcounter<=maxoptnumber then optvalue=alldata(18,rowcounter) else optvalue=""
				%></span><span id="swaltlgim<%=rowcounter%>" style="display:none"><input type="text" name="<%=IIfVr(optvalue<>"","altlimg" & rowcounter,"xxx")%>" id="altlimg<%=rowcounter%>" size="20" value="<% print optvalue %>" onchange="this.name=this.id" /></span>
					<span id="depopts<%=rowcounter%>" style="display:none">
<%					if rowcounter<=maxoptnumber then optDependants=commaseplist(alldata(23,rowcounter)) else optDependants=""
					optionindex=1
					if optDependants<>"" then
						sSQL="SELECT optGrpID,optGrpWorkingName FROM optiongroup WHERE optGrpID IN ("&optDependants&")"
						rs2.open sSQL,cnn,0,1
						if NOT rs2.EOF then
							depsarray=split(optDependants,",")
							alldependants=rs2.getrows()
							for index2=0 to UBOUND(depsarray)
								if is_numeric(depsarray(index2)) then
									for index=0 to UBOUND(alldependants,2)
										if int(depsarray(index2))=alldependants(0,index) then
											print "<select id=""depopts"&rowcounter&"_"&optionindex&""" name=""depopts"&rowcounter&"_"&optionindex&""" onmouseover=""populateoptionsselect(this)"" style=""width:140px""><option value=""x"">"&yySelect&"</option><option value="""&alldependants(0,index)&""" selected=""selected"">"&alldependants(1,index)&"</option></select>&nbsp;"
											optionindex=optionindex+1
										end if
									next
								end if
							next
						end if
						rs2.close
					end if
					print "<select id=""depopts"&rowcounter&"_"&optionindex&""" onmouseover=""populateoptionsselect(this)"" onchange=""addoptionselect(this)"" style=""width:140px""><option value=""x"">"&yySelect&"</option></select>"
%>					</span>
				</td>
				<td>&raquo;</td>
				<td align="center" id="clscell<%=rowcounter%>">
					<span id="optclass<%=rowcounter%>">
						<input type="text" name="<%=IIfVr(optclass<>"","cls" & rowcounter,"xxx")%>" id="cls<%=rowcounter%>" size="10" value="<%=optclass %>" onchange="this.name=this.id" />
					</span>
				</td>
			  </tr>
<%			next %>
			</table>
			<input type="hidden" name="maxoptnumber" id="maxoptnumber" value="<%=rowcounter%>" />
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
<%		end if %>
			  <tr>
                <td width="100%" colspan="4" align="center"><br />
<%		if abs(optType)<>3 then %>
				<input type="text" name="numextrarows" id="numextrarows" value="10" size="4" /> <input type="button" value="<%=yyMore & " " & yyPOOpts%>" onclick="addmorerows()" />&nbsp;&nbsp;&nbsp;&nbsp;
<%		end if %>
				<input type="submit" value="<%=yySubmit%>" /><% if getpost("act")="modify" OR iscloning then %>&nbsp;&nbsp;&nbsp;<input type="reset" value="<%=yyReset%>" /><% end if %><br />&nbsp;</td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="4" align="center"><br /><a href="admin.asp"><%=yyAdmHom%></a><br />&nbsp;</td>
			  </tr>
            </table>
		  </td>
		</tr>
	  </table>
	</form>
<%		if abs(optType)<>3 then %>
<script>
/* <![CDATA[ */
for(var ti=0; ti<=<%=maxoptnumber%>; ti++) checkre(ti);
/* ]]> */
</script>
<%		end if
	elseif getpost("posted")="1" AND success then %>
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><%=yyUpdSuc%><br /><br /><%=yyNowFrd%><br /><br />
                        <%=yyNoAuto%> <a href="adminprodopts.asp"><%=yyClkHer%></a>.<br /><br />&nbsp;
                </td>
			  </tr>
			</table>
<%	elseif getpost("posted")="1" then %>
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><span style="color:#FF0000;font-weight:bold"><%=yyOpFai%></span><br /><br /><%=errmsg%><br /><br />
				<a href="javascript:history.go(-1)"><%=yyClkBac%></a></td>
			  </tr>
			</table>
<%	else
		pract=request.cookies("practopt")
		modclone=request.cookies("modclone")
%>
<script>
/* <![CDATA[ */
function setCookie(c_name,value,expiredays){
	var exdate=new Date();
	exdate.setDate(exdate.getDate()+expiredays);
	document.cookie=c_name+ "=" +escape(value)+((expiredays==null) ? "" : ";expires="+exdate.toGMTString());
}
function mr(id){
	document.mainform.id.value=id;
	document.mainform.act.value="modify";
	document.mainform.submit();
}
function cr(id){
	document.mainform.id.value=id;
	document.mainform.act.value="clone";
	document.mainform.submit();
}
function newtextrec(id) {
	document.mainform.id.value=id;
	document.mainform.act.value="addnew";
	document.mainform.optType.value="3";
	document.mainform.submit();
}
function newrec(id) {
	document.mainform.id.value=id;
	document.mainform.act.value="addnew";
	document.mainform.optType.value="2";
	document.mainform.submit();
}
function quickupdate(){
	if(document.mainform.pract.value=="del"){
		if(!confirm("<%=jscheck(yyConDel)%>\n"))
			return;
	}
	document.mainform.action="adminprodopts.asp";
	document.mainform.act.value="quickupdate";
	document.mainform.posted.value="1";
	document.mainform.submit();
}
function dr(id){
if (confirm("<%=jscheck(yyConDel)%>\n")){
	document.mainform.id.value=id;
	document.mainform.act.value="delete";
	document.mainform.submit();
}
}
function startsearch(){
	document.mainform.action="adminprodopts.asp";
	document.mainform.act.value="search";
	document.mainform.posted.value="";
	document.mainform.submit();
}
function changepract(obj){
	setCookie('practopt',obj[obj.selectedIndex].value,600);
	startsearch();
}
function changemodclone(modclone){
	setCookie('modclone',modclone[modclone.selectedIndex].value,600);
	startsearch();
}
function checkboxes(docheck){
	maxitems=document.getElementById("resultcounter").value;
	for(index=0;index<maxitems;index++){
		document.getElementById("chkbx"+index).checked=docheck;
	}
}
function setselects(tsmen){
	if(tsmen.selectedIndex==0){
		document.forms.mainform.reset();
	}else{
		maxitems=document.getElementById("resultcounter").value;
		for(index=0;index<maxitems;index++){
			if(document.getElementById("selbx"+index)) document.getElementById("selbx"+index).selectedIndex=tsmen.selectedIndex-1;
		}
	}
}
/* ]]> */
</script>
<h2><%=YYAdmPrO%></h2>
		  <form name="mainform" method="post" action="adminprodopts.asp">
			<input type="hidden" name="posted" value="1" />
			<input type="hidden" name="act" value="xxxxx" />
			<input type="hidden" name="id" value="xxxxx" />
			<input type="hidden" name="optType" value="xxxxx" />
			<table class="cobtbl" width="100%" border="0" cellspacing="1" cellpadding="3">
				  <tr><td class="cobhl" align="center" colspan="4" height="22"><strong><%=yyPOAdm%></strong></td></tr>
				  <tr> 
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
				  <tr>
				    <td class="cobhl" align="center"><%
					if getpost("act")="search" OR getget("pg")<>"" then
						if pract="del" then %>
						<input type="button" value="<%=yyCheckA%>" onclick="checkboxes(true);" /> <input type="button" value="<%=yyUCheck%>" onclick="checkboxes(false);" />
<%						elseif pract="oty" then
							print "<select size=""1"" onchange=""setselects(this)""><option value="""">Change All Options...</option><option value=""2"">"&yyDDMen&"</option><option value=""1"">"&yyRadBut&"</option><option value=""4"">"&yyMulPur&"</option></select>"
						end if
					else
						print "&nbsp;"
					end if %></td>
				    <td class="cobll" colspan="3"><table width="100%" cellspacing="0" cellpadding="0" border="0">
					    <tr>
						  <td class="cobll" align="center" style="white-space:nowrap">
							<select name="disp" size="1" style="vertical-align:middle">
							<option value="">All Options</option>
							<option value="2"<% if request("disp")="2" then print " selected"%>>Text Options</option>
							<option value="3"<% if request("disp")="3" then print " selected"%>>Multiple Purchase Options</option>
							<option value="4"<% if request("disp")="4" then print " selected"%>>Dropdown Options</option>
							<option value="5"<% if request("disp")="5" then print " selected"%>>Radio Options</option>
<%					if useStockManagement then print "<option value=""6"""&IIfVr(request("disp")="6", " selected", "")&">"&yyOOStoc&"</option>" %>
							<option value="7"<% if request("disp")="7" then print " selected"%>>Unused Options</option>
							</select>
							<input type="submit" value="List Options" onclick="startsearch();" />
						  </td>
						  <td class="cobll" height="26" width="20%" align="right" style="white-space:nowrap">
							<input type="button" value="<%=yyPONew%>" onclick="newrec()" />&nbsp;&nbsp;
							<input type="button" value="<%=yyPONewT%>" onclick="newtextrec()" />
						  </td>
						</tr>
					  </table></td>
				  </tr>
				</table>
<br />
            <table width="100%" class="stackable admin-table-a sta-white">
<%
jscript=""
if getpost("act")="search" OR getget("pg")<>"" then
	sSQL="SELECT optGrpID,optGrpName,optGrpName2,optGrpName3,optGrpWorkingName,optType FROM optiongroup"
	whereand=" WHERE "
	if request("disp")="6" then
		sSQL="SELECT DISTINCT optGrpID,optGrpName,optGrpName2,optGrpName3,optGrpWorkingName FROM optiongroup INNER JOIN (options INNER JOIN (prodoptions INNER JOIN products ON prodoptions.poProdID=products.pID) ON options.optGroup=prodoptions.poOptionGroup) ON optiongroup.optGrpID=options.optGroup WHERE options.optStock<=0 AND (optRegExp='' OR optRegExp IS NULL) AND products.pStockByOpts<>0 AND optType IN (-4,-2,-1,1,2,4)"
	elseif request("disp")="7" then
		sSQL="SELECT optGrpID,optGrpName,optGrpName2,optGrpName3,optGrpWorkingName,poProdID FROM optiongroup LEFT JOIN prodoptions ON optiongroup.optGrpID=prodoptions.poOptionGroup WHERE poProdID IS NULL"
	elseif request("disp")="2" then
		sSQL=sSQL & " WHERE optType IN (-5,-3,3,5)"
	elseif request("disp")="3" then
		sSQL=sSQL & " WHERE optType IN (-4,4)"
	elseif request("disp")="4" then
		sSQL=sSQL & " WHERE optType IN (-2,2)"
	elseif request("disp")="5" then
		sSQL=sSQL & " WHERE optType IN (-1,1)"
	end if
	if request("disp")<>"" then whereand=" AND "
	if trim(request("stext"))<>"" then
		sText=escape_string(request("stext"))
		aText=Split(sText)
		maxsearchindex=1
		aFields(0)="optGrpWorkingName"
		aFields(1)="optGrpName"
		if request("stype")="exact" then
			sSQL=sSQL & whereand & "(optGrpName LIKE '%"&sText&"%' OR optGrpWorkingName LIKE '%"&sText&"%') "
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
	sSQL=sSQL&" ORDER BY optGrpWorkingName,optGrpName"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then %>
			  <tr>
				<th class="minicell">
					<select name="pract" id="pract" size="1" onchange="changepract(this)">
					<option value="none">Quick Entry...</option>
					<option value="opn"<% if pract="opn" then print " selected=""selected"""%>><%=yyPOName%></option>
<%					for index=2 to adminlanguages+1
						if (adminlangsettings AND 16)=16 then print "<option value=""opn"&index&""""&IIfVr(pract="opn"&index," selected=""selected""","")&">"&yyPOName&" "&index&"</option>"
					next %>
					<option value="own"<% if pract="own" then print " selected=""selected"""%>><%=yyWrkNam%></option>
					<option value="oty"<% if pract="oty" then print " selected=""selected"""%>><%=yyOptSty%></option>
					<option value="" disabled="disabled">------------------</option>
					<option value="del"<% if pract="del" then print " selected=""selected"""%>><%=yyDelete%></option>
					</select></th>
				<th class="maincell"><%=yyPOName%></th>
				<th class="maincell"><%=yyWrkNam%></th>
				<th class="minicell"><%=yyModify%></th>
			  </tr>
<%		do while NOT rs.EOF
			jscript=jscript&"pa["&resultcounter&"]=[" %>
<tr id="tr<%=resultcounter%>"><td class="minicell"><%
				if pract="opn" then
					print "<input type=""text"" id=""chkbx"&resultcounter&""" size=""18"" name=""pra_"&rs("optGrpID")&""" value=""" & rs("optGrpName") & """ tabindex="""&(resultcounter+1)&"""/>"
				elseif pract="opn2" then
					print "<input type=""text"" id=""chkbx"&resultcounter&""" size=""18"" name=""pra_"&rs("optGrpID")&""" value=""" & rs("optGrpName2") & """ tabindex="""&(resultcounter+1)&"""/>"
				elseif pract="opn3" then
					print "<input type=""text"" id=""chkbx"&resultcounter&""" size=""18"" name=""pra_"&rs("optGrpID")&""" value=""" & rs("optGrpName3") & """ tabindex="""&(resultcounter+1)&"""/>"
				elseif pract="own" then
					print "<input type=""text"" id=""chkbx"&resultcounter&""" size=""18"" name=""pra_"&rs("optGrpID")&""" value=""" & rs("optGrpWorkingName") & """ tabindex="""&(resultcounter+1)&"""/>"
				elseif pract="oty" then
					opttype=abs(rs("optType"))
					if opttype=3 then
						print "-"
					else
						print "<select id=""selbx"&resultcounter&""" size=""1"" name=""pra_"&rs("optGrpID")&""" tabindex="""&(resultcounter+1)&"""><option value=""2"""&IIfVs(opttype=2," selected=""selected""")&">DROPDOWN</option><option value=""1"""&IIfVs(opttype=1," selected=""selected""")&">RADIO</option><option value=""4"""&IIfVs(opttype=4," selected=""selected""")&">MULTIPLE</option></select>"
					end if
				elseif pract="del" then
					print "<input type=""checkbox"" id=""chkbx"&resultcounter&""" name=""pra_"&rs("optGrpID")&""" value=""del"" tabindex="""&(resultcounter+1)&"""/>"
				else
					print "&nbsp;"
				end if
%></td><td><%=rs("optGrpName")%></td><td><%=rs("optGrpWorkingName")%></td><td>-</td></tr>
<%			jscript=jscript&rs("optGrpID")&"];"&vbCrLf
			resultcounter=resultcounter + 1
			rs.movenext
		loop
	else
%>			  <tr>
                <td width="100%" colspan="4" align="center"><br /><%=yyItNone%><br />&nbsp;</td>
			  </tr>
<%	end if
	rs.close
else
	numitems=0
	sSQL="SELECT COUNT(*) as totcount FROM optiongroup"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		numitems=rs("totcount")
	end if
	rs.close
	print "<tr><td colspan=""4""><div class=""itemsdefine"">You have " & numitems & " product options defined.</div></td></tr>"
end if ' getpost("act")="search" OR getget("pg")<>"" then
%>			  <tr>
				<td align="center" style="white-space:nowrap"><% if resultcounter>0 AND pract<>"" AND pract<>"none" then print "<input type=""hidden"" name=""resultcounter"" id=""resultcounter"" value="""&resultcounter&""" /><input type=""button"" value="""&yyUpdate&""" onclick=""quickupdate()"" /> <input type=""reset"" value="""&yyReset&""" />" else print "&nbsp;"%></td>
                <td width="100%" colspan="2" align="center"><br /><a href="admin.asp"><%=yyAdmHom%></a><br />&nbsp;<br /></td>
				<td>&nbsp;</td>
			  </tr>
            </table>
		  </form>
<script>
/* <![CDATA[ */
var pa=[];
<%=jscript%>
for(var pidind in pa){
	var ttr=document.getElementById('tr'+pidind);
	ttr.cells[3].style.textAlign='center';
	ttr.cells[3].style.whiteSpace='nowrap';
	ttr.cells[3].innerHTML='<input type="button" value="M" style="width:30px;margin-right:4px" onclick="mr(\''+pa[pidind][0]+'\')" title="<%=jsescape(htmlspecials(yyModify))%>" />' +
		'<input type="button" value="C" style="width:30px;margin-right:4px" onclick="cr(\''+pa[pidind][0]+'\')" title="<%=jsescape(htmlspecials(yyClone))%>" />' +
		'<input type="button" value="X" style="width:30px" onclick="dr(\''+pa[pidind][0]+'\')" title="<%=jsescape(htmlspecials(yyDelete))%>" />';
}
/* ]]> */
</script>
<%
end if
cnn.close
set rs=nothing
set cnn=nothing
%>