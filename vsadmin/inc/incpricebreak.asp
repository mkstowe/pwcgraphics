<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
Dim sSQL,rs,alldata,alladmin,success,cnn,rowcounter,errmsg,aFields(1)
success=true
if maxbreaksperpage="" then maxbreaksperpage = 200
maxpricebreaks = 25
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
sSQL = ""
dropdown = (getpost("ddown")="1")
dorefresh=FALSE
if getpost("posted")="1" then
	theprod=getpost("pid")
	if getpost("act")="delete" then
		sSQL = "DELETE FROM pricebreaks WHERE pbProdID='" & escape_string(getpost("id")) & "'"
		ect_query(sSQL)
		dorefresh=TRUE
	elseif getpost("act")="domodify" then
		sSQL = "SELECT pID FROM products WHERE pID='" & escape_string(theprod) & "'"
		rs.open sSQL,cnn,0,1
		if rs.EOF then
			success=false
			errmsg = "The specified product id (" & theprod & ") does not exist."
		end if
		rs.close
		if success then
			ect_query("DELETE FROM pricebreaks WHERE pbProdID='" & escape_string(theprod) & "'")
			for index=1 to maxpricebreaks
				thequant=getpost("quant"&index)
				if NOT is_numeric(thequant) then thequant=0
				price=getpost("price"&index)
				if NOT is_numeric(price) then price=0
				wprice=getpost("wprice"&index)
				if NOT is_numeric(wprice) then wprice=0
				wpercent=IIfVr(getpost("wpercent"&index)="1","1","0")
				wholesalepercent=IIfVr(getpost("wholesalepercent"&index)="1","1","0")
				if thequant<>0 AND (price<>0 OR wprice<>0) then
					sSQL = "INSERT INTO pricebreaks (pbProdID,pbQuantity,pPrice,pWholesalePrice,pbPercent,pbWholesalePercent) VALUES ('"&escape_string(theprod)&"',"
					sSQL = sSQL & thequant & ","
					sSQL = sSQL & price & ","
					sSQL = sSQL & wprice & ","
					sSQL = sSQL & wpercent & ","
					sSQL = sSQL & wholesalepercent & ")"
					on error resume next
					ect_query(sSQL)
					on error goto 0
				end if
			next
			dorefresh=TRUE
		end if
	elseif getpost("act")="doaddnew" AND theprod<>"" then
		sSQL = "SELECT pbProdID FROM pricebreaks WHERE pbProdID='" & escape_string(theprod) & "'"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			success=false
			errmsg = "Price breaks already exist for this product id. You should use the ""Modify"" option on the price breaks admin page"
		end if
		rs.close
		sSQL = "SELECT pID FROM products WHERE pID='" & escape_string(theprod) & "'"
		rs.open sSQL,cnn,0,1
		if rs.EOF then
			success=false
			errmsg = "The specified product id (" & theprod & ") does not exist."
		end if
		rs.close
		if success then
			for index=1 to maxpricebreaks
				thequant=getpost("quant"&index)
				if NOT is_numeric(thequant) then thequant=0
				price=getpost("price"&index)
				if NOT is_numeric(price) then price=0
				wprice=getpost("wprice"&index)
				if NOT is_numeric(wprice) then wprice=0
				wpercent=IIfVr(getpost("wpercent"&index)="1","1","0")
				wholesalepercent=IIfVr(getpost("wholesalepercent"&index)="1","1","0")
				if thequant<>0 AND (price<>0 OR wprice<>0) then
					sSQL = "INSERT INTO pricebreaks (pbProdID,pbQuantity,pPrice,pWholesalePrice,pbPercent,pbWholesalePercent) VALUES ('"&escape_string(theprod)&"',"
					sSQL = sSQL & thequant & ","
					sSQL = sSQL & price & ","
					sSQL = sSQL & wprice & ","
					sSQL = sSQL & wpercent & ","
					sSQL = sSQL & wholesalepercent & ")"
					ect_query(sSQL)
				end if
			next
			dorefresh=TRUE
		end if
	end if
	if dorefresh then
		print "<meta http-equiv=""refresh"" content=""1; url=adminpricebreak.asp?stext=" & urlencode(request("stext")) & "&sort=" & request("sort") & "&stype=" & request("stype") & "&ddown=" & request("ddown") & "&pg=" & request("pg") & """ />"
	end if
end if
%>
<script>
<!--
function setCookie(c_name,value,expiredays){
	var exdate=new Date();
	exdate.setDate(exdate.getDate()+expiredays);
	document.cookie=c_name+ "=" +escape(value)+((expiredays==null) ? "" : ";expires="+exdate.toGMTString());
}
var opensels=[];
var tmrid;
var gtxt;
document.getElementById('main').onclick=function(){
	for(var ii=0; ii<opensels.length; ii++)
		document.getElementById(opensels[ii]).style.display='none';
};
function getvalsfromserver(oSelect){
	oText=document.getElementById('pid');
	if(oSelect.selectedIndex != -1){
		oText.value=oSelect.options[oSelect.selectedIndex].value;
		oSelect.style.display='none';
		oText.focus();
	}
	return false;
}
function comboselect_onclick(oSelect){
	return(getvalsfromserver(oSelect));
}
function comboselect_onchange(oSelect){
	oText=document.getElementById('pid');
	if(oSelect.selectedIndex != -1){
		oText.value=oSelect.options[oSelect.selectedIndex].value;
	}
}
function comboselect_onkeyup(keyCode,oSelect){
	if(keyCode==13){
		getvalsfromserver(oSelect);
	}
	return(false);
}
function plajaxcallback(){
	if(ajaxobj.readyState==4){
		var resarr=ajaxobj.responseText.replace(/^\s+|\s+$/g,"").split('==LISTOBJ==');
		var index,isname=false;
		oSelect=document.getElementById(resarr[0]);
		var act=resarr[0].replace(/\d/g,'');
		for(index=0; index<resarr.length-2; index++){
			var splitelem=resarr[index+1].split('==LISTELM==');
			var val1=splitelem[0];
			var val2=splitelem[1];
			if(index<oSelect.length)
				var y=oSelect.options[index];
			else
				var y=document.createElement('option');
			y.text=val1+(val2!='----------------'?' / '+val2:'');
			y.value=val1;
			if(y.text=='----------------') y.disabled=true; else y.disabled=false;
			if(index>=oSelect.length){
				try{oSelect.add(y, null);} // FF etc
				catch(ex){oSelect.add(y);} // IE
			}
		}
		if(oSelect){
			for(var ii=oSelect.length;ii>=index;ii--){
				oSelect.remove(ii);
			}
		}
	}
}
function populatelist(){
	var stext=gtxt;
	ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
	ajaxobj.onreadystatechange=plajaxcallback;
	ajaxobj.open("POST", "ajaxservice.asp?action=getlist&objid=comboselect&listtype=prodid", true);
	ajaxobj.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
	ajaxobj.send('listtext='+stext);
}
function addopensel(id){
	for(var ii=0; ii<opensels.length; ii++)
		if(id==opensels[ii]) return;
	opensels.push(id);
}
function combochange(oText,e){
	if(document.getElementById("autocomplete").checked==false)
		return;
	keyCode=e.keyCode;
	if(keyCode<32&&keyCode!=8)return true;
	oSelect=document.getElementById('comboselect');
	addopensel(oSelect.id);
	oSelect.style.display='';
	toFind=oText.value.toLowerCase();
	gtxt=toFind;
	clearTimeout(tmrid);
	tmrid=setTimeout("populatelist()",800);
}
function combokey(oText,e){
	if(document.getElementById("autocomplete").checked==false)
		return
	oSelect=document.getElementById('comboselect');
	keyCode=e.keyCode;
	if(keyCode==40 || keyCode==38){ // Up / down arrows
		addopensel(oSelect.id);
		oSelect.style.display='';
		oSelect.focus();
		comboselect_onchange(oSelect);
	}
	else if(keyCode==13){
		oSelect.style.display='none';
		oText.focus();
		return getvalsfromserver(oSelect);
	}
	return true;
}
function focusfield(tfield,tmsg){
tfield.focus();
alert(tmsg);
return(false);
}
function formvalidator(theForm){
var patternprice=/[^0-9\.]/
var patternquant=/[^0-9]/
<% if dropdown then %>
  if (theForm.pid.selectedIndex == 0){
    alert("<%=jscheck(yyPlsSel&" """&yyPrId)%>\".");
<% else %>
  if (theForm.pid.value == ""){
    alert("<%=jscheck(yyPlsEntr&" """&yyPrId)%>\".");
<% end if %>
    theForm.pid.focus();
    return (false);
  }
  for(var index=0;index<<%=maxpricebreaks%>;index++){
	if(document.getElementById('quant'+index)){
		if(patternquant.test(document.getElementById('quant'+index).value)) return(focusfield(document.getElementById('quant'+index),"<%=jscheck(yyOnDig)%>"));
		if(patternprice.test(document.getElementById('price'+index).value)) return(focusfield(document.getElementById('price'+index),"<%=jscheck(yyDecFld)%>"));
		if(patternprice.test(document.getElementById('wprice'+index).value)) return(focusfield(document.getElementById('wprice'+index),"<%=jscheck(yyDecFld)%>"));
		if(document.getElementById('wpercent'+index).checked){
			if(parseFloat(document.getElementById('price'+index).value)>100) return(focusfield(document.getElementById('price'+index),"<%=jscheck("Percentage breaks should be between 0 and 100")%>"));
		}
		if(document.getElementById('wholesalepercent'+index).checked){
			if(parseFloat(document.getElementById('wprice'+index).value)>100) return(focusfield(document.getElementById('wprice'+index),"<%=jscheck("Percentage breaks should be between 0 and 100")%>"));
		}
	}
  }
  return (true);
}
//-->
</script>
<% if getpost("posted")="1" AND (getpost("act")="modify" OR getpost("act")="clone" OR getpost("act")="addnew") then
		if dropdown AND (getpost("act")="clone" OR getpost("act")="addnew") then
			allprodids=""
			sSQL = "SELECT pID FROM products LEFT JOIN pricebreaks ON products.pID=pricebreaks.pbProdID WHERE pbProdID IS NULL ORDER BY pID"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				allprodids=rs.getrows
			end if
			rs.close
		end if
%>
	<table width="100%" border="0" cellspacing="0" cellpadding="1">
	  <tr> 
		<td align="center">
		  <form name="mainform" method="post" action="adminpricebreak.asp" onsubmit="return formvalidator(this)">
			<input type="hidden" name="posted" value="1" />
			<% if getpost("act")="clone" OR getpost("act")="addnew" then %>
			<input type="hidden" name="act" value="doaddnew" />
			<% else %>
			<input type="hidden" name="act" value="domodify" />
			<input type="hidden" name="pid" value="<%=getpost("id")%>" />
			<% end if
			call writehiddenvar("ddown", getpost("ddown"))
			call writehiddenvar("stext", getpost("stext"))
			call writehiddenvar("sort", getpost("sort"))
			call writehiddenvar("stype", getpost("stype"))
			call writehiddenvar("pg", getpost("pg")) %>
            <table width="320" border="0" cellspacing="0" cellpadding="1">
			  <tr> 
                <td colspan="5" align="center"><strong><%=yyPBKAdm%></strong><br />&nbsp;</td>
			  </tr>
			  <tr>
				<td style="text-align:right"><%=yyPrId%>:</td>
				<td colspan="4"><%
				if dropdown AND (getpost("act")="clone" OR getpost("act")="addnew") then
					print "<select size=""1"" name=""pid""><option value="""">"&yySelect&"</option>"
					if IsArray(allprodids) then
						for index=0 to UBOUND(allprodids, 2)
							print "<option value="""&allprodids(0,index)&""">"&allprodids(0,index)&"</option>"&vbCrLf
						next
					end if
					print "</select>"
				elseif (getpost("act")="clone" OR getpost("act")="addnew") then
					print "<input type=""text"" name=""pid"" id=""pid"" size=""30"" autocomplete=""off"" onkeydown=""return combokey(this,event)"" onkeyup=""combochange(this,event)"" />"
					print "<div style=""position:absolute""><select id=""comboselect"" size=""15"" " & _
		"style=""display:none;position:absolute;min-width:280px;top:0px;left:0px;"" " & _
		"onblur=""this.style.display='none'"" " & _
		"onchange=""comboselect_onchange(this)"" " & _
		"onclick=""comboselect_onclick(this)"" " & _
		"onkeyup=""comboselect_onkeyup(event.keyCode,this)"">" & _
		"<option value="""">Populating...</option>" & _
		"</select></div>"
				else
					print htmlspecials(getpost("id"))
				end if %></td>
			  </tr>
<%				if NOT dropdown AND (getpost("act")="clone" OR getpost("act")="addnew") then %>
			  <tr>
				<td style="text-align:right"><%="<input type=""checkbox"" value=""ON"" name=""autocomplete"" id=""autocomplete"" onclick=""setCookie('ectautocomp',this.checked?1:0,600)"" "&IIfVr(request.cookies("ectautocomp")="1","checked=""checked"" ","")&"/>"%></td>
				<td colspan="4"><%=yyUsAuCo%></td>
			  </tr>
<%				end if %>
			  <tr>
				<td align="center"><span style="font-size:10px;font-weight:bold"><%=yyQuaFro%></span></td>
				<td align="center"><span style="font-size:10px;font-weight:bold"><%=yyPrPri%></span></td>
				<td align="left" style="padding-left:10px"><span style="font-size:10px;font-weight:bold" title="Treat as percentage price reduction. Eg &quot;10&quot; would mean 10% off the regular price.">%</span></td>
				<td align="center"><span style="font-size:10px;font-weight:bold"><%=yyWhoPri%></span></td>
				<td align="left" style="padding-left:10px"><span style="font-size:10px;font-weight:bold" title="Treat as percentage price reduction. Eg &quot;10&quot; would mean 10% off the regular price.">%</span></td>
			  </tr>
<%			sSQL = "SELECT pbQuantity,pPrice,pWholesalePrice,pbPercent,pbWholesalePercent FROM pricebreaks WHERE pbProdID='"&escape_string(getpost("id"))&"' ORDER BY pbQuantity"
			index=1
			rs.open sSQL,cnn,0,1
			do while NOT rs.EOF %>
			  <tr>
				<td align="center" style="padding-right:15px"><input type="text" name="quant<%=index%>" id="quant<%=index%>" size="12" value="<%=rs("pbQuantity")%>" /></td>
				<td align="center"><input type="text" name="price<%=index%>" id="price<%=index%>" size="12" value="<%=rs("pPrice")%>" /></td>
				<td align="left" style="padding-right:5px"><input type="checkbox" name="wpercent<%=index%>" id="wpercent<%=index%>" value="1" <% if rs("pbPercent")<>0 then print "checked=""checked"" "%>/></td>
				<td align="center"><input type="text" name="wprice<%=index%>" id="wprice<%=index%>" size="12" value="<%=rs("pWholesalePrice")%>" /></td>
				<td align="left"><input type="checkbox" name="wholesalepercent<%=index%>" id="wholesalepercent<%=index%>" value="1" <% if rs("pbWholesalePercent")<>0 then print "checked=""checked"" "%>/></td>
			  </tr>
<%				rs.MoveNext
				index=index+1
			loop
			rs.close
			for index2=index to maxpricebreaks %>
			  <tr>
				<td align="center" style="padding-right:15px"><input type="text" name="quant<%=index2%>" id="quant<%=index2%>" size="12" value="" /></td>
				<td align="center"><input type="text" name="price<%=index2%>" id="price<%=index2%>" size="12" value="" /></td>
				<td align="left" style="padding-right:15px"><input type="checkbox" name="wpercent<%=index2%>" id="wpercent<%=index2%>" value="1" /></td>
				<td align="center"><input type="text" name="wprice<%=index2%>" id="wprice<%=index2%>" size="12" value="" /></td>
				<td align="left"><input type="checkbox" name="wholesalepercent<%=index2%>" id="wholesalepercent<%=index2%>" value="1" /></td>
			  </tr>
<%			next %>
			  <tr>
                <td width="100%" colspan="5" align="center"><br /><input type="submit" value="<%=yySubmit%>" /></td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="5" align="center"><br /><a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
			  </tr>
            </table>
		  </form>
		</td>
	  </tr>
	</table>
<% elseif getpost("posted")="1" AND success then %>
		<table width="100%" border="0" cellspacing="0" cellpadding="3">
		  <tr> 
			<td width="100%" colspan="2" align="center"><br /><strong><%=yyUpdSuc%></strong><br /><br /><%=yyNowFrd%><br /><br />
					<%=yyNoAuto%><a href="adminpricebreak.asp"><strong><%=yyClkHer%></strong></a>.<br />
					<br />&nbsp;</td>
		  </tr>
		</table>
<% elseif getpost("posted")="1" then %>
		<table width="100%" border="0" cellspacing="0" cellpadding="3">
		  <tr> 
			<td width="100%" colspan="2" align="center"><br /><span style="color:#FF0000;font-weight:bold"><%=yyOpFai%></span><br /><br /><%=errmsg%><br /><br />
			<a href="javascript:history.go(-1)"><strong><%=yyClkBac%></strong></a></td>
		  </tr>
		</table>
<% else
	jscript=""
	sortorder=request("sort")
	modclone=request.cookies("modclone") %>
<script>
<!--
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
function newrec(){
	document.mainform.act.value="addnew";
	document.mainform.submit();
}
function cr(id){
	document.mainform.id.value=id;
	document.mainform.act.value="clone";
	document.mainform.submit();
}
function dr(id){
if(confirm("<%=jscheck(yyConDel)%>\n")){
	document.mainform.id.value=id;
	document.mainform.act.value="delete";
	document.mainform.submit();
}
}
function startsearch(){
	document.mainform.action="adminpricebreak.asp";
	document.mainform.act.value="search";
	document.mainform.posted.value="";
	document.mainform.submit();
}
function changemodclone(modclone){
	setCookie('modclone',modclone[modclone.selectedIndex].value,600);
	startsearch();
}
// -->
</script>
<h2><%=yyAdmQua%></h2>
		  <form name="mainform" method="post" action="adminpricebreak.asp">
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
		<td class="cobhl">&nbsp;</td>
		<td class="cobll" colspan="3" align="center">
				<select name="sort" size="1" style="vertical-align:middle">
				<option value="bid">Sort - Product ID</option>
				<option value="bna"<% if sortorder="bna" then print " selected=""selected"""%>>Sort - Product Name</option>
				</select>
				<input type="submit" value="List Quantity Discounts" onclick="startsearch();" />
				<input type="button" value="<%=yyNewPBK%>" onclick="newrec()" />
				<select name="ddown" size="1" style="vertical-align:middle"><option value="">Text Entry</option><option value="1"<% if request("ddown")="1" then print " selected=""selected"""%>>Dropdown Menu</option></select>
	  </tr>
	</table>
<br />
            <table width="100%" class="stackable admin-table-a sta-white">
<%	if getpost("act")="search" OR getget("pg")<>"" then
		CurPage = 1
		if is_numeric(getget("pg")) then CurPage=int(getget("pg"))
		sSQL = "SELECT DISTINCT pbProdID,pName FROM pricebreaks INNER JOIN products ON pricebreaks.pbProdID=products.pID"
		whereand=" WHERE "
		if trim(request("stext"))<>"" then
			hassearch=TRUE
			Xstext = escape_string(request("stext"))
			aText = Split(Xstext)
			maxsearchindex=1
			aFields(0)="pbProdID"
			aFields(1)="pName"
			if request("stype")="exact" then
				sSQL=sSQL & whereand & "(pbProdID LIKE '%"&Xstext&"%' OR pName LIKE '%"&Xstext&"%') "
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
		if sortorder="bna" then
			sSQL = sSQL & " ORDER BY pName"
		else
			sSQL = sSQL & " ORDER BY pbProdID"
		end if
		rs2.CursorLocation = 3 ' adUseClient
		rs2.CacheSize = maxbreaksperpage
		rs2.Open sSQL,cnn
		if NOT rs2.EOF then
			rs2.MoveFirst
			rs2.PageSize = maxbreaksperpage
			rs2.AbsolutePage = CurPage
			islooping=false
			noproducts=false
			hascatinprodsection=false
			rowcounter=0
			totnumrows=rs2.RecordCount
			iNumOfPages=int((totnumrows + (maxbreaksperpage-1)) / maxbreaksperpage)
			pblink="<a href=""adminpricebreak.asp?stext="&urlencode(request("stext"))&"&stype="&request("stype")&"&ddown="&request("ddown")&"&sort="&sortorder&"&pg="
			if iNumOfPages > 1 then print "<tr><td align=""center"" colspan=""3"">" & writepagebar(CurPage,iNumOfPages,yyPrev,yyNext,pblink,FALSE) & "<br /><br /></td></tr>"
%>				  <tr>
					<th class="maincell"><strong><%=yyPrId%></strong></th>
					<th class="maincell"><strong><%=yyPrName%></strong></th>
					<th class="minicell"><%=yyModify%></th>
				  </tr>
<%			do while NOT rs2.EOF AND rowcounter < maxbreaksperpage
				jscript=jscript&"pa["&rowcounter&"]=[" %>
	<tr id="tr<%=rowcounter%>">
	<td class="maincell"><%=rs2("pbProdID")%></td>
	<td class="maincell"><%=rs2("pName")%></td>
	<td>-</td>
	</tr><%		jscript=jscript&"'"&rs2("pbProdID")&"'];"&vbCrLf
				rowcounter=rowcounter+1
				rs2.MoveNext
			loop
			if iNumOfPages > 1 then print "<tr><td align=""center"" colspan=""3""><br />" & writepagebar(CurPage,iNumOfPages,yyPrev,yyNext,pblink,FALSE) & "</td></tr>"
		else %>
				  <tr><td width="100%" colspan="3" align="center"><br /><%=yyItNone%><br />&nbsp;</td></tr>
<%		end if
		rs2.Close
	else
		numitems=0
		if mysqlserver OR sqlserver then
			sSQL="SELECT COUNT(DISTINCT pbProdID) AS totcount FROM pricebreaks"
		else
			sSQL="SELECT COUNT(*) AS totcount FROM (SELECT DISTINCT pbProdID FROM pricebreaks)"
		end if
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			numitems=rs("totcount")
		end if
		rs.close
		print "<tr><td colspan=""3""><div class=""itemsdefine"">You have " & numitems & " quantity discounts defined.</div></td></tr>"
	end if %>
			  <tr> 
                <td width="100%" colspan="3" align="center"><br /><a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
			  </tr>
            </table>
		  </form>
<script>
/* <![CDATA[ */
var pa=[];
<%=jscript%>
for(var pidind in pa){
	var ttr=document.getElementById('tr'+pidind);
	ttr.cells[2].style.textAlign='center';
	ttr.cells[2].style.whiteSpace='nowrap';
	ttr.cells[2].innerHTML='<input type="button" value="M" style="width:30px;margin-right:4px" onclick="mr(\''+pa[pidind][0]+'\')" title="<%=jsescape(htmlspecials(yyModify))%>" />' +
		'<input type="button" value="C" style="width:30px;margin-right:4px" onclick="cr(\''+pa[pidind][0]+'\')" title="<%=jsescape(htmlspecials(yyClone))%>" />' +
		'<input type="button" value="X" style="width:30px" onclick="dr(\''+pa[pidind][0]+'\')" title="<%=jsescape(htmlspecials(yyDelete))%>" />';
}
/* ]]> */
</script>
<% end if
cnn.Close
set rs = nothing
set rs2 = nothing
set cnn = nothing
%>