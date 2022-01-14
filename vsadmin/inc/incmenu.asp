<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if menupoplimit="" then menupoplimit=9
if menuid="" then menuid=""
if vertical2scripts="" then vertical2scripts=""
if popmenuscripts="" then popmenuscripts=""
if commonmenuscripts="" then commonmenuscripts=""
if ormenuroot="" then ormenuroot=0
mAlldata=""
rs=""
cnn=""
arrayem=""
incstoreurl=""
function join2paths(stourl,securl)
	if instr(securl,"://")>0 then
		join2paths=securl
	elseif right(stourl,1)="/" AND left(securl,1)="/" then
		join2paths=stourl&mid(securl,2)
	else
		join2paths=stourl&securl
	end if
end function
if lcase(adminencoding)="iso-8859-1" then raquo="»" else raquo="&raquo;"
if NOT isempty(menuraquo) then raquo=menuraquo
sub mwritemenulevel(id,itlevel,incatalogroot)
	hassub=FALSE
	if itlevel<=menupoplimit then
		if NOT (menucategoriesatroot=2 AND id=ormenuroot) then
			for mIndex=0 TO ubound(mAlldata,2)
				if mAlldata(2,mIndex)=id then
					if (menustyle="horizontalmenu1" OR menustyle="verticalmenu3") AND NOT hassub then
						response.write "<ul id=""ecttop"&menuid&"_"&id&""" style=""list-style:none;margin:0px;border:0px;"&IIfVs(id<>ormenuroot,"display:none;position:absolute")&""" class=""ect"&IIfVs(id<>ormenuroot,"sub")&menustyle&" ectmenu"&(menuid+1)&IIfVs(id<>ormenuroot," ectsub"&(menuid+1))&""">" & vbCrLf
					end if
					hassub=TRUE
					mTID=mAlldata(2,mIndex)
					if mTID=0 then mTID=""
					if menustyle="horizontalmenu1" OR menustyle="verticalmenu3" then
						if NOT (menucategoriesatroot AND mAlldata(0,mIndex)=catalogroot) then
							arrayem=arrayem&mAlldata(0,mIndex)&":"&mAlldata(2,mIndex)&","
							response.write "<li id=""ect"&menuid&"_"&mAlldata(0,mIndex)&""" class=""ectmenu"&(menuid+1)&IIfVs(id<>ormenuroot," ectsub"&(menuid+1))&""""&IIfVs(menustyle<>"verticalmenu3"," style="""&IIfVr(id<>ormenuroot,"margin-bottom:-1px","display:inline;margin-right:-1px")&"""")&">"
							if trim(mAlldata(4,mIndex)&"")<>"" then
								if incatalogroot AND instr(mAlldata(4,mIndex),"://")=0 then caturl=getcatid(mAlldata(4,mIndex),IIfVs(seocategoryurls,mAlldata(4,mIndex)),IIfVr(mAlldata(3,mIndex)=1,seoprodurlpattern,seocaturlpattern)) else caturl=mAlldata(4,mIndex)
								response.write "<a href="""&join2paths(incstoreurl,caturl)&""">"&replace(mAlldata(1,mIndex)&"","<","&lt;")&"</a>"
							else
								if mAlldata(3,mIndex)=0 then
									response.write "<a href="""&join2paths(incstoreurl,IIfVs(NOT seocategoryurls,"categories"&extension&"?cat=")&getcatid(mAlldata(0,mIndex),mAlldata(1,mIndex),seocaturlpattern))&""">"&replace(mAlldata(1,mIndex)&"","<","&lt;")&"</a>"
								else
									response.write "<a href="""&join2paths(incstoreurl,IIfVs(NOT seocategoryurls,"products"&extension&"?cat=")&getcatid(mAlldata(0,mIndex),mAlldata(1,mIndex),seoprodurlpattern))&""">"&replace(mAlldata(1,mIndex)&"","<","&lt;")&"</a>"
								end if
							end if
							response.write " </li>" & vbLf
						end if
					else
						menuheadsec="mymenu.addSubMenu(""products"&mTID&""","
						if menucategoriesatroot=1 then menuheadsec="mymenu.addMenu("
						if trim(mAlldata(4,mIndex)&"")<>"" then
							response.write menuheadsec&"""products"&mAlldata(0,mIndex)&""","""&menuprestr&replace(mAlldata(1,mIndex)&"","""","\""")&menupoststr&""","""&join2paths(incstoreurl,mAlldata(4,mIndex))&""");"&vbCrLf
						else
							if mAlldata(3,mIndex)=0 then
								response.write menuheadsec&"""products"&mAlldata(0,mIndex)&""","""&menuprestr&replace(mAlldata(1,mIndex)&"","""","\""")&menupoststr&""","""&join2paths(incstoreurl,IIfVs(NOT seocategoryurls,"categories"&extension&"?cat=")&getcatid(mAlldata(0,mIndex),mAlldata(1,mIndex),seocaturlpattern))&""");"&vbCrLf
							else
								response.write menuheadsec&"""products"&mAlldata(0,mIndex)&""","""&menuprestr&replace(mAlldata(1,mIndex)&"","""","\""")&menupoststr&""","""&join2paths(incstoreurl,IIfVs(NOT seocategoryurls,"products"&extension&"?cat=")&getcatid(mAlldata(0,mIndex),mAlldata(1,mIndex),seoprodurlpattern))&""");"&vbCrLf
							end if
						end if
					end if
				end if
			next
			if (menustyle="horizontalmenu1" OR menustyle="verticalmenu3") AND hassub then response.write "</ul>"
		end if
		for mIndex=0 to ubound(mAlldata,2)
			if mAlldata(2,mIndex)=id AND mAlldata(3,mIndex)=0 AND menucategoriesatroot<>1 then call mwritemenulevel(mAlldata(0,mIndex),itlevel+1,incatalogroot OR mAlldata(0,mIndex)=catalogroot)
		next
	end if
end sub
function mstrdpth(mstr,dep)
mstrdpth=""
for index=2 to dep
	mstrdpth=mstrdpth&mstr&" "
next
end function
sub cssmenulevel(id,itlevel,incatalogroot)
	Dim mIndex
	if itlevel<=menupoplimit then
		for mIndex=0 TO ubound(mAlldata,2)
			if mAlldata(2,mIndex)=id AND NOT (menucategoriesatroot AND mAlldata(0,mIndex)=catalogroot) then
				arrayem=arrayem&mAlldata(0,mIndex)&":"&mAlldata(2,mIndex)&","
				if mAlldata(3,mIndex)=0 then
					if itlevel=menupoplimit then mlink=join2paths(incstoreurl,IIfVs(NOT seocategoryurls,"categories"&extension&"?cat=")&getcatid(mAlldata(0,mIndex),mAlldata(1,mIndex),seocaturlpattern)) else mlink="#"
				elseif trim(mAlldata(4,mIndex)&"")<>"" then
					if incatalogroot AND instr(mAlldata(4,mIndex),"://")=0 then caturl=getcatid(mAlldata(4,mIndex),IIfVs(seocategoryurls,mAlldata(4,mIndex)),seoprodurlpattern) else caturl=mAlldata(4,mIndex)
					mlink=join2paths(incstoreurl,caturl)
				else
					mlink=join2paths(incstoreurl,IIfVs(NOT seocategoryurls,"products"&extension&"?cat=")&getcatid(mAlldata(0,mIndex),mAlldata(1,mIndex),seoprodurlpattern))
				end if
				response.write "<li"&IIfVs(id<>ormenuroot," class=""ectsub ectsub"&(menuid+1)&"""")&" id=""ect"&menuid&"_"&mAlldata(0,mIndex)&"""><a href=""" & mlink & """>" & IIfVs(id<>ormenuroot,mstrdpth(raquo,itlevel)) & replace(mAlldata(1,mIndex)&"","<","&lt;") & "</a></li>" & vbCrLf
				call cssmenulevel(mAlldata(0,mIndex),itlevel+1,incatalogroot OR mAlldata(0,mIndex)=catalogroot)
			end if
		next
	end if
end sub
sub writesubmenus()
	menucategoriesatroot=2
	call mwritemenulevel(ormenuroot,2,catalogroot=ormenuroot)
end sub
function displayectmenu(menstyle)
	if menuid="" then menuid=0 else menuid=menuid+1
	menustyle=menstyle
	if sDSN<>"" then
		Set rs=Server.CreateObject("ADODB.RecordSet")
		Set cnn=Server.CreateObject("ADODB.Connection")
		cnn.open sDSN
		alreadygotadmin=getadminsettings()
		if (request.servervariables("HTTPS")="on" OR request.servervariables("SERVER_PORT_SECURE")="1" OR request.servervariables("SERVER_PORT")="443") AND instr(storeurl,"https:")=0 then incstoreurl=storeurl else incstoreurl=""
		if SESSION("clientLoginLevel")<>"" then minloglevel=SESSION("clientLoginLevel") else minloglevel=0
		rs.open "SELECT sectionID,"&getlangid("sectionName",256)&",topSection,rootSection,"&getlangid("sectionurl",2048)&" FROM sections WHERE sectionID<>0 AND sectionDisabled<="&minloglevel&IIfVs(menupoplimit<=1," AND topSection=0")&" ORDER BY "&IIfVr(sortcategoriesalphabetically=TRUE, getlangid("sectionName",256), "sectionOrder")&IIfVs(menustyle="verticalmenu2",",topSection"),cnn,0,1
		if NOT rs.EOF then mAlldata=rs.getrows
		rs.close
		cnn.Close
		set rs=nothing
		set cnn=nothing
		if isarray(mAlldata) then
			if menucategoriesatroot AND (menustyle="verticalmenu2" OR menustyle="horizontalmenu1" OR menustyle="verticalmenu3") then
				theroot=catalogroot
				for mIndex=0 to ubound(mAlldata,2)
					if mAlldata(0,mIndex)=catalogroot then theroot=mAlldata(2,mIndex) : exit for
				next
				for mIndex=0 to ubound(mAlldata,2)
					if mAlldata(2,mIndex)=catalogroot then mAlldata(2,mIndex)=theroot
				next
			end if
			if menustyle="verticalmenu2" then
				response.write "<ul class=""ect"&menustyle&" ectmenu"&(menuid+1)&""" style=""list-style:none"">"
				call cssmenulevel(ormenuroot,1,catalogroot=ormenuroot)
				response.write "</ul>"
			elseif menustyle="horizontalmenu1" OR menustyle="verticalmenu3" then
				call mwritemenulevel(ormenuroot,1,catalogroot=ormenuroot)
			else
				call mwritemenulevel(ormenuroot,1,catalogroot=ormenuroot)
			end if
		end if
	end if
	if menustyle="horizontalmenu1" OR menustyle="verticalmenu2" OR menustyle="verticalmenu3" then %>
<script>
/* <![CDATA[ */
<%		if menuid=0 then response.write "var curmen=[];var lastmen=[];var em=[];var emt=[];"
		writemenuscripts()
		response.write "em["&menuid&"]={"&left(arrayem,len(arrayem)-1)&"};emt["&menuid&"]=[];curmen["&menuid&"]=0;"&vbCrLf
		response.write "addsubsclass("&menuid&",0,'"&menustyle&"')"&vbCrLf
		response.write "/* ]]> */</script>"
	end if
	displayectmenu=""
end function
function writemenuscripts()
	if menustyle<>"verticalmenu2" AND popmenuscripts="" then
		popmenuscripts=TRUE %>
function closepopdelay(menid){
	var re=new RegExp('ect\\d+_');
	var theid=menid.replace(re,'');
	var mennum=menid.replace('ect','').replace(/_\d+/,'');
	for(var ei in emt[mennum]){
		if(ei!=0&&emt[mennum][ei]==true&&!insubmenu(ei,mennum)){
			document.getElementById('ecttop'+mennum+"_"+ei).style.display='none';
			emt[mennum][ei]=false; // closed
		}
	}
}
function closepop(men){
	var mennum=men.id.replace('ect','').replace(/_\d+/,'');
	lastmen[mennum]=curmen[mennum];
	curmen[mennum]=0;
	setTimeout("closepopdelay('"+men.id+"')",1000);
}
function getPos(el){
	for (var lx=0,ly=0; el!=null; lx+=el.offsetLeft,ly+=el.offsetTop, el=el.offsetParent){
	};
	return{x:lx,y:ly};
}
function openpop(men,ispopout){
	var re=new RegExp('ect\\d+_');
	var theid=men.id.replace(re,'');
	var mennum=men.id.replace('ect','').replace(/_\d+/,'');
	curmen[mennum]=theid;
	if(lastmen[mennum]!=0)
		closepopdelay('ect'+mennum+'_'+lastmen[mennum]);
	if(mentop=document.getElementById('ecttop'+mennum+'_'+theid)){
		var px=getPos(men);
		if(em[mennum][theid]==0&&!ispopout){
			mentop.style.left=px.x+'px';
			mentop.style.top=(px.y+men.offsetHeight-1)+'px';
			mentop.style.display='';
		}else{
			mentop.style.left=(px.x+men.offsetWidth-1)+'px';
			mentop.style.top=px.y+'px';
			mentop.style.display='';
		}
		emt[mennum][theid]=true; // open
	}
}
<%	end if
	if menustyle="verticalmenu2" AND vertical2scripts="" then
		vertical2scripts=TRUE
%>
function closecascade(men){
	var re=new RegExp('ect\\d+_');
	var theid=men.id.replace(re,'');
	var mennum=men.id.replace('ect','').replace(/_\d+/,'');
	curmen[mennum]=0;
	for(var ei in emt[mennum]){
		if(ei!=0&&emt[mennum][ei]==true&&!insubmenu(ei,mennum)){
			for(var ei2 in em[mennum]){
				if(em[mennum][ei2]==ei){
					document.getElementById('ect'+mennum+"_"+ei2).style.display='none';
				}
			}
		}
	}
	emt[mennum][theid]=false; // closed
	return(false);
}
function opencascade(men){
	var re=new RegExp('ect\\d+_');
	var theid=men.id.replace(re,'');
	var mennum=men.id.replace('ect','').replace(/_\d+/,'');
	if(emt[mennum][theid]==true) return(closecascade(men));
	curmen[mennum]=theid;
	for(var ei in em[mennum]){
		if(em[mennum][ei]==theid){
			document.getElementById('ect'+mennum+'_'+ei).style.display='block';
			emt[mennum][theid]=true; // open
		}
	}
	return(false);
}
function ectChCk(men){
return(hassubs(men)?opencascade(men):true)
}
<%	end if
	if commonmenuscripts="" then
		commonmenuscripts=TRUE %>
function hassubs(men){
	var re=new RegExp('ect\\d+_');
	var theid=men.id.replace(re,'');
	var mennum=men.id.replace('ect','').replace(/_\d+/,'');
	for(var ei in em[mennum]){
		if(em[mennum][ei]==theid)
			return(true);
	}
	return(false);
}
function insubmenu(mei,mid){
	if(curmen[mid]==0)return(false);
	curm=curmen[mid];
	maxloops=0;
	while(curm!=0){
		if(mei==curm)return(true);
		curm=em[mid][curm];
		if(maxloops++>10) break;
	}
	return(false);
}
function addsubsclass(mennum,menid,menutype){
	for(var ei in em[mennum]){
		men=document.getElementById('ect'+mennum+'_'+ei);
		if(menutype=='verticalmenu2')
			men.onclick=function(){return(ectChCk(this))};
		else{
			men.onmouseover=function(){openpop(this,menutype=='verticalmenu3'?true:false)};
			men.onmouseout=function(){closepop(this)};
		}
		if(hassubs(men)){
			if(men.className.indexOf('ectmenuhassub')==-1)men.className+=' ectmenuhassub'+(mennum+1);
		}
	}
}
<%	end if
end function ' writemenuscripts
displayectmenu(menustyle)
%>