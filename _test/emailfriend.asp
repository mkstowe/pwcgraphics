<!--#include file="vsadmin/db_conn_open.asp"-->
<%	storelang=trim(request.querystring("lang"))
	if storelang="de" then %>
<!--#include file="vsadmin/inc/languagefile_de.asp"-->
<% elseif storelang="dk" then %>
<!--#include file="vsadmin/inc/languagefile_dk.asp"-->
<% elseif storelang="es" then %>
<!--#include file="vsadmin/inc/languagefile_es.asp"-->
<% elseif storelang="fr" then %>
<!--#include file="vsadmin/inc/languagefile_fr.asp"-->
<% elseif storelang="it" then %>
<!--#include file="vsadmin/inc/languagefile_it.asp"-->
<% elseif storelang="nl" then %>
<!--#include file="vsadmin/inc/languagefile_nl.asp"-->
<% elseif storelang="pt" then %>
<!--#include file="vsadmin/inc/languagefile_pt.asp"-->
<% else
	storelang="en" %>
<!--#include file="vsadmin/inc/languagefile_en.asp"-->
<% end if %>
<!--#include file="vsadmin/includes.asp"-->
<%
savecodepage=response.codepage
if lcase(adminencoding)<>"utf-8" then response.codepage=65001
response.charset="utf-8"
%>
<!--#include file="vsadmin/inc/incfunctions.asp"-->
<!--#include file="vsadmin/inc/incemail.asp"-->
<%
set rs =Server.CreateObject("ADODB.RecordSet")
set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
getadminsettings()
Response.Buffer=True
Response.Expires=60
Response.Expiresabsolute=Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl="no-cache"
if getpost("posted")<>"1" then print "<div class=""emfmaindiv"" id=""efrcell"">"
if multiemfblockmessage="" then multiemfblockmessage="I'm sorry. We are experiencing temporary difficulties at the moment. Please try again later."
hascustomlayout=FALSE
if detailpagelayout="" OR NOT usecsslayout then detailpagelayout="productimage,productid,manufacturer,sku,productname,discounts,instock,description,listprice,price,currency,options,addtocart,previousnext,emailfriend"&IIfVs(showsearchwords,",searchwords") else hascustomlayout=TRUE
if socialmediabuttons="" then
	if hascustomlayout AND instr(detailpagelayout,"socialmedia")>0 then
		socialmediabuttons="facebook,linkedin,twitter,google,askaquestion"
	elseif useemailfriend OR useaskaquestion then
		socialmediabuttons=IIfVs(useemailfriend,"emailfriend")&IIfVs(useaskaquestion,IIfVs(useemailfriend,",")&"askaquestion")
	end if
end if
if instr(socialmediabuttons,"askaquestion")>0 then useaskaquestion=TRUE
if instr(socialmediabuttons,"emailfriend")>0 then useemailfriend=TRUE
if request("askq")="1" AND useaskaquestion=TRUE then isaskquestion=TRUE else isaskquestion=FALSE
extraparams=0
function checkemfuserblock()
	if blockmultiemf="" then blockmultiemf=20
	multiemfblocked=FALSE
	theip=trim(replace(left(request.servervariables("REMOTE_ADDR"), 48), "'", ""))
	if theip="" then theip="none"
	if blockmultiemf<>"" then
		cnn.Execute("DELETE FROM multibuyblock WHERE lastaccess<" & datedelim & VSUSDateTime(Now()-1) & datedelim)
		sSQL="SELECT ssdenyid,sstimesaccess FROM multibuyblock WHERE ssdenyip='" & "EMF " & theip & "'"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			cnn.Execute("UPDATE multibuyblock SET sstimesaccess=sstimesaccess+1,lastaccess=" & datedelim & VSUSDateTime(Now()) & datedelim & " WHERE ssdenyid=" & rs("ssdenyid"))
			if rs("sstimesaccess") >= blockmultiemf then multiemfblocked=TRUE
		else
			cnn.Execute("INSERT INTO multibuyblock (ssdenyip,lastaccess) VALUES ('" & "EMF " & theip & "'," & datedelim & VSUSDateTime(Now()) & datedelim & ")")
		end if
		rs.Close
	end if
	if theip="none" then
		sSQL="SELECT "&IIfVr(mysqlserver<>true,"TOP 1","")&" dcid FROM ipblocking"&IIfVr(mysqlserver=true," LIMIT 0,1","")
	else
		sSQL="SELECT dcid FROM ipblocking WHERE (dcip1=" & ip2long(theip) & " AND dcip2=0) OR (dcip1 <= " & ip2long(theip) & " AND " & ip2long(theip) & " <= dcip2 AND dcip2 <> 0)"
	end if
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then multiemfblocked=TRUE
	rs.Close
	checkemfuserblock=multiemfblocked
end function
sub checkaqp(aqpindex,taqp)
	if taqp<>"" then
		seBody=seBody & taqp & ": " & replace(left(getpost("askquestionparam" & aqpindex),2000),vbCrLf,emlNl) & emlNl
	end if
end sub
if getpost("posted")="1" then
	success=TRUE
	recaptchasuccess=TRUE
	errormsg=""
	referer=request.servervariables("HTTP_REFERER")
	host=request.servervariables("HTTP_HOST")
	if recaptchaenabled(2) then recaptchasuccess=checkrecaptcha(errormsg)
	if instr(referer, host)=0 then
		xxEFThk="<strong><font color=""#FF0000"">I'm sorry but your email could not be sent at this time.</font></strong>"
		success=FALSE
	elseif NOT recaptchasuccess then
		xxEFThk="reCAPTCHA failure. If you inadvertently refreshed this page, your message will have already been sent.<br />" & errormsg
		success=FALSE
	else
		if htmlemails=true then emlNl="<br />" else emlNl=vbCrLf
		theprodid=left(getpost("efid"),50)
		if useemailfriend<>TRUE AND useaskaquestion<>TRUE then
			xxEFThk="<strong><font color=""#FF0000"">Email Friend / Ask a Question not enabled.</font></strong>"
			success=FALSE
		elseif checkemfuserblock() then
			xxEFThk="<strong><font color=""#FF0000"">" & multiemfblockmessage & "</font></strong>"
			response.status="403 Forbidden"
			response.end
		else
			if isaskquestion AND useaskaquestion=TRUE then
				friendsemail=emailAddr
			elseif useemailfriend=TRUE AND len(getpost("friendsemail"))<50 then
				friendsemail=left(getpost("friendsemail"),50)
			else
				friendsemail=""
			end if
			yourname=left(getpost("yourname"),50)
			youremail=left(getpost("youremail"),50)
			yourcomments=replace(left(getpost("yourcomments"),2000),vbCr,"")
			yourcomments=replace(yourcomments,vbLf,emlNl)
			if isaskquestion then
				seBody="PID: " & getpost("origprodid") & emlNl
				sSQL="SELECT "&getlangid("pName",1)&" FROM products WHERE pID='" & escape_string(getpost("origprodid")) & "'"
				rs.Open sSQL,cnn,0,1
				if NOT rs.EOF then seBody=seBody & "Product: " & rs(getlangid("pName",1)) & emlNl
				rs.Close
				seBody=seBody & xxAskQue & ": " & yourname & emlNl & emlNl & yourcomments & emlNl
				call checkaqp(1,askquestionparam1)
				call checkaqp(2,askquestionparam2)
				call checkaqp(3,askquestionparam3)
				call checkaqp(4,askquestionparam4)
				call checkaqp(5,askquestionparam5)
				call checkaqp(6,askquestionparam6)
				call checkaqp(7,askquestionparam7)
				call checkaqp(8,askquestionparam8)
				call checkaqp(9,askquestionparam9)
				thesubject=xxAsqSub & " from " & yourname & " about " & theprodid
			else
				seBody=xxEFYF1 & yourname & " (" & youremail & ")" & xxEFYF2
				if yourcomments<>"" then
					seBody=seBody & xxEFYF3 & emlNl
					seBody=seBody & yourcomments & emlNl
				else
					seBody=seBody & "." & emlNl
				end if
				produrl=""
				if theprodid<>"" then
					sSQL="SELECT pID,"&getlangid("pName",1)&",pStaticPage,pStaticURL FROM products WHERE pID='" & escape_string(theprodid) & "'"
					rs.Open sSQL,cnn,0,1
					if lcase(adminencoding)<>"utf-8" then response.codepage=savecodepage
					if NOT rs.EOF then produrl=getdetailsurl(rs("pID"),rs("pStaticPage"),rs(getlangid("pName",1)),trim(rs("pStaticURL")&""),"","")
					if lcase(adminencoding)<>"utf-8" then response.codepage=65001
					rs.Close
				end if
				if htmlemails=TRUE then
					storeLink=storeurl
					if getpost("efid") <> "" then storeLink=storeLink & produrl
					seBody=seBody & emlNl & "<a href=""" & storeLink & """>" & storeLink & "</a>"
				else
					seBody=seBody & emlNl & storeurl
					if getpost("efid") <> "" then seBody=seBody & produrl
				end if
				thesubject=yourname & xxEFRec
			end if
			seBody=seBody & emlNl
			if friendsemail<>"" then call DoSendEmailEO(friendsemail,emailAddr,youremail,thesubject,seBody,emailObject,themailhost,theuser,thepass)
		end if
	end if
%>
  <table class="cobtbl emfsubtable" border="0" cellspacing="1" cellpadding="3" width="100%">
	<tr>
	  <td class="cobll emfll"colspan="2" align="center" width="100%"><p>&nbsp;</p>
	  <p><%=IIfVr(isaskquestion AND success,xxAsqThk,xxEFThk)%></p>
	  <p><%=xxClkClo%></p>
	  <p>&nbsp;</p>
	  <%=imageorbutton(imgefclose,xxClsWin,"efclose","document.body.removeChild(document.getElementById('efrdiv'))",TRUE)%>
	  <p>&nbsp;</p>
	  </td>
	</tr>
  </table>
<%
else %>
<form id="efform" method="post" action="emailfriend.asp" onsubmit="return efformvalidator(this)">
  <input type="hidden" name="posted" value="1" />
  <input type="hidden" id="efid" name="efid" value="<%=server.htmlencode(Request.QueryString("id"))%>" />
  <input type="hidden" id="askq" name="askq" value="<%=IIfVr(isaskquestion,"1","")%>" />
  <table class="cobtbl emfsubtable" border="0" cellspacing="1" cellpadding="7" width="100%">
	<tr>
	  <td class="cobhl emfhl" align="center" width="100%" height="30"><%=IIfVr(isaskquestion,xxAskQue,xxEmFrnd)%></td>
	</tr>
	<tr>
		<td class="cobll emfll" width="100%" align="left"><%=IIfVr(isaskquestion,xxAQBlr,xxEFBlr)%><br />
		<br /><%=redstar & xxEFNam%><br /><input type="text" id="yourname" name="yourname" size="30" /><br />
		<%=redstar & xxEFEm%><br /><input type="text" id="youremail" name="youremail" size="30" /><br />
<%	if NOT isaskquestion then %>
		<%=redstar & xxEFFEm%><br /><input type="text" id="friendsemail" name="friendsemail" size="30" /><br />
<%	else
		for index=1 to 9
			execute("askquestionparam=askquestionparam"&index)
			execute("askquestionrequired=askquestionrequired"&index&"")
			execute("askquestionhtml=askquestionhtml"&index)
			if askquestionparam<>"" then
				extraparams=extraparams+1
				if askquestionrequired then print redstar
				print askquestionparam & "<br />"
				if askquestionhtml<>"" then print replace(askquestionhtml,"ectfield","askquestionparam"&index) & "<br />" else print "<input type=""text"" id=""askquestionparam"&index&""" name=""askquestionparam"&index&""" size=""30"" /><br />"
			end if
		next
		theproduct=trim(left(request.querystring("id"),50))
		call writehiddenidvar("origprodid",theproduct)
		sSQL="SELECT "&getlangid("pName",1)&" FROM products WHERE pID='" & escape_string(theproduct) & "'"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			theproduct=rs(getlangid("pName",1))
		end if
		rs.Close
	end if %>
		<%=redstar & xxEFCmt%><br /><textarea id="yourcomments" name="yourcomments" cols="46" rows="6"><%=IIfVr(isaskquestion,htmlspecials(replace(xxAskCom,"%nl%",vbCrLf) & theproduct) & vbCrLf,"")%></textarea><%
	if recaptchaenabled(2) then
		print "<div id=""emfCaptcha"" class=""reCAPTCHAemf""></div>"
	end if
%>		<p align="center"><%
	print imageorbutton(imgefsend,xxSend,"efsend","dosendefdata()",TRUE)
	print "&nbsp;&nbsp;"
	print imageorbutton(imgefclose,xxClsWin,"efclose","document.body.removeChild(document.getElementById('efrdiv'))",TRUE)
%></p>
      </td>
	</tr>
  </table>
</form>
<%
end if
if getpost("posted")<>"1" then print "</div>"
cnn.Close
set rs=nothing
set cnn=nothing
%>
