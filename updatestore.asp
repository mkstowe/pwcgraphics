<%@ Language="VBSCRIPT"%>
<%
' This code is copyright ViciSoft SL.
' Unauthorized copying, use or transmittal without the
' express permission of ViciSoft SL is strictly prohibited.
' Author: Vince Reid, vince@virtualred.net
Response.Buffer = True
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Cache-Control", "must-revalidate"
Response.AddHeader "Cache-Control", "no-cache"
Dim sVersion,rs,cnn,sSQL,errnum,index
sVersion="v7.4.4"
success=true
%><!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">
<head>
<title>Update Ecommerce Plus to <%=sVersion%></title>
<style type="text/css">
<!--
p {  font: 11pt  Arial, Helvetica, sans-serif}
BODY {  font: 11pt Arial, Helvetica, sans-serif}
-->
</style>
<!--#include file="vsadmin/inc/md5.asp"-->
<!--#include file="vsadmin/db_conn_open.asp"-->
<!--#include file="vsadmin/includes.asp"-->
<% if adminencoding="" then adminencoding="iso-8859-1" %>
<meta http-equiv="Content-Type" content="text/html; charset=<%=adminencoding%>"/>
</head>
<body style="background:#f5f7fa">
<div style="padding:24px;border:1px solid #e0e0e0;width:680px;margin:0 auto;background:#FFF;-moz-border-radius:10px;-webkit-border-radius:10px;margin-top:40px;">
<%
if mysqlserver=true then
	addcl = "ADD COLUMN"
	txtcl = "VARCHAR"
	idtxtcl = "VARCHAR"
	smallcl = "SMALLINT"
	bytecl = "TINYINT"
	dblcl = "DOUBLE"
	memocl = "TEXT"
	datecl = "DATETIME"
	smalldatecl = "DATE"
	autoinc = "INT NOT NULL AUTO_INCREMENT PRIMARY KEY"
	datedelim="'"
	bitfield="TINYINT(1)"
	altcl="MODIFY"
	txtcollen=8000
elseif sqlserver=true then
	addcl = "ADD"
	txtcl = "VARCHAR"
	idtxtcl = "VARCHAR"
	smallcl = "SMALLINT"
	bytecl = "TINYINT"
	dblcl = "FLOAT(32)"
	memocl = "VARCHAR(MAX)"
	datecl = "DATETIME"
	smalldatecl = "DATE"
	autoinc = "INT NOT NULL IDENTITY(1, 1) PRIMARY KEY"
	datedelim="'"
	bitfield="bit"
	altcl="ALTER COLUMN"
	txtcollen=8000
else
	addcl = "ADD COLUMN"
	txtcl = "TEXT"
	idtxtcl = "TEXT"
	smallcl = "SHORT"
	bytecl = "BYTE"
	dblcl = "DOUBLE"
	memocl = "MEMO"
	datecl = "DATETIME"
	smalldatecl = "SMALLDATETIME"
	autoinc = "AUTOINCREMENT PRIMARY KEY"
	datedelim="#"
	bitfield="bit"
	altcl="ALTER COLUMN"
	txtcollen=255
end if
sub dropconstraintcolumn(tbl,tcol)
	on error resume next
	call drop_constraints(tbl,tcol)
	cnn.execute("ALTER TABLE " & tbl & " DROP COLUMN " & tcol)
	on error goto 0
end sub
sub drop_constraints(thetable,thecolumn)
	if sqlserver=TRUE then
		sSQL = "SELECT OBJECT_NAME(PARENT_OBJECT_ID) AS table_name,COL_NAME(PARENT_OBJECT_ID, PARENT_COLUMN_ID) AS column_name,NAME AS the_constraint_name FROM SYS.DEFAULT_CONSTRAINTS WHERE OBJECT_NAME(PARENT_OBJECT_ID)='"&thetable&"' AND COL_NAME(PARENT_OBJECT_ID, PARENT_COLUMN_ID)='"&thecolumn&"'"
		rs.Open sSQL,cnn,0,1
		do while NOT rs.EOF
			sSQL="ALTER TABLE "&thetable&" DROP CONSTRAINT " & rs("the_constraint_name")
			cnn.execute(sSQL)
			rs.movenext
		loop
		rs.close
	end if
end sub
function is_numeric(tstr)
	is_numeric=isnumeric(trim(tstr&"")) AND instr(trim(tstr&""),",")=0
end function
function VSUSDate(thedate)
	if mysqlserver=true then
		VSUSDate = DatePart("yyyy",thedate) & "-" & DatePart("m",thedate) & "-" & DatePart("d",thedate)
	elseif sqlserver=true then
		VSUSDate = right(DatePart("yyyy",thedate),2) & IIfVr(DatePart("m",thedate)<10,"0","") & DatePart("m",thedate) & IIfVr(DatePart("d",thedate)<10,"0","") & DatePart("d",thedate)
	else
		VSUSDate = DatePart("m",thedate) & "/" & DatePart("d",thedate) & "/" & DatePart("yyyy",thedate)
	end if
end function
function VSUSDateTime(thedate)
	if mysqlserver=true then
		VSUSDateTime = DatePart("yyyy",thedate) & "-" & DatePart("m",thedate) & "-" & DatePart("d",thedate) & " " & DatePart("h",thedate) & ":" & DatePart("n",thedate) & ":" & DatePart("s",thedate)
	elseif sqlserver=true then
		VSUSDateTime = right(DatePart("yyyy",thedate),2) & IIfVr(DatePart("m",thedate)<10,"0","") & DatePart("m",thedate) & IIfVr(DatePart("d",thedate)<10,"0","") & DatePart("d",thedate) & " " & DatePart("h",thedate) & ":" & DatePart("n",thedate) & ":" & DatePart("s",thedate)
	else
		VSUSDateTime = DatePart("m",thedate) & "/" & DatePart("d",thedate) & "/" & DatePart("yyyy",thedate) & " " & DatePart("h",thedate) & ":" & DatePart("n",thedate) & ":" & DatePart("s",thedate)
	end if
end function
function escape_string(str)
	escape_string = trim(replace(str&"","'","''"))
end function
function dohashpw(thepw)
	if trim(thepw&"")="" then dohashpw="" else dohashpw=calcmd5("ECT IS BEST"&trim(thepw))
end function
function IIfVr(theExp,theTrue,theFalse)
if theExp then IIfVr=theTrue else IIfVr=theFalse
end function
function IIfVs(theExp,theTrue)
if theExp then IIfVs=theTrue else IIfVs=""
end function
sub printtick(tstr)
	response.write "<script type=""text/javascript"">iqueue.push('B" & replace(tstr,"'","\'") & "');</script>" & vbCrLf
end sub
sub printtickdiv(tstr)
	response.write "<script type=""text/javascript"">iqueue.push('A" & replace(tstr,"'","\'") & "');</script>" & vbCrLf
	response.flush
end sub
function checkaddcolumn(tblname,colname,notnull,dtype,dlen,setdef)
	printtickdiv("Checking for " & colname & " in table " & tblname)
	sSQL="SELECT " & IIfVr(mysqlserver<>TRUE,"TOP 1 ","") & colname & " FROM " & tblname & IIfVr(mysqlserver=TRUE," LIMIT 0,1","")
	on error resume next
	err.number = 0
	rs.Open sSQL,cnn,0,1
	errnum=err.number
	rs.Close
	on error goto 0
	if errnum<>0 then
		printtick("Adding " & colname & " column to " & tblname & " table")
		defval=""
		if dtype="INT" OR dtype=dblcl OR dtype=smallcl OR dtype=bytecl OR dtype=bitfield then defval="DEFAULT 0" : setdef=IIfVr(setdef="","0",setdef) : notnull=TRUE
		if dtype=txtcl OR dtype=idtxtcl then defval="DEFAULT ''" : setdef=IIfVr(setdef="","''",setdef)
		if dtype=memocl then defval="x" : setdef=IIfVr(setdef="","''",setdef)
		if dtype=datecl OR dtype=smalldatecl then
			if sqlserver=TRUE then
				defval="DEFAULT GETDATE()"
			elseif mysqlserver=TRUE then
				defval="x"
			else
				defval="DEFAULT NOW()"
			end if
			setdef=datedelim&vsusdate(Now()-10)&datedelim
		end if
		if defval="" then response.write "<font color=""#FF0000"">" & dtype & " not supported!!<br /></div></body></html>" : response.end
		if defval="x" then defval=""
		sSQL = "ALTER TABLE " & tblname & " " & addcl & " " & colname & " " & dtype & dlen & " " & IIfVR(notnull,"NOT NULL ","") & defval
		' response.write sSQL & "<br />"
		cnn.execute(sSQL)
		if setdef<>"" then cnn.execute("UPDATE " & tblname & " SET " & colname & "=" & setdef)
		checkaddcolumn=TRUE
	else
		checkaddcolumn=FALSE
	end if
	response.flush
end function
if sqlserver=TRUE OR mysqlserver=TRUE then
	if instr(sDSN,"Provider=Microsoft.Jet")>0 then response.write "<br />&nbsp;<br /><div style=""font-weight:bold;color:#FF0000;text-align:center;"">WARNING!! You have the switch set for an SQL Server or mySQL Server database, but appear to be using a MS Access database!!</div>"
end if

Set rs =Server.CreateObject("ADODB.RecordSet")
Set rs2=Server.CreateObject("ADODB.RecordSet")
Set rs3=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN

if request.form("posted")="1" then

on error resume next
Server.ScriptTimeout = 1800
%>
<script type="text/javascript">
var iqueue = [];
function writetickitem(sitm){
	var msgtype=sitm.substr(0,1);
	if(msgtype=='A')
		document.getElementById('checkdiv').innerHTML=sitm.substr(1);
	else if(msgtype=='B'){
		var thetable = document.getElementById('resulttable');
		newrow = thetable.insertRow(-1);
		newcell = newrow.insertCell(0);
		newcell.innerHTML='<img src="https://www.ecommercetemplates.com/images/ecttick.gif"> ';
		newcell = newrow.insertCell(1);
		newcell.innerHTML=sitm.substr(1);
	}else if(msgtype=='C'){
		clearInterval(intid);
		setTimeout("document.location='updatestore.asp?posted=2'",8000)
	}
}
function checkqueue(){
	if(iqueue.length>0){
		writetickitem(iqueue.shift());
	}
}
var intid=setInterval("checkqueue()", 30);
</script>
<table>
	<tr>
		<td width="20" align="right"><img src="https://www.ecommercetemplates.com/images/ecttick.gif"> </td>
		<td><div id="checkdiv">Checking for Ecommerce Plus Template...</div></td>
	</tr>
</table>
<table id="resulttable">
	<tr>
		<td width="20" align="right">&nbsp;</td><td>&nbsp;</td>
	</tr>
</table>
<%
err.number = 0
sSQL = "SELECT * FROM postalzones"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	response.write "ERROR ! ! ERROR ! ! ERROR ! ! <br>This does not seem to be an Ecommerce Plus templates. Quitting...<br>ERROR ! ! ERROR ! ! ERROR ! ! <br></body></html>"
	response.end
end if

on error resume next
printtickdiv("Checking for Email Object upgrade")
err.number = 0
sSQL = "SELECT emailObject FROM admin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Email Object columns")
	sSQL = "ALTER TABLE admin "&addcl&" emailObject "&bytecl&" DEFAULT 0"
	cnn.execute(sSQL)
	mailissent = false
	emailobject = 0
	on error resume next
	err.number = 0
	' First check for CDONTS.
	if mailhost = "your.mailserver.com" OR mailhost = "" then
		Set EmailObj = Server.CreateObject("CDO.Message")
		if err.number = 0 then
			emailobject = 1
			mailissent = true
		end if
		if NOT mailissent then
			err.number = 0
			Set EmailObj = Server.CreateObject("CDONTS.NewMail")
			if err.number = 0 then
				emailobject = 0
				mailissent = true
			end if
		end if
	end if
	if NOT mailissent AND mailhost <> "your.mailserver.com" AND mailhost <> "" then
		err.number = 0
		Set EmailObj = Server.CreateObject("Persits.MailSender")
		if err.number = 0 then
			emailobject = 2
			mailissent = true
		end if
	end if
	if NOT mailissent AND mailhost <> "your.mailserver.com" AND mailhost <> "" then
		err.number = 0
		Set EmailObj = Server.CreateObject("SMTPsvg.Mailer")
		if err.number = 0 then
			emailobject = 3
			mailissent = true
		end if
	end if
	if NOT mailissent AND mailhost <> "your.mailserver.com" AND mailhost <> "" then
		err.number = 0
		Set EmailObj = Server.CreateObject("JMail.SMTPMail")
		if err.number = 0 then
			emailobject = 4
			mailissent = true
		end if
	end if
	Set EmailObj = nothing
	on error goto 0
	cnn.execute("UPDATE admin SET emailObject=" & emailobject)
end if

on error resume next
printtickdiv("Checking for Order Status upgrade")
err.number = 0
sSQL = "SELECT * FROM orderstatus"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Order Status table")
	cnn.execute("CREATE TABLE orderstatus (statID INT PRIMARY KEY,statPrivate "&txtcl&"(255) NULL,statPublic "&txtcl&"(255) NULL)")
	cnn.execute("INSERT INTO orderstatus (statID,statPrivate,statPublic) VALUES (0,'Cancelled','Order Cancelled')")
	cnn.execute("INSERT INTO orderstatus (statID,statPrivate,statPublic) VALUES (1,'Deleted','Order Deleted')")
	cnn.execute("INSERT INTO orderstatus (statID,statPrivate,statPublic) VALUES (2,'Unauthorized','Awaiting Payment')")
	cnn.execute("INSERT INTO orderstatus (statID,statPrivate,statPublic) VALUES (3,'Authorized','Payment Received')")
	cnn.execute("INSERT INTO orderstatus (statID,statPrivate,statPublic) VALUES (4,'Packing','In Packing')")
	cnn.execute("INSERT INTO orderstatus (statID,statPrivate,statPublic) VALUES (5,'Shipping','In Shipping')")
	cnn.execute("INSERT INTO orderstatus (statID,statPrivate,statPublic) VALUES (6,'Shipped','Order Shipped')")
	cnn.execute("INSERT INTO orderstatus (statID,statPrivate,statPublic) VALUES (7,'Completed','Order Completed')")
	cnn.execute("INSERT INTO orderstatus (statID,statPrivate,statPublic) VALUES (8,'','')")
	cnn.execute("INSERT INTO orderstatus (statID,statPrivate,statPublic) VALUES (9,'','')")
	cnn.execute("INSERT INTO orderstatus (statID,statPrivate,statPublic) VALUES (10,'','')")
	cnn.execute("INSERT INTO orderstatus (statID,statPrivate,statPublic) VALUES (11,'','')")
	cnn.execute("INSERT INTO orderstatus (statID,statPrivate,statPublic) VALUES (12,'','')")
	cnn.execute("INSERT INTO orderstatus (statID,statPrivate,statPublic) VALUES (13,'','')")
	cnn.execute("INSERT INTO orderstatus (statID,statPrivate,statPublic) VALUES (14,'','')")
	cnn.execute("INSERT INTO orderstatus (statID,statPrivate,statPublic) VALUES (15,'','')")
	cnn.execute("INSERT INTO orderstatus (statID,statPrivate,statPublic) VALUES (16,'','')")
	cnn.execute("INSERT INTO orderstatus (statID,statPrivate,statPublic) VALUES (17,'','')")
end if
response.flush

on error resume next
printtickdiv("Checking for multisections table")
err.number = 0
sSQL = "SELECT * FROM multisections"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding multisections table")
	cnn.execute("CREATE TABLE multisections (pID "&txtcl&"(128) NOT NULL,pSection INT DEFAULT 0 NOT NULL, PRIMARY KEY(pID,pSection))")
end if

call checkaddcolumn("products","pWholesalePrice",FALSE,dblcl,"","")
call checkaddcolumn("orders","ordExtra1",FALSE,txtcl,"(255)","")
call checkaddcolumn("orders","ordExtra2",FALSE,txtcl,"(255)","")
call checkaddcolumn("orders","ordHSTTax",FALSE,dblcl,"","")

call checkaddcolumn("admin","smtpserver",FALSE,txtcl,"(100)","")
call checkaddcolumn("admin","emailUser",FALSE,txtcl,"(50)","")
call checkaddcolumn("admin","emailPass",FALSE,txtcl,"(50)","")

on error resume next
printtickdiv("Checking for Order Status orders upgrade")
err.number = 0
sSQL = "SELECT ordStatus FROM orders"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Order Status orders columns")
	cnn.execute("ALTER TABLE orders "&addcl&" ordStatus "&bytecl&" DEFAULT 0")
	cnn.execute("ALTER TABLE orders "&addcl&" ordStatusDate "&datecl)
	cnn.execute("ALTER TABLE orders "&addcl&" ordStatusInfo "&memocl&" NULL")
	cnn.execute("UPDATE orders SET ordStatus=2")
	cnn.execute("UPDATE orders SET ordStatus=3 WHERE ordAuthNumber<>''")
	cnn.execute("UPDATE orders SET ordStatusDate=ordDate")
end if
response.flush

on error resume next
printtickdiv("Checking for Currency Conversions upgrade")
err.number = 0
sSQL = "SELECT currRate1 FROM admin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Currency Conversions columns")
	cnn.execute("ALTER TABLE admin "&addcl&" currRate1 "&dblcl&" DEFAULT 0")
	cnn.execute("ALTER TABLE admin "&addcl&" currRate2 "&dblcl&" DEFAULT 0")
	cnn.execute("ALTER TABLE admin "&addcl&" currRate3 "&dblcl&" DEFAULT 0")
	cnn.execute("ALTER TABLE admin "&addcl&" currSymbol1 "&txtcl&"(50) NULL")
	cnn.execute("ALTER TABLE admin "&addcl&" currSymbol2 "&txtcl&"(50) NULL")
	cnn.execute("ALTER TABLE admin "&addcl&" currSymbol3 "&txtcl&"(50) NULL")
	cnn.execute("UPDATE admin SET currRate1=0,currRate2=0,currRate3=0,currSymbol1='',currSymbol2='',currSymbol3=''")
end if
response.flush

call checkaddcolumn("admin","currConvUser",FALSE,txtcl,"(50)","")
call checkaddcolumn("admin","currConvPw",FALSE,txtcl,"(50)","")
call checkaddcolumn("admin","currLastUpdate",FALSE,datecl,"","")

on error resume next
printtickdiv("Checking for pay provider method upgrade")
err.number = 0
sSQL = "SELECT payProvMethod FROM payprovider"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding pay provider method column")
	cnn.execute("ALTER TABLE payprovider "&addcl&" payProvMethod INT DEFAULT 0")
	cnn.execute("UPDATE payprovider SET payProvMethod=0")
	sSQL = "SELECT payProvData2 FROM payprovider WHERE payProvID=11"
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then ppdata2 = Trim(rs("payProvData2"))
	rs.Close
	if ppdata2<>"" then cnn.execute("UPDATE payprovider SET payProvMethod=" & ppdata2 & " WHERE payProvID=11")
	sSQL = "SELECT payProvData2 FROM payprovider WHERE payProvID=12"
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then ppdata2 = Trim(rs("payProvData2"))
	rs.Close
	if ppdata2<>"" then cnn.execute("UPDATE payprovider SET payProvMethod=" & ppdata2 & " WHERE payProvID=12")
	cnn.execute("UPDATE payprovider SET payProvData2='' WHERE payProvID=11 OR payProvID=12")
end if
response.flush

' call checkaddcolumn("admin","adminUPSLicense",FALSE,memocl,"","")

call checkaddcolumn("orders","ordComLoc",FALSE,bytecl,"","")

on error resume next
printtickdiv("Checking for admin cert downgrade")
err.number = 0
sSQL = "SELECT adminCert FROM admin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum=0 then ' check if EXISTS
	printtick("Adding admin cert column")
	cnn.execute("UPDATE admin SET adminCert='01010101010101010101010101010101010101010101010101010101010101'")
	cnn.execute("UPDATE admin SET adminCert='10101010101010101010101010101010101010101010101010101010101010'")
	cnn.execute("UPDATE admin SET adminCert='ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890'")
	on error resume next
	cnn.execute("ALTER TABLE admin DROP COLUMN adminCert")
	on error goto 0
end if

on error resume next
printtickdiv("Checking for admin cert downgrade")
err.number = 0
sSQL = "SELECT adminDelCC FROM admin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	cnn.execute("ALTER TABLE admin "&addcl&" adminDelCC INT DEFAULT 0")
	cnn.execute("UPDATE admin SET adminDelCC=7")
end if

call checkaddcolumn("orders","ordCNum",FALSE,memocl,"","")
call checkaddcolumn("admin","adminTweaks",FALSE,"INT","","")
call checkaddcolumn("admin","adminHandling",FALSE,dblcl,"","")
call checkaddcolumn("orders","ordHandling",FALSE,dblcl,"","")

on error resume next
printtickdiv("Checking for discount table")
err.number = 0
sSQL = "SELECT * FROM coupons"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding discount table")
	cnn.execute("CREATE TABLE coupons (cpnID "&autoinc&",cpnWorkingName "&txtcl&"(255) NULL,cpnNumber "&txtcl&"(255) NULL,cpnType INT DEFAULT 0,cpnEndDate "&datecl&",cpnDiscount "&dblcl&" DEFAULT 0,cpnThreshold "&dblcl&" DEFAULT 0,cpnQuantity INT DEFAULT 0,cpnNumAvail INT DEFAULT 0,cpnCntry "&bytecl&" DEFAULT 0,cpnIsCoupon "&bytecl&" DEFAULT 0,cpnSitewide "&bytecl&" DEFAULT 0)")
end if

on error resume next
printtickdiv("Checking for discount max upgrade")
err.number = 0
sSQL = "SELECT cpnThresholdMax FROM coupons"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding discount max columns")
	cnn.execute("ALTER TABLE coupons "&addcl&" cpnThresholdMax "&dblcl&" DEFAULT 0")
	cnn.execute("ALTER TABLE coupons "&addcl&" cpnQuantityMax INT DEFAULT 0")
	cnn.execute("UPDATE coupons SET cpnThresholdMax=0, cpnQuantityMax=0")
end if

on error resume next
printtickdiv("Checking for discount assignment table")
err.number = 0
sSQL = "SELECT * FROM cpnassign"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding discount assignment table")
	cnn.execute("CREATE TABLE cpnassign (cpaID "&autoinc&",cpaCpnID INT DEFAULT 0,cpaType "&bytecl&",cpaAssignment "&txtcl&"(255) NULL)")
end if
response.flush

on error resume next
printtickdiv("Checking for Discounts upgrade")
err.number = 0
sSQL = "SELECT ordDiscount FROM orders"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding Discounts columns")
	cnn.execute("ALTER TABLE orders "&addcl&" ordDiscount "&dblcl&" DEFAULT 0")
	cnn.execute("UPDATE orders SET ordDiscount=0")
	cnn.execute("ALTER TABLE orders "&addcl&" ordDiscountText "&txtcl&"(255) NULL")
	cnn.execute("UPDATE orders SET ordDiscountText=''")
end if

on error resume next
printtickdiv("Checking for Country Free Shipping upgrade")
err.number = 0
sSQL = "SELECT countryFreeShip FROM countries"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Country Free Shipping column")
	cnn.execute("ALTER TABLE countries "&addcl&" countryFreeShip "&bytecl&" DEFAULT 0")
	cnn.execute("UPDATE countries SET countryFreeShip=0")
end if

on error resume next
printtickdiv("Checking for State Free Shipping upgrade")
err.number = 0
sSQL = "SELECT stateFreeShip FROM states"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding State Free Shipping column")
	cnn.execute("ALTER TABLE states "&addcl&" stateFreeShip "&bytecl&" DEFAULT 1")
	cnn.execute("UPDATE states SET stateFreeShip=1")
end if

on error resume next
printtickdiv("Checking for List Price upgrade")
err.number = 0
sSQL = "SELECT pListPrice FROM products WHERE pID='xyxyx'"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding List Price column")
	cnn.execute("ALTER TABLE products "&addcl&" pListPrice "&dblcl&" DEFAULT 0")
	cnn.execute("UPDATE products SET pListPrice=0")
end if
response.flush

on error resume next
printtickdiv("Checking for USPS Methods upgrade")
err.number = 0
sSQL = "SELECT * FROM uspsmethods"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding USPS Methods table")
	sSQL = "CREATE TABLE uspsmethods (uspsID INT PRIMARY KEY,uspsMethod "&txtcl&"(150) NOT NULL,uspsShowAs "&txtcl&"(150) NOT NULL,uspsUseMethod "&bytecl&" DEFAULT 0,uspsLocal "&bytecl&" DEFAULT 0)"
	cnn.execute(sSQL)
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (1,'EXPRESS','Express Mail',0,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (2,'PRIORITY','Priority Mail',0,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (3,'PARCEL','Parcel Post',1,1)")
end if

printtickdiv("Checking for UPS Methods upgrade")
sSQL = "SELECT * FROM uspsmethods WHERE uspsID = 101"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	printtick("Adding UPS Methods information")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (101,'01','UPS Next Day Air&reg;',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (102,'02','UPS 2nd Day Air&reg;',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (103,'03','UPS Ground',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (104,'07','UPS Worldwide Express',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (105,'08','UPS Worldwide Expedited',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (106,'11','UPS Standard To Canada',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (107,'12','UPS 3 Day Select&reg;',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (108,'13','UPS Next Day Air Saver&reg;',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (109,'14','UPS Next Day Air&reg; Early A.M.&reg;',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (110,'54','UPS Worldwide Express Plus',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (111,'59','UPS 2nd Day Air A.M.&reg;',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (112,'65','UPS Express Saver',1,1)")
end if
rs.Close

on error resume next
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (14,'Media','Media Mail',0,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (15,'BPM','Bound Printed Matter',0,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (16,'FIRST CLASS','First-Class Mail',0,1)")
on error goto 0
response.flush

on error resume next
printtickdiv("Checking for U(S)PS FSA upgrade")
err.number = 0
sSQL = "SELECT uspsFSA FROM uspsmethods"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding U(S)PS FSA columns")
	cnn.execute("ALTER TABLE uspsmethods "&addcl&" uspsFSA "&bytecl&" DEFAULT 0")
	cnn.execute("UPDATE uspsmethods SET uspsFSA=0")
	cnn.execute("UPDATE uspsmethods SET uspsFSA=1 WHERE uspsID=103 OR uspsID=3")
	cnn.execute("ALTER TABLE postalzones "&addcl&" pzFSA INT DEFAULT 1")
	cnn.execute("UPDATE postalzones SET pzFSA=1")
end if

on error resume next
printtickdiv("Checking for pay provider order upgrade")
err.number = 0
sSQL = "SELECT payProvOrder FROM payprovider"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding pay provider order column")
	sSQL = "ALTER TABLE payprovider "&addcl&" payProvOrder INT DEFAULT 0"
	cnn.execute(sSQL)
	sSQL = "SELECT payProvID FROM payprovider"
	rs.Open sSQL,cnn,0,1
	index=0
	do while NOT rs.EOF
		cnn.execute("UPDATE payprovider SET payProvOrder="&index&" WHERE payProvID=" & rs("payProvID"))
		index=index+1
		rs.MoveNext
	loop
	rs.Close
end if

on error resume next
printtickdiv("Checking for top category order upgrade")
err.number = 0
sSQL = "SELECT * FROM topsections"
cnn.execute(sSQL)
errnum=err.number
if errnum=0 then ' The table is going to be destroyed later anyway, but we need these columns if it exists
	err.number = 0
	cnn.execute("SELECT tsOrder FROM topsections")
	errnum=err.number
	on error goto 0
	if errnum<>0 then
		printtick("Adding top category order column")
		cnn.execute("ALTER TABLE topsections "&addcl&" tsOrder INT DEFAULT 0")
		sSQL = "SELECT tsID FROM topsections ORDER BY tsName"
		rs.Open sSQL,cnn,0,1
		index=0
		do while NOT rs.EOF
			cnn.execute("UPDATE topsections SET tsOrder="&index&" WHERE tsID=" & rs("tsID"))
			index=index+1
			rs.MoveNext
		loop
		rs.Close
	end if
end if

on error resume next
printtickdiv("Checking for category order upgrade")
err.number = 0
sSQL = "SELECT sectionOrder FROM sections"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding category order column")
	cnn.execute("ALTER TABLE sections "&addcl&" sectionOrder INT DEFAULT 0")
	sSQL = "SELECT sectionID FROM sections ORDER BY sectionName"
	rs.Open sSQL,cnn,0,1
	index=0
	do while NOT rs.EOF
		cnn.execute("UPDATE sections SET sectionOrder="&index&" WHERE sectionID=" & rs("sectionID"))
		index=index+1
		rs.MoveNext
	loop
	rs.Close
end if
response.flush

on error resume next
printtickdiv("Checking for disabled section upgrade")
err.number = 0
sSQL = "SELECT sectionDisabled FROM sections"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding disabled section column")
	cnn.execute("ALTER TABLE sections "&addcl&" sectionDisabled "&bytecl&" DEFAULT 0")
	cnn.execute("UPDATE sections SET sectionDisabled=0")
end if
response.flush

on error resume next
printtickdiv("Checking for options weight difference upgrade")
err.number = 0
sSQL = "SELECT optWeightDiff FROM options"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding options weight difference column")
	cnn.execute("ALTER TABLE options "&addcl&" optWeightDiff "&dblcl&" DEFAULT 0")
	cnn.execute("UPDATE options SET optWeightDiff=0")
end if

on error resume next
printtickdiv("Checking for options wholesale price difference upgrade")
err.number = 0
sSQL = "SELECT optWholesalePriceDiff FROM options"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding options wholesale price difference column")
	cnn.execute("ALTER TABLE options "&addcl&" optWholesalePriceDiff "&dblcl&" DEFAULT 0")
	cnn.execute("UPDATE options SET optWholesalePriceDiff=optPriceDiff")
end if

on error resume next
printtickdiv("Checking for stock options upgrade")
err.number = 0
sSQL = "SELECT optStock FROM options"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding stock options column")
	cnn.execute("ALTER TABLE options "&addcl&" optStock INT DEFAULT 0")
	cnn.execute("UPDATE options SET optStock=0")
end if

printtickdiv("Checking for VeriSign PayFlow Link upgrade")
sSQL = "SELECT * FROM payprovider WHERE payProvID=8"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	sSQL = "INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (8,'Payflow Link','Credit Card',0,1,0,'','',8)"
	cnn.execute(sSQL)
end if
rs.Close

printtickdiv("Checking for PayPoint.net upgrade")
sSQL = "SELECT * FROM payprovider WHERE payProvID=9"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	sSQL = "INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (9,'PayPoint.net','Credit Card',0,1,0,'','',9)"
	cnn.execute(sSQL)
end if
rs.Close
cnn.execute("UPDATE payprovider SET payProvName='PayPoint.net' WHERE payProvID=9")

printtickdiv("Checking for 'Capture Card' upgrade")
sSQL = "SELECT * FROM payprovider WHERE payProvID=10"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	sSQL = "INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (10,'Capture Card','Credit Card',0,0,0,'XXXXXOOOOOOO','',10)"
	cnn.execute(sSQL)
end if
rs.Close

printtickdiv("Checking for 'PSiGate' upgrade")
sSQL = "SELECT * FROM payprovider WHERE payProvID=11"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	sSQL = "INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (11,'PSiGate','Credit Card',0,1,0,'','',11)"
	cnn.execute(sSQL)
end if
rs.Close
response.flush

printtickdiv("Checking for 'PSiGate SSL' upgrade")
sSQL = "SELECT * FROM payprovider WHERE payProvID=12"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	sSQL = "INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES ("
	sSQL = sSQL & "12,'PSiGate SSL','Credit Card',0,1,0,'','',12)"
	cnn.execute(sSQL)
end if
rs.Close

printtickdiv("Checking for 'Authorize.NET AIM' upgrade")
sSQL = "SELECT * FROM payprovider WHERE payProvID=13"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	sSQL = "INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES ("
	sSQL = sSQL & "13,'Auth.NET AIM','Credit Card',0,1,0,'','',13)"
	cnn.execute(sSQL)
end if
rs.Close

printtickdiv("Checking for 'Custom PayProc' upgrade")
sSQL = "SELECT * FROM payprovider WHERE payProvID=14"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	sSQL = "INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES ("
	sSQL = sSQL & "14,'Custom','Credit Card',0,1,0,'','',14)"
	cnn.execute(sSQL)
end if
rs.Close

printtickdiv("Checking for 'Netbanx' upgrade")
sSQL = "SELECT * FROM payprovider WHERE payProvID=15"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	sSQL = "INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES ("
	sSQL = sSQL & "15,'Netbanx','Credit Card',0,1,0,'','',15)"
	cnn.execute(sSQL)
end if
rs.Close
response.flush

printtickdiv("Checking for 'Linkpoint' upgrade")
sSQL = "SELECT * FROM payprovider WHERE payProvID=16"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	sSQL = "INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES ("
	sSQL = sSQL & "16,'Linkpoint','Credit Card',0,1,0,'','',16)"
	cnn.execute(sSQL)
end if
rs.Close
response.flush

on error resume next
printtickdiv("Checking for Option type upgrade")
err.number = 0
sSQL = "SELECT optType FROM optiongroup"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding Option type column")
	cnn.execute("ALTER TABLE optiongroup "&addcl&" optType INT DEFAULT 0")
	cnn.execute("UPDATE optiongroup SET optType=2")
end if

on error resume next
printtickdiv("Checking for category image upgrade")
err.number = 0
sSQL = "SELECT sectionImage FROM sections"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding category image column")
	cnn.execute("ALTER TABLE sections "&addcl&" sectionImage "&txtcl&"(255) NULL")
	cnn.execute("UPDATE sections SET sectionImage=''")
end if

on error resume next
printtickdiv("Checking for top category image upgrade")
err.number = 0
sSQL = "SELECT * FROM topsections"
cnn.execute(sSQL)
errnum=err.number
if errnum=0 then ' The table is going to be destroyed later anyway, but we need these columns if it exists
	err.number = 0
	sSQL = "SELECT tsImage FROM topsections"
	rs.Open sSQL,cnn,0,1
	errnum=err.number
	rs.Close
	on error goto 0
	if errnum<>0 then
		printtick("Adding category image column")
		cnn.execute("ALTER TABLE topsections "&addcl&" tsImage "&txtcl&"(255) NULL")
		cnn.execute("UPDATE topsections SET tsImage=''")
	end if
	response.flush
end if

call dropconstraintcolumn("admin","adminEuro")

on error resume next
printtickdiv("Checking for admin upgrade")
err.number = 0
sSQL = "SELECT adminVersion FROM admin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding admin columns")
	cnn.execute("ALTER TABLE admin "&addcl&" adminVersion "&txtcl&"(100) NULL")
	cnn.execute("ALTER TABLE admin "&addcl&" adminDelUncompleted INT DEFAULT 0")
	cnn.execute("UPDATE admin SET adminVersion='xxx',adminEuro=0,adminDelUncompleted=4")
end if

on error resume next
printtickdiv("Checking for adminUnits upgrade")
err.number = 0
sSQL = "SELECT adminUnits FROM admin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding adminUnits columns")
	cnn.execute("ALTER TABLE admin "&addcl&" adminUnits "&bytecl&" DEFAULT 0")
	cnn.execute("UPDATE admin SET adminUnits=0")
end if
response.flush

on error resume next
printtickdiv("Checking for adminUSZones upgrade")
err.number = 0
sSQL = "SELECT adminUSZones FROM admin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding adminUSZones columns")
	sSQL = "ALTER TABLE admin "&addcl&" adminUSZones "&smallcl&" NULL"
	cnn.execute(sSQL)
	sSQL = "UPDATE admin SET adminUSZones=0"
	cnn.execute(sSQL)
	on error resume next
	for index=11 to 24
		sSQL = "INSERT INTO postalzones (pzID,pzName) VALUES ("&index&",'')"
		cnn.execute(sSQL)
	next
	for index=101 to 124
		sSQL = "INSERT INTO postalzones (pzID,pzName) VALUES ("&index&",'')"
		cnn.execute(sSQL)
	next
	sSQL = "UPDATE postalzones SET pzName='All US States' WHERE pzID=101"
	cnn.execute(sSQL)
	sSQL = "INSERT INTO zonecharges (zcZone,zcWeight,zcRate) VALUES (101,-1,1)"
	cnn.execute(sSQL)
	sSQL = "INSERT INTO zonecharges (zcZone,zcWeight,zcRate) VALUES (101,1,1)"
	cnn.execute(sSQL)
	on error goto 0
end if

on error resume next
printtickdiv("Checking for US state zones upgrade")
err.number = 0
sSQL = "SELECT stateZone FROM states"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding US state zone column")
	sSQL = "ALTER TABLE states "&addcl&" stateZone INT DEFAULT 0"
	cnn.execute(sSQL)
	sSQL = "UPDATE states SET stateZone=101"
	cnn.execute(sSQL)
end if

on error resume next
printtickdiv("Checking for Exemptions upgrade")
err.number = 0
sSQL = "SELECT pExemptions FROM products WHERE pID='xyxyx'"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Exemptions column")
	cnn.execute("ALTER TABLE products "&addcl&" pExemptions int NOT NULL")
	cnn.execute("UPDATE products SET pExemptions=0")
end if
response.flush

on error resume next
printtickdiv("Checking for Options Percentage upgrade")
err.number = 0
sSQL = "SELECT optFlags FROM optiongroup"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding Options Percentage column")
	sSQL = "ALTER TABLE optiongroup "&addcl&" optFlags INT DEFAULT 0"
	cnn.execute(sSQL)
	sSQL = "UPDATE optiongroup SET optFlags=0"
	cnn.execute(sSQL)
	' This change can only be done once and is necessary for the v3.6.5 upgrade
	cnn.execute("UPDATE products SET pExemptions=7 WHERE pExemptions=3")
	cnn.execute("UPDATE products SET pExemptions=4 WHERE pExemptions=2")
	cnn.execute("UPDATE products SET pExemptions=3 WHERE pExemptions=1")
end if

on error resume next
printtickdiv("Checking for Unlimited Product Option upgrade")
err.number = 0
sSQL = "SELECT * FROM prodoptions"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Product Option table")
	sSQL = "CREATE TABLE prodoptions (poID AUTOINCREMENT PRIMARY KEY,poProdID "&txtcl&"(255) NOT NULL,poOptionGroup INT NOT NULL)"
	cnn.execute(sSQL)
	sSQL = "SELECT pID,pOptionGroup0,pOptionGroup1 FROM products"
	rs.Open sSQL,cnn,0,1
	do while NOT rs.EOF
		if rs("pOptionGroup0")<>0 then
			sSQL = "INSERT INTO prodoptions (poProdID,poOptionGroup) VALUES ('"&rs("pID")&"',"&rs("pOptionGroup0")&")"
			cnn.execute(sSQL)
		end if
		if rs("pOptionGroup1")<>0 then
			sSQL = "INSERT INTO prodoptions (poProdID,poOptionGroup) VALUES ('"&rs("pID")&"',"&rs("pOptionGroup1")&")"
			cnn.execute(sSQL)
		end if
		rs.MoveNext
	loop
	rs.Close
	sSQL = "ALTER TABLE products DROP COLUMN pOptionGroup0"
	cnn.execute(sSQL)
	sSQL = "ALTER TABLE products DROP COLUMN pOptionGroup1"
	cnn.execute(sSQL)
end if

on error resume next
printtickdiv("Checking for Cart Options Table")
err.number = 0
sSQL = "SELECT * FROM cartoptions"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Cart Options table")
	sSQL = "CREATE TABLE cartoptions (coID AUTOINCREMENT PRIMARY KEY,coCartID int NOT NULL,coOptID int NOT NULL,coOptGroup "&txtcl&"(255) NULL, coCartOption "&txtcl&"("&txtcollen&") NULL, coPriceDiff "&dblcl&")"
	cnn.execute(sSQL)
	sSQL = "SELECT cartID,cartOptGroup0,cartOption0,cartPriceDiff0,cartOptGroup1,cartOption1,cartPriceDiff1 FROM cart"
	rs.Open sSQL,cnn,0,1
	do while NOT rs.EOF
		if rs("cartOptGroup0")<>"" then
			sSQL = "INSERT INTO cartoptions (coCartID,coOptGroup,coCartOption,coPriceDiff,coOptID) VALUES ("&rs("cartID")&",'"&rs("cartOptGroup0")&"','"&rs("cartOption0")&"',"&rs("cartPriceDiff0")&",0)"
			cnn.execute(sSQL)
		end if
		if rs("cartOptGroup1")<>0 then
			sSQL = "INSERT INTO cartoptions (coCartID,coOptGroup,coCartOption,coPriceDiff,coOptID) VALUES ("&rs("cartID")&",'"&rs("cartOptGroup1")&"','"&rs("cartOption1")&"',"&rs("cartPriceDiff1")&",0)"
			cnn.execute(sSQL)
		end if
		rs.MoveNext
	loop
	rs.Close
	sSQL = "ALTER TABLE cart DROP COLUMN cartOptGroup0, cartOption0, cartPriceDiff0, cartOptGroup1, cartOption1, cartPriceDiff1"
	cnn.execute(sSQL)
end if

on error resume next
printtickdiv("Checking for cartoptions weight difference upgrade")
err.number = 0
sSQL = "SELECT coWeightDiff FROM cartoptions"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding cartoptions weight difference column")
	sSQL = "ALTER TABLE cartoptions "&addcl&" coWeightDiff "&dblcl&" DEFAULT 0"
	cnn.execute(sSQL)
	sSQL = "UPDATE cartoptions SET coWeightDiff=0"
	cnn.execute(sSQL)
end if
response.flush

on error resume next
printtickdiv("Checking for Affiliates Table")
err.number = 0
sSQL = "SELECT * FROM affiliates"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding affiliates table")
	sSQL = "CREATE TABLE affiliates (affilID "&txtcl&"(32) PRIMARY KEY,affilPW "&txtcl&"(255) NOT NULL,affilEmail "&txtcl&"(255),affilName "&txtcl&"(255) NULL,affilAddress "&txtcl&"(255) NULL,affilCity "&txtcl&"(255) NULL,affilState "&txtcl&"(255) NULL,affilZip "&txtcl&"(255) NULL,affilCountry "&txtcl&"(255) NULL,affilInform "&bytecl&" DEFAULT 0)"
	cnn.execute(sSQL)
end if

on error resume next
printtickdiv("Checking for Affiliate Commission Column")
err.number = 0
sSQL = "SELECT affilCommision FROM affiliates"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Affiliate Commission Column")
	cnn.execute("ALTER TABLE affiliates "&addcl&" affilCommision "&dblcl&" DEFAULT 0")
	cnn.execute("UPDATE affiliates SET affilCommision=0")
end if

sSQL = "UPDATE payprovider SET payProvAvailable=1 WHERE payProvID=3"
cnn.execute(sSQL)

on error resume next
printtickdiv("Checking for Multiple Shipping Method upgrade")
err.number = 0
sSQL = "SELECT pzMultiShipping FROM postalzones"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding pzMultiShipping column")
	cnn.execute("ALTER TABLE postalzones "&addcl&" pzMultiShipping "&smallcl&" DEFAULT 0")
	cnn.execute("UPDATE postalzones SET pzMultiShipping=0")
end if
if NOT sqlserver then cnn.execute("ALTER TABLE postalzones "&altcl&" pzMultiShipping "&smallcl&" DEFAULT 0")
cnn.execute("UPDATE postalzones SET pzMultiShipping=0 WHERE pzMultiShipping IS NULL")

on error resume next
printtickdiv("Checking for Extra Shipping Methods upgrade")
err.number = 0
sSQL = "SELECT pzMethodName1 FROM postalzones"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Extra Shipping Methods columns")
	cnn.execute("ALTER TABLE postalzones "&addcl&" pzMethodName1 "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE postalzones "&addcl&" pzMethodName2 "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE postalzones "&addcl&" pzMethodName3 "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE postalzones "&addcl&" pzMethodName4 "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE postalzones "&addcl&" pzMethodName5 "&txtcl&"(255) NULL")
	cnn.execute("UPDATE postalzones SET pzMethodName1='Standard Shipping', pzMethodName2='Express Shipping'")
	cnn.execute("ALTER TABLE zonecharges "&addcl&" zcRate3 "&dblcl&" DEFAULT 0")
	cnn.execute("ALTER TABLE zonecharges "&addcl&" zcRate4 "&dblcl&" DEFAULT 0")
	cnn.execute("ALTER TABLE zonecharges "&addcl&" zcRate5 "&dblcl&" DEFAULT 0")
	cnn.execute("UPDATE zonecharges SET zcRate3=0,zcRate4=0,zcRate5=0")
end if

on error resume next
printtickdiv("Checking for Multiple Shipping Method Charges upgrade")
err.number = 0
sSQL = "SELECT zcRate2 FROM zonecharges"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding zcRate2 column")
	cnn.execute("ALTER TABLE zonecharges "&addcl&" zcRate2 "&dblcl&" DEFAULT 0")
	cnn.execute("UPDATE zonecharges SET zonecharges.zcRate2=zonecharges.zcRate")
end if
response.flush

on error resume next
printtickdiv("Checking for countries upgrade")
err.number = 0
sSQL = "SELECT countryLCID FROM countries"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding countries columns")
	cnn.execute("ALTER TABLE countries "&addcl&" countryLCID INT DEFAULT 0")
	cnn.execute("ALTER TABLE countries "&addcl&" countryCurrency "&txtcl&"(100) NULL")
	cnn.execute("UPDATE countries SET countryLCID=0,countryCurrency=''")
end if

on error resume next
printtickdiv("Checking for Country Code upgrade")
err.number = 0
sSQL = "SELECT countryCode FROM countries"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding countryCode column")
	cnn.execute("ALTER TABLE countries "&addcl&" countryCode "&txtcl&"(10) NULL")
end if

printtickdiv("Checking for worldpay upgrade")
sSQL = "SELECT * FROM payprovider WHERE payProvID=5"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	cnn.execute("INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (5,'World Pay','Credit Card',0,1,1,'','',5)")
end if
rs.Close

printtickdiv("Checking for NOCHEX upgrade")
sSQL = "SELECT * FROM payprovider WHERE payProvID=6"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	sSQL = "INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (6,'NOCHEX','NOCHEX',0,1,1,'','',6)"
	cnn.execute(sSQL)
end if
rs.Close

printtickdiv("Checking for Verisign Payflow Pro upgrade")
sSQL = "SELECT * FROM payprovider WHERE payProvID=7"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	sSQL = "INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (7,'Payflow Pro','Credit Card',0,1,1,'','',7)"
	cnn.execute(sSQL)
end if
rs.Close
response.flush

on error resume next
printtickdiv("Checking for admin stock management upgrade")
err.number = 0
sSQL = "SELECT adminStockManage FROM admin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding admin columns")
	cnn.execute("ALTER TABLE admin "&addcl&" adminStockManage INT DEFAULT 0")
	cnn.execute("UPDATE admin SET adminStockManage=0")
end if

on error resume next
printtickdiv("Checking for products stock management upgrade")
err.number = 0
sSQL = "SELECT pInStock FROM products WHERE pID='xyxyx'"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding products columns")
	cnn.execute("ALTER TABLE products "&addcl&" pInStock INT DEFAULT 0")
	cnn.execute("UPDATE products SET pInStock=0")
end if

on error resume next
printtickdiv("Checking for IP address upgrade")
err.number = 0
sSQL = "SELECT ordIP FROM orders"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding IP address column")
	cnn.execute("ALTER TABLE orders "&addcl&" ordIP "&txtcl&"(50) NULL")
end if

on error resume next
printtickdiv("Checking for Affiliate upgrade")
err.number = 0
sSQL = "SELECT ordAffiliate FROM orders"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding ordAffiliate column")
	cnn.execute("ALTER TABLE orders "&addcl&" ordAffiliate "&txtcl&"(50) NULL")
end if

on error resume next
' cnn.execute("ALTER TABLE products "&altcl&" pDescription "&memocl&" NULL")
cnn.execute("ALTER TABLE products "&altcl&" pName "&txtcl&"(255) NOT NULL")
response.flush

on error resume next
err.number = 0
if sqlserver<>true then
	sSQL = "ALTER TABLE states "&altcl&" stateID INT NOT NULL"
	cnn.execute(sSQL)
	sSQL = "ALTER TABLE countries "&altcl&" countryID INT NOT NULL"
	cnn.execute(sSQL)
end if
errnum=err.number
on error goto 0
if errnum<>0 then
	response.write "<font color='#FF0000'>Could not remove autonumber from states and countries.</font><br />"
end if

on error resume next
printtickdiv("Checking for Unlimited Top Categories upgrade")
err.number = 0
sSQL = "SELECT rootSection FROM sections"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Unlimited Top Categories column")
	cnn.execute("ALTER TABLE sections "&addcl&" sectionWorkingName "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE sections "&addcl&" rootSection "&bytecl&" DEFAULT 0")
	cnn.execute("UPDATE sections SET rootSection=1")
	sSQL = "SELECT adminSubCats FROM admin"
	rs.Open sSQL,cnn,0,1
	subCats=(Int(rs("adminSubCats"))=1)
	rs.Close
	if subCats then
		tslist = ""
		addcomma = ""
		sSQL = "SELECT DISTINCT topSection FROM sections"
		rs.Open sSQL,cnn,0,1
		do while NOT rs.EOF
			tslist = rs("topSection") & addcomma & tslist
			addcomma = ","
			rs.MoveNext
		loop
		rs.Close
		if tslist<>"" then
			sSQL = "SELECT tsID,tsName,tsImage,tsOrder,tsDescription FROM topsections WHERE tsID IN (" & tslist & ")"
			rs.Open sSQL,cnn,0,1
			do while NOT rs.EOF
				rs2.Open "sections",cnn,1,3,&H0002
				rs2.AddNew
				rs2.Fields("sectionName") = rs("tsName")
				rs2.Fields("sectionImage") = rs("tsImage")
				rs2.Fields("sectionOrder") = rs("tsOrder")
				rs2.Fields("sectionDescription") = rs("tsDescription")
				rs2.Fields("rootSection") = 0
				rs2.Fields("topSection") = 0
				rs2.Update
				iID  = rs2.Fields("sectionID")
				rs2.Close
				cnn.execute("UPDATE sections SET rootSection=2,topSection=" & iID & " WHERE topSection=" & rs("tsID") & " AND rootSection<>2")
				cnn.execute("UPDATE cpnassign SET cpaType=1,cpaAssignment='" & iID & "' WHERE cpaAssignment='" & rs("tsID") & "' AND cpaType=0")
				rs.MoveNext
			loop
			rs.Close
			cnn.execute("UPDATE sections SET rootSection=1 WHERE rootSection=2")
		end if
	else
		cnn.execute("UPDATE sections SET topSection=0")
	end if
	cnn.execute("DELETE FROM cpnassign WHERE cpaType=0")
	cnn.execute("DROP TABLE topsections")
	cnn.execute("UPDATE sections SET sectionWorkingName=sectionName")
	call dropconstraintcolumn("admin","adminSubCats")
end if
response.flush

' All updates for version v4.7.0 and above below here

on error resume next
printtickdiv("Checking for Price Break upgrade")
err.number = 0
sSQL = "SELECT * FROM pricebreaks"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Price Break table")
	cnn.execute("CREATE TABLE pricebreaks (pbQuantity INT NOT NULL,pbProdID "&txtcl&"(255) NOT NULL,pPrice "&dblcl&" DEFAULT 0,pWholesalePrice "&dblcl&" DEFAULT 0,PRIMARY KEY(pbProdID,pbQuantity))")
end if
response.flush

on error resume next
printtickdiv("Checking for product dimensions upgrade")
err.number = 0
sSQL = "SELECT pDims FROM products WHERE pID='xyxyx'"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding product dimensions column")
	cnn.execute("ALTER TABLE products "&addcl&" pDims "&txtcl&"(255) NULL")
end if
response.flush

printtickdiv("Checking for Canada Post Methods upgrade")
sSQL = "SELECT * FROM uspsmethods WHERE uspsID = 201"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	printtick("Adding Canada Post Methods info")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal,uspsFSA) VALUES (201,'1010','Regular',1,1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (202,'1020','Expedited',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (203,'1030','Xpresspost',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (204,'1040','Priority Courier',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (205,'1120','Expedited Evening',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (206,'1130','XpressPost Evening',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (207,'1220','Expedited Saturday',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (208,'1230','XpressPost Saturday',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (210,'2005','Small Packets Surface',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (211,'2010','Surface USA',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (212,'2015','Small Packets Air USA',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (213,'2020','Air USA',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (214,'2025','Expedited USA Commercial',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (215,'2030','XPressPost USA',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (216,'2040','Purolator USA',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (217,'2050','PuroPak USA',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (218,'3005','Small Packets Surface International',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (221,'3010','Parcel Surface International',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (222,'3015','Small Packets Air International',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (223,'3020','Air International',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (224,'3025','XPressPost International',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (225,'3040','Purolator International',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (226,'3050','PuroPak International',1,0)")
end if
rs.Close

on error resume next
printtickdiv("Checking for IP Deny upgrade")
err.number = 0
sSQL = "SELECT * FROM multibuyblock"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding IP Deny table")
	cnn.execute("CREATE TABLE multibuyblock (ssdenyid "&autoinc&",ssdenyip "&txtcl&"(255) NOT NULL,sstimesaccess INT DEFAULT 0,lastaccess "&datecl&")")
	cnn.execute("CREATE TABLE ipblocking (dcid "&autoinc&",dcip1 "&dblcl&" DEFAULT 0,dcip2 "&dblcl&" DEFAULT 0)")
	cnn.execute("UPDATE sections SET sectionDisabled=127 WHERE sectionDisabled=1")
end if
response.flush

call checkaddcolumn("admin","adminlanguages",FALSE,"INT","","")
call checkaddcolumn("admin","adminlangsettings",FALSE,"INT","","")
call checkaddcolumn("countries","countryName2",FALSE,txtcl,"(255)","countryName")
call checkaddcolumn("countries","countryName3",FALSE,txtcl,"(255)","countryName")
call checkaddcolumn("optiongroup","optGrpName2",FALSE,txtcl,"(255)","")
call checkaddcolumn("optiongroup","optGrpName3",FALSE,txtcl,"(255)","")
call checkaddcolumn("options","optName2",FALSE,txtcl,"(255)","")
call checkaddcolumn("options","optName3",FALSE,txtcl,"(255)","")
call checkaddcolumn("orderstatus","statPublic2",FALSE,txtcl,"(255)","statPublic")
call checkaddcolumn("orderstatus","statPublic3",FALSE,txtcl,"(255)","statPublic")
call checkaddcolumn("payprovider","payProvShow2",FALSE,txtcl,"(255)","payProvShow")
call checkaddcolumn("payprovider","payProvShow3",FALSE,txtcl,"(255)","payProvShow")
call checkaddcolumn("products","pName2",FALSE,txtcl,"(255)","")
call checkaddcolumn("products","pName3",FALSE,txtcl,"(255)","")
call checkaddcolumn("products","pDescription2",FALSE,memocl,"","")
call checkaddcolumn("products","pDescription3",FALSE,memocl,"","")
call checkaddcolumn("products","pLongDescription2",FALSE,memocl,"","")
call checkaddcolumn("products","pLongDescription3",FALSE,memocl,"","")
call checkaddcolumn("products","pTax",FALSE,dblcl,"","")
call checkaddcolumn("sections","sectionName2",FALSE,txtcl,"(255)","")
call checkaddcolumn("sections","sectionName3",FALSE,txtcl,"(255)","")
call checkaddcolumn("sections","sectionDescription2",FALSE,memocl,"","")
call checkaddcolumn("sections","sectionDescription3",FALSE,memocl,"","")

on error resume next
printtickdiv("Checking for multi language upgrade part 2")
err.number = 0
sSQL = "SELECT cpnName FROM coupons"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	cnn.execute("ALTER TABLE coupons "&addcl&" cpnName "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE coupons "&addcl&" cpnName2 "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE coupons "&addcl&" cpnName3 "&txtcl&"(255) NULL")
	cnn.execute("UPDATE coupons SET cpnName=cpnWorkingName")
	sSQL = "SELECT adminShipping FROM admin WHERE adminID=1"
	rs.Open sSQL,cnn,0,1
	shipType = Int(rs("adminShipping"))
	rs.Close
	if shipType=3 then
		' Convert lbs + Oz to lbs.Oz
		sSQL = "SELECT pID,pWeight FROM products"
		rs.Open sSQL,cnn,0,1
		do while NOT rs.EOF
			pWeight = rs("pWeight")
			pWeight = Int(pWeight) + ((pWeight - Int(pWeight)) / 0.16)
			cnn.execute("UPDATE products SET pWeight="&pWeight&" WHERE pID='"&replace(rs("pID"),"'","''")&"'")
			rs.MoveNext
		loop
		rs.Close
		sSQL = "SELECT optID,optWeightDiff FROM options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE (optType=2 OR optType=-2) AND optFlags<2"
		rs.Open sSQL,cnn,0,1
		do while NOT rs.EOF
			iPounds=Int(rs("optWeightDiff"))
			iOunces = (iPounds*16)+Round((cDbl(rs("optWeightDiff"))-cDbl(iPounds))*100,2)
			cnn.execute("UPDATE options SET optWeightDiff="&(iOunces/16.0)&" WHERE optID="&rs("optID"))
			rs.MoveNext
		loop
		rs.Close
	end if
end if
response.flush

on error resume next
printtickdiv("Checking for dropshipper upgrade")
err.number = 0
sSQL = "SELECT * FROM dropshipper"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding dropshipper table")
	cnn.execute("CREATE TABLE dropshipper (dsID "&autoinc&",dsName "&txtcl&"(255) NULL,dsEmail "&txtcl&"(255) NULL,dsAddress "&txtcl&"(255) NULL,dsCity "&txtcl&"(255) NULL,dsState "&txtcl&"(255) NULL,dsZip "&txtcl&"(255) NULL,dsCountry "&txtcl&"(255) NULL,dsAction INT DEFAULT 0)")
	cnn.execute("ALTER TABLE products "&addcl&" pDropship INT DEFAULT 0")
	cnn.execute("UPDATE products SET pDropship=0")
end if
response.flush

printtickdiv("Checking for Email 2 upgrade")
sSQL = "SELECT * FROM payprovider WHERE payProvID=17"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	sSQL = "INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (17,'Email 2','Email 2',0,1,0,'','',17)"
	cnn.execute(sSQL)
end if
rs.Close

on error resume next
printtickdiv("Checking for Trans ID upgrade")
err.number = 0
sSQL = "SELECT ordTransID FROM orders WHERE ordID=1"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Trans ID column")
	if sqlserver<>TRUE then ' SQL Server will have to drop this column by hand
		cnn.execute("ALTER TABLE orders DROP COLUMN ordDemoMode")
	end if
	cnn.execute("ALTER TABLE orders "&addcl&" ordTransID "&txtcl&"(255) NULL")
end if
response.flush

cnn.execute("DELETE FROM admin WHERE adminID<>1")

on error resume next
printtickdiv("Checking for discount repeat upgrade")
err.number = 0
sSQL = "SELECT cpnThresholdRepeat FROM coupons"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding discount repeat columns")
	cnn.execute("ALTER TABLE coupons "&addcl&" cpnThresholdRepeat "&dblcl&" DEFAULT 0")
	cnn.execute("ALTER TABLE coupons "&addcl&" cpnQuantityRepeat INT DEFAULT 0")
	cnn.execute("UPDATE coupons SET cpnThresholdRepeat=0,cpnQuantityRepeat=0")
end if

on error resume next
printtickdiv("Checking for option modifyer upgrade")
err.number = 0
sSQL = "SELECT optRegExp FROM options"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding option modifyer column")
	cnn.execute("ALTER TABLE options "&addcl&" optRegExp "&txtcl&"(255) NULL")
end if
response.flush

on error resume next
printtickdiv("Checking for Address line 2 upgrade")
err.number = 0
sSQL = "SELECT ordAddress2 FROM orders"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Address line 2 columns")
	cnn.execute("ALTER TABLE orders "&addcl&" ordAddress2 "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE orders "&addcl&" ordShipAddress2 "&txtcl&"(255) NULL")
end if
response.flush

printtickdiv("Checking for PayPal Direct Payment upgrade")
sSQL = "SELECT * FROM payprovider WHERE payProvID=18"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	printtick("Adding PayPal Express Payment info")
	cnn.execute("INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (18,'PayPal Direct','Credit Card',0,1,0,'','',18)")
end if
rs.Close

printtickdiv("Checking for PayPal Express Payment upgrade")
sSQL = "SELECT * FROM payprovider WHERE payProvID=19"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	printtick("Adding PayPal Express Payment info")
	cnn.execute("INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (19,'PayPal Express','PayPal Express',0,0,0,'','',19)")
end if
rs.Close

on error resume next
printtickdiv("Checking for pay provider login level upgrade")
err.number = 0
sSQL = "SELECT payProvLevel FROM payprovider"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding pay provider login level column")
	cnn.execute("ALTER TABLE payprovider "&addcl&" payProvLevel INT DEFAULT 0")
	cnn.execute("UPDATE payprovider SET payProvLevel=0")
end if

response.flush

printtickdiv("Checking for FedEx Methods upgrade")
sSQL = "SELECT * FROM uspsmethods WHERE uspsID = 301"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	printtick("Adding FedEx Shipping Methods info")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (301,'PRIORITYOVERNIGHT','FedEx Priority Overnight&reg;',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (302,'STANDARDOVERNIGHT','FedEx Standard Overnight&reg;',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (303,'FIRSTOVERNIGHT','FedEx First Overnight&reg;',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (304,'FEDEX2DAY','FedEx 2Day&reg;',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (305,'FEDEXEXPRESSSAVER','FedEx Express Saver&reg;',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (306,'INTERNATIONALPRIORITY','FedEx International Priority&reg;',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (307,'INTERNATIONALECONOMY','FedEx International Economy&reg;',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (308,'INTERNATIONALFIRST','FedEx International First&reg;',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (310,'FEDEX1DAYFREIGHT','FedEx 1Day Freight&reg;',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (311,'FEDEX2DAYFREIGHT','FedEx 2Day Freight&reg;',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (312,'FEDEX3DAYFREIGHT','FedEx 3Day Freight&reg;',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal,uspsFSA) VALUES (313,'FEDEXGROUND','FedEx Ground&reg;',1,0,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (314,'GROUNDHOMEDELIVERY','FedEx Home Delivery&reg;',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (315,'INTERNATIONALPRIORITYFREIGHT','FedEx International Priority Freight&reg;',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (316,'INTERNATIONALECONOMYFREIGHT','FedEx International Economy Freight&reg;',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (317,'EUROPEFIRSTINTERNATIONALPRIORITY','FedEx Europe First&reg; - Int''l Priority',1,1)")
end if
rs.Close
response.flush

on error resume next
printtickdiv("Checking for bitfield upgrades")
err.number = 0
sSQL = "SELECT pStockByOpts FROM products WHERE pID='xyxyx'"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding bitfield columns")
	cnn.execute("ALTER TABLE products "&addcl&" pStaticPage "&bitfield&" DEFAULT 0")
	cnn.execute("ALTER TABLE products "&addcl&" pStockByOpts "&bitfield&" DEFAULT 0")
	cnn.execute("UPDATE products SET pStaticPage=0,pStockByOpts=0")
	cnn.execute("UPDATE products SET pStockByOpts=1 WHERE pSell=2 OR pSell=3 OR pSell=6 OR pSell=7")
	cnn.execute("UPDATE products SET pStaticPage=1 WHERE pSell=4 OR pSell=5 OR pSell=6 OR pSell=7")
	cnn.execute("UPDATE products SET pSell=1 WHERE pSell<>0")
end if

on error resume next
printtickdiv("Checking for Order Tracking Number upgrade")
err.number = 0
sSQL = "SELECT ordTrackNum FROM orders WHERE ordID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Order Tracking Number column")
	cnn.execute("ALTER TABLE orders "&addcl&" ordTrackNum "&txtcl&"(255) NULL")
end if

on error resume next
printtickdiv("Checking for Order AVS / CVV upgrade")
err.number = 0
sSQL = "SELECT ordAVS FROM orders WHERE ordID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Order AVS / CVV columns")
	cnn.execute("ALTER TABLE orders "&addcl&" ordAVS "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE orders "&addcl&" ordCVV "&txtcl&"(255) NULL")
end if
cnn.execute("UPDATE uspsmethods SET uspsMethod='FIRST CLASS' WHERE uspsID=16")

on error resume next
printtickdiv("Checking for international shipping upgrade")
err.number = 0
sSQL = "SELECT adminIntShipping FROM admin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding international shipping columns")
	cnn.execute("ALTER TABLE admin "&addcl&" adminIntShipping INT DEFAULT 0")
	cnn.execute("UPDATE admin SET adminIntShipping=0")
end if
response.flush

cnn.execute("UPDATE countries SET countryCode='GB' WHERE countryID=107 OR countryID=142")
cnn.execute("UPDATE countries SET countryName='Great Britain' WHERE countryName='Great Britain and Northern Ireland'")

on error resume next
	cnn.execute("INSERT INTO countries (countryID,countryName,countryEnabled,countryFreeShip,countryTax,countryOrder,countryZone,countryLCID,countryCurrency,countryCode) VALUES (214,'Channel Islands',0,0,0,0,3,0,'GBP','GB')")
	cnn.execute("INSERT INTO countries (countryID,countryName,countryEnabled,countryFreeShip,countryTax,countryOrder,countryZone,countryLCID,countryCurrency,countryCode) VALUES (215,'Puerto Rico',0,0,0,0,3,0,'USD','PR')")
	cnn.execute("INSERT INTO countries (countryID,countryName,countryEnabled,countryTax,countryOrder,countryZone,countryLCID,countryCurrency,countryCode) VALUES (216,'Isle of Man',0,0,0,3,0,'GBP','GB')")
	cnn.execute("INSERT INTO countries (countryID,countryName,countryEnabled,countryTax,countryOrder,countryZone,countryLCID,countryCurrency,countryCode) VALUES (217,'Azores',0,0,0,3,0,'EUR','PT')")
	cnn.execute("INSERT INTO countries (countryID,countryName,countryEnabled,countryTax,countryOrder,countryZone,countryLCID,countryCurrency,countryCode) VALUES (218,'Corsica',0,0,0,3,0,'EUR','FR')")
	cnn.execute("INSERT INTO countries (countryID,countryName,countryEnabled,countryTax,countryOrder,countryZone,countryLCID,countryCurrency,countryCode) VALUES (219,'Balearic Islands',0,0,0,3,0,'EUR','ES')")
	cnn.execute("INSERT INTO countries (countryID,countryName,countryEnabled,countryTax,countryOrder,countryZone,countryLCID,countryCurrency,countryCode) VALUES (221,'Serbia',0,0,0,3,0,'SRD','SR')")
	cnn.execute("INSERT INTO countries (countryID,countryName,countryEnabled,countryTax,countryOrder,countryZone,countryLCID,countryCurrency,countryCode) VALUES (222,'Ivory Coast',0,0,0,3,0,'XOF','CI')")
	cnn.execute("INSERT INTO countries (countryID,countryName,countryEnabled,countryTax,countryOrder,countryZone,countryLCID,countryCurrency,countryCode) VALUES (223,'Montenegro',0,0,0,3,0,'EUR','ME')")
	cnn.execute("INSERT INTO countries (countryID,countryName,countryEnabled,countryTax,countryOrder,countryZone,countryLCID,countryCurrency,countryCode,countryNumCurrency) VALUES (224,'American Samoa',0,0,0,3,0,'USD','AS',840)")
on error goto 0
cnn.execute("UPDATE countries SET countryName2=countryName WHERE countryName2='' OR countryName2 IS NULL")
cnn.execute("UPDATE countries SET countryName3=countryName WHERE countryName3='' OR countryName3 IS NULL")
cnn.execute("UPDATE countries SET countryLCID=1093 WHERE countryID=88")

on error resume next
cnn.execute("ALTER TABLE payprovider "&altcl&" payProvData1 "&txtcl&"(255) NULL")
cnn.execute("ALTER TABLE payprovider "&altcl&" payProvData2 "&txtcl&"(255) NULL")
on error goto 0

on error resume next
printtickdiv("Checking for product order upgrade")
err.number = 0
sSQL = "SELECT pOrder FROM products WHERE pID='xyxyx'"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding product order column")
	cnn.execute("ALTER TABLE products "&addcl&" pOrder INT DEFAULT 0")
	cnn.execute("UPDATE products SET pOrder=0")
	cnn.execute("CREATE INDEX pOrder_Indx ON products(pOrder)")
end if

on error resume next
printtickdiv("Checking for recommended products upgrade")
err.number = 0
sSQL = "SELECT pRecommend FROM products WHERE pID='xyxyx'"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding recommended products columns")
	cnn.execute("ALTER TABLE products "&addcl&" pRecommend "&bitfield&" DEFAULT 0")
	cnn.execute("UPDATE products SET pRecommend=0")
end if

on error resume next
printtickdiv("Checking for related products table")
err.number = 0
sSQL = "SELECT * FROM relatedprods WHERE rpProdID='xyxyx'"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding related products table")
	cnn.execute("CREATE TABLE relatedprods (rpProdID "&txtcl&"(128) NOT NULL,rpRelProdID "&txtcl&"(128) NOT NULL, PRIMARY KEY(rpProdID,rpRelProdID))")
end if
response.flush

on error resume next
printtickdiv("Checking for payprovider data3 upgrade")
err.number = 0
sSQL = "SELECT payProvData3 FROM payprovider WHERE payProvID=1"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding payprovider data3 column")
	cnn.execute("ALTER TABLE payprovider "&addcl&" payProvData3 "&txtcl&"(255) NULL")
end if

on error resume next
printtickdiv("Checking for Percentage Shipping Methods upgrade")
err.number = 0
sSQL = "SELECT zcRatePC FROM zonecharges"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Percentage Shipping Methods columns")
	cnn.execute("ALTER TABLE zonecharges "&addcl&" zcRatePC "&bitfield&" DEFAULT 0")
	cnn.execute("ALTER TABLE zonecharges "&addcl&" zcRatePC2 "&bitfield&" DEFAULT 0")
	cnn.execute("ALTER TABLE zonecharges "&addcl&" zcRatePC3 "&bitfield&" DEFAULT 0")
	cnn.execute("ALTER TABLE zonecharges "&addcl&" zcRatePC4 "&bitfield&" DEFAULT 0")
	cnn.execute("ALTER TABLE zonecharges "&addcl&" zcRatePC5 "&bitfield&" DEFAULT 0")
	cnn.execute("UPDATE zonecharges SET zcRatePC=0,zcRatePC2=0,zcRatePC3=0,zcRatePC4=0,zcRatePC5=0")
end if
response.flush

on error resume next
printtickdiv("Checking for default option upgrade")
err.number = 0
sSQL = "SELECT optDefault FROM options"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding default option column")
	cnn.execute("ALTER TABLE options "&addcl&" optDefault "&bitfield&" DEFAULT 0")
	cnn.execute("UPDATE options SET optDefault=0")
end if

on error resume next
printtickdiv("Checking for option select upgrade")
err.number = 0
sSQL = "SELECT optGrpSelect FROM optiongroup"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding option select column")
	cnn.execute("ALTER TABLE optiongroup "&addcl&" optGrpSelect "&bitfield&" DEFAULT 0")
	cnn.execute("UPDATE optiongroup SET optGrpSelect=0")
	cnn.execute("UPDATE optiongroup SET optGrpSelect=1 WHERE optType=2")
end if

on error resume next
printtickdiv("Checking for Order Invoice upgrade")
err.number = 0
sSQL = "SELECT ordInvoice FROM orders WHERE ordID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Order Invoice columns")
	cnn.execute("ALTER TABLE orders "&addcl&" ordInvoice "&txtcl&"(255) NULL")
end if
response.flush
'printtickdiv("Checking for Google Checkout upgrade")
'sSQL = "SELECT * FROM payprovider WHERE payProvID=20"
'rs.Open sSQL,cnn,0,1
'if rs.EOF then
'	cnn.execute("INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (20,'Google Checkout','Google Checkout',0,1,0,'','',20)")
'end if
'rs.Close

printtickdiv("Checking for PayPal Express Available upgrade")
sSQL = "SELECT payProvAvailable FROM payprovider WHERE payProvID=19"
rs.Open sSQL,cnn,0,1
if rs("payProvAvailable")=0 then
	rs2.Open "SELECT payProvEnabled,payProvMethod,payProvLevel,payProvData1,payProvData2,payProvData3 FROM payprovider WHERE payProvID=18",cnn,0,1
		cnn.execute("UPDATE payprovider SET payProvAvailable=1,payProvEnabled='"&rs2("payProvEnabled")&"',payProvMethod='"&rs2("payProvMethod")&"',payProvLevel='"&rs2("payProvLevel")&"',payProvData1='"&rs2("payProvData1")&"',payProvData2='"&rs2("payProvData2")&"',payProvData3='"&rs2("payProvData3")&"' WHERE payProvID=19")
	rs2.Close
end if
rs.Close

on error resume next
printtickdiv("Checking for Shipping Carrier upgrade")
err.number = 0
sSQL = "SELECT ordShipCarrier FROM orders WHERE ordID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding shipping carrier column")
	cnn.execute("ALTER TABLE orders "&addcl&" ordShipCarrier INT DEFAULT 0")
	cnn.execute("UPDATE orders SET ordShipCarrier=0")
end if

cnn.execute("UPDATE countries SET countryCurrency='RUB' WHERE countryID=157")

on error resume next
printtickdiv("Checking for Customer Login upgrade")
err.number = 0
sSQL = "SELECT * FROM customerlogin WHERE clID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Customer Login table")
	cnn.execute("CREATE TABLE customerlogin (clID "&autoinc&", clUserName "&txtcl&"(255),clPW "&txtcl&"(255) NULL,clLoginLevel "&bytecl&" DEFAULT 0,clActions INT DEFAULT 0,clPercentDiscount "&dblcl&" DEFAULT 0,clEmail "&txtcl&"(255) NULL,clDateCreated "&datecl&")")
	
	on error resume next
	err.number = 0
	sSQL = "SELECT * FROM clientlogin WHERE clientUser='xyxyxyxyx'"
	rs.Open sSQL,cnn,0,1
	errnum=err.number
	rs.Close
	if errnum<>0 then hasclientlogin=FALSE else hasclientlogin=TRUE
	err.number = 0
	sSQL = "SELECT clientPercentDiscount FROM clientlogin WHERE clientUser='xyxyxyxyx'"
	rs.Open sSQL,cnn,0,1
	errnum=err.number
	rs.Close
	if errnum<>0 then haspercentdiscount=FALSE else haspercentdiscount=TRUE
	on error goto 0
	
	if hasclientlogin then
		if haspercentdiscount then
			sSQL = "SELECT clientUser,clientPW,clientLoginLevel,clientActions,clientPercentDiscount,clientEmail FROM clientlogin"
		else
			sSQL = "SELECT clientUser,clientPW,clientLoginLevel,clientActions,clientEmail FROM clientlogin"
		end if
		rs.Open sSQL,cnn,0,1
		do while NOT rs.EOF
			if haspercentdiscount then percentdisc=rs("clientPercentDiscount") else percentdisc=0
			cnn.execute("INSERT INTO customerlogin (clUserName,clPW,clLoginLevel,clActions,clPercentDiscount,clEmail,clDateCreated) VALUES ('"&replace(rs("clientUser"),"'","''")&"','"&replace(rs("clientPw"),"'","''")&"',"&rs("clientLoginLevel")&","&rs("clientActions")&","&percentdisc&",''," & datedelim & vsusdate(Date()) & datedelim & ")")
			rs.movenext
		loop
		rs.Close
		cnn.execute("DROP TABLE clientlogin")
	end if
end if

on error resume next
printtickdiv("Checking for Order Table Customer Login upgrade")
err.number = 0
sSQL = "SELECT ordClientID FROM orders WHERE ordID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Order Table Customer Login ID column")
	cnn.execute("ALTER TABLE orders "&addcl&" ordClientID INT DEFAULT 0")
	cnn.execute("ALTER TABLE cart "&addcl&" cartClientID INT DEFAULT 0")
	cnn.execute("UPDATE orders SET ordClientID=0")
	cnn.execute("UPDATE cart SET cartClientID=0")
end if
response.flush

on error resume next
printtickdiv("Checking for customer address table")
err.number = 0
sSQL = "SELECT * FROM address WHERE addID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding address table")
	sSQL = "CREATE TABLE address (addID "&autoinc&"," & _
		"addCustID INT DEFAULT 0," & _
		"addIsDefault "&bytecl&" DEFAULT 0," & _
		"addName "&txtcl&"(255) NULL," & _
		"addAddress "&txtcl&"(255) NULL," & _
		"addAddress2 "&txtcl&"(255) NULL," & _
		"addCity "&txtcl&"(255) NULL," & _
		"addState "&txtcl&"(255) NULL," & _
		"addZip "&txtcl&"(255) NULL," & _
		"addCountry "&txtcl&"(255) NULL," & _
		"addPhone "&txtcl&"(255) NULL," & _
		"addShipFlags "&bytecl&" DEFAULT 0," & _
		"addExtra1 "&txtcl&"(255) NULL," & _
		"addExtra2 "&txtcl&"(255) NULL)"
	cnn.execute(sSQL)
	cnn.execute("CREATE INDEX addCustID_Indx ON address(addCustID)")
end if

on error resume next
printtickdiv("Checking for Order Shipping Phone upgrade")
err.number = 0
sSQL = "SELECT ordShipPhone FROM orders WHERE ordID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Order Shipping Phone column")
	cnn.execute("ALTER TABLE orders "&addcl&" ordShipPhone "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE orders "&addcl&" ordShipExtra1 "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE orders "&addcl&" ordShipExtra2 "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE orders "&addcl&" ordCheckoutExtra1 "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE orders "&addcl&" ordCheckoutExtra2 "&txtcl&"(255) NULL")
	on error resume next
	cnn.execute("UPDATE orders SET ordShipExtra1=ordExtra3")
	cnn.execute("ALTER TABLE orders DROP COLUMN ordExtra3")
	on error goto 0
end if
response.flush

on error resume next
printtickdiv("Checking for admin clear cart upgrade")
err.number = 0
sSQL = "SELECT adminClearCart FROM admin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding admin clear cart column")
	cnn.execute("ALTER TABLE admin "&addcl&" adminClearCart int DEFAULT 0")
	cnn.execute("UPDATE admin SET adminClearCart=364")
end if

on error resume next
printtickdiv("Checking for installedmods table")
err.number = 0
sSQL = "SELECT * FROM installedmods"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding installedmods table")
	cnn.execute("CREATE TABLE installedmods (modkey "&txtcl&"(255) PRIMARY KEY,modtitle "&txtcl&"(255) NOT NULL, modauthor "&txtcl&"(255) NULL, modauthorlink "&txtcl&"(255) NULL, modversion "&txtcl&"(255) NULL, modectversion "&txtcl&"(255) NULL, modlink "&txtcl&"(255) NULL, moddate "&datecl&" NOT NULL, modnotes "&memocl&" NULL)")
end if

on error resume next
printtickdiv("Checking for mailing list upgrade")
err.number = 0
sSQL = "SELECT * FROM mailinglist WHERE email='xyxyx'"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding mailing list table")
	cnn.execute("CREATE TABLE mailinglist (email "&txtcl&"(255) PRIMARY KEY,emailFormat "&bytecl&" DEFAULT 0)")
end if

printtickdiv("Checking for new UPS Methods upgrade")
sSQL = "SELECT * FROM uspsmethods WHERE uspsID = 30"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	printtick("Adding new USPS Methods info")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (30,'Global Express Guaranteed','Global Express Guaranteed',0,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (31,'Global Express Guaranteed Non-Document Rectangular','Global Express Guaranteed',0,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (32,'Global Express Guaranteed Non-Document Non-Rectangular','Global Express Guaranteed',0,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (33,'Express Mail International','Express Mail International',0,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (34,'Express Mail International Flat Rate Envelope','Express Mail International',0,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (35,'Priority Mail International','Priority Mail International',0,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (36,'Priority Mail International Flat Rate Envelope','Priority Mail International',0,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (37,'Priority Mail International Regular Flat-Rate Boxes','Priority Mail International',0,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (38,'First-Class Mail International','First-Class Mail',1,0)")
end if
rs.Close

cnn.execute("DELETE FROM uspsmethods WHERE uspsID>=4 AND uspsID<=13")

on error resume next
if sqlserver=TRUE then
	call drop_constraints("tmplogin","tmploginchk")
	cnn.execute("ALTER TABLE tmplogin "&altcl&" tmploginchk "&dblcl)
	cnn.execute("ALTER TABLE tmplogin ADD CONSTRAINT DF__tmploginchk DEFAULT (0) FOR tmploginchk")
	
	cnn.execute("DROP INDEX cartProdID ON cart")
	cnn.execute("DROP INDEX cartSessionID ON cart")
	cnn.execute("DROP INDEX ordSessionID ON orders")
else
	cnn.execute("ALTER TABLE tmplogin "&altcl&" tmploginchk "&dblcl&" DEFAULT 0")
end if
on error goto 0

on error resume next
printtickdiv("Checking for Order Authorization Status upgrade")
err.number = 0
sSQL = "SELECT ordAuthStatus FROM orders WHERE ordID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Order Authorization Status column")
	cnn.execute("ALTER TABLE orders "&addcl&" ordAuthStatus "&txtcl&"(255) NULL")
end if

on error resume next
printtickdiv("Checking for Admin Login upgrade")
err.number = 0
sSQL = "SELECT * FROM adminlogin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Admin Login table")
	cnn.execute("CREATE TABLE adminlogin (adminloginid "&autoinc&",adminloginname "&txtcl&"(255) NOT NULL,adminloginpassword "&txtcl&"(255) NOT NULL,adminloginpermissions "&txtcl&"(255) NOT NULL)")
	Session("loggedon")="" ' Force relogin
end if

on error resume next
printtickdiv("Checking for mailinglist confirmation upgrade")
err.number = 0
sSQL = "SELECT isconfirmed FROM mailinglist"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding mailinglist confirmation column")
	cnn.execute("ALTER TABLE mailinglist "&addcl&" isconfirmed "&bitfield&" DEFAULT 0")
	cnn.execute("UPDATE mailinglist SET isconfirmed=1")
end if

'on error resume next
'printtickdiv("Checking for manufacturer table upgrade")
'err.number = 0
'sSQL = "SELECT * FROM manufacturer WHERE mfID=0"
'rs.Open sSQL,cnn,0,1
'errnum=err.number
'rs.Close
'on error goto 0
'if errnum<>0 then
'	printtick("Adding manufacturer table")
'	cnn.execute("CREATE TABLE manufacturer (mfID "&autoinc&",mfName "&txtcl&"(255) NULL,mfEmail "&txtcl&"(255) NULL,mfAddress "&txtcl&"(255) NULL,mfCity "&txtcl&"(255) NULL,mfState "&txtcl&"(255) NULL,mfZip "&txtcl&"(255) NULL,mfCountry "&txtcl&"(255) NULL)")
'end if

on error resume next
printtickdiv("Checking for manufacturer column upgrade")
err.number = 0
sSQL = "SELECT pManufacturer FROM products WHERE pID='xyxyx'"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding pManufacturer column")
	cnn.execute("ALTER TABLE products "&addcl&" pManufacturer INT DEFAULT 0")
	cnn.execute("UPDATE products SET pManufacturer=0")
end if

on error resume next
printtickdiv("Checking for Product SKU upgrade")
err.number = 0
sSQL = "SELECT pSKU FROM products WHERE pID='xyxyx'"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding Product SKU column")
	cnn.execute("ALTER TABLE products "&addcl&" pSKU "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE products "&addcl&" pDateAdded "&datecl)
	cnn.execute("UPDATE products SET pDateAdded="&datedelim&(vsusdate(date()-1))&datedelim)
end if

on error resume next
printtickdiv("Checking for option alternate image upgrade")
err.number = 0
sSQL = "SELECT optAltImage FROM options"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding option alternate image column")
	cnn.execute("ALTER TABLE options "&addcl&" optAltImage "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE options "&addcl&" optAltLargeImage "&txtcl&"(255) NULL")
end if

on error resume next
printtickdiv("Checking for admin secret upgrade")
err.number = 0
sSQL = "SELECT adminSecret FROM admin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding admin secret columns")
	cnn.execute("ALTER TABLE admin "&addcl&" adminSecret "&txtcl&"(255) NULL")
	randomize
	cnn.execute("UPDATE admin SET adminSecret='its a big secret "&(Int(1000000 * Rnd) + 1000000)&"'")
end if

on error resume next
printtickdiv("Checking for mailing list confirmation date upgrade")
err.number = 0
sSQL = "SELECT mlIPAddress FROM mailinglist"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding mailing list confirmation date columns")
	cnn.execute("ALTER TABLE mailinglist "&addcl&" mlConfirmDate "&smalldatecl)
	cnn.execute("ALTER TABLE mailinglist "&addcl&" mlIPAddress "&txtcl&"(255) NULL")
	cnn.execute("UPDATE mailinglist SET mlConfirmDate="&datedelim&vsusdate(Date())&datedelim)
end if

on error resume next
printtickdiv("Checking for ratings table")
err.number = 0
sSQL = "SELECT * FROM ratings"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding ratings table")
	cnn.execute("CREATE TABLE ratings (rtID "&autoinc&",rtProdID "&txtcl&"(255) NOT NULL,rtRating "&bytecl&" DEFAULT 0,rtLanguage "&bytecl&" DEFAULT 0,rtDate "&datecl&",rtApproved "&bitfield&" DEFAULT 0,rtIPAddress "&txtcl&"(255) NULL,rtPosterName "&txtcl&"(255),rtPosterLoginID INT DEFAULT 0,rtPosterEmail "&txtcl&"(255) NULL,rtHeader "&txtcl&"(255) NULL,rtComments "&memocl&" NULL)")
	cnn.execute("CREATE INDEX rtProdID_Indx ON ratings(rtProdID)")
end if

on error resume next
printtickdiv("Checking for Unlimited Multiple Image upgrade")
response.flush
err.number = 0
sSQL = "SELECT imageProduct FROM productimages WHERE imageProduct='xyxyx'"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding Unlimited Multiple Image table")
	cnn.execute("CREATE TABLE productimages (imageProduct "&txtcl&"(128),imageSrc "&txtcl&"(255) NOT NULL,imageNumber INT DEFAULT 0 NOT NULL,imageType "&smallcl&" DEFAULT 0 NOT NULL, PRIMARY KEY(imageProduct,imageType,imageNumber))")
	openbr="" : closebr=""
	if sqlserver=TRUE then openbr="(" : closebr=")"
	for index=1 to 5
		if index=1 then tim="" else tim=index
		on error resume next
		cnn.execute("INSERT INTO productimages (imageProduct,imageSrc,imageNumber,imageType) " & openbr & "SELECT pId,pImage"&tim&","&(index-1)&",0 FROM products WHERE pImage"&tim&"<>'' AND pImage"&tim&"<>'prodimages/' AND NOT (pImage"&tim&" IS NULL)" & closebr)
		cnn.execute("INSERT INTO productimages (imageProduct,imageSrc,imageNumber,imageType) " & openbr & "SELECT pId,pLargeImage"&tim&","&(index-1)&",1 FROM products WHERE pLargeImage"&tim&"<>'' AND pLargeImage"&tim&"<>'prodimages/' AND NOT (pLargeImage"&tim&" IS NULL)" & closebr)
		cnn.execute("INSERT INTO productimages (imageProduct,imageSrc,imageNumber,imageType) " & openbr & "SELECT pId,pGiantImage"&tim&","&(index-1)&",2 FROM products WHERE pGiantImage"&tim&"<>'' AND pGiantImage"&tim&"<>'prodimages/' AND NOT (pGiantImage"&tim&" IS NULL)" & closebr)
		cnn.execute("ALTER TABLE products DROP COLUMN pImage"&tim)
		cnn.execute("ALTER TABLE products DROP COLUMN pLargeImage"&tim)
		cnn.execute("ALTER TABLE products DROP COLUMN pGiantImage"&tim)
		on error goto 0
	next
	cnn.execute("CREATE INDEX imageProduct_Indx ON productimages(imageProduct)")
	cnn.execute("CREATE INDEX imageType_Indx ON productimages(imageType)")
end if

on error resume next
printtickdiv("Checking for shipping discount handling upgrade")
err.number = 0
sSQL = "SELECT cpnHandling FROM coupons"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding shipping discount handling columns")
	cnn.execute("ALTER TABLE coupons "&addcl&" cpnHandling "&bitfield&" DEFAULT 0")
	cnn.execute("UPDATE coupons SET cpnHandling=0")
end if

on error resume next
cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (39,'14','First-Class Mail',0,0)")
cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (40,'15','First-Class Mail',0,0)")
cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (41,'11','Priority Mail International',0,0)")

cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (42,'16','Priority Mail International',0,0)")
cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (43,'17','Express Mail International',0,0)")
cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (44,'20','Priority Mail International',0,0)")
cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (45,'24','Priority Mail International',0,0)")
cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (46,'26','Express Mail International',0,0)")
cnn.execute("UPDATE uspsmethods SET uspsShowAs='Priority Mail International' WHERE uspsID IN (41,42,44,45) AND uspsShowAs='Priority Mail'")
cnn.execute("UPDATE uspsmethods SET uspsShowAs='Express Mail International' WHERE uspsID IN (43,46) AND uspsShowAs='Express Mail'")
on error goto 0

on error resume next
printtickdiv("Checking for Gift Certificate upgrade")
err.number = 0
sSQL = "SELECT * FROM giftcertificate"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Gift Certificate table")
	cnn.execute("CREATE TABLE giftcertificate (gcID "&txtcl&"(255) PRIMARY KEY,gcTo "&txtcl&"(255) NULL,gcFrom "&txtcl&"(255) NULL,gcEmail "&txtcl&"(255) NULL,gcOrigAmount "&dblcl&" DEFAULT 0,gcRemaining "&dblcl&" DEFAULT 0,gcDateCreated "&datecl&",gcDateUsed "&datecl&",gcCartID INT DEFAULT 0 NOT NULL,gcOrderID INT DEFAULT 0 NOT NULL,gcAuthorized "&bitfield&" DEFAULT 0,gcMessage "&memocl&" NULL)")
	cnn.execute("CREATE INDEX gcCartID_Indx ON giftcertificate(gcCartID)")
	cnn.execute("CREATE INDEX gcOrderID_Indx ON giftcertificate(gcOrderID)")
end if

on error resume next
printtickdiv("Checking for Gift Cert Applied upgrade")
err.number = 0
sSQL = "SELECT * FROM giftcertsapplied"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Gift Cert Applied table")
	cnn.execute("CREATE TABLE giftcertsapplied (gcaGCID "&txtcl&"(255) NOT NULL,gcaOrdID INT DEFAULT 0 NOT NULL,gcaAmount "&dblcl&" DEFAULT 0, PRIMARY KEY(gcaGCID,gcaOrdID))")
end if

on error resume next
printtickdiv("Checking for Email Messages upgrade")
err.number = 0
sSQL = "SELECT emailID FROM emailmessages"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Email Messages table")
	cnn.execute("CREATE TABLE emailmessages (emailID INT PRIMARY KEY,giftcertsubject "&txtcl&"(255) NULL,giftcertsubject2 "&txtcl&"(255) NULL,giftcertsubject3 "&txtcl&"(255) NULL, giftcertemail "&memocl&" NULL, giftcertemail2 "&memocl&" NULL, giftcertemail3 "&memocl&" NULL,giftcertsendersubject "&txtcl&"(255) NULL,giftcertsendersubject2 "&txtcl&"(255) NULL,giftcertsendersubject3 "&txtcl&"(255) NULL,giftcertsender "&memocl&" NULL,giftcertsender2 "&memocl&" NULL,giftcertsender3 "&memocl&" NULL,emailsubject "&txtcl&"(255) NULL,emailsubject2 "&txtcl&"(255) NULL,emailsubject3 "&txtcl&"(255) NULL,emailheaders "&memocl&" NULL,emailheaders2 "&memocl&" NULL,emailheaders3 "&memocl&" NULL,dropshipsubject "&txtcl&"(255) NULL,dropshipsubject2 "&txtcl&"(255) NULL,dropshipsubject3 "&txtcl&"(255) NULL,dropshipheaders "&memocl&" NULL,dropshipheaders2 "&memocl&" NULL,dropshipheaders3 "&memocl&" NULL,orderstatussubject "&txtcl&"(255) NULL,orderstatussubject2 "&txtcl&"(255) NULL,orderstatussubject3 "&txtcl&"(255) NULL,orderstatusemail "&memocl&" NULL,orderstatusemail2 "&memocl&" NULL,orderstatusemail3 "&memocl&" NULL)")
	cnn.execute("INSERT INTO emailmessages (emailID) VALUES (1)")

	cnn.execute("UPDATE emailmessages SET giftcertsubject='You received a gift certificate from %fromname%'")
	cnn.execute("UPDATE emailmessages SET giftcertemail='Hi %toname%, %fromname% has sent you a gift certificate to the value of %value%!<br />{Your friend left the following message: %message%}<br />To redeem your gift certificate, simply pop along to our online store at:<br />%storeurl%<br />Then select the goods you require and when checking out enter the gift certificate code below:<br />%certificateid%'")
	cnn.execute("UPDATE emailmessages SET giftcertsendersubject='You sent a gift certificate to %toname%'")
	cnn.execute("UPDATE emailmessages SET giftcertsender='You sent a gift certificate to %toname%.<br />Below is a copy of the email they will receive. You may want to check it was delivered.'")
	
	cnn.execute("UPDATE emailmessages SET emailsubject='Thank you for your order'")
	themessage = replace(replace(emailheader&"%emailmessage%<br />"&emailfooter, "<br>", "<br />"), "<br/>", "<br />")
	cnn.execute("UPDATE emailmessages SET emailheaders='"&replace(themessage, "'", "''")&"'")
	
	if dropshipsubject="" then dropshipsubject="We have received the following order"
	cnn.execute("UPDATE emailmessages SET dropshipsubject='"&replace(dropshipsubject, "'", "''")&"'")
	themessage = replace(replace(dropshipheader&"%emailmessage%<br />"&dropshipfooter, "<br>", "<br />"), "<br/>", "<br />")
	cnn.execute("UPDATE emailmessages SET dropshipheaders='"&replace(themessage, "'", "''")&"'")
	
	if orderstatussubject="" then orderstatussubject="Order status updated"
	cnn.execute("UPDATE emailmessages SET orderstatussubject='"&replace(orderstatussubject, "'", "''")&"'")
	
	if trackingnumtext<>"" then orderstatusemail=replace(orderstatusemail, "%trackingnum%", "{" & replace(trackingnumtext, "%s", "%trackingnum%") & "}")
	orderstatusemail = replace(replace(replace(orderstatusemail, "<br>", "<br />"), "<br/>", "<br />"), "%nl%", "<br />")
	cnn.execute("UPDATE emailmessages SET orderstatusemail='"&replace(orderstatusemail, "'", "''")&"'")

	cnn.execute("UPDATE emailmessages SET giftcertsubject3=giftcertsubject,giftcertsubject2=giftcertsubject")
	cnn.execute("UPDATE emailmessages SET giftcertemail3=giftcertemail,giftcertemail2=giftcertemail")
	cnn.execute("UPDATE emailmessages SET giftcertsendersubject3=giftcertsendersubject,giftcertsendersubject2=giftcertsendersubject")
	cnn.execute("UPDATE emailmessages SET giftcertsender3=giftcertsender,giftcertsender2=giftcertsender")
	cnn.execute("UPDATE emailmessages SET emailsubject3=emailsubject,emailsubject2=emailsubject")
	cnn.execute("UPDATE emailmessages SET emailheaders3=emailheaders,emailheaders2=emailheaders")
	cnn.execute("UPDATE emailmessages SET dropshipsubject3=dropshipsubject,dropshipsubject2=dropshipsubject")
	cnn.execute("UPDATE emailmessages SET dropshipheaders3=dropshipheaders,dropshipheaders2=dropshipheaders")
	cnn.execute("UPDATE emailmessages SET orderstatussubject3=orderstatussubject,orderstatussubject2=orderstatussubject")
	cnn.execute("UPDATE emailmessages SET orderstatusemail3=orderstatusemail,orderstatusemail2=orderstatusemail")
end if

on error resume next
printtickdiv("Checking for Payment Provider Headers upgrade")
err.number = 0
sSQL = "SELECT pProvHeaders FROM payprovider"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding Payment Provider Headers columns")
	cnn.execute("ALTER TABLE payprovider "&addcl&" ppHandlingCharge "&dblcl&" DEFAULT 0")
	cnn.execute("ALTER TABLE payprovider "&addcl&" ppHandlingPercent "&dblcl&" DEFAULT 0")
	cnn.execute("ALTER TABLE payprovider "&addcl&" pProvHeaders "&memocl&" NULL")
	cnn.execute("ALTER TABLE payprovider "&addcl&" pProvHeaders2 "&memocl&" NULL")
	cnn.execute("ALTER TABLE payprovider "&addcl&" pProvHeaders3 "&memocl&" NULL")
	cnn.execute("ALTER TABLE payprovider "&addcl&" pProvDropShipHeaders "&memocl&" NULL")
	cnn.execute("ALTER TABLE payprovider "&addcl&" pProvDropShipHeaders2 "&memocl&" NULL")
	cnn.execute("ALTER TABLE payprovider "&addcl&" pProvDropShipHeaders3 "&memocl&" NULL")
	cnn.execute("UPDATE payprovider SET ppHandlingCharge=0,ppHandlingPercent=0")
	for index=1 to 20
		execute("handlingcharge = handlingcharge" & index & " : handlingchargepercent = handlingchargepercent" & index)
		if handlingcharge="" then handlingcharge=0
		if handlingchargepercent="" then handlingchargepercent=0
		execute("emailheaders = emailheader" & index & "&""%emailmessage%<br />""& emailfooter" & index)
		execute("dropshipheaders = dropshipheader" & index & "&""%emailmessage%<br />""& dropshipfooter" & index)
		cnn.execute("UPDATE payprovider SET ppHandlingCharge=" & handlingcharge & ",ppHandlingPercent=" & handlingchargepercent & " WHERE payProvID="&index)
		cnn.execute("UPDATE payprovider SET pProvHeaders='"&replace(emailheaders,"'","''")&"',pProvDropShipHeaders='"&replace(dropshipheaders,"'","''")&"' WHERE payProvID="&index)
		cnn.execute("UPDATE payprovider SET pProvHeaders2='"&replace(emailheaders,"'","''")&"',pProvDropShipHeaders2='"&replace(dropshipheaders,"'","''")&"' WHERE payProvID="&index)
		cnn.execute("UPDATE payprovider SET pProvHeaders3='"&replace(emailheaders,"'","''")&"',pProvDropShipHeaders3='"&replace(dropshipheaders,"'","''")&"' WHERE payProvID="&index)
	next
end if

on error resume next
printtickdiv("Checking for handling charge percentage upgrade")
err.number = 0
sSQL = "SELECT adminHandlingPercent FROM admin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding handling charge percentage columns")
	if handlingchargepercent="" then handlingchargepercent=0
	cnn.execute("ALTER TABLE admin "&addcl&" adminHandlingPercent "&dblcl&" DEFAULT 0")
	cnn.execute("UPDATE admin SET adminHandlingPercent=" & handlingchargepercent)
end if

on error resume next
printtickdiv("Checking for product ratings upgrade")
err.number = 0
sSQL = "SELECT pNumRatings FROM products WHERE pID='xyxyx'"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding product ratings columns")
	cnn.execute("ALTER TABLE products "&addcl&" pTotRating INT DEFAULT 0")
	cnn.execute("ALTER TABLE products "&addcl&" pNumRatings INT DEFAULT 0")
	cnn.execute("UPDATE products SET pTotRating=0,pNumRatings=0")
end if

on error resume next
	cnn.execute("ALTER TABLE manufacturer "&addcl&" mfLogo "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE manufacturer "&addcl&" mfURL "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE manufacturer "&addcl&" mfURL2 "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE manufacturer "&addcl&" mfURL3 "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE manufacturer "&addcl&" mfDescription "&memocl&" NULL")
	cnn.execute("ALTER TABLE manufacturer "&addcl&" mfDescription2 "&memocl&" NULL")
	cnn.execute("ALTER TABLE manufacturer "&addcl&" mfDescription3 "&memocl&" NULL")
	cnn.execute("ALTER TABLE manufacturer "&addcl&" mfOrder INT DEFAULT 0")
	cnn.execute("UPDATE manufacturer SET mfOrder=0")
on error goto 0

on error resume next
printtickdiv("Checking for Search Params upgrade")
err.number = 0
sSQL = "SELECT pSearchParams FROM products WHERE pID='xyxyx'"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding Search Params column")
	cnn.execute("ALTER TABLE products "&addcl&" pSearchParams "&txtcl&"(255) NULL")
end if
response.flush

on error resume next
printtickdiv("Checking for Referer upgrade")
err.number = 0
sSQL = "SELECT ordReferer FROM orders WHERE ordID=0"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding Referer column")
	cnn.execute("ALTER TABLE orders "&addcl&" ordReferer "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE orders "&addcl&" ordQuerystr "&txtcl&"(255) NULL")
end if

printtickdiv("Checking for Amazon Simple upgrade")
sSQL = "SELECT * FROM payprovider WHERE payProvID=21"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	sSQL = "INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (21,'Amazon Pay','Amazon Pay',0,1,0,'','',21)"
	cnn.execute(sSQL)
end if
rs.Close

on error resume next
printtickdiv("Checking for split order name upgrade")
err.number = 0
sSQL = "SELECT ordLastName FROM orders WHERE ordID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding split order name column")
	cnn.execute("ALTER TABLE orders "&addcl&" ordLastName "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE orders "&addcl&" ordShipLastName "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE address "&addcl&" addLastName "&txtcl&"(255) NULL")
end if

on error resume next
printtickdiv("Checking for drop shipper email header upgrade")
err.number = 0
sSQL = "SELECT dsEmailHeader FROM dropshipper"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding drop shipper email header column")
	cnn.execute("ALTER TABLE dropshipper "&addcl&" dsEmailHeader "&memocl&" NULL")
end if

on error resume next
printtickdiv("Checking for Order Language upgrade")
err.number = 0
sSQL = "SELECT ordLang FROM orders WHERE ordID=0"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding Order Language column")
	cnn.execute("ALTER TABLE orders "&addcl&" ordLang "&bytecl&" DEFAULT 0")
	cnn.execute("UPDATE orders SET ordLang=0")
end if

on error resume next
printtickdiv("Checking for order status email upgrade")
err.number = 0
sSQL = "SELECT emailstatus FROM orderstatus"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding order status email column")
	cnn.execute("ALTER TABLE orderstatus "&addcl&" emailstatus "&bytecl&" DEFAULT 0")
	cnn.execute("UPDATE orderstatus SET emailstatus=0")
	cnn.execute("UPDATE orderstatus SET emailstatus=1 WHERE statID>=4")
end if

on error resume next
printtickdiv("Checking for mailinglist name upgrade")
err.number = 0
sSQL = "SELECT mlName FROM mailinglist"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding mailinglist name column")
	cnn.execute("ALTER TABLE mailinglist "&addcl&" mlName "&txtcl&"(255) NULL")
	sSQL = "SELECT email,ordName FROM mailinglist INNER JOIN orders ON mailinglist.email=orders.ordEmail"
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		cnn.execute("UPDATE mailinglist SET mlName='"&replace(trim(rs("ordName")&""),"'","''")&"' WHERE email='"&replace(rs("email"),"'","''")&"'")
		rs.movenext
	loop
end if

on error resume next
printtickdiv("Checking for Affiliate Date Column")
err.number = 0
sSQL = "SELECT affilDate FROM affiliates"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Affiliate Date Column")
	cnn.execute("ALTER TABLE affiliates "&addcl&" affilDate "&smalldatecl)
	cnn.execute("UPDATE affiliates SET affilDate="&datedelim&vsusdate(Date()-10)&datedelim)
end if

on error resume next
printtickdiv("Checking for Updates Column")
err.number = 0
sSQL = "SELECT updLastCheck FROM admin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Updates Columns")
	cnn.execute("ALTER TABLE admin "&addcl&" updLastCheck "&smalldatecl)
	cnn.execute("ALTER TABLE admin "&addcl&" updRecommended "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE admin "&addcl&" updSecurity "&bitfield&" DEFAULT 0")
	cnn.execute("ALTER TABLE admin "&addcl&" updShouldUpd "&bitfield&" DEFAULT 0")
end if

cnn.execute("UPDATE uspsmethods SET uspsMethod='4' WHERE uspsID=30")
cnn.execute("UPDATE uspsmethods SET uspsMethod='6' WHERE uspsID=31")
cnn.execute("UPDATE uspsmethods SET uspsMethod='7' WHERE uspsID=32")
cnn.execute("UPDATE uspsmethods SET uspsMethod='1' WHERE uspsID=33")
cnn.execute("UPDATE uspsmethods SET uspsMethod='10' WHERE uspsID=34")
cnn.execute("UPDATE uspsmethods SET uspsMethod='2' WHERE uspsID=35")
cnn.execute("UPDATE uspsmethods SET uspsMethod='8' WHERE uspsID=36")
cnn.execute("UPDATE uspsmethods SET uspsMethod='9' WHERE uspsID=37")
cnn.execute("UPDATE uspsmethods SET uspsMethod='13' WHERE uspsID=38")
cnn.execute("UPDATE uspsmethods SET uspsMethod='14' WHERE uspsID=39")
cnn.execute("UPDATE uspsmethods SET uspsMethod='15' WHERE uspsID=40")
cnn.execute("UPDATE uspsmethods SET uspsMethod='11' WHERE uspsID=41")

on error resume next
printtickdiv("Checking for new receipt upgrade")
err.number = 0
sSQL = "SELECT receiptheaders FROM emailmessages"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding new receipt columns")
	cnn.execute("ALTER TABLE emailmessages "&addcl&" receiptheaders "&memocl&" NULL")
	cnn.execute("ALTER TABLE emailmessages "&addcl&" receiptheaders2 "&memocl&" NULL")
	cnn.execute("ALTER TABLE emailmessages "&addcl&" receiptheaders3 "&memocl&" NULL")
	sSQL = "SELECT emailheaders,emailheaders2,emailheaders3 FROM emailmessages"
	rs.Open sSQL,cnn,0,1
	sSQL = "UPDATE emailmessages SET receiptheaders='" & replace(replace(rs("emailheaders")&"","%emailmessage%","%messagebody%"),"'","''") & "',receiptheaders2='" & replace(replace(rs("emailheaders2")&"","%emailmessage%","%messagebody%"),"'","''") & "',receiptheaders3='" & replace(replace(rs("emailheaders3")&"","%emailmessage%","%messagebody%"),"'","''") & "'"
	cnn.execute(sSQL)
	rs.Close
end if

cnn.execute("UPDATE uspsmethods SET uspsShowAs='UPS Worldwide Express&reg;' WHERE uspsID=104")
cnn.execute("UPDATE uspsmethods SET uspsShowAs='UPS Worldwide Expedited&reg;' WHERE uspsID=105")
cnn.execute("UPDATE uspsmethods SET uspsShowAs='UPS Worldwide Express Plus&reg;' WHERE uspsID=110")
cnn.execute("UPDATE uspsmethods SET uspsShowAs='UPS Saver&reg;' WHERE uspsID=112")

on error resume next
printtickdiv("Checking for Customer Lists upgrade")
err.number = 0
sSQL = "SELECT * FROM customerlists"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Customer Lists table")
	cnn.execute("CREATE TABLE customerlists (listID "&autoinc&",listName "&txtcl&"(255) NOT NULL,listOwner INT NOT NULL DEFAULT 0,listAccess "&txtcl&"(255) NOT NULL)")
	cnn.execute("CREATE INDEX listOwner_Indx ON customerlists(listOwner)")
end if

on error resume next
printtickdiv("Checking for Customer List Cart Column")
err.number = 0
sSQL = "SELECT cartListID FROM cart WHERE cartID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Customer List Cart Column")
	cnn.execute("ALTER TABLE cart "&addcl&" cartListID INT DEFAULT 0")
	cnn.execute("UPDATE cart SET cartListID=0")
	cnn.execute("CREATE INDEX cartListID_Indx ON cart(cartListID)")
end if

if FALSE AND sqlserver=TRUE then
	on error resume next
	cnn.execute("ALTER TABLE products "&altcl&" pDescription VARCHAR(MAX)")
	cnn.execute("ALTER TABLE products "&altcl&" pDescription2 VARCHAR(MAX)")
	cnn.execute("ALTER TABLE products "&altcl&" pDescription3 VARCHAR(MAX)")
	cnn.execute("ALTER TABLE products "&altcl&" pLongDescription VARCHAR(MAX)")
	cnn.execute("ALTER TABLE products "&altcl&" pLongDescription2 VARCHAR(MAX)")
	cnn.execute("ALTER TABLE products "&altcl&" pLongDescription3 VARCHAR(MAX)")
	on error goto 0
end if

cnn.execute("UPDATE countries SET countryCurrency='EUR' WHERE countryID IN (48,118,124,171,205)")
cnn.execute("UPDATE countries SET countryCurrency='TRY' WHERE countryID=194")
cnn.execute("UPDATE countries SET countryCurrency='CLP' WHERE countryID=41")
cnn.execute("UPDATE countries SET countryCurrency='CRC' WHERE countryID=45")
cnn.execute("UPDATE countries SET countryCurrency='BGN' WHERE countryID=32")
cnn.execute("UPDATE countries SET countryCurrency='VEF' WHERE countryID=206")
cnn.execute("UPDATE countries SET countryCurrency='USD' WHERE countryID=57")

on error resume next
printtickdiv("Checking for ISO 4217 Column")
err.number = 0
sSQL = "SELECT countryNumCurrency FROM countries WHERE countryID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding ISO 4217 Column")
	cnn.execute("ALTER TABLE countries "&addcl&" countryNumCurrency INT DEFAULT 0")
	cnn.execute("UPDATE countries SET countryNumCurrency=0")
	cnn.execute("UPDATE countries SET countryNumCurrency=784 WHERE countryCurrency='AED'")
	cnn.execute("UPDATE countries SET countryNumCurrency=971 WHERE countryCurrency='AFN'")
	cnn.execute("UPDATE countries SET countryNumCurrency=008 WHERE countryCurrency='ALL'")
	cnn.execute("UPDATE countries SET countryNumCurrency=051 WHERE countryCurrency='AMD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=532 WHERE countryCurrency='ANG'")
	cnn.execute("UPDATE countries SET countryNumCurrency=973 WHERE countryCurrency='AOA'")
	cnn.execute("UPDATE countries SET countryNumCurrency=032 WHERE countryCurrency='ARS'")
	cnn.execute("UPDATE countries SET countryNumCurrency=036 WHERE countryCurrency='AUD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=533 WHERE countryCurrency='AWG'")
	cnn.execute("UPDATE countries SET countryNumCurrency=944 WHERE countryCurrency='AZN'")
	cnn.execute("UPDATE countries SET countryNumCurrency=977 WHERE countryCurrency='BAM'")
	cnn.execute("UPDATE countries SET countryNumCurrency=052 WHERE countryCurrency='BBD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=050 WHERE countryCurrency='BDT'")
	cnn.execute("UPDATE countries SET countryNumCurrency=975 WHERE countryCurrency='BGN'")
	cnn.execute("UPDATE countries SET countryNumCurrency=048 WHERE countryCurrency='BHD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=108 WHERE countryCurrency='BIF'")
	cnn.execute("UPDATE countries SET countryNumCurrency=060 WHERE countryCurrency='BMD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=096 WHERE countryCurrency='BND'")
	cnn.execute("UPDATE countries SET countryNumCurrency=068 WHERE countryCurrency='BOB'")
	cnn.execute("UPDATE countries SET countryNumCurrency=984 WHERE countryCurrency='BOV'")
	cnn.execute("UPDATE countries SET countryNumCurrency=986 WHERE countryCurrency='BRL'")
	cnn.execute("UPDATE countries SET countryNumCurrency=044 WHERE countryCurrency='BSD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=064 WHERE countryCurrency='BTN'")
	cnn.execute("UPDATE countries SET countryNumCurrency=072 WHERE countryCurrency='BWP'")
	cnn.execute("UPDATE countries SET countryNumCurrency=974 WHERE countryCurrency='BYR'")
	cnn.execute("UPDATE countries SET countryNumCurrency=084 WHERE countryCurrency='BZD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=124 WHERE countryCurrency='CAD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=976 WHERE countryCurrency='CDF'")
	cnn.execute("UPDATE countries SET countryNumCurrency=947 WHERE countryCurrency='CHE'")
	cnn.execute("UPDATE countries SET countryNumCurrency=756 WHERE countryCurrency='CHF'")
	cnn.execute("UPDATE countries SET countryNumCurrency=948 WHERE countryCurrency='CHW'")
	cnn.execute("UPDATE countries SET countryNumCurrency=990 WHERE countryCurrency='CLF'")
	cnn.execute("UPDATE countries SET countryNumCurrency=152 WHERE countryCurrency='CLP'")
	cnn.execute("UPDATE countries SET countryNumCurrency=156 WHERE countryCurrency='CNY'")
	cnn.execute("UPDATE countries SET countryNumCurrency=170 WHERE countryCurrency='COP'")
	cnn.execute("UPDATE countries SET countryNumCurrency=970 WHERE countryCurrency='COU'")
	cnn.execute("UPDATE countries SET countryNumCurrency=188 WHERE countryCurrency='CRC'")
	cnn.execute("UPDATE countries SET countryNumCurrency=931 WHERE countryCurrency='CUC'")
	cnn.execute("UPDATE countries SET countryNumCurrency=192 WHERE countryCurrency='CUP'")
	cnn.execute("UPDATE countries SET countryNumCurrency=132 WHERE countryCurrency='CVE'")
	cnn.execute("UPDATE countries SET countryNumCurrency=203 WHERE countryCurrency='CZK'")
	cnn.execute("UPDATE countries SET countryNumCurrency=262 WHERE countryCurrency='DJF'")
	cnn.execute("UPDATE countries SET countryNumCurrency=208 WHERE countryCurrency='DKK'")
	cnn.execute("UPDATE countries SET countryNumCurrency=214 WHERE countryCurrency='DOP'")
	cnn.execute("UPDATE countries SET countryNumCurrency=012 WHERE countryCurrency='DZD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=233 WHERE countryCurrency='EEK'")
	cnn.execute("UPDATE countries SET countryNumCurrency=818 WHERE countryCurrency='EGP'")
	cnn.execute("UPDATE countries SET countryNumCurrency=232 WHERE countryCurrency='ERN'")
	cnn.execute("UPDATE countries SET countryNumCurrency=230 WHERE countryCurrency='ETB'")
	cnn.execute("UPDATE countries SET countryNumCurrency=978 WHERE countryCurrency='EUR'")
	cnn.execute("UPDATE countries SET countryNumCurrency=242 WHERE countryCurrency='FJD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=238 WHERE countryCurrency='FKP'")
	cnn.execute("UPDATE countries SET countryNumCurrency=826 WHERE countryCurrency='GBP'")
	cnn.execute("UPDATE countries SET countryNumCurrency=981 WHERE countryCurrency='GEL'")
	cnn.execute("UPDATE countries SET countryNumCurrency=936 WHERE countryCurrency='GHS'")
	cnn.execute("UPDATE countries SET countryNumCurrency=292 WHERE countryCurrency='GIP'")
	cnn.execute("UPDATE countries SET countryNumCurrency=270 WHERE countryCurrency='GMD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=324 WHERE countryCurrency='GNF'")
	cnn.execute("UPDATE countries SET countryNumCurrency=320 WHERE countryCurrency='GTQ'")
	cnn.execute("UPDATE countries SET countryNumCurrency=328 WHERE countryCurrency='GYD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=344 WHERE countryCurrency='HKD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=340 WHERE countryCurrency='HNL'")
	cnn.execute("UPDATE countries SET countryNumCurrency=191 WHERE countryCurrency='HRK'")
	cnn.execute("UPDATE countries SET countryNumCurrency=332 WHERE countryCurrency='HTG'")
	cnn.execute("UPDATE countries SET countryNumCurrency=348 WHERE countryCurrency='HUF'")
	cnn.execute("UPDATE countries SET countryNumCurrency=360 WHERE countryCurrency='IDR'")
	cnn.execute("UPDATE countries SET countryNumCurrency=376 WHERE countryCurrency='ILS'")
	cnn.execute("UPDATE countries SET countryNumCurrency=356 WHERE countryCurrency='INR'")
	cnn.execute("UPDATE countries SET countryNumCurrency=368 WHERE countryCurrency='IQD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=364 WHERE countryCurrency='IRR'")
	cnn.execute("UPDATE countries SET countryNumCurrency=352 WHERE countryCurrency='ISK'")
	cnn.execute("UPDATE countries SET countryNumCurrency=388 WHERE countryCurrency='JMD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=400 WHERE countryCurrency='JOD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=392 WHERE countryCurrency='JPY'")
	cnn.execute("UPDATE countries SET countryNumCurrency=404 WHERE countryCurrency='KES'")
	cnn.execute("UPDATE countries SET countryNumCurrency=417 WHERE countryCurrency='KGS'")
	cnn.execute("UPDATE countries SET countryNumCurrency=116 WHERE countryCurrency='KHR'")
	cnn.execute("UPDATE countries SET countryNumCurrency=174 WHERE countryCurrency='KMF'")
	cnn.execute("UPDATE countries SET countryNumCurrency=408 WHERE countryCurrency='KPW'")
	cnn.execute("UPDATE countries SET countryNumCurrency=410 WHERE countryCurrency='KRW'")
	cnn.execute("UPDATE countries SET countryNumCurrency=414 WHERE countryCurrency='KWD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=136 WHERE countryCurrency='KYD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=398 WHERE countryCurrency='KZT'")
	cnn.execute("UPDATE countries SET countryNumCurrency=418 WHERE countryCurrency='LAK'")
	cnn.execute("UPDATE countries SET countryNumCurrency=422 WHERE countryCurrency='LBP'")
	cnn.execute("UPDATE countries SET countryNumCurrency=144 WHERE countryCurrency='LKR'")
	cnn.execute("UPDATE countries SET countryNumCurrency=430 WHERE countryCurrency='LRD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=426 WHERE countryCurrency='LSL'")
	cnn.execute("UPDATE countries SET countryNumCurrency=440 WHERE countryCurrency='LTL'")
	cnn.execute("UPDATE countries SET countryNumCurrency=428 WHERE countryCurrency='LVL'")
	cnn.execute("UPDATE countries SET countryNumCurrency=434 WHERE countryCurrency='LYD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=504 WHERE countryCurrency='MAD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=498 WHERE countryCurrency='MDL'")
	cnn.execute("UPDATE countries SET countryNumCurrency=969 WHERE countryCurrency='MGA'")
	cnn.execute("UPDATE countries SET countryNumCurrency=807 WHERE countryCurrency='MKD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=104 WHERE countryCurrency='MMK'")
	cnn.execute("UPDATE countries SET countryNumCurrency=496 WHERE countryCurrency='MNT'")
	cnn.execute("UPDATE countries SET countryNumCurrency=446 WHERE countryCurrency='MOP'")
	cnn.execute("UPDATE countries SET countryNumCurrency=478 WHERE countryCurrency='MRO'")
	cnn.execute("UPDATE countries SET countryNumCurrency=480 WHERE countryCurrency='MUR'")
	cnn.execute("UPDATE countries SET countryNumCurrency=462 WHERE countryCurrency='MVR'")
	cnn.execute("UPDATE countries SET countryNumCurrency=454 WHERE countryCurrency='MWK'")
	cnn.execute("UPDATE countries SET countryNumCurrency=484 WHERE countryCurrency='MXN'")
	cnn.execute("UPDATE countries SET countryNumCurrency=979 WHERE countryCurrency='MXV'")
	cnn.execute("UPDATE countries SET countryNumCurrency=458 WHERE countryCurrency='MYR'")
	cnn.execute("UPDATE countries SET countryNumCurrency=943 WHERE countryCurrency='MZN'")
	cnn.execute("UPDATE countries SET countryNumCurrency=516 WHERE countryCurrency='NAD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=566 WHERE countryCurrency='NGN'")
	cnn.execute("UPDATE countries SET countryNumCurrency=558 WHERE countryCurrency='NIO'")
	cnn.execute("UPDATE countries SET countryNumCurrency=578 WHERE countryCurrency='NOK'")
	cnn.execute("UPDATE countries SET countryNumCurrency=524 WHERE countryCurrency='NPR'")
	cnn.execute("UPDATE countries SET countryNumCurrency=554 WHERE countryCurrency='NZD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=512 WHERE countryCurrency='OMR'")
	cnn.execute("UPDATE countries SET countryNumCurrency=590 WHERE countryCurrency='PAB'")
	cnn.execute("UPDATE countries SET countryNumCurrency=604 WHERE countryCurrency='PEN'")
	cnn.execute("UPDATE countries SET countryNumCurrency=598 WHERE countryCurrency='PGK'")
	cnn.execute("UPDATE countries SET countryNumCurrency=608 WHERE countryCurrency='PHP'")
	cnn.execute("UPDATE countries SET countryNumCurrency=586 WHERE countryCurrency='PKR'")
	cnn.execute("UPDATE countries SET countryNumCurrency=985 WHERE countryCurrency='PLN'")
	cnn.execute("UPDATE countries SET countryNumCurrency=600 WHERE countryCurrency='PYG'")
	cnn.execute("UPDATE countries SET countryNumCurrency=634 WHERE countryCurrency='QAR'")
	cnn.execute("UPDATE countries SET countryNumCurrency=946 WHERE countryCurrency='RON'")
	cnn.execute("UPDATE countries SET countryNumCurrency=941 WHERE countryCurrency='RSD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=643 WHERE countryCurrency='RUB'")
	cnn.execute("UPDATE countries SET countryNumCurrency=646 WHERE countryCurrency='RWF'")
	cnn.execute("UPDATE countries SET countryNumCurrency=682 WHERE countryCurrency='SAR'")
	cnn.execute("UPDATE countries SET countryNumCurrency=090 WHERE countryCurrency='SBD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=690 WHERE countryCurrency='SCR'")
	cnn.execute("UPDATE countries SET countryNumCurrency=938 WHERE countryCurrency='SDG'")
	cnn.execute("UPDATE countries SET countryNumCurrency=752 WHERE countryCurrency='SEK'")
	cnn.execute("UPDATE countries SET countryNumCurrency=702 WHERE countryCurrency='SGD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=654 WHERE countryCurrency='SHP'")
	cnn.execute("UPDATE countries SET countryNumCurrency=694 WHERE countryCurrency='SLL'")
	cnn.execute("UPDATE countries SET countryNumCurrency=706 WHERE countryCurrency='SOS'")
	cnn.execute("UPDATE countries SET countryNumCurrency=968 WHERE countryCurrency='SRD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=678 WHERE countryCurrency='STD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=760 WHERE countryCurrency='SYP'")
	cnn.execute("UPDATE countries SET countryNumCurrency=748 WHERE countryCurrency='SZL'")
	cnn.execute("UPDATE countries SET countryNumCurrency=764 WHERE countryCurrency='THB'")
	cnn.execute("UPDATE countries SET countryNumCurrency=972 WHERE countryCurrency='TJS'")
	cnn.execute("UPDATE countries SET countryNumCurrency=934 WHERE countryCurrency='TMT'")
	cnn.execute("UPDATE countries SET countryNumCurrency=788 WHERE countryCurrency='TND'")
	cnn.execute("UPDATE countries SET countryNumCurrency=776 WHERE countryCurrency='TOP'")
	cnn.execute("UPDATE countries SET countryNumCurrency=949 WHERE countryCurrency='TRY'")
	cnn.execute("UPDATE countries SET countryNumCurrency=780 WHERE countryCurrency='TTD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=901 WHERE countryCurrency='TWD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=834 WHERE countryCurrency='TZS'")
	cnn.execute("UPDATE countries SET countryNumCurrency=980 WHERE countryCurrency='UAH'")
	cnn.execute("UPDATE countries SET countryNumCurrency=800 WHERE countryCurrency='UGX'")
	cnn.execute("UPDATE countries SET countryNumCurrency=840 WHERE countryCurrency='USD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=997 WHERE countryCurrency='USN'")
	cnn.execute("UPDATE countries SET countryNumCurrency=998 WHERE countryCurrency='USS'")
	cnn.execute("UPDATE countries SET countryNumCurrency=858 WHERE countryCurrency='UYU'")
	cnn.execute("UPDATE countries SET countryNumCurrency=860 WHERE countryCurrency='UZS'")
	cnn.execute("UPDATE countries SET countryNumCurrency=937 WHERE countryCurrency='VEF'")
	cnn.execute("UPDATE countries SET countryNumCurrency=704 WHERE countryCurrency='VND'")
	cnn.execute("UPDATE countries SET countryNumCurrency=548 WHERE countryCurrency='VUV'")
	cnn.execute("UPDATE countries SET countryNumCurrency=882 WHERE countryCurrency='WST'")
	cnn.execute("UPDATE countries SET countryNumCurrency=950 WHERE countryCurrency='XAF'")
	cnn.execute("UPDATE countries SET countryNumCurrency=961 WHERE countryCurrency='XAG'")
	cnn.execute("UPDATE countries SET countryNumCurrency=959 WHERE countryCurrency='XAU'")
	cnn.execute("UPDATE countries SET countryNumCurrency=955 WHERE countryCurrency='XBA'")
	cnn.execute("UPDATE countries SET countryNumCurrency=956 WHERE countryCurrency='XBB'")
	cnn.execute("UPDATE countries SET countryNumCurrency=957 WHERE countryCurrency='XBC'")
	cnn.execute("UPDATE countries SET countryNumCurrency=958 WHERE countryCurrency='XBD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=951 WHERE countryCurrency='XCD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=960 WHERE countryCurrency='XDR'")
	cnn.execute("UPDATE countries SET countryNumCurrency=952 WHERE countryCurrency='XOF'")
	cnn.execute("UPDATE countries SET countryNumCurrency=964 WHERE countryCurrency='XPD'")
	cnn.execute("UPDATE countries SET countryNumCurrency=953 WHERE countryCurrency='XPF'")
	cnn.execute("UPDATE countries SET countryNumCurrency=962 WHERE countryCurrency='XPT'")
	cnn.execute("UPDATE countries SET countryNumCurrency=886 WHERE countryCurrency='YER'")
	cnn.execute("UPDATE countries SET countryNumCurrency=710 WHERE countryCurrency='ZAR'")
	cnn.execute("UPDATE countries SET countryNumCurrency=894 WHERE countryCurrency='ZMK'")
	cnn.execute("UPDATE countries SET countryNumCurrency=932 WHERE countryCurrency='ZWL'")
end if

on error resume next
printtickdiv("Checking for Cardinal Commerce authentication upgrade")
err.number = 0
sSQL = "SELECT cardinalProcessor FROM admin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Cardinal Commerce authentication columns")
	cnn.execute("ALTER TABLE admin "&addcl&" cardinalProcessor "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE admin "&addcl&" cardinalMerchant "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE admin "&addcl&" cardinalPwd "&txtcl&"(255) NULL")
	cnn.execute("UPDATE admin SET cardinalProcessor='',cardinalMerchant='',cardinalPwd=''")
end if

on error resume next
printtickdiv("Checking for catalog root upgrade")
err.number = 0
sSQL = "SELECT catalogRoot FROM admin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding catalog root column")
	cnn.execute("ALTER TABLE admin "&addcl&" catalogRoot INT DEFAULT 0")
	cnn.execute("UPDATE admin SET catalogRoot=0")
end if

on error resume next
printtickdiv("Checking for CMS Content Region upgrade")
err.number = 0
sSQL = "SELECT * FROM contentregions WHERE contentID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding CMS Content Region table")
	cnn.execute("CREATE TABLE contentregions (contentID "&autoinc&",contentName "&txtcl&"(255) NULL,contentX INT DEFAULT 0,contentY INT DEFAULT 0,contentData "&memocl&" NULL,contentData2 "&memocl&" NULL,contentData3 "&memocl&" NULL)")
end if

on error resume next
printtickdiv("Checking for mailinglist sent upgrade")
err.number = 0
sSQL = "SELECT emailsent FROM mailinglist WHERE email='xyxyxy'"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding mailinglist sent column")
	cnn.execute("ALTER TABLE mailinglist "&addcl&" emailsent "&bitfield&" DEFAULT 0")
	cnn.execute("UPDATE mailinglist SET emailsent=0")
end if

cnn.execute("UPDATE uspsmethods SET uspsShowAs='FedEx Priority Overnight&reg;' WHERE uspsID=301")
cnn.execute("UPDATE uspsmethods SET uspsShowAs='FedEx Standard Overnight&reg;' WHERE uspsID=302")
cnn.execute("UPDATE uspsmethods SET uspsShowAs='FedEx First Overnight&reg;' WHERE uspsID=303")
cnn.execute("UPDATE uspsmethods SET uspsShowAs='FedEx 2Day&reg;' WHERE uspsID=304")
cnn.execute("UPDATE uspsmethods SET uspsShowAs='FedEx Express Saver&reg;' WHERE uspsID=305")
cnn.execute("UPDATE uspsmethods SET uspsShowAs='FedEx International Priority&reg;' WHERE uspsID=306")
cnn.execute("UPDATE uspsmethods SET uspsShowAs='FedEx International Economy&reg;' WHERE uspsID=307")
cnn.execute("UPDATE uspsmethods SET uspsShowAs='FedEx International First&reg;' WHERE uspsID=308")
cnn.execute("UPDATE uspsmethods SET uspsShowAs='FedEx 1Day&reg; Freight' WHERE uspsID=310")
cnn.execute("UPDATE uspsmethods SET uspsShowAs='FedEx 2Day&reg; Freight' WHERE uspsID=311")
cnn.execute("UPDATE uspsmethods SET uspsShowAs='FedEx 3Day&reg; Freight' WHERE uspsID=312")
cnn.execute("UPDATE uspsmethods SET uspsShowAs='FedEx Ground&reg;' WHERE uspsID=313")
cnn.execute("UPDATE uspsmethods SET uspsShowAs='FedEx Home Delivery&reg;' WHERE uspsID=314")
cnn.execute("UPDATE uspsmethods SET uspsShowAs='FedEx International Priority&reg; Freight' WHERE uspsID=315")
cnn.execute("UPDATE uspsmethods SET uspsShowAs='FedEx International Economy&reg; Freight' WHERE uspsID=316")
cnn.execute("UPDATE uspsmethods SET uspsShowAs='FedEx Europe First&reg; - Int''l Priority' WHERE uspsID=317")

on error resume next
if sqlserver=TRUE then
	cnn.execute("ALTER TABLE admin ADD CONSTRAINT DF__admin__adminDelUnc DEFAULT 0 FOR adminDelUncompleted")
	cnn.execute("ALTER TABLE admin ADD CONSTRAINT DF__admin__adminStockM DEFAULT 0 FOR adminStockManage")
	cnn.execute("ALTER TABLE optiongroup ADD CONSTRAINT DF__optiongroup__optType DEFAULT 0 FOR optType")
	cnn.execute("ALTER TABLE products ADD CONSTRAINT DF__products__pInStock DEFAULT 0 FOR pInStock")
	cnn.execute("ALTER TABLE states ADD CONSTRAINT DF__states__stateZone DEFAULT 0 FOR stateZone")
else
	cnn.execute("ALTER TABLE admin "&altcl&" adminDelUncompleted INT DEFAULT 0")
	cnn.execute("ALTER TABLE admin "&altcl&" adminStockManage INT DEFAULT 0")
	cnn.execute("ALTER TABLE optiongroup "&altcl&" optType INT DEFAULT 0")
	cnn.execute("ALTER TABLE products "&altcl&" pInStock INT DEFAULT 0")
	cnn.execute("ALTER TABLE states "&altcl&" stateZone INT DEFAULT 0")
end if
cnn.execute("UPDATE countries SET countrylcid='' WHERE countrylcid IS NULL")
cnn.execute("UPDATE optiongroup SET opttype=0 WHERE opttype IS NULL")
cnn.execute("UPDATE products SET pinstock=0 WHERE pinstock IS NULL")
cnn.execute("UPDATE states SET statezone=0 WHERE statezone IS NULL")
on error goto 0

on error resume next
printtickdiv("Checking for Alternate Rates Admin upgrade")
err.number = 0
sSQL = "SELECT adminAltRates FROM admin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Alternate Rates Admin column")
	cnn.execute("ALTER TABLE admin "&addcl&" adminAltRates INT DEFAULT 0")
	cnn.execute("UPDATE admin SET adminAltRates=0")
end if

yyFlatShp="Flat Rate Shipping"
yyWghtShp="Weight Based Shipping"
yyPriShp="Price Based Shipping"
yyUSPS="U.S.P.S. Shipping"
yyUPS="UPS Shipping"
yyFedex="FedEx Shipping"
yyCanPos="Canada Post"

on error resume next
printtickdiv("Checking for Alternate Rates upgrade")
err.number = 0
sSQL = "SELECT * FROM alternaterates"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Alternate Rates table")
	cnn.execute("CREATE TABLE alternaterates (altrateid INT PRIMARY KEY,altratename "&txtcl&"(255) NOT NULL,altratetext "&txtcl&"(255) NULL,altratetext2 "&txtcl&"(255) NULL,altratetext3 "&txtcl&"(255) NULL, usealtmethod INT DEFAULT 0, usealtmethodintl INT DEFAULT 0, altrateorder INT DEFAULT 0)")
	for index=2 to 7
		if index=2 then altratetext=alternateratesweightbased
		if index=3 then altratetext=alternateratesusps
		if index=4 then altratetext=alternateratesups
		if index=5 then altratetext=alternateratespricebased
		if index=6 then altratetext=alternateratescanadapost
		if index=7 then altratetext=alternateratesfedex
		if altratetext<>"" then cnn.execute("UPDATE admin SET adminAltRates=1")
	next
end if

for index=1 to 7
	sSQL = "INSERT INTO alternaterates (altrateid,altratename,altratetext,altratetext2,altratetext3,usealtmethod,usealtmethodintl) VALUES ("

	if index=1 then altratetext=""
	if index=2 then altratetext=alternateratesweightbased
	if index=3 then altratetext=alternateratesusps
	if index=4 then altratetext=alternateratesups
	if index=5 then altratetext=alternateratespricebased
	if index=6 then altratetext=alternateratescanadapost
	if index=7 then altratetext=alternateratesfedex
	if altratetext<>"" then usealtmethod=1 else usealtmethod=0

	if index=1 then altratename=yyFlatShp : altratetext=yyFlatShp
	if index=2 then altratename=yyWghtShp : altratetext=IIfVr(alternateratesweightbased<>"",alternateratesweightbased,yyWghtShp)
	if index=3 then altratename=yyUSPS : altratetext=IIfVr(alternateratesusps<>"",alternateratesusps,yyUSPS)
	if index=4 then altratename=yyUPS : altratetext=IIfVr(alternateratesups<>"",alternateratesups,yyUPS)
	if index=5 then altratename=yyPriShp : altratetext=IIfVr(alternateratespricebased<>"",alternateratespricebased,yyPriShp)
	if index=6 then altratename=yyCanPos : altratetext=IIfVr(alternateratescanadapost<>"",alternateratescanadapost,yyCanPos)
	if index=7 then altratename=yyFedex : altratetext=IIfVr(alternateratesfedex<>"",alternateratesfedex,yyFedex)

	sSQL = sSQL & index & ",'" & escape_string(altratename) & "','" & escape_string(altratetext) & "','" & escape_string(altratetext) & "','" & escape_string(altratetext) & "'," & usealtmethod & "," & usealtmethod & ")"
	on error resume next
	cnn.execute(sSQL)
	on error goto 0
next

sSQL = "SELECT altrateid FROM alternaterates WHERE altrateid=8"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	sSQL = "INSERT INTO alternaterates (altrateid,altratename,altratetext,altratetext2,altratetext3,usealtmethod,usealtmethodintl) VALUES (" & _
		"8,'FedEx SmartPost&reg;','FedEx SmartPost&reg;','FedEx SmartPost&reg;','FedEx SmartPost&reg;',0,0)"
	cnn.execute(sSQL)
end if
rs.close

sSQL = "SELECT altrateid FROM alternaterates WHERE altrateid=9"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	sSQL = "INSERT INTO alternaterates (altrateid,altratename,altratetext,altratetext2,altratetext3,usealtmethod,usealtmethodintl) VALUES (" & _
		"9,'DHL Shipping','DHL Shipping','DHL Shipping','DHL Shipping',0,0)"
	cnn.execute(sSQL)
end if
rs.close

sSQL = "SELECT altrateid FROM alternaterates WHERE altrateid=10"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	sSQL = "INSERT INTO alternaterates (altrateid,altratename,altratetext,altratetext2,altratetext3,usealtmethod,usealtmethodintl) VALUES (" & _
		"10,'Australia Post','Australia Post','Australia Post','Australia Post',0,0)"
	cnn.execute(sSQL)
end if
rs.close

for index=1 to 10
	if index=1 then altratename=yyFlatShp
	if index=2 then altratename=yyWghtShp
	if index=3 then altratename=yyUSPS
	if index=4 then altratename=yyUPS
	if index=5 then altratename=yyPriShp
	if index=6 then altratename=yyCanPos
	if index=7 then altratename=yyFedex
	if index=8 then altratename="FedEx SmartPost&reg;"
	if index=9 then altratename="DHL Shipping"
	if index=10 then altratename="Australia Post"
	sSQL = "UPDATE alternaterates SET altratename='" & escape_string(altratename) & "' WHERE altratename='' AND altrateid=" & index
	cnn.execute(sSQL)
next

on error resume next
printtickdiv("Checking for Shipping Options upgrade")
err.number = 0
sSQL = "SELECT * FROM shipoptions WHERE soOrderID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Shipping Options table")
	cnn.execute("CREATE TABLE shipoptions (soIndex INT NOT NULL DEFAULT 0,soOrderID INT NOT NULL DEFAULT 0,soMethodName "&txtcl&"(255) NULL,soCost "&dblcl&" DEFAULT 0,soFreeShip "&bytecl&" DEFAULT 0,soShipType INT DEFAULT 0,soDeliveryTime "&txtcl&"(255) NULL,soDateAdded "&datecl&" NOT NULL, PRIMARY KEY(soIndex,soOrderID))")
	cnn.execute("CREATE INDEX soDateAdded_Indx ON shipoptions(soDateAdded)")
	
	on error resume next
	call drop_constraints("cart","cartsessionid")
	cnn.execute("DROP INDEX cartSessionID_Indx ON cart")
	cnn.execute("UPDATE cart SET cartSessionID=111111 WHERE cartSessionID IS NULL")
	cnn.execute("ALTER TABLE cart "&altcl&" cartSessionID "&idtxtcl&"(100) NOT NULL")
	cnn.execute("CREATE INDEX cartSessionID_Indx ON cart(cartSessionID)")

	call drop_constraints("orders","ordsessionid")
	cnn.execute("DROP INDEX ordSessionID_Indx ON orders")
	cnn.execute("UPDATE orders SET ordSessionID=111111 WHERE ordSessionID IS NULL")
	cnn.execute("ALTER TABLE orders "&altcl&" ordSessionID "&idtxtcl&"(100) NOT NULL")
	cnn.execute("CREATE INDEX ordSessionID_Indx ON orders(ordSessionID)")
	
	cnn.execute("DROP TABLE recentlyviewed")
	
	cnn.execute("DROP TABLE tmplogin")
	on error goto 0
end if

call drop_constraints("countries","countryLCID")
cnn.execute("ALTER TABLE countries "&altcl&" countryLCID "&idtxtcl&"(50) NOT NULL")
cnn.execute("UPDATE countries SET countryLCID='' WHERE countryLCID='0'")
cnn.execute("UPDATE countries SET countryLCID='18441' WHERE countryID=169")
cnn.execute("UPDATE countries SET countryLCID='1049' WHERE countryID=157")

on error resume next
printtickdiv("Checking for Recently Viewed upgrade")
err.number = 0
sSQL = "SELECT * FROM recentlyviewed"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Recently Viewed table")
	cnn.execute("CREATE TABLE recentlyviewed (rvID "&autoinc&",rvProdID "&txtcl&"(255) NOT NULL,rvProdName "&txtcl&"(255) NOT NULL,rvProdSection INT NOT NULL DEFAULT 0,rvProdURL "&txtcl&"(255) NOT NULL,rvSessionID "&idtxtcl&"(50) NOT NULL,rvCustomerID INT NOT NULL DEFAULT 0, rvDate "&datecl&" NOT NULL)")
end if

on error resume next
printtickdiv("Checking for Tmp Login upgrade")
err.number = 0
sSQL = "SELECT * FROM tmplogin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Tmp Login table")
	cnn.execute("CREATE TABLE tmplogin (tmploginid "&idtxtcl&"(100) PRIMARY KEY,tmploginname "&txtcl&"(50) NULL,tmplogindate "&datecl&",tmploginchk "&dblcl&" DEFAULT 0)")
end if

on error resume next
printtickdiv("Checking for temporary login table upgrade")
err.number = 0
sSQL = "SELECT tmploginchk FROM tmplogin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding temporary login table check column")
	cnn.execute("ALTER TABLE tmplogin "&addcl&" tmploginchk "&dblcl&" DEFAULT 0")
	cnn.execute("UPDATE tmplogin SET tmploginchk=0")
end if

on error resume next
printtickdiv("Checking for mailinglist selected upgrade")
err.number = 0
sSQL = "SELECT selected FROM mailinglist WHERE email='xyxyx'"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding mailinglist selected column")
	cnn.execute("ALTER TABLE mailinglist "&addcl&" selected "&bitfield&" DEFAULT 0")
	cnn.execute("UPDATE mailinglist SET selected=0")
end if

on error resume next
printtickdiv("Checking for discount login level")
err.number = 0
sSQL = "SELECT cpnLoginLevel FROM coupons"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding discount login level columns")
	cnn.execute("ALTER TABLE coupons "&addcl&" cpnLoginLevel INT DEFAULT 0")
	cnn.execute("UPDATE coupons SET cpnLoginLevel=0")
end if

on error resume next
printtickdiv("Checking for product filter upgrade")
err.number = 0
sSQL = "SELECT prodFilter FROM admin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding product filter columns")
	cnn.execute("ALTER TABLE admin "&addcl&" prodFilter INT DEFAULT 0")
	cnn.execute("ALTER TABLE admin "&addcl&" prodFilterText "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE admin "&addcl&" prodFilterText2 "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE admin "&addcl&" prodFilterText3 "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE admin "&addcl&" sortOrder INT DEFAULT 0")
	cnn.execute("ALTER TABLE admin "&addcl&" sortOptions INT DEFAULT 0")
	if filterresults<>"" then prodfilter=32 else prodfilter=0
	if sortBy="" OR NOT isnumeric(sortBy) then sortBy=0
	if sortBy=0 then sortOptions=0 else sortOptions=(2 ^ (sortBy-1))
	cnn.execute("UPDATE admin SET sortOrder="&sortBy&",sortOptions="&sortOptions&",prodFilter=" & prodfilter & ",prodFilterText='&&&&&"&escape_string(replace(filterresults,"&","%26"))&"',prodFilterText2='&&&&&"&escape_string(replace(filterresults,"&","%26"))&"',prodFilterText3='&&&&&"&escape_string(replace(filterresults,"&","%26"))&"'")
end if

newsearchcriteriatable=FALSE
searchcriteriatable="(scID INT PRIMARY KEY,scOrder INT DEFAULT 0,scGroup INT DEFAULT 0,scWorkingName "&txtcl&"(255) NULL,scName "&txtcl&"(255) NULL,scName2 "&txtcl&"(255) NULL,scName3 "&txtcl&"(255) NULL,scLogo "&txtcl&"(255) NULL,scURL "&txtcl&"(255) NULL,scURL2 "&txtcl&"(255) NULL,scURL3 "&txtcl&"(255) NULL,scEmail "&txtcl&"(255) NULL,scDescription "&memocl&" NULL,scDescription2 "&memocl&" NULL,scDescription3 "&memocl&" NULL,scNotes "&memocl&" NULL)"
on error resume next
printtickdiv("Checking for searchcriteria table upgrade")
err.number = 0
sSQL = "SELECT * FROM searchcriteria WHERE scID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	newsearchcriteriatable=TRUE
	printtick("Adding searchcriteria table")
	cnn.execute("CREATE TABLE searchcriteria "&searchcriteriatable)
end if

call checkaddcolumn("searchcriteria","scLogo",FALSE,txtcl,"(255)","")
call checkaddcolumn("searchcriteria","scURL",FALSE,txtcl,"(255)","")
call checkaddcolumn("searchcriteria","scURL2",FALSE,txtcl,"(255)","")
call checkaddcolumn("searchcriteria","scURL3",FALSE,txtcl,"(255)","")
call checkaddcolumn("searchcriteria","scEmail",FALSE,txtcl,"(255)","")
call checkaddcolumn("searchcriteria","scDescription",FALSE,memocl,"","")
call checkaddcolumn("searchcriteria","scDescription2",FALSE,memocl,"","")
call checkaddcolumn("searchcriteria","scDescription3",FALSE,memocl,"","")
call checkaddcolumn("searchcriteria","scNotes",FALSE,memocl,"","")

'on error resume next
'printtickdiv("Checking for Search Criteria upgrade")
'err.number = 0
'sSQL = "SELECT pSearchCriteria FROM products WHERE pID='xyxyx'"
'rs.Open sSQL,cnn,0,1
'errnum=err.number
'rs.Close
'on error goto 0
'if errnum<>0 then
'	printtick("Adding Search Criteria column")
'	cnn.execute("ALTER TABLE products "&addcl&" pSearchCriteria INT DEFAULT 0")
'	cnn.execute("UPDATE products SET pSearchCriteria=0")
'end if

if TRUE then ' Hash passwords
	sSQL = "SELECT adminPassword FROM admin WHERE adminID=1"
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then
		if len(rs("adminPassword")&"")<>32 then
			sSQL = "UPDATE admin SET adminPassword='"&escape_string(dohashpw(rs("adminPassword")))&"' WHERE adminID=1"
			cnn.execute(sSQL)
		end if
	end if
	rs.Close
	sSQL = "SELECT adminloginid,adminLoginPassword FROM adminlogin"
	rs.Open sSQL,cnn,0,1
	do while NOT rs.EOF
		if len(rs("adminLoginPassword")&"")<>32 then
			sSQL = "UPDATE adminlogin SET adminLoginPassword='"&escape_string(dohashpw(rs("adminLoginPassword")))&"' WHERE adminloginid="&rs("adminloginid")
			cnn.execute(sSQL)
		end if
		rs.movenext
	loop
	rs.Close
	sSQL = "SELECT affilID,affilPW FROM affiliates"
	rs.Open sSQL,cnn,0,1
	do while NOT rs.EOF
		if len(rs("affilPW")&"")<>32 then
			sSQL = "UPDATE affiliates SET affilPW='"&escape_string(dohashpw(rs("affilPW")))&"' WHERE affilID='"&escape_string(rs("affilID"))&"'"
			cnn.execute(sSQL)
		end if
		rs.movenext
	loop
	rs.Close
	sSQL = "SELECT clID,clPW FROM customerlogin"
	rs.Open sSQL,cnn,0,1
	do while NOT rs.EOF
		if len(rs("clPW")&"")<>32 then
			sSQL = "UPDATE customerlogin SET clPW='"&escape_string(dohashpw(rs("clPW")))&"' WHERE clID="&rs("clID")
			cnn.execute(sSQL)
		end if
		rs.movenext
	loop
	rs.Close
end if

on error resume next
printtickdiv("Checking for notifyinstock table upgrade")
err.number = 0
sSQL = "SELECT * FROM notifyinstock WHERE nsProdID='xyxyxyxyx'"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding notifyinstock table")
	cnn.execute("CREATE TABLE notifyinstock (nsProdID "&txtcl&"(150) NOT NULL,nsOptID INT DEFAULT 0,nsTriggerProdID "&txtcl&"(150) NOT NULL,nsEmail "&txtcl&"(150) NOT NULL,nsDate "&datecl&", PRIMARY KEY(nsTriggerProdID,nsEmail))")
end if

on error resume next
printtickdiv("Checking for notify back in stock email upgrade")
err.number = 0
sSQL = "SELECT notifystocksubject FROM emailmessages"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding notify back in stock email columns")
	cnn.execute("ALTER TABLE emailmessages "&addcl&" notifystocksubject "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE emailmessages "&addcl&" notifystocksubject2 "&txtcl&"(255) NULL")
	cnn.execute("ALTER TABLE emailmessages "&addcl&" notifystocksubject3 "&txtcl&"(255) NULL")

	cnn.execute("ALTER TABLE emailmessages "&addcl&" notifystockemail "&memocl&" NULL")
	cnn.execute("ALTER TABLE emailmessages "&addcl&" notifystockemail2 "&memocl&" NULL")
	cnn.execute("ALTER TABLE emailmessages "&addcl&" notifystockemail3 "&memocl&" NULL")

	cnn.execute("UPDATE emailmessages SET notifystocksubject='" & escape_string("We now have stock for %pname%") & "'")
	cnn.execute("UPDATE emailmessages SET notifystockemail='" & escape_string("The product %pid% / %pname% is now back in stock.%nl%%nl%You can find this in our store at the following location:%nl%%link%%nl%%nl%Many Thanks%nl%%nl%%storeurl%%nl%") & "'")
	cnn.execute("UPDATE emailmessages SET notifystocksubject2=notifystocksubject,notifystocksubject3=notifystocksubject,notifystockemail2=notifystockemail,notifystockemail3=notifystockemail")
end if

on error resume next
printtickdiv("Checking for admin login password upgrade")
err.number = 0
sSQL = "SELECT adminLoginLastChange FROM adminlogin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding admin login password columns")
	cnn.execute("ALTER TABLE adminlogin "&addcl&" adminLoginLastChange "&datecl)
	cnn.execute("ALTER TABLE adminlogin "&addcl&" adminLoginLock INT DEFAULT 0")
	cnn.execute("UPDATE adminlogin SET adminLoginLock=0,adminLoginLastChange="&datedelim&vsusdate(Now())&datedelim)
end if

on error resume next
printtickdiv("Checking for order cnum downgrade")
err.number = 0
sSQL = "SELECT ordCNum FROM orders WHERE ordID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum=0 then ' Note checking to see if EXISTS
	printtickdiv("Applying order cnum downgrade")
	cnn.execute("UPDATE orders SET ordCNum='01010101010101010101010101010101010101010101010101010101010101' WHERE ordPayProvider=10")
	cnn.execute("UPDATE orders SET ordCNum='10101010101010101010101010101010101010101010101010101010101010' WHERE ordPayProvider=10")
	cnn.execute("UPDATE orders SET ordCNum='ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890' WHERE ordPayProvider=10")
end if

on error resume next
printtickdiv("Checking for admin password upgrade")
err.number = 0
sSQL = "SELECT adminPWLastChange FROM admin"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding admin password columns")
	cnn.execute("ALTER TABLE admin "&addcl&" adminPWLastChange "&datecl)
	cnn.execute("ALTER TABLE admin "&addcl&" adminUserLock INT DEFAULT 0")
	cnn.execute("UPDATE admin SET adminUserLock=0,adminPWLastChange="&datedelim&vsusdate(Now())&datedelim)
end if

on error resume next
printtickdiv("Checking for Password History upgrade")
err.number = 0
sSQL = "SELECT * FROM passwordhistory"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Password History table")
	cnn.execute("CREATE TABLE passwordhistory (pwhID "&autoinc&",liID INT DEFAULT 0,pwhPwd "&txtcl&"(50),datePWChanged "&datecl&")")
end if

on error resume next
printtickdiv("Checking for Audit Log upgrade")
err.number = 0
sSQL = "SELECT * FROM auditlog WHERE logID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Audit Log table")
	cnn.execute("CREATE TABLE auditlog (logID "&autoinc&",userID "&txtcl&"(50),eventType "&txtcl&"(50),eventDate "&datecl&",eventSuccess "&bitfield&" DEFAULT 0,eventOrigin "&txtcl&"(50),areaAffected "&txtcl&"(50))")
end if

on error resume next
printtickdiv("Checking for Option Group upgrade")
err.number = 0
sSQL = "SELECT optTxtCharge FROM optiongroup WHERE optGrpID=0"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding Option Group column")
	cnn.execute("ALTER TABLE optiongroup "&addcl&" optTxtMaxLen INT DEFAULT 0")
	cnn.execute("ALTER TABLE optiongroup "&addcl&" optTxtCharge "&dblcl&" DEFAULT 0")
	cnn.execute("UPDATE optiongroup SET optTxtMaxLen=0,optTxtCharge=0")
end if

on error resume next
printtickdiv("Checking for Option Multiplier upgrade")
err.number = 0
sSQL = "SELECT optMultiply FROM optiongroup WHERE optGrpID=0"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding Option Group column")
	cnn.execute("ALTER TABLE optiongroup "&addcl&" optMultiply "&bitfield&" DEFAULT 0")
	cnn.execute("ALTER TABLE optiongroup "&addcl&" optAcceptChars "&txtcl&"(255)")
	cnn.execute("UPDATE optiongroup SET optMultiply=0")
end if

on error resume next
printtickdiv("Checking for cartoptions multiplier upgrade")
err.number = 0
sSQL = "SELECT coMultiply FROM cartoptions WHERE coID=0"
cnn.execute(sSQL)
errnum=err.number
on error goto 0
if errnum<>0 then
	printtick("Adding cartoptions multiplier column")
	sSQL = "ALTER TABLE cartoptions "&addcl&" coMultiply "&bitfield&" NOT NULL DEFAULT 0"
	cnn.execute(sSQL)
end if

on error resume next
printtickdiv("Checking for Customer Login Loyalty Points upgrade")
err.number = 0
sSQL = "SELECT loyaltyPoints FROM customerlogin WHERE clID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Customer Login Loyalty Points column")
	cnn.execute("ALTER TABLE customerlogin "&addcl&" loyaltyPoints INT DEFAULT 0")
	cnn.execute("UPDATE customerlogin SET loyaltyPoints=0")
end if

on error resume next
printtickdiv("Checking for Orders Loyalty Points upgrade")
err.number = 0
sSQL = "SELECT loyaltyPoints FROM orders WHERE ordID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Orders Loyalty Points column")
	cnn.execute("ALTER TABLE orders "&addcl&" loyaltyPoints INT DEFAULT 0")
	cnn.execute("ALTER TABLE orders "&addcl&" pointsRedeemed INT DEFAULT 0")
	cnn.execute("UPDATE orders SET loyaltyPoints=0,pointsRedeemed=0")
end if

' printtick("Disable Capture Card")
sSQL = "UPDATE payprovider SET payProvEnabled=0,payProvAvailable=0 WHERE payProvID=10"
cnn.execute(sSQL)

idlist="0"
for index=1 to 10
	sSQL="SELECT sectionID,sectionDisabled,rootSection FROM sections WHERE rootSection=0 AND topSection IN (" & idlist & ")"
	idlist=""
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		sSQL="UPDATE sections SET sectionDisabled=" & rs("sectionDisabled") & " WHERE topSection=" & rs("sectionID") & " AND sectionDisabled<" & rs("sectionDisabled")
		cnn.execute(sSQL)
		idlist=idlist&rs("sectionID")&","
		rs.movenext
	loop
	rs.close
	if idlist<>"" then idlist=left(idlist,len(idlist)-1) else exit for
next

call checkaddcolumn("shipoptions","soFreeShipExempt",FALSE,"INT","","")
call checkaddcolumn("orders","ordPrivateStatus",FALSE,memocl,"","")
call checkaddcolumn("sections","sectionHeader",FALSE,memocl,"","")
call checkaddcolumn("sections","sectionHeader2",FALSE,memocl,"","")
call checkaddcolumn("sections","sectionHeader3",FALSE,memocl,"","")
call checkaddcolumn("products","pGiftWrap",FALSE,bitfield,"","")
call checkaddcolumn("products","pBackOrder",FALSE,bitfield,"","")
call checkaddcolumn("cart","cartGiftWrap",FALSE,bitfield,"","")
call checkaddcolumn("cart","cartGiftMessage",FALSE,memocl,"","")
call checkaddcolumn("admin","adminlang",FALSE,txtcl,"(10)","")
call checkaddcolumn("admin","storelang",FALSE,txtcl,"(10)","")

if checkaddcolumn("countries","loadStates",FALSE,"INT","","") then
	cnn.execute("UPDATE countries SET loadStates=2")
end if

nextfreeid=1
sub addstate(stateCountryID,stateName,stateAbbrev)
	call doaddstate(stateCountryID,stateName,stateAbbrev,1)
end sub
sub adddisabledstate(stateCountryID,stateName,stateAbbrev)
	call doaddstate(stateCountryID,stateName,stateAbbrev,0)
end sub
sub doaddstate(stateCountryID,stateName,stateAbbrev,stateEnabled)
	gotstateid=FALSE
	do while NOT gotstateid
		rs.open "SELECT stateID FROM states WHERE stateID=" & nextfreeid,cnn,0,1
		if rs.EOF then gotstateid=TRUE else nextfreeid=nextfreeid+1
		rs.close
	loop
	cnn.execute("INSERT INTO states (stateID,stateCountryID,stateName,stateAbbrev,stateTax,stateEnabled,stateZone,stateFreeShip) VALUES (" & nextfreeid & "," & stateCountryID & ",'" & escape_string(stateName) & "','" & escape_string(stateAbbrev) & "',0," & stateEnabled & ",0,0)")
end sub
if checkaddcolumn("states","stateCountryID",FALSE,"INT","","") then
	on error resume next
	call drop_constraints("states","stateName")
	if sqlserver then
		cnn.execute("DROP INDEX states.stateName")
	elseif mysqlserver then
		cnn.execute("ALTER TABLE states DROP INDEX stateName")
	else
		cnn.execute("DROP INDEX stateName ON states")
	end if
	on error goto 0
end if
sSQL = "SELECT stateID FROM states WHERE stateCountryID=0"
rs.open sSQL,cnn,0,1
updatestates=NOT rs.EOF
rs.close
if updatestates then ' {
	' USA
	statelist="'AL','AK','AS','AZ','AR','CA','CO','CT','DE','DC','FM','FL','GA','GU','HI','ID','IL','IN','IA','KS','KY','LA','ME','MH','MD','MA','MI','MN','MS','MO','MT','NE','NV','NH','NJ','NM','NY','NC','ND','MP','OH','OK','OR','PW','PA','PR','RI','SC','SD','TN','TX','UT','VT','VI','VA','WA','WV','WI','WY','AE','AA','AE','AE','AE','AP'"
	rs.open "SELECT COUNT(*) AS numstates FROM states WHERE stateAbbrev IN (" & statelist & ")",cnn,0,1
	if isnull(rs("numstates")) then numstates=0 else numstates=cint(rs("numstates"))
	rs.close
	if numstates<60 then
		call addstate(1,"Alabama","AL")
		call addstate(1,"Alaska","AK")
		call addstate(1,"American Samoa","AS")
		call addstate(1,"Arizona","AZ")
		call addstate(1,"Arkansas","AR")
		call addstate(1,"California","CA")
		call addstate(1,"Colorado","CO")
		call addstate(1,"Connecticut","CT")
		call addstate(1,"Delaware","DE")
		call addstate(1,"District Of Columbia","DC")
		call addstate(1,"Fdr. States Of Micronesia","FM")
		call addstate(1,"Florida","FL")
		call addstate(1,"Georgia","GA")
		call addstate(1,"Guam","GU")
		call addstate(1,"Hawaii","HI")
		call addstate(1,"Idaho","ID")
		call addstate(1,"Illinois","IL")
		call addstate(1,"Indiana","IN")
		call addstate(1,"Iowa","IA")
		call addstate(1,"Kansas","KS")
		call addstate(1,"Kentucky","KY")
		call addstate(1,"Louisiana","LA")
		call addstate(1,"Maine","ME")
		call addstate(1,"Marshall Islands","MH")
		call addstate(1,"Maryland","MD")
		call addstate(1,"Massachusetts","MA")
		call addstate(1,"Michigan","MI")
		call addstate(1,"Minnesota","MN")
		call addstate(1,"Mississippi","MS")
		call addstate(1,"Missouri","MO")
		call addstate(1,"Montana","MT")
		call addstate(1,"Nebraska","NE")
		call addstate(1,"Nevada","NV")
		call addstate(1,"New Hampshire","NH")
		call addstate(1,"New Jersey","NJ")
		call addstate(1,"New Mexico","NM")
		call addstate(1,"New York","NY")
		call addstate(1,"North Carolina","NC")
		call addstate(1,"North Dakota","ND")
		call addstate(1,"Northern Mariana Islands","MP")
		call addstate(1,"Ohio","OH")
		call addstate(1,"Oklahoma","OK")
		call addstate(1,"Oregon","OR")
		call addstate(1,"Palau","PW")
		call addstate(1,"Pennsylvania","PA")
		call addstate(1,"Puerto Rico","PR")
		call addstate(1,"Rhode Island","RI")
		call addstate(1,"South Carolina","SC")
		call addstate(1,"South Dakota","SD")
		call addstate(1,"Tennessee","TN")
		call addstate(1,"Texas","TX")
		call addstate(1,"Utah","UT")
		call addstate(1,"Vermont","VT")
		call addstate(1,"Virgin Islands","VI")
		call addstate(1,"Virginia","VA")
		call addstate(1,"Washington","WA")
		call addstate(1,"West Virginia","WV")
		call addstate(1,"Wisconsin","WI")
		call addstate(1,"Wyoming","WY")
		call adddisabledstate(1,"Armed Forces Africa","AE")
		call adddisabledstate(1,"Armed Forces Americas","AA")
		call adddisabledstate(1,"Armed Forces Canada","AE")
		call adddisabledstate(1,"Armed Forces Europe","AE")
		call adddisabledstate(1,"Armed Forces Middle East","AE")
		call adddisabledstate(1,"Armed Forces Pacific","AP")
	else
		cnn.execute("UPDATE states SET stateCountryID=1 WHERE stateCountryID=0 AND stateAbbrev IN (" & statelist & ")")
	end if
	' Canada
	statelist="'AB','BC','MB','NB','NF','NT','NS','NU','ON','PE','QC','SK','YT'"
	rs.open "SELECT COUNT(*) AS numstates FROM states WHERE stateAbbrev IN (" & statelist & ")",cnn,0,1
	if isnull(rs("numstates")) then numstates=0 else numstates=cint(rs("numstates"))
	rs.close
	if numstates<12 then
		call addstate(2,"Alberta","AB")
		call addstate(2,"British Columbia","BC")
		call addstate(2,"Manitoba","MB")
		call addstate(2,"New Brunswick","NB")
		call addstate(2,"Newfoundland","NF")
		call addstate(2,"North West Territories","NT")
		call addstate(2,"Nova Scotia","NS")
		call addstate(2,"Nunavut","NU")
		call addstate(2,"Ontario","ON")
		call addstate(2,"Prince Edward Island","PE")
		call addstate(2,"Quebec","QC")
		call addstate(2,"Saskatchewan","SK")
		call addstate(2,"Yukon Territory","YT")
	else
		cnn.execute("UPDATE states SET stateCountryID=2 WHERE stateCountryID=0 AND stateAbbrev IN (" & statelist & ")")
	end if
	' Australia
	statelist="'ACT','NSW','NT','QLD','SA','TA','VIC','WA'"
	rs.open "SELECT COUNT(*) AS numstates FROM states WHERE stateAbbrev IN (" & statelist & ")",cnn,0,1
	if isnull(rs("numstates")) then numstates=0 else numstates=cint(rs("numstates"))
	rs.close
	if numstates<8 then
		call addstate(14,"Australian Capital Territory","ACT")
		call addstate(14,"New South Wales","NSW")
		call addstate(14,"Northern Territory","NT")
		call addstate(14,"Queensland","QLD")
		call addstate(14,"South Australia","SA")
		call addstate(14,"Tasmania","TA")
		call addstate(14,"Victoria","VIC")
		call addstate(14,"Western Australia","WA")
	else
		cnn.execute("UPDATE states SET stateCountryID=14 WHERE stateCountryID=0")
	end if
	' Ireland
	statelist="'Carlow','Cavan','Clare','Cork','Donegal','Dublin','Galway','Kerry','Kildare','Kilkenny'"
	rs.open "SELECT COUNT(*) AS numstates FROM states WHERE stateName IN (" & statelist & ")",cnn,0,1
	if isnull(rs("numstates")) then numstates=0 else numstates=cint(rs("numstates"))
	rs.close
	if numstates<10 then
		call addstate(91,"Carlow","CA")
		call addstate(91,"Cavan","CV")
		call addstate(91,"Clare","CL")
		call addstate(91,"Cork","CO")
		call addstate(91,"Donegal","DO")
		call addstate(91,"Dublin","DU")
		call addstate(91,"Galway","GA")
		call addstate(91,"Kerry","KE")
		call addstate(91,"Kildare","KI")
		call addstate(91,"Kilkenny","KL")
		call addstate(91,"Laois","LA")
		call addstate(91,"Leitrim","LE")
		call addstate(91,"Limerick","LI")
		call addstate(91,"Longford","LO")
		call addstate(91,"Louth","LU")
		call addstate(91,"Mayo","MA")
		call addstate(91,"Meath","ME")
		call addstate(91,"Monaghan","MO")
		call addstate(91,"Offaly","OF")
		call addstate(91,"Roscommon","RO")
		call addstate(91,"Sligo","SL")
		call addstate(91,"Tipperary","TI")
		call addstate(91,"Waterford","WA")
		call addstate(91,"Westmeath","WE")
		call addstate(91,"Wexford","WX")
		call addstate(91,"Wicklow","WI")
	else
		cnn.execute("UPDATE states SET stateCountryID=91 WHERE stateCountryID=0")
	end if
	' New Zealand
	statelist="'Southland','Westland','Waikato','Marlborough','Canterbury'"
	rs.open "SELECT COUNT(*) AS numstates FROM states WHERE stateName IN (" & statelist & ")",cnn,0,1
	if isnull(rs("numstates")) then numstates=0 else numstates=cint(rs("numstates"))
	rs.close
	if numstates<5 then
		call addstate(136,"Ashburton","AS")
		call addstate(136,"Auckland","AU")
		call addstate(136,"Bay of Plenty","BP")
		call addstate(136,"Buller","BU")
		call addstate(136,"Canterbury","CB")
		call addstate(136,"Carterton","CA")
		call addstate(136,"Central Otago","CO")
		call addstate(136,"Clutha","CL")
		call addstate(136,"Counties Manukau","CM")
		call addstate(136,"Dunedin City","DC")
		call addstate(136,"Far North","FN")
		call addstate(136,"Franklin","FR")
		call addstate(136,"Gisborne","GS")
		call addstate(136,"Gore","GO")
		call addstate(136,"Grey","GR")
		call addstate(136,"Hamilton City","HC")
		call addstate(136,"Hastings","HS")
		call addstate(136,"Hauraki","HI")
		call addstate(136,"Hawke's Bay","HB")
		call addstate(136,"Horowhenua","HW")
		call addstate(136,"Hurunui","HU")
		call addstate(136,"Hutt Valley","HV")
		call addstate(136,"Invercargill","IC")
		call addstate(136,"Kaikoura","KK")
		call addstate(136,"Kaipara","KP")
		call addstate(136,"Kapiti Coast","KC")
		call addstate(136,"Kawerau","KW")
		call addstate(136,"Manawatu","MW")
		call addstate(136,"Marlborough","MB")
		call addstate(136,"Masteron","MS")
		call addstate(136,"Matamata Piako","MP")
		call addstate(136,"New Plymouth","NP")
		call addstate(136,"North Shore City","NS")
		call addstate(136,"Otaki","OT")
		call addstate(136,"Otorohanga","OT")
		call addstate(136,"Palmerston North","PN")
		call addstate(136,"Papakura","PK")
		call addstate(136,"Porirua City","PC")
		call addstate(136,"Queenstown Lakes","QL")
		call addstate(136,"Rotorua","RT")
		call addstate(136,"Ruapehu","RU")
		call addstate(136,"Selwyn","SN")
		call addstate(136,"South Taranaki","ST")
		call addstate(136,"South Waikato","SW")
		call addstate(136,"South Wairarapa","SA")
		call addstate(136,"Southland","SL")
		call addstate(136,"Stratford","SF")
		call addstate(136,"Tasman","TM")
		call addstate(136,"Taupo","TP")
		call addstate(136,"Tauranga","TR")
		call addstate(136,"Thames Coromandel","TC")
		call addstate(136,"Timaru","TM")
		call addstate(136,"Waikato","WK")
		call addstate(136,"Waimakariri","WM")
		call addstate(136,"Waimate","WE")
		call addstate(136,"Waiora","WO")
		call addstate(136,"Waipa","WP")
		call addstate(136,"Waitakere","WT")
		call addstate(136,"Waitaki","WI")
		call addstate(136,"Waitomo","Wa")
		call addstate(136,"Wellington City","WC")
		call addstate(136,"Western Bay of Plenty","WB")
		call addstate(136,"Westland","WL")
		call addstate(136,"Whakatane","WH")
		call addstate(136,"Whanganui","WG")
		call addstate(136,"Whangarei","WE")
	else
		cnn.execute("UPDATE states SET stateCountryID=136 WHERE stateCountryID=0")
	end if
	' South Africa
	statelist="'Eastern Cape','Free State','Gauteng','Kwazulu-Natal','Mpumalanga'"
	rs.open "SELECT COUNT(*) AS numstates FROM states WHERE stateName IN (" & statelist & ")",cnn,0,1
	if isnull(rs("numstates")) then numstates=0 else numstates=cint(rs("numstates"))
	rs.close
	if numstates<5 then
		call addstate(174,"Eastern Cape","EP")
		call addstate(174,"Free State","OFS")
		call addstate(174,"Gauteng","GA")
		call addstate(174,"Kwazulu-Natal","KZN")
		call addstate(174,"Mpumalanga","MP")
		call addstate(174,"Northern Cape","NC")
		call addstate(174,"Limpopo","LI")
		call addstate(174,"North West Province","NWP")
		call addstate(174,"Western Cape","WC")
	else
		cnn.execute("UPDATE states SET stateCountryID=174 WHERE stateCountryID=0")
	end if
	' UK
	statelist="'Aberdeenshire','Angus','Argyll','Avon','Ayrshire'"
	rs.open "SELECT COUNT(*) AS numstates FROM states WHERE stateName IN (" & statelist & ")",cnn,0,1
	if isnull(rs("numstates")) then numstates=0 else numstates=cint(rs("numstates"))
	rs.close
	if numstates<5 then
		call addstate(201,"Aberdeenshire","AB")
		call addstate(201,"Angus","AG")
		call addstate(201,"Argyll","AR")
		call addstate(201,"Avon","AV")
		call addstate(201,"Ayrshire","AY")
		call addstate(201,"Banffshire","BF")
		call addstate(201,"Bedfordshire","Beds")
		call addstate(201,"Berkshire","Berks")
		call addstate(201,"Buckinghamshire","Bucks")
		call addstate(201,"Caithness","CN")
		call addstate(201,"Cambridgeshire","Cambs")
		call addstate(201,"Ceredigion","CE")
		call addstate(201,"Cheshire","CH")
		call addstate(201,"Clackmannanshire","CL")
		call addstate(201,"Cleveland","CV")
		call addstate(201,"Clwyd","CW")
		call addstate(201,"County Antrim","Co Antrim")
		call addstate(201,"County Armagh","Co Armagh")
		call addstate(201,"County Down","Co Down")
		call addstate(201,"County Durham","Co Durham")
		call addstate(201,"County Fermanagh","Co Fermanagh")
		call addstate(201,"County Londonderry","Co Londonderry")
		call addstate(201,"County Tyrone","Co Tyrone")
		call addstate(201,"Cornwall","CO")
		call addstate(201,"Cumbria","CU")
		call addstate(201,"Derbyshire","DB")
		call addstate(201,"Devon","DV")
		call addstate(201,"Dorset","DO")
		call addstate(201,"Dumfriesshire","DF")
		call addstate(201,"Dunbartonshire","DU")
		call addstate(201,"Dyfed","DY")
		call addstate(201,"East Lothian","EL")
		call addstate(201,"East Sussex","E Sussex")
		call addstate(201,"Essex","EX")
		call addstate(201,"Fife","FI")
		call addstate(201,"Gloucestershire","Glos")
		call addstate(201,"Gwent","GW")
		call addstate(201,"Gwynedd","GY")
		call addstate(201,"Hampshire","Hants")
		call addstate(201,"Herefordshire","HE")
		call addstate(201,"Hertfordshire","Herts")
		call addstate(201,"Inverness-shire","IS")
		call addstate(201,"Isle of Mull","IsMu")
		call addstate(201,"Isle of Shetland","IsSh")
		call addstate(201,"Isle of Skye","IsSk")
		call addstate(201,"Isle of Wight","IsWi")
		call addstate(201,"Isles of Scilly","IsSc")
		call addstate(201,"Kent","KE")
		call addstate(201,"Kincardineshire","KI")
		call addstate(201,"Kinross-shire","KR")
		call addstate(201,"Kirkudbrightshire","KK")
		call addstate(201,"Lanarkshire","LK")
		call addstate(201,"Lancashire","Lancs")
		call addstate(201,"Leicestershire","Leics")
		call addstate(201,"Lincolnshire","Lincs")
		call addstate(201,"London","LO")
		call addstate(201,"Merseyside","ME")
		call addstate(201,"Mid Glamorgan","M Glam")
		call addstate(201,"Midlothian","MI")
		call addstate(201,"Middlesex","Middx")
		call addstate(201,"Morayshire","MO")
		call addstate(201,"Nairnshire","NA")
		call addstate(201,"Norfolk","NO")
		call addstate(201,"North Humberside","N Humberside")
		call addstate(201,"North Yorkshire","N Yorkshire")
		call addstate(201,"Northamptonshire","Northants")
		call addstate(201,"Northumberland","Northd")
		call addstate(201,"Nottinghamshire","Notts")
		call addstate(201,"Oxfordshire","Oxon")
		call addstate(201,"Peebleshire","PE")
		call addstate(201,"Perthshire","PR")
		call addstate(201,"Powys","PO")
		call addstate(201,"Renfrewshire","RE")
		call addstate(201,"Ross-shire","RO")
		call addstate(201,"Roxburghshire","RX")
		call addstate(201,"Selkirkshire","SK")
		call addstate(201,"Shropshire","SR")
		call addstate(201,"Somerset","SO")
		call addstate(201,"South Glamorgan","S Glam")
		call addstate(201,"South Humberside","S Humberside")
		call addstate(201,"South Yorkshire","S Yorkshire")
		call addstate(201,"Staffordshire","Staffs")
		call addstate(201,"Stirlingshire","SS")
		call addstate(201,"Suffolk","SF")
		call addstate(201,"Surrey","SY")
		call addstate(201,"Sutherland","SU")
		call addstate(201,"Tyne and Wear","Tyne & Wear")
		call addstate(201,"Warwickshire","Warks")
		call addstate(201,"West Glamorgan","W Glam")
		call addstate(201,"West Lothian","WL")
		call addstate(201,"West Midlands","W Midlands")
		call addstate(201,"West Sussex","W Sussex")
		call addstate(201,"West Yorkshire","W Yorkshire")
		call addstate(201,"Wigtownshire","WT")
		call addstate(201,"Wiltshire","Wilts")
		call addstate(201,"Worcestershire","Worcs")
		call addstate(201,"East Yorkshire","EY")
		call addstate(201,"Carmarthenshire","CS")
		call addstate(201,"Berwickshire","BS")
		call addstate(201,"Anglesey","AN")
		call addstate(201,"Pembrokeshire","PK")
		call addstate(201,"Flintshire","FS")
		call addstate(201,"Rutland","RD")
		call addstate(201,"Glamorgan","AA")

		call addstate(201,"Cardiff","AA")
		call addstate(201,"Bristol","AA")
		call addstate(201,"Manchester","AA")
		call addstate(201,"Birmingham","AA")
		call addstate(201,"Glasgow","AA")
		call addstate(201,"Edinburgh","AA")
		
		call adddisabledstate(201,"BFPO","FO")
		call adddisabledstate(201,"APO/FPO","AO")
	else
		cnn.execute("UPDATE states SET stateCountryID=201 WHERE stateCountryID=0")
	end if
	' Denmark
	statelist="'Bornholm','Falster','Fyn','Jylland','Sjaelland'"
	rs.open "SELECT COUNT(*) AS numstates FROM states WHERE stateName IN (" & statelist & ")",cnn,0,1
	if isnull(rs("numstates")) then numstates=0 else numstates=cint(rs("numstates"))
	rs.close
	if numstates<5 then
		call addstate(50,"Bornholm","BH")
		call addstate(50,"Falster","FA")
		call addstate(50,"Fyn","FY")
		call addstate(50,"Jylland","JY")
		call addstate(50,"Sjaelland","SJ")
	else
		cnn.execute("UPDATE states SET stateCountryID=50 WHERE stateCountryID=0")
	end if
	' France
	statelist="'Ain','Aisne','Allier','Ardennes','Averyon'"
	rs.open "SELECT COUNT(*) AS numstates FROM states WHERE stateName IN (" & statelist & ")",cnn,0,1
	if isnull(rs("numstates")) then numstates=0 else numstates=cint(rs("numstates"))
	rs.close
	if numstates<5 then
		call addstate(65,"Ain","01")
		call addstate(65,"Aisne","02")
		call addstate(65,"Allier","03")
		call addstate(65,"Alpes de Haute Provence","04")
		call addstate(65,"Hautes Alpes","05")
		call addstate(65,"Alpes Maritimes","06")
		call addstate(65,"Ard&egrave;che","07")
		call addstate(65,"Ardennes","08")
		call addstate(65,"Ari&egrave;ge","09")
		call addstate(65,"Aube","10")
		call addstate(65,"Aude","11")
		call addstate(65,"Averyon","12")
		call addstate(65,"Bouche du Rh&ocirc;ne","13")
		call addstate(65,"Calvados","14")
		call addstate(65,"Cantal","15")
		call addstate(65,"Charente","16")
		call addstate(65,"Charente Maritime","17")
		call addstate(65,"Cher","18")
		call addstate(65,"Corr&egrave;ze","19")
		call addstate(65,"Corse du Sud","2a")
		call addstate(65,"Haute Corse","2b")
		call addstate(65,"C&ocirc;te d'Or","21")
		call addstate(65,"C&ocirc;tes d'Armor","22")
		call addstate(65,"Creuse","23")
		call addstate(65,"Dordogne","24")
		call addstate(65,"Doubs","25")
		call addstate(65,"Dr&ocirc;me","26")
		call addstate(65,"Eure","27")
		call addstate(65,"Eure et Loire","28")
		call addstate(65,"Finist&egrave;re","29")
		call addstate(65,"Gard","30")
		call addstate(65,"Haute Garonne","31")
		call addstate(65,"Gers","32")
		call addstate(65,"Gironde","33")
		call addstate(65,"Herault","34")
		call addstate(65,"Ille et Vilaine","35")
		call addstate(65,"Indre","36")
		call addstate(65,"Indre et Loire","37")
		call addstate(65,"Is&egrave;re","38")
		call addstate(65,"Jura","39")
		call addstate(65,"Landes","40")
		call addstate(65,"Loir et Cher","41")
		call addstate(65,"Loire","42")
		call addstate(65,"Haute Loire","43")
		call addstate(65,"Loire Atlantique","44")
		call addstate(65,"Loiret","45")
		call addstate(65,"Lot","46")
		call addstate(65,"Lot et Garonne","47")
		call addstate(65,"Loz&egrave;re","48")
		call addstate(65,"Maine et Loire","49")
		call addstate(65,"Manche","50")
		call addstate(65,"Marne","51")
		call addstate(65,"Haute Marne","52")
		call addstate(65,"Mayenne","53")
		call addstate(65,"Meurthe et Moselle","54")
		call addstate(65,"Meuse","55")
		call addstate(65,"Morbihan","56")
		call addstate(65,"Moselle","57")
		call addstate(65,"Ni&egrave;vre","58")
		call addstate(65,"Nord","59")
		call addstate(65,"Oise","60")
		call addstate(65,"Orne","61")
		call addstate(65,"Pas de Calais","62")
		call addstate(65,"Puy de D&ocirc;me","63")
		call addstate(65,"Pyren&eacute;es Atlantiques","64")
		call addstate(65,"Haute Pyren&eacute;es","65")
		call addstate(65,"Pyren&eacute;es orientales","66")
		call addstate(65,"Bas Rhin","67")
		call addstate(65,"Haut Rhin","68")
		call addstate(65,"Rh&ocirc;ne","69")
		call addstate(65,"Haute Sa&ocirc;ne","70")
		call addstate(65,"Sa&ocirc;ne et Loire","71")
		call addstate(65,"Sarthe","72")
		call addstate(65,"Savoie","73")
		call addstate(65,"Haute Savoie","74")
		call addstate(65,"Paris","75")
		call addstate(65,"Seine Maritime","76")
		call addstate(65,"Seine et Marne","77")
		call addstate(65,"Yvelines","78")
		call addstate(65,"Deux S&egrave;vres","79")
		call addstate(65,"Somme","80")
		call addstate(65,"Tarn","81")
		call addstate(65,"Tarn et Garonne","82")
		call addstate(65,"Var","83")
		call addstate(65,"Vaucluse","84")
		call addstate(65,"Vend&eacute;e","85")
		call addstate(65,"Vienne","86")
		call addstate(65,"Haute Vienne","87")
		call addstate(65,"Vosges","88")
		call addstate(65,"Yonne","89")
		call addstate(65,"Territoire de Belfort","90")
		call addstate(65,"Essonne","91")
		call addstate(65,"Hauts de Seine","92")
		call addstate(65,"Seine Saint Denis","93")
		call addstate(65,"Val de Marne","94")
		call addstate(65,"Val d'Oise","95")
	else
		cnn.execute("UPDATE states SET stateCountryID=65 WHERE stateCountryID=0")
	end if
	' Germany
	statelist="'Bayern','Berlin','Brandenburg','Bremen','Hamburg'"
	rs.open "SELECT COUNT(*) AS numstates FROM states WHERE stateName IN (" & statelist & ")",cnn,0,1
	if isnull(rs("numstates")) then numstates=0 else numstates=cint(rs("numstates"))
	rs.close
	if numstates<5 then
		call addstate(71,"Baden-W&uuml;rttemberg","01")
		call addstate(71,"Bayern","02")
		call addstate(71,"Berlin","03")
		call addstate(71,"Brandenburg","04")
		call addstate(71,"Bremen","05")
		call addstate(71,"Hamburg","06")
		call addstate(71,"Hessen","07")
		call addstate(71,"Mecklenburg-Vorpommern","08")
		call addstate(71,"Niedersachsen","09")
		call addstate(71,"Nordrhein-Westfalen","10")
		call addstate(71,"Rheinland-Pfalz","11")
		call addstate(71,"Saarland","12")
		call addstate(71,"Sachsen","13")
		call addstate(71,"Sachsen Anhalt","14")
		call addstate(71,"Schleswig Holstein","15")
		call addstate(71,"Th&uuml;ringen","16")
	else
		cnn.execute("UPDATE states SET stateCountryID=71 WHERE stateCountryID=0")
	end if
	' Switzerland
	statelist="'Argovia','Ginevra','Glarona','Grigioni','Lucerna','Aargau','Bern','Luzern','Neuenburg','Nidwalden'"
	rs.open "SELECT COUNT(*) AS numstates FROM states WHERE stateName IN (" & statelist & ")",cnn,0,1
	if isnull(rs("numstates")) then numstates=0 else numstates=cint(rs("numstates"))
	rs.close
	if numstates<5 then
		call addstate(183,"Aargau","AG")
		call addstate(183,"Appenzell Innerrhoden","AI")
		call addstate(183,"Appenzell Ausserrhoden","AR")
		call addstate(183,"Basel-Stadt","BS")
		call addstate(183,"Basel-Landschaft","BL")
		call addstate(183,"Bern","BE")
		call addstate(183,"Freiburg","FR")
		call addstate(183,"Genf","GE")
		call addstate(183,"Glarus","GL")
		call addstate(183,"Graub&uuml;nden","GR")
		call addstate(183,"Jura","JU")
		call addstate(183,"Luzern","LU")
		call addstate(183,"Neuenburg","NE")
		call addstate(183,"Nidwalden","NW")
		call addstate(183,"Obwalden","OW")
		call addstate(183,"Schaffhausen","SH")
		call addstate(183,"Schwyz","SZ")
		call addstate(183,"Solothurn","SO")
		call addstate(183,"St. Gallen","SG")
		call addstate(183,"Thurgau","TG")
		call addstate(183,"Tessin","TI")
		call addstate(183,"Uri","UR")
		call addstate(183,"Wallis","VS")
		call addstate(183,"Waadt","VD")
		call addstate(183,"Zug","ZG")
		call addstate(183,"Z&uuml;rich","ZH")
	else
		cnn.execute("UPDATE states SET stateCountryID=183 WHERE stateCountryID=0")
	end if
	' Italy
	statelist="'Abruzzo','Basilicata','Calabria','Campania','Lombardia'"
	rs.open "SELECT COUNT(*) AS numstates FROM states WHERE stateName IN (" & statelist & ")",cnn,0,1
	if isnull(rs("numstates")) then numstates=0 else numstates=cint(rs("numstates"))
	rs.close
	if numstates<5 then
		call addstate(93,"Abruzzo","AL")
		call addstate(93,"Basilicata","AK")
		call addstate(93,"Calabria","AS")
		call addstate(93,"Campania","AZ")
		call addstate(93,"Emilia Romagna","AR")
		call addstate(93,"Friuli Venezia Giulia","CA")
		call addstate(93,"Lazio","CO")
		call addstate(93,"Liguria","CT")
		call addstate(93,"Lombardia","DE")
		call addstate(93,"Marche","DC")
		call addstate(93,"Piemonte","FM")
		call addstate(93,"Puglia","FL")
		call addstate(93,"Sardegna","GA")
		call addstate(93,"Sicilia","GU")
		call addstate(93,"Toscana","HI")
		call addstate(93,"Trentino Alto Adige","ID")
		call addstate(93,"Umbria","IL")
		call addstate(93,"Valle d'Aosta","IN")
		call addstate(93,"Veneto","IA")
	else
		cnn.execute("UPDATE states SET stateCountryID=93 WHERE stateCountryID=0")
	end if
	' Portugal
	statelist="'Aveiro','Beja','Braganca','Coimbra','Lisboa'"
	rs.open "SELECT COUNT(*) AS numstates FROM states WHERE stateName IN (" & statelist & ")",cnn,0,1
	if isnull(rs("numstates")) then numstates=0 else numstates=cint(rs("numstates"))
	rs.close
	if numstates<5 then
		call addstate(153,"Aveiro","AB")
		call addstate(153,"Beja","AG")
		call addstate(153,"Braga","AR")
		call addstate(153,"Braganca","AV")
		call addstate(153,"Castelo Branco","AY")
		call addstate(153,"Coimbra","BF")
		call addstate(153,"Evora","BE")
		call addstate(153,"Faro","BK")
		call addstate(153,"Guarda","BU")
		call addstate(153,"Leiria","CN")
		call addstate(153,"Lisboa","CB")
		call addstate(153,"Portalegre","CH")
		call addstate(153,"Porto","CL")
		call addstate(153,"Santarem","CV")
		call addstate(153,"Setubal","CW")
		call addstate(153,"Viana do Castelo","CAn")
		call addstate(153,"Vila Real","CL")
		call addstate(153,"Viseu","CL")
		call addstate(153,"Madeira","MA")
		call addstate(153,"A&ccedil;ores","AC")
	else
		cnn.execute("UPDATE states SET stateCountryID=153 WHERE stateCountryID=0")
	end if
	' Spain
	statelist="'Albacete','Alicante','Barcelona','Burgos','Cantabria'"
	rs.open "SELECT COUNT(*) AS numstates FROM states WHERE stateName IN (" & statelist & ")",cnn,0,1
	if isnull(rs("numstates")) then numstates=0 else numstates=cint(rs("numstates"))
	rs.close
	if numstates<5 then
		call addstate(175,"Alava","VI")
		call addstate(175,"Albacete","AB")
		call addstate(175,"Alicante","A")
		call addstate(175,"Almer&iacute;a","AL")
		call addstate(175,"Asturias","O")
		call addstate(175,"Avila","AV")
		call addstate(175,"Badajoz","BA")
		call addstate(175,"Barcelona","B")
		call addstate(175,"Burgos","BU")
		call addstate(175,"C&aacute;ceres","CC")
		call addstate(175,"C&aacute;diz","CA")
		call addstate(175,"Cantabria","S")
		call addstate(175,"Castell&oacute;n","CS")
		call addstate(175,"Ceuta","CE")
		call addstate(175,"Ciudad Real","CR")
		call addstate(175,"C&oacute;rdoba","CO")
		call addstate(175,"Cuenca","CU")
		call addstate(175,"Guip&uacute;zcoa","SS")
		call addstate(175,"Girona","GI")
		call addstate(175,"Granada","GR")
		call addstate(175,"Guadalajara","GU")
		call addstate(175,"Huelva","H")
		call addstate(175,"Huesca","HU")
		call addstate(175,"Islas Baleares","IB")
		call addstate(175,"Ja&eacute;n","J")
		call addstate(175,"La Coru&ntilde;a","C")
		call addstate(175,"La Rioja","LO")
		call addstate(175,"Las Palmas","GC")
		call addstate(175,"Le&oacute;n","LE")
		call addstate(175,"L&eacute;rida","LL")
		call addstate(175,"Lugo","LU")
		call addstate(175,"Madrid","M")
		call addstate(175,"M&aacute;laga","MA")
		call addstate(175,"Melilla","ML")
		call addstate(175,"Murcia","MU")
		call addstate(175,"Navarra","NA")
		call addstate(175,"Orense","OR")
		call addstate(175,"Palencia","P")
		call addstate(175,"Pontevedra","PO")
		call addstate(175,"Salamanca","SA")
		call addstate(175,"Tenerife","TF")
		call addstate(175,"Segovia","SG")
		call addstate(175,"Sevilla","SE")
		call addstate(175,"Soria","SO")
		call addstate(175,"Tarragona","T")
		call addstate(175,"Teruel","TE")
		call addstate(175,"Toledo","TO")
		call addstate(175,"Valencia","V")
		call addstate(175,"Valladolid","VA")
		call addstate(175,"Vizcaya","BI")
		call addstate(175,"Zamora","ZA")
		call addstate(175,"Zaragoza","Z")
	else
		cnn.execute("UPDATE states SET stateCountryID=175 WHERE stateCountryID=0")
	end if
end if ' } updatestates
if checkaddcolumn("states","stateName2",FALSE,txtcl,"(255)","") then
	cnn.execute("UPDATE states SET stateName2=stateName")
end if

if checkaddcolumn("states","stateName3",FALSE,txtcl,"(255)","") then
	cnn.execute("UPDATE states SET stateName3=stateName")
end if

printtickdiv("Checking for DHL Methods upgrade")
sSQL = "SELECT * FROM uspsmethods WHERE uspsID = 501"
rs.Open sSQL,cnn,0,1
if rs.EOF then
	printtick("Adding DHL Shipping Methods info")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (501,'3','DHL Easy Shop',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (502,'4','DHL Jetline',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (503,'8','DHL Express Easy',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (504,'E','DHL Express 9:00',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (505,'F','DHL Freight Worldwide',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (506,'H','DHL Economy Select',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (507,'J','DHL Jumbo Box',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (508,'M','DHL Express 10:30',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (509,'P','DHL Express Worldwide',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (510,'Q','DHL Medical Express',0,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (511,'V','DHL Europack',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (512,'Y','DHL Express 12:00',1,1)")
	' Document methods
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (513,'2','DHL Easy Shop',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (514,'5','DHL Sprintline',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (515,'6','DHL Secureline',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (516,'7','DHL Express Easy',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (517,'9','DHL Europack',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (518,'B','DHL Break Bulk Express',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (519,'C','DHL Medical Express',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (520,'D','DHL Express Worldwide',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (521,'G','DHL Domestic Economy Express',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (522,'I','DHL Break Bulk Economy',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (523,'K','DHL Express 9:00',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (524,'L','DHL Express 10:30',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (525,'N','DHL Domestic Express',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (526,'R','DHL Global Mail Business',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (527,'S','DHL Same Day',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (528,'T','DHL Express 12:00',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (529,'U','DHL Express Worldwide',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (530,'W','DHL Economy Select',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (531,'X','DHL Express Envelope',1,1)")
end if
rs.Close

' New Canada Post Services
cnn.execute("UPDATE uspsmethods SET uspsMethod='DOM.RP',uspsShowAs='Regular Parcel' WHERE uspsID=201") ' 1010','Regular'
cnn.execute("UPDATE uspsmethods SET uspsMethod='DOM.EP',uspsShowAs='Expedited Parcel' WHERE uspsID=202") ' 1020','Expedited'
cnn.execute("UPDATE uspsmethods SET uspsMethod='DOM.XP',uspsShowAs='Xpresspost' WHERE uspsID=203") ' 1030','Xpresspost'
cnn.execute("UPDATE uspsmethods SET uspsMethod='DOM.XP.CERT',uspsShowAs='Xpresspost Certified' WHERE uspsID=204") ' 1040','Priority Courier'
cnn.execute("UPDATE uspsmethods SET uspsMethod='DOM.PC',uspsShowAs='Priority' WHERE uspsID=205") ' 1120','Expedited Evening'
cnn.execute("UPDATE uspsmethods SET uspsMethod='DOM.LIB',uspsShowAs='Library Books' WHERE uspsID=206") ' 1130','XpressPost Evening'
cnn.execute("UPDATE uspsmethods SET uspsMethod='USA.EP',uspsShowAs='Expedited Parcel USA' WHERE uspsID=207") ' 1220','Expedited Saturday'
cnn.execute("UPDATE uspsmethods SET uspsMethod='USA.PW.ENV',uspsShowAs='Priority Worldwide Envelope USA' WHERE uspsID=208") ' 1230','XpressPost Saturday'
cnn.execute("UPDATE uspsmethods SET uspsMethod='USA.PW.PAK',uspsShowAs='Priority Worldwide pak USA' WHERE uspsID=210") ' 2005','Small Packets Surface'
cnn.execute("UPDATE uspsmethods SET uspsMethod='USA.PW.PARCEL',uspsShowAs='Priority Worldwide Parcel USA' WHERE uspsID=211") ' 2010','Surface USA'
cnn.execute("UPDATE uspsmethods SET uspsMethod='USA.SP.AIR',uspsShowAs='Small Packet USA Air' WHERE uspsID=212") ' 2015','Small Packets Air USA'
cnn.execute("UPDATE uspsmethods SET uspsMethod='USA.SP.SURF',uspsShowAs='Small Packet USA Surface' WHERE uspsID=213") ' 2020','Air USA'
cnn.execute("UPDATE uspsmethods SET uspsMethod='USA.XP',uspsShowAs='Xpresspost USA' WHERE uspsID=214") ' 2025','Expedited USA Commercial'
cnn.execute("UPDATE uspsmethods SET uspsMethod='INT.XP',uspsShowAs='Xpresspost International' WHERE uspsID=215") ' 2030','XPressPost USA'
cnn.execute("UPDATE uspsmethods SET uspsMethod='INT.IP.AIR',uspsShowAs='International Parcel Air' WHERE uspsID=216") ' 2040','Purolator USA'
cnn.execute("UPDATE uspsmethods SET uspsMethod='INT.IP.SURF',uspsShowAs='International Parcel Surface' WHERE uspsID=217") ' 2050','PuroPak USA'
cnn.execute("UPDATE uspsmethods SET uspsMethod='INT.PW.ENV',uspsShowAs='Priority Worldwide Envelope Int''l' WHERE uspsID=218") ' 3005','Small Packets Surface International'
cnn.execute("UPDATE uspsmethods SET uspsMethod='INT.PW.PAK',uspsShowAs='Priority Worldwide pak Int''l' WHERE uspsID=221") ' 3010','Parcel Surface International'
cnn.execute("UPDATE uspsmethods SET uspsMethod='INT.PW.PARCEL',uspsShowAs='Priority Worldwide parcel Int''l' WHERE uspsID=222") ' 3015','Small Packets Air International'
cnn.execute("UPDATE uspsmethods SET uspsMethod='INT.SP.AIR',uspsShowAs='Small Packet International Air' WHERE uspsID=223") ' 3020','Air International'
cnn.execute("UPDATE uspsmethods SET uspsMethod='INT.SP.SURF',uspsShowAs='Small Packet International Surface' WHERE uspsID=224") ' 3025','XPressPost International'
'cnn.execute("DELETE FROM uspsmethods WHERE uspsID=225") ' 3040','Purolator International'
on error resume next
cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (225,'INT.TP','Tracked Packet - International',0,0)")
on error goto 0
cnn.execute("UPDATE uspsmethods SET uspsMethod='INT.TP',uspsShowAs='Tracked Packet - International' WHERE uspsID=225") ' 3025','XPressPost International'
cnn.execute("DELETE FROM uspsmethods WHERE uspsID=226") ' 3050','PuroPak International'

on error resume next
printtickdiv("PayPal Advanced")
sSQL = "INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvShow2,payProvShow3,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (22,'PayPal Advanced','Credit Card','Credit Card','Credit Card',0,1,0,'','',22)"
cnn.execute(sSQL)
printtickdiv("Stripe")
sSQL = "INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvShow2,payProvShow3,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (23,'Stripe','Credit Card','Credit Card','Credit Card',0,1,0,'','',23)"
cnn.execute(sSQL)
sSQL = "INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvShow2,payProvShow3,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (24,'Opayo','Credit Card','Credit Card','Credit Card',0,1,0,'','',24)"
cnn.execute(sSQL)
sSQL = "INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvShow2,payProvShow3,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (27,'PayPal Checkout','PayPal Checkout','PayPal Checkout','PayPal Checkout',0,0,0,'','',1)"
cnn.execute(sSQL)
cnn.execute("UPDATE payprovider SET payProvName='PayPal Checkout' WHERE payProvID=27")
cnn.execute("INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvShow2,payProvShow3,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (28,'SquareUp','Credit Card','Credit Card','Credit Card',0,1,0,'','',1)")
cnn.execute("INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvShow2,payProvShow3,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (29,'NMI','Credit Card','Credit Card','Credit Card',0,1,0,'','',1)")
cnn.execute("INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvShow2,payProvShow3,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (30,'eWay','Credit Card','Credit Card','Credit Card',0,1,0,'','',1)")
cnn.execute("INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvShow2,payProvShow3,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (31,'Pay360','Credit Card','Credit Card','Credit Card',0,1,0,'','',1)")
cnn.execute("INSERT INTO payprovider (payProvID,payProvName,payProvShow,payProvShow2,payProvShow3,payProvEnabled,payProvAvailable,payProvDemo,payProvData1,payProvData2,payProvOrder) VALUES (32,'Global Payments','Credit Card','Credit Card','Credit Card',0,1,0,'','',1)")
on error goto 0

call checkaddcolumn("cart","cartOrigProdID",FALSE,txtcl,"(255)","")

on error resume next
if sqlserver OR mysqlserver then
	cnn.execute("ALTER TABLE cartoptions "&altcl&" coCartOption "&txtcl&"("&txtcollen&")")
end if
cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (401,'SMARTPOST','FedEx SmartPost&reg;',1,1)")

call checkaddcolumn("products","pCustom1",FALSE,txtcl,"(255)","")
call checkaddcolumn("products","pCustom2",FALSE,txtcl,"(255)","")
call checkaddcolumn("products","pCustom3",FALSE,txtcl,"(255)","")

call checkaddcolumn("customerlogin","clientAdminNotes",FALSE,memocl,"","")
call checkaddcolumn("customerlogin","clientCustom1",FALSE,txtcl,"(255)","")
call checkaddcolumn("customerlogin","clientCustom2",FALSE,txtcl,"(255)","")

' call checkaddcolumn("mailinglist","errorsend",FALSE,"INT","","")

cnn.execute("UPDATE states SET stateName='Yorkshire' WHERE stateCountryID=201 AND stateName='East Yorkshire'")
cnn.execute("UPDATE states SET stateName='Durham' WHERE stateCountryID=201 AND stateName='County Durham'")
cnn.execute("UPDATE states SET stateName='Moray' WHERE stateCountryID=201 AND stateName='Morayshire'")
cnn.execute("UPDATE states SET stateName='Shetland' WHERE stateCountryID=201 AND stateName='Isle of Shetland'")
cnn.execute("UPDATE states SET stateName='Kirkcudbrightshire' WHERE stateCountryID=201 AND stateName='Kirkudbrightshire'")
rs2.open "SELECT stateID FROM states WHERE stateCountryID=201 AND stateName='Orkney'",cnn,0,1
if rs2.EOF then call addstate(201,"Orkney","ORK")
rs2.close
rs2.open "SELECT stateID FROM states WHERE stateCountryID=201 AND stateName='Denbighshire'",cnn,0,1
if rs2.EOF then call addstate(201,"Denbighshire","DEN")
rs2.close
rs2.open "SELECT stateID FROM states WHERE stateCountryID=201 AND stateName='Monmouthshire'",cnn,0,1
if rs2.EOF then call addstate(201,"Monmouthshire","MON")
rs2.close
rs2.open "SELECT stateID FROM states WHERE stateCountryID=201 AND stateName='Rhondda Cynon Taff'",cnn,0,1
if rs2.EOF then call addstate(201,"Rhondda Cynon Taff","RON")
rs2.close
rs2.open "SELECT stateID FROM states WHERE stateCountryID=201 AND stateName='Channel Islands'",cnn,0,1
if rs2.EOF then call adddisabledstate(201,"Channel Islands","CHI")
rs2.close
rs2.open "SELECT stateID FROM states WHERE stateCountryID=201 AND stateName='Isle of Man'",cnn,0,1
if rs2.EOF then call adddisabledstate(201,"Isle of Man","ISM")
rs2.close

cnn.execute("UPDATE payprovider SET payProvAvailable=0,payProvEnabled=0 WHERE payProvID=20")

call checkaddcolumn("products","pStaticURL",FALSE,txtcl,"(255)","")

call checkaddcolumn("uspsmethods","uspsOrder",TRUE,"INT","","")

on error resume next
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal,uspsFSA) VALUES (601,'AUS_PARCEL_REGULAR','Parcel Post',1,1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (602,'AUS_PARCEL_REGULAR_SATCHEL_LARGE','Parcel Post',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (603,'AUS_PARCEL_EXPRESS','Express Post',1,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (604,'AUS_PARCEL_EXPRESS_SATCHEL_LARGE','Express Post',1,1)")
	cnn.execute("UPDATE uspsmethods SET uspsMethod='AUS_PARCEL_EXPRESS_SATCHEL_LARGE' WHERE uspsMethod='AUS_PARCEL_EXPRESS_SATCHEL_3KG'")
	cnn.execute("DELETE FROM uspsmethods WHERE uspsID IN (607,609,610,613,616)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (617,'AUS_PARCEL_EXPRESS_SATCHEL_SMALL','Express Post',0,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (618,'AUS_PARCEL_EXPRESS_SATCHEL_MEDIUM','Express Post',0,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (619,'AUS_PARCEL_EXPRESS_SATCHEL_EXTRA_LARGE','Express Post',0,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (620,'AUS_PARCEL_REGULAR_PACKAGE_SMALL','Parcel Post',0,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (621,'AUS_PARCEL_REGULAR_PACKAGE_MEDIUM','Parcel Post',0,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (622,'AUS_PARCEL_REGULAR_PACKAGE_LARGE','Parcel Post',0,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (630,'AUS_PARCEL_REGULAR_PACKAGE_EXTRA_LARGE','Parcel Post',0,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (623,'AUS_PARCEL_EXPRESS_PACKAGE_SMALL','Express Post',0,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (624,'AUS_PARCEL_EXPRESS_PACKAGE_MEDIUM','Express Post',0,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (625,'AUS_PARCEL_EXPRESS_PACKAGE_LARGE','Express Post',0,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (629,'AUS_PARCEL_EXPRESS_PACKAGE_EXTRA_LARGE','Express Post',0,1)")

	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (626,'AUS_PARCEL_REGULAR_SATCHEL_SMALL','Parcel Post',0,1)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (627,'AUS_PARCEL_REGULAR_SATCHEL_MEDIUM','Parcel Post',0,1)")
	cnn.execute("UPDATE uspsmethods SET uspsMethod='AUS_PARCEL_REGULAR_SATCHEL_LARGE' WHERE uspsMethod='AUS_PARCEL_REGULAR_SATCHEL_3KG'")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (628,'AUS_PARCEL_REGULAR_SATCHEL_EXTRA_LARGE','Parcel Post',0,1)")

	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (605,'INT_PARCEL_COR_OWN_PACKAGING','International Post Courier',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (606,'INT_PARCEL_EXP_OWN_PACKAGING','International Post Express',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (608,'INT_PARCEL_STD_OWN_PACKAGING','International Post Standard',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (611,'INT_PARCEL_AIR_OWN_PACKAGING','International Post Economy',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (612,'INT_PARCEL_SEA_OWN_PACKAGING','International Post Economy',1,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (614,'INT_LETTER_REG_SMALL','International Economy Letter',0,0)")
	cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (615,'INT_LETTER_REG_LARGE','International Economy Letter',0,0)")
on error goto 0

cnn.execute("UPDATE uspsmethods SET uspsOrder=0 WHERE uspsMethod='AUS_PARCEL_REGULAR'")
cnn.execute("UPDATE uspsmethods SET uspsOrder=1 WHERE uspsMethod='AUS_PARCEL_REGULAR_PACKAGE_SMALL'")
cnn.execute("UPDATE uspsmethods SET uspsOrder=2 WHERE uspsMethod='AUS_PARCEL_REGULAR_PACKAGE_MEDIUM'")
cnn.execute("UPDATE uspsmethods SET uspsOrder=3 WHERE uspsMethod='AUS_PARCEL_REGULAR_PACKAGE_LARGE'")
cnn.execute("UPDATE uspsmethods SET uspsOrder=4 WHERE uspsMethod='AUS_PARCEL_REGULAR_PACKAGE_EXTRA_LARGE'")
cnn.execute("UPDATE uspsmethods SET uspsOrder=5 WHERE uspsMethod='AUS_PARCEL_REGULAR_SATCHEL_SMALL'")
cnn.execute("UPDATE uspsmethods SET uspsOrder=6 WHERE uspsMethod='AUS_PARCEL_REGULAR_SATCHEL_MEDIUM'")
cnn.execute("UPDATE uspsmethods SET uspsOrder=7 WHERE uspsMethod='AUS_PARCEL_REGULAR_SATCHEL_LARGE'")
cnn.execute("UPDATE uspsmethods SET uspsOrder=8 WHERE uspsMethod='AUS_PARCEL_REGULAR_SATCHEL_EXTRA_LARGE'")
cnn.execute("UPDATE uspsmethods SET uspsOrder=9 WHERE uspsMethod='AUS_PARCEL_EXPRESS'")
cnn.execute("UPDATE uspsmethods SET uspsOrder=10 WHERE uspsMethod='AUS_PARCEL_EXPRESS_PACKAGE_SMALL'")
cnn.execute("UPDATE uspsmethods SET uspsOrder=11 WHERE uspsMethod='AUS_PARCEL_EXPRESS_PACKAGE_MEDIUM'")
cnn.execute("UPDATE uspsmethods SET uspsOrder=12 WHERE uspsMethod='AUS_PARCEL_EXPRESS_PACKAGE_LARGE'")
cnn.execute("UPDATE uspsmethods SET uspsOrder=13 WHERE uspsMethod='AUS_PARCEL_EXPRESS_PACKAGE_EXTRA_LARGE'")

cnn.execute("UPDATE uspsmethods SET uspsOrder=14 WHERE uspsMethod='AUS_PARCEL_EXPRESS_SATCHEL_SMALL'")
cnn.execute("UPDATE uspsmethods SET uspsOrder=15 WHERE uspsMethod='AUS_PARCEL_EXPRESS_SATCHEL_MEDIUM'")
cnn.execute("UPDATE uspsmethods SET uspsOrder=16 WHERE uspsMethod='AUS_PARCEL_EXPRESS_SATCHEL_LARGE'")
cnn.execute("UPDATE uspsmethods SET uspsOrder=17 WHERE uspsMethod='AUS_PARCEL_EXPRESS_SATCHEL_EXTRA_LARGE'")

cnn.execute("UPDATE uspsmethods SET uspsMethod='INT_PARCEL_COR_OWN_PACKAGING',uspsShowAs='International Post Courier' WHERE uspsMethod='INTL_SERVICE_ECI_PLATINUM'")
cnn.execute("UPDATE uspsmethods SET uspsMethod='INT_PARCEL_EXP_OWN_PACKAGING',uspsShowAs='International Post Express' WHERE uspsMethod='INTL_SERVICE_ECI_M'")
cnn.execute("UPDATE uspsmethods SET uspsMethod='INT_PARCEL_STD_OWN_PACKAGING',uspsShowAs='International Post Standard' WHERE uspsMethod='INTL_SERVICE_EPI'")
cnn.execute("UPDATE uspsmethods SET uspsMethod='INT_PARCEL_AIR_OWN_PACKAGING',uspsShowAs='International Post Economy' WHERE uspsMethod='INTL_SERVICE_AIR_MAIL'")
cnn.execute("UPDATE uspsmethods SET uspsMethod='INT_PARCEL_SEA_OWN_PACKAGING',uspsShowAs='International Post Economy' WHERE uspsMethod='INTL_SERVICE_SEA_MAIL'")
cnn.execute("UPDATE uspsmethods SET uspsMethod='INT_LETTER_REG_SMALL',uspsShowAs='International Economy Letter' WHERE uspsMethod='INTL_SERVICE_RPI_DLE'")
cnn.execute("UPDATE uspsmethods SET uspsMethod='INT_LETTER_REG_LARGE',uspsShowAs='International Economy Letter' WHERE uspsMethod='INTL_SERVICE_RPI_B4'")

on error resume next
printtickdiv("Checking for multisearchcriteria table")
err.number = 0
sSQL = "SELECT * FROM multisearchcriteria"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding multisearchcriteria table")
	cnn.execute("CREATE TABLE multisearchcriteria (mSCpID "&txtcl&"(128) NOT NULL,mSCscID INT DEFAULT 0 NOT NULL, PRIMARY KEY(mSCpID,mSCscID))")
	on error resume next
	sSQL = "SELECT pSearchCriteria FROM products WHERE pID='xyxyx'"
	rs.open sSQL,cnn,0,1
	errnum=err.number
	rs.close
	on error goto 0
	if errnum=0 then
		rs.open "SELECT pID,pSearchCriteria FROM products WHERE pSearchCriteria<>0"
		do while NOT rs.EOF
			cnn.execute("INSERT INTO multisearchcriteria (mSCpID,mSCscID) VALUES ('"&escape_string(rs("pID"))&"',"&rs("pSearchCriteria")&")")
			rs.movenext
		loop
		rs.close
		on error resume next
		call drop_constraints("products","pSearchCriteria")
		if sqlserver then
			cnn.execute("DROP INDEX products.pSearchCriteria")
		elseif mysqlserver then
			cnn.execute("ALTER TABLE products DROP INDEX pSearchCriteria")
		else
			cnn.execute("DROP INDEX pSearchCriteria ON products")
		end if
		on error goto 0
		cnn.execute("ALTER TABLE products DROP COLUMN pSearchCriteria")
	end if
end if

function checktableexists(tablename)
	on error resume next
	printtickdiv("Checking for "&tablename&" table")
	err.number = 0
	sSQL="SELECT " & IIfVr(mysqlserver<>TRUE,"TOP 1 ","") & " * FROM "&tablename& IIfVr(mysqlserver=TRUE," LIMIT 0,1","")
	rs.open sSQL,cnn,0,1
	checktableexists=err.number=0
	rs.close
	on error goto 0
end function

if NOT checktableexists("searchcriteriagroup") then
	printtick("Adding searchcriteriagroup table")
	if manufacturerfield="" then manufacturerfield="Manufacturer"
	cnn.execute("CREATE TABLE searchcriteriagroup (scgID INT PRIMARY KEY,scgOrder INT DEFAULT 0,scgTitle "&txtcl&"(128) NOT NULL,scgTitle2 "&txtcl&"(128),scgTitle3 "&txtcl&"(128),scgWorkingName "&txtcl&"(128))")
	cnn.execute("INSERT INTO searchcriteriagroup (scgID,scgTitle,scgWorkingName) VALUES (0,'"&escape_string(manufacturerfield)&"','"&escape_string(manufacturerfield)&"')")
	rs.open "SELECT DISTINCT sc2.scID,sc2.scName,sc2.scWorkingName FROM searchcriteria sc1 INNER JOIN searchcriteria sc2 ON sc1.scGroup=sc2.scID",cnn,0,1
	do while NOT rs.EOF
		cnn.execute("INSERT INTO searchcriteriagroup (scgID,scgTitle,scgWorkingName) VALUES ("&rs("scID")&",'"&escape_string(rs("scName"))&"','"&escape_string(rs("scWorkingName"))&"')")
		rs.movenext
	loop
	rs.close
	cnn.execute("UPDATE searchcriteriagroup SET scgTitle2=scgTitle,scgTitle3=scgTitle")
	if NOT newsearchcriteriatable then
		cnn.execute("CREATE TABLE searchcriteria2 "&searchcriteriatable)
		rs.open "SELECT scID,scOrder,scGroup,scWorkingName,scName,scName2,scName3 FROM searchcriteria",cnn,0,1
		do while NOT rs.EOF
			sSQL="INSERT INTO searchcriteria2 (scID,scOrder,scGroup,scWorkingName,scName,scName2,scName3) VALUES ("&rs("scID")&","&rs("scOrder")&","&rs("scGroup")&",'"&escape_string(rs("scWorkingName"))&"','"&escape_string(rs("scName"))&"','"&escape_string(rs("scName2"))&"','"&escape_string(rs("scName3"))&"')"
			cnn.execute(sSQL)
			rs.movenext
		loop
		rs.close
		cnn.execute("DROP TABLE searchcriteria")
		cnn.execute("CREATE TABLE searchcriteria "&searchcriteriatable)
		rs.open "SELECT scID,scOrder,scGroup,scWorkingName,scName,scName2,scName3 FROM searchcriteria2",cnn,0,1
		do while NOT rs.EOF
			sSQL="INSERT INTO searchcriteria (scID,scOrder,scGroup,scWorkingName,scName,scName2,scName3) VALUES ("&rs("scID")&","&rs("scOrder")&","&rs("scGroup")&",'"&escape_string(rs("scWorkingName"))&"','"&escape_string(rs("scName"))&"','"&escape_string(rs("scName2"))&"','"&escape_string(rs("scName3"))&"')"
			cnn.execute(sSQL)
			rs.movenext
		loop
		rs.close
		cnn.execute("DROP TABLE searchcriteria2")
	end if
	if checktableexists("manufacturer") then
		rs.open "SELECT mfID,mfName,mfOrder,mfLogo,mfURL,mfURL2,mfURL3,mfEmail,mfAddress,mfCity,mfState,mfZip,mfCountry FROM manufacturer ORDER BY mfOrder",cnn,0,1
		uniqueindex=1
		do while NOT rs.EOF
			rs2.open "SELECT scID FROM searchcriteria WHERE scid="&rs("mfID"),cnn,0,1
			if NOT rs2.EOF then
				haveuniqueindex=FALSE
				do while NOT haveuniqueindex
					rs3.open "SELECT scID FROM searchcriteria WHERE scID="&uniqueindex,cnn,0,1
					if rs3.EOF then haveuniqueindex=TRUE else uniqueindex=uniqueindex+1
					rs3.close
				loop
				cnn.execute("UPDATE searchcriteria SET scID="&uniqueindex&" WHERE scid="&rs("mfID"))
				on error resume next
				cnn.execute("UPDATE multisearchcriteria SET mSCscID="&uniqueindex&" WHERE mSCscID="&rs("mfID"))
				on error goto 0
			end if
			rs2.close
			taddress=""
			if trim(rs("mfAddress"))<>"" then taddress=taddress&trim(rs("mfAddress"))&vbCrLf
			if trim(rs("mfCity"))<>"" then taddress=taddress&trim(rs("mfCity"))&vbCrLf
			if trim(rs("mfState"))<>"" then taddress=taddress&trim(rs("mfState"))&vbCrLf
			if trim(rs("mfZip"))<>"" then taddress=taddress&trim(rs("mfZip"))&vbCrLf
			if trim(rs("mfCountry"))<>"" then taddress=taddress&trim(rs("mfCountry"))&vbCrLf
			if trim(rs("mfEmail"))<>"" then taddress=taddress&trim(rs("mfEmail"))&vbCrLf
			cnn.execute("INSERT INTO searchcriteria (scID,scGroup,scOrder,scWorkingName,scName,scName2,scName3,scLogo,scURL,scURL2,scURL3,scNotes) VALUES ("&rs("mfID")&",0,"&rs("mfOrder")&",'"&escape_string(rs("mfName"))&"','"&escape_string(rs("mfName"))&"','"&escape_string(rs("mfName"))&"','"&escape_string(rs("mfName"))&"','"&escape_string(rs("mfLogo"))&"','"&escape_string(rs("mfURL"))&"','"&escape_string(rs("mfURL2"))&"','"&escape_string(rs("mfURL3"))&"','"&escape_string(taddress)&"')")
			rs2.open "SELECT mfDescription FROM manufacturer WHERE mfID="&rs("mfID"),cnn,0,1
			if NOT rs2.EOF then cnn.execute("UPDATE searchcriteria SET scDescription='"&escape_string(rs2("mfDescription"))&"' WHERE scID="&rs("mfID"))
			rs2.close
			rs2.open "SELECT mfDescription2 FROM manufacturer WHERE mfID="&rs("mfID"),cnn,0,1
			if NOT rs2.EOF then cnn.execute("UPDATE searchcriteria SET scDescription2='"&escape_string(rs2("mfDescription2"))&"' WHERE scID="&rs("mfID"))
			rs2.close
			rs2.open "SELECT mfDescription3 FROM manufacturer WHERE mfID="&rs("mfID"),cnn,0,1
			if NOT rs2.EOF then cnn.execute("UPDATE searchcriteria SET scDescription3='"&escape_string(rs2("mfDescription3"))&"' WHERE scID="&rs("mfID"))
			rs2.close
			rs2.open "SELECT pID FROM products WHERE pManufacturer="&rs("mfID")
			do while NOT rs2.EOF
				on error resume next
				cnn.execute("INSERT INTO multisearchcriteria (mSCpID,mSCscID) VALUES ('"&escape_string(rs2("pID"))&"',"&rs("mfID")&")")
				on error goto 0
				rs2.movenext
			loop
			rs2.close
			rs.movenext
		loop
		rs.close
		cnn.execute("DROP TABLE manufacturer")
	end if
	on error resume next
	rs.open "SELECT prodFilter FROM admin WHERE adminID=1",cnn,0,1
	prodFilter=rs("prodFilter")
	rs.close
	if (prodFilter AND 1)=1 then cnn.execute("UPDATE admin SET prodFilter="&(prodFilter OR 2))
	cnn.execute("UPDATE sections SET sectionName='' WHERE sectionName IS NULL")
	cnn.execute("UPDATE sections SET sectionName2='' WHERE sectionName2 IS NULL")
	cnn.execute("UPDATE sections SET sectionName3='' WHERE sectionName3 IS NULL")
	cnn.execute("UPDATE sections SET sectionurl='' WHERE sectionurl IS NULL")
	cnn.execute("UPDATE sections SET sectionurl2='' WHERE sectionurl2 IS NULL")
	cnn.execute("UPDATE sections SET sectionurl3='' WHERE sectionurl3 IS NULL")
	sectionstablestructure="(sectionID INT PRIMARY KEY,sectionName "&txtcl&"(255) NOT NULL,sectionName2 "&txtcl&"(255) NOT NULL,sectionName3 "&txtcl&"(255) NOT NULL,sectionWorkingName "&txtcl&"(255),sectionImage "&txtcl&"(255),topSection INT DEFAULT 0,sectionOrder INT DEFAULT 0,rootSection "&bytecl&" DEFAULT 0,sectionDisabled "&bytecl&" DEFAULT 0,sectionurl "&txtcl&"(255) NOT NULL,sectionurl2 "&txtcl&"(255) NOT NULL,sectionurl3 "&txtcl&"(255) NOT NULL,sectionHeader "&memocl&" NULL,sectionHeader2 "&memocl&" NULL,sectionHeader3 "&memocl&" NULL,sectionDescription "&memocl&" NULL,sectionDescription2 "&memocl&" NULL,sectionDescription3 "&memocl&" NULL)"
	sectioncolumns="sectionID,sectionName,sectionName2,sectionName3,sectionWorkingName,sectionImage,topSection,sectionOrder,rootSection,sectionDisabled,sectionurl,sectionurl2,sectionurl3,sectionHeader,sectionHeader2,sectionHeader3,sectionDescription,sectionDescription2,sectionDescription3"
	on error goto 0
	if mysqlserver=TRUE then
		cnn.execute("ALTER TABLE sections MODIFY sectionID INT")
		cnn.execute("ALTER TABLE sections MODIFY sectionName VARCHAR(255) NOT NULL")
		cnn.execute("ALTER TABLE sections MODIFY sectionName2 VARCHAR(255) NOT NULL")
		cnn.execute("ALTER TABLE sections MODIFY sectionName3 VARCHAR(255) NOT NULL")
	else
		cnn.execute("CREATE TABLE sections2 " & sectionstablestructure)
		cnn.execute("INSERT INTO sections2 ("&sectioncolumns&") SELECT "&sectioncolumns&" FROM sections")
		cnn.execute("DROP TABLE sections")
		cnn.execute("CREATE TABLE sections " & sectionstablestructure)
		cnn.execute("INSERT INTO sections ("&sectioncolumns&") SELECT "&sectioncolumns&" FROM sections2")
		cnn.execute("DROP TABLE sections2")
	end if
end if

call checkaddcolumn("admin","prodFilterOrder",FALSE,txtcl,"(255)","")

call checkaddcolumn("admin","sideFilter",FALSE,"INT","","")
call checkaddcolumn("admin","sideFilterText",FALSE,txtcl,"(255)","")
call checkaddcolumn("admin","sideFilterText2",FALSE,txtcl,"(255)","")
call checkaddcolumn("admin","sideFilterText3",FALSE,txtcl,"(255)","")
if checkaddcolumn("admin","sideFilterOrder",FALSE,txtcl,"(255)","") then
	cnn.execute("UPDATE admin SET sideFilter=127,sideFilterText='&Attributes&Price&Sort Order&Per Page&Filter By',sideFilterText2='&Attributes&Price&Sort Order&Per Page&Filter By',sideFilterText3='&Attributes&Price&Sort Order&Per Page&Filter By'")
end if
call checkaddcolumn("options","optDependants",FALSE,txtcl,"(255)","")
call checkaddcolumn("products","pTitle",FALSE,txtcl,"(255)","")
call checkaddcolumn("products","pMetaDesc",FALSE,txtcl,"(255)","")
call checkaddcolumn("sections","sTitle",FALSE,txtcl,"(255)","")
call checkaddcolumn("sections","sMetaDesc",FALSE,txtcl,"(255)","")

call checkaddcolumn("sections","sectionurl",TRUE,txtcl,"(255)","")
call checkaddcolumn("sections","sectionurl2",TRUE,txtcl,"(255)","")
call checkaddcolumn("sections","sectionurl3",TRUE,txtcl,"(255)","")

rs.open "SELECT stateID FROM states WHERE stateCountryID=2 AND stateTax=9.5 AND stateAbbrev='QC'",cnn,0,1
if NOT rs.EOF then
	printtickdiv("Updating Quebec Tax Rate to 9.975%")
	cnn.execute("UPDATE states SET stateTax=9.975 WHERE stateCountryID=2 AND stateTax=9.5 AND stateAbbrev='QC'")
end if
rs.close

call checkaddcolumn("optiongroup","optTooltip",FALSE,memocl,"","")

on error resume next
printtickdiv("Checking for productpackages table")
err.number = 0
sSQL = "SELECT pID FROM productpackages WHERE pID='xyxyx'"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding productpackages table")
	cnn.execute("CREATE TABLE productpackages (packageID "&txtcl&"(128) NOT NULL,pID "&txtcl&"(128) NOT NULL,quantity INT NOT NULL DEFAULT 0,PRIMARY KEY(packageID,pID))")
	cnn.execute("CREATE INDEX packageID_Indx ON productpackages(packageID)")
	cnn.execute("CREATE INDEX pID_Indx ON productpackages(pID)")
end if

call checkaddcolumn("options","optPlaceholder",FALSE,txtcl,"(255)","")
call checkaddcolumn("options","optPlaceholder2",FALSE,txtcl,"(255)","")
call checkaddcolumn("options","optPlaceholder3",FALSE,txtcl,"(255)","")

cnn.execute("UPDATE payprovider SET payProvName='Amazon Pay',payProvShow='Amazon Pay',payProvShow2='Amazon Pay',payProvShow3='Amazon Pay' WHERE payProvID=21 AND payProvName='Amazon Simple Pay'")
cnn.execute("UPDATE payprovider SET payProvName='Opayo' WHERE payProvID=24")

on error resume next
cnn.execute("ALTER TABLE admin "&altcl&" adminEmail "&txtcl&"(255) NULL")

cnn.execute("UPDATE sections SET sectionurl='' WHERE sectionurl IS NULL")
cnn.execute("UPDATE sections SET sectionurl2='' WHERE sectionurl2 IS NULL")
cnn.execute("UPDATE sections SET sectionurl3='' WHERE sectionurl3 IS NULL")
cnn.execute("UPDATE products SET pSKU='' WHERE pSKU IS NULL")

cnn.execute("ALTER TABLE sections "&altcl&" sectionurl "&txtcl&"(255) NOT NULL")
cnn.execute("ALTER TABLE sections "&altcl&" sectionurl2 "&txtcl&"(255) NOT NULL")
cnn.execute("ALTER TABLE sections "&altcl&" sectionurl3 "&txtcl&"(255) NOT NULL")
on error goto 0

call checkaddcolumn("searchcriteria","scHeader",FALSE,memocl,"","")
call checkaddcolumn("searchcriteria","scHeader2",FALSE,memocl,"","")
call checkaddcolumn("searchcriteria","scHeader3",FALSE,memocl,"","")

call checkaddcolumn("countries","currSymbolText",FALSE,txtcl,"(255)","")
call checkaddcolumn("countries","currDecimalSep",FALSE,txtcl,"(255)","")
call checkaddcolumn("countries","currThousandsSep",FALSE,txtcl,"(255)","")
call checkaddcolumn("countries","currPostAmount",FALSE,bitfield,"","")
call checkaddcolumn("countries","currDecimals",FALSE,"INT","","")
if checkaddcolumn("countries","currSymbolHTML",FALSE,txtcl,"(255)","") then
	cnn.execute("UPDATE countries SET currSymbolHTML=countryCurrency,currSymbolText=countryCurrency,currDecimalSep=',',currThousandsSep='.',currDecimals=2,currPostAmount=0")
	cnn.execute("UPDATE countries SET currDecimalSep='.',currThousandsSep=',' WHERE countryCode IN ('AU','BD','BW','BN','KH','CA','CN','HK','MO','DO','EG','SV','GH','GT','HN','IN','IE','IL','JP','JO','KE','KP','KR','LB','LU','MY','MT','MX','MN','MM','NP','NZ','NI','NG','PK','PA','PH','SG','LK','CH','TW','TZ','TH','UG','GB','US','ZW')")
	cnn.execute("UPDATE countries SET currDecimals=0 WHERE countryCurrency IN ('JPY','TWD')")
	cnn.execute("UPDATE countries SET currSymbolHTML='&pound;' WHERE countryCurrency='GBP'")
	cnn.execute("UPDATE countries SET currSymbolHTML='&euro;',currPostAmount=1 WHERE countryCurrency='EUR'")
	cnn.execute("UPDATE countries SET currSymbolHTML='&yen;' WHERE countryCurrency='JPY'")
	cnn.execute("UPDATE countries SET currSymbolHTML='R$',currSymbolText='R$' WHERE countryCurrency='BRL'")
	cnn.execute("UPDATE countries SET currSymbolHTML='AU$ ',currSymbolText='AU$ ' WHERE countryCurrency='AUD'")
	cnn.execute("UPDATE countries SET currSymbolHTML='CDN$ ',currSymbolText='CDN$ ' WHERE countryCurrency='CAD'")
	cnn.execute("UPDATE countries SET currSymbolHTML='HK$',currSymbolText='HK$' WHERE countryCurrency='HKD'")
	cnn.execute("UPDATE countries SET currSymbolHTML='NZ$',currSymbolText='NZ$' WHERE countryCurrency='NZD'")
	cnn.execute("UPDATE countries SET currSymbolHTML='$',currSymbolText='$' WHERE countryCurrency IN ('USD')")
	if overridecurrency then
		rs.open "SELECT adminCountry FROM admin WHERE adminID=1",cnn,0,1
		countryID=rs("adminCountry")
		rs.close
		cnn.execute("UPDATE countries SET currSymbolHTML='" & escape_string(orcsymbol) & "',currSymbolText='" & escape_string(orcemailsymbol) & "',currDecimalSep='" & escape_string(IIfVr(orcdecimals<>"",orcdecimals,".")) & "',currThousandsSep='" & escape_string(orcthousands) & "',currDecimals='" & escape_string(IIfVr(isnumeric(orcdecplaces) AND trim(orcdecplaces)<>"",orcdecplaces,0)) & "',currPostAmount=" & IIfVr(orcpreamount,0,1) & " WHERE countryID=" & countryID)
	end if
end if

if checkaddcolumn("alternaterates","altrateorderintl",FALSE,"INT","","") then
	cnn.execute("UPDATE alternaterates SET altrateorderintl=altrateorder")
	rs.open "SELECT adminShipping,adminIntShipping FROM admin WHERE adminID=1",cnn,0,1
	cnn.execute("UPDATE alternaterates SET usealtmethod=1 WHERE altrateid="&rs("adminShipping"))
	cnn.execute("UPDATE alternaterates SET usealtmethodintl=1 WHERE altrateid="&rs("adminIntShipping"))
	rs.close
end if

printtickdiv("Checking for ajaxfloodcontrol table")
if NOT checktableexists("ajaxfloodcontrol") then
	printtick("Adding ajaxfloodcontrol table")
	cnn.execute("CREATE TABLE ajaxfloodcontrol (afcID "&autoinc&",afcAction INT NOT NULL DEFAULT 0,afcIP "&txtcl&"(128),afcSession "&txtcl&"(128),afcDate "&datecl&")")
	cnn.execute("CREATE INDEX afcAction_Indx ON ajaxfloodcontrol(afcAction)")
	cnn.execute("CREATE INDEX afcIP_Indx ON ajaxfloodcontrol(afcIP)")
	cnn.execute("CREATE INDEX afcSession_Indx ON ajaxfloodcontrol(afcSession)")
	cnn.execute("CREATE INDEX afcDate_Indx ON ajaxfloodcontrol(afcDate)")
end if

call checkaddcolumn("admin","reCAPTCHAsitekey",FALSE,txtcl,"(255)","")
call checkaddcolumn("admin","reCAPTCHAsecret",FALSE,txtcl,"(255)","")
call checkaddcolumn("admin","reCAPTCHAuseon",FALSE,"INT","","")

if checkaddcolumn("admin","adminStoreURLSSL",FALSE,txtcl,"(255)","") then
	storeurlssl=pathtossl
	if storeurlssl="" then
		sSQL="SELECT adminStoreURL FROM admin WHERE adminID=1"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then storeurl=rs("adminStoreURL")
		rs.close
		requiressl=FALSE
		sSQL="SELECT payProvID FROM payprovider WHERE payProvEnabled=1 AND (payProvID IN (7,10,12,13" & IIfVs(NOT paypalhostedsolution,",18") & ") OR (payProvID=16 AND payProvData2='1'))" ' All the ones that require SSL
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then storeurlssl=replace(storeurl,"http:","https:")
		rs.close
	end if
	cnn.execute("UPDATE admin SET adminStoreURLSSL='" & escape_string(storeurlssl) & "' WHERE adminID=1")
end if

call checkaddcolumn("admin","packingslipuseinvoice",FALSE,bytecl,"","")
call checkaddcolumn("emailmessages","invoiceheader",FALSE,memocl,"","")
call checkaddcolumn("emailmessages","invoiceaddress",FALSE,memocl,"","")
call checkaddcolumn("emailmessages","invoicefooter",FALSE,memocl,"","")
call checkaddcolumn("emailmessages","packingslipheader",FALSE,memocl,"","")
call checkaddcolumn("emailmessages","packingslipaddress",FALSE,memocl,"","")
if checkaddcolumn("emailmessages","packingslipfooter",FALSE,memocl,"","") then
	cnn.execute("UPDATE emailmessages SET invoiceheader='" & escape_string(invoiceheader) & "'")
	cnn.execute("UPDATE emailmessages SET invoiceaddress='" & escape_string(invoiceaddress) & "'")
	cnn.execute("UPDATE emailmessages SET invoicefooter='" & escape_string(invoicefooter) & "'")
	if packingslipheader<>"" then
		cnn.execute("UPDATE emailmessages SET packingslipheader='" & escape_string(packingslipheader) & "'")
		cnn.execute("UPDATE emailmessages SET packingslipaddress='" & escape_string(packingslipaddress) & "'")
		cnn.execute("UPDATE emailmessages SET packingslipfooter='" & escape_string(packingslipfooter) & "'")
		cnn.execute("UPDATE admin SET packingslipuseinvoice=0")
	else
		cnn.execute("UPDATE admin SET packingslipuseinvoice=1")
	end if
end if

call checkaddcolumn("admin","vacationmessage",FALSE,memocl,"","")
call checkaddcolumn("admin","onvacation",FALSE,bytecl,"","")
call checkaddcolumn("products","pMinQuant",FALSE,"INT","","")

on error resume next
printtickdiv("Checking for Abandoned Cart Email upgrade")
err.number = 0
sSQL = "SELECT aceID FROM abandonedcartemail WHERE aceID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Order Abandoned Cart Email table")
	cnn.execute("CREATE TABLE abandonedcartemail (aceID "&autoinc&",aceOrderID INT NOT NULL DEFAULT 0,aceDateSent "&datecl&",aceKey "&txtcl&"(255) NULL)")
end if
response.flush

call checkaddcolumn("emailmessages","abandonedcartsubject2",FALSE,txtcl,"(255)","")
call checkaddcolumn("emailmessages","abandonedcartsubject3",FALSE,txtcl,"(255)","")
if checkaddcolumn("emailmessages","abandonedcartsubject",FALSE,txtcl,"(255)","") then
	cnn.execute("UPDATE emailmessages SET abandonedcartsubject='" & escape_string("Trouble purchasing from our store?") & "'")
	cnn.execute("UPDATE emailmessages SET abandonedcartsubject2=abandonedcartsubject,abandonedcartsubject3=abandonedcartsubject")
end if

call checkaddcolumn("emailmessages","abandonedcartemail2",FALSE,memocl,"","")
call checkaddcolumn("emailmessages","abandonedcartemail3",FALSE,memocl,"","")
if checkaddcolumn("emailmessages","abandonedcartemail",FALSE,memocl,"","") then
	cnn.execute("UPDATE emailmessages SET abandonedcartemail='" & escape_string("You recently started, but did not complete an order at our store. If we can help in any way or if you are having trouble with your purchase then please let us know. If you would like to continue with your order you can do so by clicking the link below.%nl%%abandonedcartid%") & "'")
	cnn.execute("UPDATE emailmessages SET abandonedcartemail2=abandonedcartemail,abandonedcartemail3=abandonedcartemail")
end if

if FALSE then
	on error resume next
	sSQL="SELECT pID,pSection FROM products"
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		sSQL="INSERT INTO multisections (pID,pSection) VALUES ('"&escape_string(rs("pID"))&"',"&rs("pSection")&")"
		cnn.execute(sSQL)
		rs.movenext
	loop
	rs.close
	on error goto 0

	call checkaddcolumn("searchcriteria","scGroupTitle",TRUE,txtcl,"(255)","")
	call checkaddcolumn("searchcriteria","scGroupTitle2",TRUE,txtcl,"(255)","")
	call checkaddcolumn("searchcriteria","scGroupTitle3",TRUE,txtcl,"(255)","")
	if checkaddcolumn("searchcriteria","scGroupOrder",TRUE,"INT","","") then
		sSQL="UPDATE searchcriteria SET scGroupOrder=(SELECT scgOrder FROM searchcriteriagroup WHERE searchcriteriagroup.scgID=searchcriteria.scGroup),scGroupTitle=(SELECT scgTitle FROM searchcriteriagroup WHERE searchcriteriagroup.scgID=searchcriteria.scGroup),scGroupTitle2=(SELECT scgTitle2 FROM searchcriteriagroup WHERE searchcriteriagroup.scgID=searchcriteria.scGroup),scGroupTitle3=(SELECT scgTitle3 FROM searchcriteriagroup WHERE searchcriteriagroup.scgID=searchcriteria.scGroup) WHERE EXISTS (SELECT * FROM searchcriteriagroup WHERE searchcriteriagroup.scgID=searchcriteria.scGroup)"
		cnn.execute(sSQL)
	end if

	if checkaddcolumn("multisearchcriteria","mscDisplay",TRUE,bitfield,"","") then
		sSQL="UPDATE multisearchcriteria SET mscDisplay=(SELECT pDisplay FROM products WHERE products.pID=multisearchcriteria.mSCpID) WHERE EXISTS (SELECT * FROM products WHERE products.pID=multisearchcriteria.mSCpID)"
		cnn.execute(sSQL)
	end if

	on error resume next
	cnn.execute("CREATE INDEX scGroupOrder_Indx ON searchcriteria(scGroupOrder)")
	cnn.execute("CREATE INDEX scGroupTitle_Indx ON searchcriteria(scGroupTitle)")
	cnn.execute("CREATE INDEX scGroupTitle2_Indx ON searchcriteria(scGroupTitle2)")
	cnn.execute("CREATE INDEX scGroupTitle3_Indx ON searchcriteria(scGroupTitle3)")
	cnn.execute("CREATE INDEX mscDisplay_Indx ON multisearchcriteria(mscDisplay)")
	on error goto 0
else
	on error resume next
	cnn.execute("ALTER TABLE searchcriteria DROP COLUMN scGroupTitle")
	cnn.execute("ALTER TABLE searchcriteria DROP COLUMN scGroupTitle2")
	cnn.execute("ALTER TABLE searchcriteria DROP COLUMN scGroupTitle3")
	cnn.execute("ALTER TABLE searchcriteria DROP COLUMN scGroupOrder")
	cnn.execute("ALTER TABLE multisearchcriteria DROP COLUMN mscDisplay")
	on error goto 0
end if

call checkaddcolumn("admin","mailchimpAPIKey",TRUE,txtcl,"(255)","")
call checkaddcolumn("admin","mailchimpList",TRUE,txtcl,"(255)","")

on error resume next
cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (17,'First Class Commercial','First Class',0,1)")
cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (18,'Priority Commercial','Priority Mail',0,1)")
cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (19,'Priority Cpp','Priority Mail',0,1)")
cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (20,'Priority Mail Express Commercial','Express Mail',0,1)")
cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (21,'Priority Mail Express CPP','Express Mail',0,1)")
cnn.execute("INSERT INTO uspsmethods (uspsID,uspsMethod,uspsShowAs,uspsUseMethod,uspsLocal) VALUES (22,'Retail Ground','Parcel Post',0,1)")
cnn.execute("UPDATE uspsmethods SET uspsMethod='Priority Mail Express' WHERE uspsID=1")
cnn.execute("UPDATE uspsmethods SET uspsMethod='Parcel Select Ground' WHERE uspsID=3")
on error goto 0

call checkaddcolumn("products","pPopularity",TRUE,"INT","","")
call checkaddcolumn("products","pNumSales",TRUE,"INT","","")

call checkaddcolumn("payprovider","payProvFlag1",TRUE,bitfield,"","")
call checkaddcolumn("payprovider","payProvFlag2",TRUE,bitfield,"","")
call checkaddcolumn("payprovider","payProvFlag3",TRUE,bitfield,"","")
call checkaddcolumn("payprovider","payProvData4",FALSE,txtcl,"(255)","")
call checkaddcolumn("payprovider","payProvData5",FALSE,txtcl,"(255)","")
call checkaddcolumn("payprovider","payProvData6",FALSE,txtcl,"(255)","")
call checkaddcolumn("payprovider","payProvBits",TRUE,"INT","","")

call checkaddcolumn("pricebreaks","pbPercent",TRUE,bitfield,"","")
call checkaddcolumn("pricebreaks","pbWholesalePercent",TRUE,bitfield,"","")

if checkaddcolumn("admin","smtpport",FALSE,txtcl,"(64)","") then cnn.execute("UPDATE admin SET smtpport='" & escape_string(smtpserverport) & "'")
if checkaddcolumn("admin","smtpsecure",FALSE,txtcl,"(64)","") then cnn.execute("UPDATE admin SET smtpsecure='" & escape_string(IIfVs(smtpusessl,"ssl")) & "'")
if checkaddcolumn("admin","emailfromname",FALSE,txtcl,"(128)","") then cnn.execute("UPDATE admin SET emailfromname='" & escape_string(emailfromname) & "'")
if checkaddcolumn("admin","htmlemails",FALSE,bitfield,"","") then cnn.execute("UPDATE admin SET htmlemails=" & IIfVr(htmlemails,"1","0"))

call checkaddcolumn("options","optClass",FALSE,txtcl,"(255)","")

if checkaddcolumn("coupons","cpnStartDate",FALSE,datecl,"","") then
	cnn.execute("UPDATE coupons SET cpnStartDate=" & datedelim & vsusdate(cdate("2000-01-01"))) & datedelim 
end if

call checkaddcolumn("orders","ordUserAgent",FALSE,txtcl,"(255)","")

printtickdiv("Checking for ISO 3166 Column")
if checkaddcolumn("countries","countryCode3",FALSE,txtcl,"(8)","") then
	cnn.execute("UPDATE countries SET countryCode3=''")
	cnn.execute("UPDATE countries SET countryCode3='AND' WHERE countryCode='AD'")
	cnn.execute("UPDATE countries SET countryCode3='ARE' WHERE countryCode='AE'")
	cnn.execute("UPDATE countries SET countryCode3='AFG' WHERE countryCode='AF'")
	cnn.execute("UPDATE countries SET countryCode3='ATG' WHERE countryCode='AG'")
	cnn.execute("UPDATE countries SET countryCode3='AIA' WHERE countryCode='AI'")
	cnn.execute("UPDATE countries SET countryCode3='ALB' WHERE countryCode='AL'")
	cnn.execute("UPDATE countries SET countryCode3='ARM' WHERE countryCode='AM'")
	cnn.execute("UPDATE countries SET countryCode3='AGO' WHERE countryCode='AO'")
	cnn.execute("UPDATE countries SET countryCode3='ATA' WHERE countryCode='AQ'")
	cnn.execute("UPDATE countries SET countryCode3='ARG' WHERE countryCode='AR'")
	cnn.execute("UPDATE countries SET countryCode3='ASM' WHERE countryCode='AS'")
	cnn.execute("UPDATE countries SET countryCode3='AUT' WHERE countryCode='AT'")
	cnn.execute("UPDATE countries SET countryCode3='AUS' WHERE countryCode='AU'")
	cnn.execute("UPDATE countries SET countryCode3='ABW' WHERE countryCode='AW'")
	cnn.execute("UPDATE countries SET countryCode3='ALA' WHERE countryCode='AX'")
	cnn.execute("UPDATE countries SET countryCode3='AZE' WHERE countryCode='AZ'")
	cnn.execute("UPDATE countries SET countryCode3='BIH' WHERE countryCode='BA'")
	cnn.execute("UPDATE countries SET countryCode3='BRB' WHERE countryCode='BB'")
	cnn.execute("UPDATE countries SET countryCode3='BGD' WHERE countryCode='BD'")
	cnn.execute("UPDATE countries SET countryCode3='BEL' WHERE countryCode='BE'")
	cnn.execute("UPDATE countries SET countryCode3='BFA' WHERE countryCode='BF'")
	cnn.execute("UPDATE countries SET countryCode3='BGR' WHERE countryCode='BG'")
	cnn.execute("UPDATE countries SET countryCode3='BHR' WHERE countryCode='BH'")
	cnn.execute("UPDATE countries SET countryCode3='BDI' WHERE countryCode='BI'")
	cnn.execute("UPDATE countries SET countryCode3='BEN' WHERE countryCode='BJ'")
	cnn.execute("UPDATE countries SET countryCode3='BLM' WHERE countryCode='BL'")
	cnn.execute("UPDATE countries SET countryCode3='BMU' WHERE countryCode='BM'")
	cnn.execute("UPDATE countries SET countryCode3='BRN' WHERE countryCode='BN'")
	cnn.execute("UPDATE countries SET countryCode3='BOL' WHERE countryCode='BO'")
	cnn.execute("UPDATE countries SET countryCode3='BES' WHERE countryCode='BQ'")
	cnn.execute("UPDATE countries SET countryCode3='BRA' WHERE countryCode='BR'")
	cnn.execute("UPDATE countries SET countryCode3='BHS' WHERE countryCode='BS'")
	cnn.execute("UPDATE countries SET countryCode3='BTN' WHERE countryCode='BT'")
	cnn.execute("UPDATE countries SET countryCode3='BVT' WHERE countryCode='BV'")
	cnn.execute("UPDATE countries SET countryCode3='BWA' WHERE countryCode='BW'")
	cnn.execute("UPDATE countries SET countryCode3='BLR' WHERE countryCode='BY'")
	cnn.execute("UPDATE countries SET countryCode3='BLZ' WHERE countryCode='BZ'")
	cnn.execute("UPDATE countries SET countryCode3='CAN' WHERE countryCode='CA'")
	cnn.execute("UPDATE countries SET countryCode3='CCK' WHERE countryCode='CC'")
	cnn.execute("UPDATE countries SET countryCode3='COD' WHERE countryCode='CD'")
	cnn.execute("UPDATE countries SET countryCode3='CAF' WHERE countryCode='CF'")
	cnn.execute("UPDATE countries SET countryCode3='COG' WHERE countryCode='CG'")
	cnn.execute("UPDATE countries SET countryCode3='CHE' WHERE countryCode='CH'")
	cnn.execute("UPDATE countries SET countryCode3='CIV' WHERE countryCode='CI'")
	cnn.execute("UPDATE countries SET countryCode3='COK' WHERE countryCode='CK'")
	cnn.execute("UPDATE countries SET countryCode3='CHL' WHERE countryCode='CL'")
	cnn.execute("UPDATE countries SET countryCode3='CMR' WHERE countryCode='CM'")
	cnn.execute("UPDATE countries SET countryCode3='CHN' WHERE countryCode='CN'")
	cnn.execute("UPDATE countries SET countryCode3='COL' WHERE countryCode='CO'")
	cnn.execute("UPDATE countries SET countryCode3='CRI' WHERE countryCode='CR'")
	cnn.execute("UPDATE countries SET countryCode3='CUB' WHERE countryCode='CU'")
	cnn.execute("UPDATE countries SET countryCode3='CPV' WHERE countryCode='CV'")
	cnn.execute("UPDATE countries SET countryCode3='CUW' WHERE countryCode='CW'")
	cnn.execute("UPDATE countries SET countryCode3='CXR' WHERE countryCode='CX'")
	cnn.execute("UPDATE countries SET countryCode3='CYP' WHERE countryCode='CY'")
	cnn.execute("UPDATE countries SET countryCode3='CZE' WHERE countryCode='CZ'")
	cnn.execute("UPDATE countries SET countryCode3='DEU' WHERE countryCode='DE'")
	cnn.execute("UPDATE countries SET countryCode3='DJI' WHERE countryCode='DJ'")
	cnn.execute("UPDATE countries SET countryCode3='DNK' WHERE countryCode='DK'")
	cnn.execute("UPDATE countries SET countryCode3='DMA' WHERE countryCode='DM'")
	cnn.execute("UPDATE countries SET countryCode3='DOM' WHERE countryCode='DO'")
	cnn.execute("UPDATE countries SET countryCode3='DZA' WHERE countryCode='DZ'")
	cnn.execute("UPDATE countries SET countryCode3='ECU' WHERE countryCode='EC'")
	cnn.execute("UPDATE countries SET countryCode3='EST' WHERE countryCode='EE'")
	cnn.execute("UPDATE countries SET countryCode3='EGY' WHERE countryCode='EG'")
	cnn.execute("UPDATE countries SET countryCode3='ESH' WHERE countryCode='EH'")
	cnn.execute("UPDATE countries SET countryCode3='ERI' WHERE countryCode='ER'")
	cnn.execute("UPDATE countries SET countryCode3='ESP' WHERE countryCode='ES'")
	cnn.execute("UPDATE countries SET countryCode3='ETH' WHERE countryCode='ET'")
	cnn.execute("UPDATE countries SET countryCode3='FIN' WHERE countryCode='FI'")
	cnn.execute("UPDATE countries SET countryCode3='FJI' WHERE countryCode='FJ'")
	cnn.execute("UPDATE countries SET countryCode3='FLK' WHERE countryCode='FK'")
	cnn.execute("UPDATE countries SET countryCode3='FSM' WHERE countryCode='FM'")
	cnn.execute("UPDATE countries SET countryCode3='FRO' WHERE countryCode='FO'")
	cnn.execute("UPDATE countries SET countryCode3='FRA' WHERE countryCode='FR'")
	cnn.execute("UPDATE countries SET countryCode3='GAB' WHERE countryCode='GA'")
	cnn.execute("UPDATE countries SET countryCode3='GBR' WHERE countryCode='GB'")
	cnn.execute("UPDATE countries SET countryCode3='GRD' WHERE countryCode='GD'")
	cnn.execute("UPDATE countries SET countryCode3='GEO' WHERE countryCode='GE'")
	cnn.execute("UPDATE countries SET countryCode3='GUF' WHERE countryCode='GF'")
	cnn.execute("UPDATE countries SET countryCode3='GGY' WHERE countryCode='GG'")
	cnn.execute("UPDATE countries SET countryCode3='GHA' WHERE countryCode='GH'")
	cnn.execute("UPDATE countries SET countryCode3='GIB' WHERE countryCode='GI'")
	cnn.execute("UPDATE countries SET countryCode3='GRL' WHERE countryCode='GL'")
	cnn.execute("UPDATE countries SET countryCode3='GMB' WHERE countryCode='GM'")
	cnn.execute("UPDATE countries SET countryCode3='GIN' WHERE countryCode='GN'")
	cnn.execute("UPDATE countries SET countryCode3='GLP' WHERE countryCode='GP'")
	cnn.execute("UPDATE countries SET countryCode3='GNQ' WHERE countryCode='GQ'")
	cnn.execute("UPDATE countries SET countryCode3='GRC' WHERE countryCode='GR'")
	cnn.execute("UPDATE countries SET countryCode3='SGS' WHERE countryCode='GS'")
	cnn.execute("UPDATE countries SET countryCode3='GTM' WHERE countryCode='GT'")
	cnn.execute("UPDATE countries SET countryCode3='GUM' WHERE countryCode='GU'")
	cnn.execute("UPDATE countries SET countryCode3='GNB' WHERE countryCode='GW'")
	cnn.execute("UPDATE countries SET countryCode3='GUY' WHERE countryCode='GY'")
	cnn.execute("UPDATE countries SET countryCode3='HKG' WHERE countryCode='HK'")
	cnn.execute("UPDATE countries SET countryCode3='HMD' WHERE countryCode='HM'")
	cnn.execute("UPDATE countries SET countryCode3='HND' WHERE countryCode='HN'")
	cnn.execute("UPDATE countries SET countryCode3='HRV' WHERE countryCode='HR'")
	cnn.execute("UPDATE countries SET countryCode3='HTI' WHERE countryCode='HT'")
	cnn.execute("UPDATE countries SET countryCode3='HUN' WHERE countryCode='HU'")
	cnn.execute("UPDATE countries SET countryCode3='IDN' WHERE countryCode='ID'")
	cnn.execute("UPDATE countries SET countryCode3='IRL' WHERE countryCode='IE'")
	cnn.execute("UPDATE countries SET countryCode3='ISR' WHERE countryCode='IL'")
	cnn.execute("UPDATE countries SET countryCode3='IMN' WHERE countryCode='IM'")
	cnn.execute("UPDATE countries SET countryCode3='IND' WHERE countryCode='IN'")
	cnn.execute("UPDATE countries SET countryCode3='IOT' WHERE countryCode='IO'")
	cnn.execute("UPDATE countries SET countryCode3='IRQ' WHERE countryCode='IQ'")
	cnn.execute("UPDATE countries SET countryCode3='IRN' WHERE countryCode='IR'")
	cnn.execute("UPDATE countries SET countryCode3='ISL' WHERE countryCode='IS'")
	cnn.execute("UPDATE countries SET countryCode3='ITA' WHERE countryCode='IT'")
	cnn.execute("UPDATE countries SET countryCode3='JEY' WHERE countryCode='JE'")
	cnn.execute("UPDATE countries SET countryCode3='JAM' WHERE countryCode='JM'")
	cnn.execute("UPDATE countries SET countryCode3='JOR' WHERE countryCode='JO'")
	cnn.execute("UPDATE countries SET countryCode3='JPN' WHERE countryCode='JP'")
	cnn.execute("UPDATE countries SET countryCode3='KEN' WHERE countryCode='KE'")
	cnn.execute("UPDATE countries SET countryCode3='KGZ' WHERE countryCode='KG'")
	cnn.execute("UPDATE countries SET countryCode3='KHM' WHERE countryCode='KH'")
	cnn.execute("UPDATE countries SET countryCode3='KIR' WHERE countryCode='KI'")
	cnn.execute("UPDATE countries SET countryCode3='COM' WHERE countryCode='KM'")
	cnn.execute("UPDATE countries SET countryCode3='KNA' WHERE countryCode='KN'")
	cnn.execute("UPDATE countries SET countryCode3='PRK' WHERE countryCode='KP'")
	cnn.execute("UPDATE countries SET countryCode3='KOR' WHERE countryCode='KR'")
	cnn.execute("UPDATE countries SET countryCode3='KWT' WHERE countryCode='KW'")
	cnn.execute("UPDATE countries SET countryCode3='CYM' WHERE countryCode='KY'")
	cnn.execute("UPDATE countries SET countryCode3='KAZ' WHERE countryCode='KZ'")
	cnn.execute("UPDATE countries SET countryCode3='LAO' WHERE countryCode='LA'")
	cnn.execute("UPDATE countries SET countryCode3='LBN' WHERE countryCode='LB'")
	cnn.execute("UPDATE countries SET countryCode3='LCA' WHERE countryCode='LC'")
	cnn.execute("UPDATE countries SET countryCode3='LIE' WHERE countryCode='LI'")
	cnn.execute("UPDATE countries SET countryCode3='LKA' WHERE countryCode='LK'")
	cnn.execute("UPDATE countries SET countryCode3='LBR' WHERE countryCode='LR'")
	cnn.execute("UPDATE countries SET countryCode3='LSO' WHERE countryCode='LS'")
	cnn.execute("UPDATE countries SET countryCode3='LTU' WHERE countryCode='LT'")
	cnn.execute("UPDATE countries SET countryCode3='LUX' WHERE countryCode='LU'")
	cnn.execute("UPDATE countries SET countryCode3='LVA' WHERE countryCode='LV'")
	cnn.execute("UPDATE countries SET countryCode3='LBY' WHERE countryCode='LY'")
	cnn.execute("UPDATE countries SET countryCode3='MAR' WHERE countryCode='MA'")
	cnn.execute("UPDATE countries SET countryCode3='MCO' WHERE countryCode='MC'")
	cnn.execute("UPDATE countries SET countryCode3='MDA' WHERE countryCode='MD'")
	cnn.execute("UPDATE countries SET countryCode3='MNE' WHERE countryCode='ME'")
	cnn.execute("UPDATE countries SET countryCode3='MAF' WHERE countryCode='MF'")
	cnn.execute("UPDATE countries SET countryCode3='MDG' WHERE countryCode='MG'")
	cnn.execute("UPDATE countries SET countryCode3='MHL' WHERE countryCode='MH'")
	cnn.execute("UPDATE countries SET countryCode3='MKD' WHERE countryCode='MK'")
	cnn.execute("UPDATE countries SET countryCode3='MLI' WHERE countryCode='ML'")
	cnn.execute("UPDATE countries SET countryCode3='MMR' WHERE countryCode='MM'")
	cnn.execute("UPDATE countries SET countryCode3='MNG' WHERE countryCode='MN'")
	cnn.execute("UPDATE countries SET countryCode3='MAC' WHERE countryCode='MO'")
	cnn.execute("UPDATE countries SET countryCode3='MNP' WHERE countryCode='MP'")
	cnn.execute("UPDATE countries SET countryCode3='MTQ' WHERE countryCode='MQ'")
	cnn.execute("UPDATE countries SET countryCode3='MRT' WHERE countryCode='MR'")
	cnn.execute("UPDATE countries SET countryCode3='MSR' WHERE countryCode='MS'")
	cnn.execute("UPDATE countries SET countryCode3='MLT' WHERE countryCode='MT'")
	cnn.execute("UPDATE countries SET countryCode3='MUS' WHERE countryCode='MU'")
	cnn.execute("UPDATE countries SET countryCode3='MDV' WHERE countryCode='MV'")
	cnn.execute("UPDATE countries SET countryCode3='MWI' WHERE countryCode='MW'")
	cnn.execute("UPDATE countries SET countryCode3='MEX' WHERE countryCode='MX'")
	cnn.execute("UPDATE countries SET countryCode3='MYS' WHERE countryCode='MY'")
	cnn.execute("UPDATE countries SET countryCode3='MOZ' WHERE countryCode='MZ'")
	cnn.execute("UPDATE countries SET countryCode3='NAM' WHERE countryCode='NA'")
	cnn.execute("UPDATE countries SET countryCode3='NCL' WHERE countryCode='NC'")
	cnn.execute("UPDATE countries SET countryCode3='NER' WHERE countryCode='NE'")
	cnn.execute("UPDATE countries SET countryCode3='NFK' WHERE countryCode='NF'")
	cnn.execute("UPDATE countries SET countryCode3='NGA' WHERE countryCode='NG'")
	cnn.execute("UPDATE countries SET countryCode3='NIC' WHERE countryCode='NI'")
	cnn.execute("UPDATE countries SET countryCode3='NLD' WHERE countryCode='NL'")
	cnn.execute("UPDATE countries SET countryCode3='NOR' WHERE countryCode='NO'")
	cnn.execute("UPDATE countries SET countryCode3='NPL' WHERE countryCode='NP'")
	cnn.execute("UPDATE countries SET countryCode3='NRU' WHERE countryCode='NR'")
	cnn.execute("UPDATE countries SET countryCode3='NIU' WHERE countryCode='NU'")
	cnn.execute("UPDATE countries SET countryCode3='NZL' WHERE countryCode='NZ'")
	cnn.execute("UPDATE countries SET countryCode3='OMN' WHERE countryCode='OM'")
	cnn.execute("UPDATE countries SET countryCode3='PAN' WHERE countryCode='PA'")
	cnn.execute("UPDATE countries SET countryCode3='PER' WHERE countryCode='PE'")
	cnn.execute("UPDATE countries SET countryCode3='PYF' WHERE countryCode='PF'")
	cnn.execute("UPDATE countries SET countryCode3='PNG' WHERE countryCode='PG'")
	cnn.execute("UPDATE countries SET countryCode3='PHL' WHERE countryCode='PH'")
	cnn.execute("UPDATE countries SET countryCode3='PAK' WHERE countryCode='PK'")
	cnn.execute("UPDATE countries SET countryCode3='POL' WHERE countryCode='PL'")
	cnn.execute("UPDATE countries SET countryCode3='SPM' WHERE countryCode='PM'")
	cnn.execute("UPDATE countries SET countryCode3='PCN' WHERE countryCode='PN'")
	cnn.execute("UPDATE countries SET countryCode3='PRI' WHERE countryCode='PR'")
	cnn.execute("UPDATE countries SET countryCode3='PSE' WHERE countryCode='PS'")
	cnn.execute("UPDATE countries SET countryCode3='PRT' WHERE countryCode='PT'")
	cnn.execute("UPDATE countries SET countryCode3='PLW' WHERE countryCode='PW'")
	cnn.execute("UPDATE countries SET countryCode3='PRY' WHERE countryCode='PY'")
	cnn.execute("UPDATE countries SET countryCode3='QAT' WHERE countryCode='QA'")
	cnn.execute("UPDATE countries SET countryCode3='REU' WHERE countryCode='RE'")
	cnn.execute("UPDATE countries SET countryCode3='ROU' WHERE countryCode='RO'")
	cnn.execute("UPDATE countries SET countryCode3='SRB' WHERE countryCode='RS'")
	cnn.execute("UPDATE countries SET countryCode3='RUS' WHERE countryCode='RU'")
	cnn.execute("UPDATE countries SET countryCode3='RWA' WHERE countryCode='RW'")
	cnn.execute("UPDATE countries SET countryCode3='SAU' WHERE countryCode='SA'")
	cnn.execute("UPDATE countries SET countryCode3='SLB' WHERE countryCode='SB'")
	cnn.execute("UPDATE countries SET countryCode3='SYC' WHERE countryCode='SC'")
	cnn.execute("UPDATE countries SET countryCode3='SDN' WHERE countryCode='SD'")
	cnn.execute("UPDATE countries SET countryCode3='SWE' WHERE countryCode='SE'")
	cnn.execute("UPDATE countries SET countryCode3='SGP' WHERE countryCode='SG'")
	cnn.execute("UPDATE countries SET countryCode3='SHN' WHERE countryCode='SH'")
	cnn.execute("UPDATE countries SET countryCode3='SVN' WHERE countryCode='SI'")
	cnn.execute("UPDATE countries SET countryCode3='SJM' WHERE countryCode='SJ'")
	cnn.execute("UPDATE countries SET countryCode3='SVK' WHERE countryCode='SK'")
	cnn.execute("UPDATE countries SET countryCode3='SLE' WHERE countryCode='SL'")
	cnn.execute("UPDATE countries SET countryCode3='SMR' WHERE countryCode='SM'")
	cnn.execute("UPDATE countries SET countryCode3='SEN' WHERE countryCode='SN'")
	cnn.execute("UPDATE countries SET countryCode3='SOM' WHERE countryCode='SO'")
	cnn.execute("UPDATE countries SET countryCode3='SUR' WHERE countryCode='SR'")
	cnn.execute("UPDATE countries SET countryCode3='SSD' WHERE countryCode='SS'")
	cnn.execute("UPDATE countries SET countryCode3='STP' WHERE countryCode='ST'")
	cnn.execute("UPDATE countries SET countryCode3='SLV' WHERE countryCode='SV'")
	cnn.execute("UPDATE countries SET countryCode3='SXM' WHERE countryCode='SX'")
	cnn.execute("UPDATE countries SET countryCode3='SYR' WHERE countryCode='SY'")
	cnn.execute("UPDATE countries SET countryCode3='SWZ' WHERE countryCode='SZ'")
	cnn.execute("UPDATE countries SET countryCode3='TCA' WHERE countryCode='TC'")
	cnn.execute("UPDATE countries SET countryCode3='TCD' WHERE countryCode='TD'")
	cnn.execute("UPDATE countries SET countryCode3='ATF' WHERE countryCode='TF'")
	cnn.execute("UPDATE countries SET countryCode3='TGO' WHERE countryCode='TG'")
	cnn.execute("UPDATE countries SET countryCode3='THA' WHERE countryCode='TH'")
	cnn.execute("UPDATE countries SET countryCode3='TJK' WHERE countryCode='TJ'")
	cnn.execute("UPDATE countries SET countryCode3='TKL' WHERE countryCode='TK'")
	cnn.execute("UPDATE countries SET countryCode3='TLS' WHERE countryCode='TL'")
	cnn.execute("UPDATE countries SET countryCode3='TKM' WHERE countryCode='TM'")
	cnn.execute("UPDATE countries SET countryCode3='TUN' WHERE countryCode='TN'")
	cnn.execute("UPDATE countries SET countryCode3='TON' WHERE countryCode='TO'")
	cnn.execute("UPDATE countries SET countryCode3='TUR' WHERE countryCode='TR'")
	cnn.execute("UPDATE countries SET countryCode3='TTO' WHERE countryCode='TT'")
	cnn.execute("UPDATE countries SET countryCode3='TUV' WHERE countryCode='TV'")
	cnn.execute("UPDATE countries SET countryCode3='TWN' WHERE countryCode='TW'")
	cnn.execute("UPDATE countries SET countryCode3='TZA' WHERE countryCode='TZ'")
	cnn.execute("UPDATE countries SET countryCode3='UKR' WHERE countryCode='UA'")
	cnn.execute("UPDATE countries SET countryCode3='UGA' WHERE countryCode='UG'")
	cnn.execute("UPDATE countries SET countryCode3='UMI' WHERE countryCode='UM'")
	cnn.execute("UPDATE countries SET countryCode3='USA' WHERE countryCode='US'")
	cnn.execute("UPDATE countries SET countryCode3='URY' WHERE countryCode='UY'")
	cnn.execute("UPDATE countries SET countryCode3='UZB' WHERE countryCode='UZ'")
	cnn.execute("UPDATE countries SET countryCode3='VAT' WHERE countryCode='VA'")
	cnn.execute("UPDATE countries SET countryCode3='VCT' WHERE countryCode='VC'")
	cnn.execute("UPDATE countries SET countryCode3='VEN' WHERE countryCode='VE'")
	cnn.execute("UPDATE countries SET countryCode3='VGB' WHERE countryCode='VG'")
	cnn.execute("UPDATE countries SET countryCode3='VIR' WHERE countryCode='VI'")
	cnn.execute("UPDATE countries SET countryCode3='VNM' WHERE countryCode='VN'")
	cnn.execute("UPDATE countries SET countryCode3='VUT' WHERE countryCode='VU'")
	cnn.execute("UPDATE countries SET countryCode3='WLF' WHERE countryCode='WF'")
	cnn.execute("UPDATE countries SET countryCode3='WSM' WHERE countryCode='WS'")
	cnn.execute("UPDATE countries SET countryCode3='YEM' WHERE countryCode='YE'")
	cnn.execute("UPDATE countries SET countryCode3='MYT' WHERE countryCode='YT'")
	cnn.execute("UPDATE countries SET countryCode3='ZAF' WHERE countryCode='ZA'")
	cnn.execute("UPDATE countries SET countryCode3='ZMB' WHERE countryCode='ZM'")
	cnn.execute("UPDATE countries SET countryCode3='ZWE' WHERE countryCode='ZW'")
end if

cnn.execute("UPDATE countries SET countryCode='AF',countryNumCurrency=0,countryCode3='AFG' WHERE countryID=3 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='AL',countryNumCurrency=8,countryCode3='ALB' WHERE countryID=4 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='DZ',countryNumCurrency=12,countryCode3='DZA' WHERE countryID=5 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='AS',countryNumCurrency=840,countryCode3='ASM' WHERE countryID=224 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='AD',countryNumCurrency=978,countryCode3='AND' WHERE countryID=6 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='AO',countryNumCurrency=973,countryCode3='AGO' WHERE countryID=7 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='AI',countryNumCurrency=951,countryCode3='AIA' WHERE countryID=8 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='AG',countryNumCurrency=951,countryCode3='ATG' WHERE countryID=10 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='AR',countryNumCurrency=32,countryCode3='ARG' WHERE countryID=11 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='AM',countryNumCurrency=51,countryCode3='ARM' WHERE countryID=12 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='AW',countryNumCurrency=533,countryCode3='ABW' WHERE countryID=13 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='AU',countryNumCurrency=36,countryCode3='AUS' WHERE countryID=14 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='AT',countryNumCurrency=978,countryCode3='AUT' WHERE countryID=15 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='AZ',countryNumCurrency=0,countryCode3='AZE' WHERE countryID=16 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='PT',countryNumCurrency=978,countryCode3='PRT' WHERE countryID=217 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='BS',countryNumCurrency=44,countryCode3='BHS' WHERE countryID=17 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='BH',countryNumCurrency=48,countryCode3='BHR' WHERE countryID=18 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='ES',countryNumCurrency=978,countryCode3='ESP' WHERE countryID=219 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='BD',countryNumCurrency=50,countryCode3='BGD' WHERE countryID=19 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='BB',countryNumCurrency=52,countryCode3='BRB' WHERE countryID=20 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='BY',countryNumCurrency=974,countryCode3='BLR' WHERE countryID=21 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='BE',countryNumCurrency=978,countryCode3='BEL' WHERE countryID=22 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='BZ',countryNumCurrency=84,countryCode3='BLZ' WHERE countryID=23 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='BJ',countryNumCurrency=952,countryCode3='BEN' WHERE countryID=24 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='BM',countryNumCurrency=60,countryCode3='BMU' WHERE countryID=25 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='BT',countryNumCurrency=64,countryCode3='BTN' WHERE countryID=26 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='BO',countryNumCurrency=68,countryCode3='BOL' WHERE countryID=27 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='BA',countryNumCurrency=977,countryCode3='BIH' WHERE countryID=28 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='BW',countryNumCurrency=72,countryCode3='BWA' WHERE countryID=29 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='BR',countryNumCurrency=986,countryCode3='BRA' WHERE countryID=30 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='VG',countryNumCurrency=840,countryCode3='VGB' WHERE countryID=208 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='BN',countryNumCurrency=96,countryCode3='BRN' WHERE countryID=31 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='BG',countryNumCurrency=975,countryCode3='BGR' WHERE countryID=32 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='BF',countryNumCurrency=952,countryCode3='BFA' WHERE countryID=33 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='BI',countryNumCurrency=108,countryCode3='BDI' WHERE countryID=34 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='KH',countryNumCurrency=116,countryCode3='KHM' WHERE countryID=35 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='CM',countryNumCurrency=950,countryCode3='CMR' WHERE countryID=36 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='CA',countryNumCurrency=124,countryCode3='CAN' WHERE countryID=2 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='CV',countryNumCurrency=132,countryCode3='CPV' WHERE countryID=37 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='KY',countryNumCurrency=136,countryCode3='CYM' WHERE countryID=38 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='CF',countryNumCurrency=950,countryCode3='CAF' WHERE countryID=39 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='TD',countryNumCurrency=950,countryCode3='TCD' WHERE countryID=40 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GB',countryNumCurrency=826,countryCode3='GBR' WHERE countryID=214 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='CL',countryNumCurrency=152,countryCode3='CHL' WHERE countryID=41 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='CN',countryNumCurrency=156,countryCode3='CHN' WHERE countryID=42 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='CO',countryNumCurrency=170,countryCode3='COL' WHERE countryID=43 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='KM',countryNumCurrency=174,countryCode3='COM' WHERE countryID=44 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='FR',countryNumCurrency=978,countryCode3='FRA' WHERE countryID=218 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='CR',countryNumCurrency=188,countryCode3='CRI' WHERE countryID=45 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='HR',countryNumCurrency=191,countryCode3='HRV' WHERE countryID=46 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='CU',countryNumCurrency=192,countryCode3='CUB' WHERE countryID=47 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='CY',countryNumCurrency=978,countryCode3='CYP' WHERE countryID=48 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='CZ',countryNumCurrency=203,countryCode3='CZE' WHERE countryID=49 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='DK',countryNumCurrency=208,countryCode3='DNK' WHERE countryID=50 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='DJ',countryNumCurrency=262,countryCode3='DJI' WHERE countryID=51 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='DM',countryNumCurrency=951,countryCode3='DMA' WHERE countryID=52 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='DO',countryNumCurrency=214,countryCode3='DOM' WHERE countryID=53 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='TP',countryNumCurrency=360,countryCode3='' WHERE countryID=54 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='EC',countryNumCurrency=840,countryCode3='ECU' WHERE countryID=55 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='EG',countryNumCurrency=818,countryCode3='EGY' WHERE countryID=56 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='SV',countryNumCurrency=840,countryCode3='SLV' WHERE countryID=57 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GB',countryNumCurrency=826,countryCode3='GBR' WHERE countryID=107 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GQ',countryNumCurrency=950,countryCode3='GNQ' WHERE countryID=58 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='EE',countryNumCurrency=233,countryCode3='EST' WHERE countryID=59 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='ET',countryNumCurrency=230,countryCode3='ETH' WHERE countryID=60 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='FK',countryNumCurrency=238,countryCode3='FLK' WHERE countryID=61 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='FO',countryNumCurrency=208,countryCode3='FRO' WHERE countryID=62 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='FJ',countryNumCurrency=242,countryCode3='FJI' WHERE countryID=63 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='FI',countryNumCurrency=978,countryCode3='FIN' WHERE countryID=64 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='FR',countryNumCurrency=978,countryCode3='FRA' WHERE countryID=65 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GF',countryNumCurrency=978,countryCode3='GUF' WHERE countryID=66 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='PF',countryNumCurrency=953,countryCode3='PYF' WHERE countryID=67 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GA',countryNumCurrency=950,countryCode3='GAB' WHERE countryID=68 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GM',countryNumCurrency=270,countryCode3='GMB' WHERE countryID=69 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GE',countryNumCurrency=981,countryCode3='GEO' WHERE countryID=70 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='DE',countryNumCurrency=978,countryCode3='DEU' WHERE countryID=71 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GH',countryNumCurrency=0,countryCode3='GHA' WHERE countryID=72 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GI',countryNumCurrency=826,countryCode3='GIB' WHERE countryID=73 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GB',countryNumCurrency=826,countryCode3='GBR' WHERE countryID=201 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GR',countryNumCurrency=978,countryCode3='GRC' WHERE countryID=74 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GL',countryNumCurrency=208,countryCode3='GRL' WHERE countryID=75 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GD',countryNumCurrency=951,countryCode3='GRD' WHERE countryID=76 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GP',countryNumCurrency=978,countryCode3='GLP' WHERE countryID=77 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GU',countryNumCurrency=840,countryCode3='GUM' WHERE countryID=78 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GT',countryNumCurrency=320,countryCode3='GTM' WHERE countryID=79 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GN',countryNumCurrency=324,countryCode3='GIN' WHERE countryID=80 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GW',countryNumCurrency=952,countryCode3='GNB' WHERE countryID=81 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GY',countryNumCurrency=328,countryCode3='GUY' WHERE countryID=82 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='HT',countryNumCurrency=840,countryCode3='HTI' WHERE countryID=83 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='HN',countryNumCurrency=340,countryCode3='HND' WHERE countryID=84 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='HK',countryNumCurrency=344,countryCode3='HKG' WHERE countryID=85 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='HU',countryNumCurrency=348,countryCode3='HUN' WHERE countryID=86 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='IS',countryNumCurrency=352,countryCode3='ISL' WHERE countryID=87 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='IN',countryNumCurrency=356,countryCode3='IND' WHERE countryID=88 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='ID',countryNumCurrency=360,countryCode3='IDN' WHERE countryID=89 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='IR',countryNumCurrency=364,countryCode3='IRN' WHERE countryID=213 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='IQ',countryNumCurrency=368,countryCode3='IRQ' WHERE countryID=90 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='IE',countryNumCurrency=978,countryCode3='IRL' WHERE countryID=91 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GB',countryNumCurrency=826,countryCode3='GBR' WHERE countryID=216 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='IL',countryNumCurrency=376,countryCode3='ISR' WHERE countryID=92 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='IT',countryNumCurrency=978,countryCode3='ITA' WHERE countryID=93 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='CI',countryNumCurrency=952,countryCode3='CIV' WHERE countryID=222 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='JM',countryNumCurrency=388,countryCode3='JAM' WHERE countryID=94 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='JP',countryNumCurrency=392,countryCode3='JPN' WHERE countryID=95 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='JO',countryNumCurrency=400,countryCode3='JOR' WHERE countryID=96 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='KZ',countryNumCurrency=398,countryCode3='KAZ' WHERE countryID=97 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='KE',countryNumCurrency=404,countryCode3='KEN' WHERE countryID=98 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='KI',countryNumCurrency=36,countryCode3='KIR' WHERE countryID=99 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='KW',countryNumCurrency=414,countryCode3='KWT' WHERE countryID=102 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='LV',countryNumCurrency=428,countryCode3='LVA' WHERE countryID=103 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='LB',countryNumCurrency=422,countryCode3='LBN' WHERE countryID=104 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='LS',countryNumCurrency=426,countryCode3='LSO' WHERE countryID=105 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='LR',countryNumCurrency=430,countryCode3='LBR' WHERE countryID=106 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='LI',countryNumCurrency=756,countryCode3='LIE' WHERE countryID=108 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='LT',countryNumCurrency=440,countryCode3='LTU' WHERE countryID=109 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='LU',countryNumCurrency=978,countryCode3='LUX' WHERE countryID=110 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='MO',countryNumCurrency=446,countryCode3='MAC' WHERE countryID=111 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='MK',countryNumCurrency=807,countryCode3='MKD' WHERE countryID=112 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='MG',countryNumCurrency=0,countryCode3='MDG' WHERE countryID=113 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='MW',countryNumCurrency=454,countryCode3='MWI' WHERE countryID=114 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='MY',countryNumCurrency=458,countryCode3='MYS' WHERE countryID=115 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='MV',countryNumCurrency=462,countryCode3='MDV' WHERE countryID=116 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='ML',countryNumCurrency=952,countryCode3='MLI' WHERE countryID=117 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='MT',countryNumCurrency=978,countryCode3='MLT' WHERE countryID=118 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='MQ',countryNumCurrency=978,countryCode3='MTQ' WHERE countryID=119 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='MR',countryNumCurrency=478,countryCode3='MRT' WHERE countryID=120 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='MU',countryNumCurrency=480,countryCode3='MUS' WHERE countryID=121 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='MX',countryNumCurrency=484,countryCode3='MEX' WHERE countryID=122 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='MD',countryNumCurrency=498,countryCode3='MDA' WHERE countryID=123 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='MC',countryNumCurrency=978,countryCode3='MCO' WHERE countryID=124 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='MN',countryNumCurrency=496,countryCode3='MNG' WHERE countryID=125 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='ME',countryNumCurrency=978,countryCode3='MNE' WHERE countryID=223 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='MS',countryNumCurrency=951,countryCode3='MSR' WHERE countryID=126 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='MA',countryNumCurrency=504,countryCode3='MAR' WHERE countryID=127 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='MZ',countryNumCurrency=0,countryCode3='MOZ' WHERE countryID=128 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='MM',countryNumCurrency=104,countryCode3='MMR' WHERE countryID=129 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='NA',countryNumCurrency=516,countryCode3='NAM' WHERE countryID=130 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='NR',countryNumCurrency=36,countryCode3='NRU' WHERE countryID=131 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='NP',countryNumCurrency=524,countryCode3='NPL' WHERE countryID=132 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='NL',countryNumCurrency=978,countryCode3='NLD' WHERE countryID=133 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='AN',countryNumCurrency=532,countryCode3='' WHERE countryID=134 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='NC',countryNumCurrency=953,countryCode3='NCL' WHERE countryID=135 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='NZ',countryNumCurrency=554,countryCode3='NZL' WHERE countryID=136 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='NI',countryNumCurrency=558,countryCode3='NIC' WHERE countryID=137 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='NE',countryNumCurrency=952,countryCode3='NER' WHERE countryID=138 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='NG',countryNumCurrency=566,countryCode3='NGA' WHERE countryID=139 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='NU',countryNumCurrency=554,countryCode3='NIU' WHERE countryID=140 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='NF',countryNumCurrency=36,countryCode3='NFK' WHERE countryID=141 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='KP',countryNumCurrency=408,countryCode3='PRK' WHERE countryID=100 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='GB',countryNumCurrency=826,countryCode3='GBR' WHERE countryID=142 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='NO',countryNumCurrency=578,countryCode3='NOR' WHERE countryID=143 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='OM',countryNumCurrency=512,countryCode3='OMN' WHERE countryID=144 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='PK',countryNumCurrency=586,countryCode3='PAK' WHERE countryID=145 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='PA',countryNumCurrency=590,countryCode3='PAN' WHERE countryID=146 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='PG',countryNumCurrency=598,countryCode3='PNG' WHERE countryID=147 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='PY',countryNumCurrency=600,countryCode3='PRY' WHERE countryID=148 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='PE',countryNumCurrency=604,countryCode3='PER' WHERE countryID=149 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='PH',countryNumCurrency=608,countryCode3='PHL' WHERE countryID=150 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='PN',countryNumCurrency=554,countryCode3='PCN' WHERE countryID=151 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='PL',countryNumCurrency=985,countryCode3='POL' WHERE countryID=152 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='PT',countryNumCurrency=978,countryCode3='PRT' WHERE countryID=153 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='PR',countryNumCurrency=840,countryCode3='PRI' WHERE countryID=215 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='QA',countryNumCurrency=634,countryCode3='QAT' WHERE countryID=154 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='RE',countryNumCurrency=978,countryCode3='REU' WHERE countryID=155 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='RO',countryNumCurrency=946,countryCode3='ROU' WHERE countryID=156 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='RU',countryNumCurrency=643,countryCode3='RUS' WHERE countryID=157 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='RW',countryNumCurrency=646,countryCode3='RWA' WHERE countryID=158 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='SH',countryNumCurrency=654,countryCode3='SHN' WHERE countryID=177 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='KN',countryNumCurrency=951,countryCode3='KNA' WHERE countryID=159 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='LC',countryNumCurrency=951,countryCode3='LCA' WHERE countryID=160 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='PM',countryNumCurrency=978,countryCode3='SPM' WHERE countryID=178 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='VC',countryNumCurrency=951,countryCode3='VCT' WHERE countryID=161 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='SM',countryNumCurrency=978,countryCode3='SMR' WHERE countryID=163 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='ST',countryNumCurrency=678,countryCode3='STP' WHERE countryID=164 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='SA',countryNumCurrency=682,countryCode3='SAU' WHERE countryID=165 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='SN',countryNumCurrency=952,countryCode3='SEN' WHERE countryID=166 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='RS',countryNumCurrency=941,countryCode3='SRB' WHERE countryID=221 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='SC',countryNumCurrency=690,countryCode3='SYC' WHERE countryID=167 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='SL',countryNumCurrency=694,countryCode3='SLE' WHERE countryID=168 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='SG',countryNumCurrency=702,countryCode3='SGP' WHERE countryID=169 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='SK',countryNumCurrency=0,countryCode3='SVK' WHERE countryID=170 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='SI',countryNumCurrency=978,countryCode3='SVN' WHERE countryID=171 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='SB',countryNumCurrency=90,countryCode3='SLB' WHERE countryID=172 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='SO',countryNumCurrency=706,countryCode3='SOM' WHERE countryID=173 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='ZA',countryNumCurrency=710,countryCode3='ZAF' WHERE countryID=174 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='KR',countryNumCurrency=410,countryCode3='KOR' WHERE countryID=101 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='ES',countryNumCurrency=978,countryCode3='ESP' WHERE countryID=175 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='LK',countryNumCurrency=144,countryCode3='LKA' WHERE countryID=176 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='SD',countryNumCurrency=0,countryCode3='SDN' WHERE countryID=179 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='SR',countryNumCurrency=0,countryCode3='SUR' WHERE countryID=180 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='SZ',countryNumCurrency=748,countryCode3='SWZ' WHERE countryID=181 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='SE',countryNumCurrency=752,countryCode3='SWE' WHERE countryID=182 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='CH',countryNumCurrency=756,countryCode3='CHE' WHERE countryID=183 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='SY',countryNumCurrency=760,countryCode3='SYR' WHERE countryID=184 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='TW',countryNumCurrency=901,countryCode3='TWN' WHERE countryID=185 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='TJ',countryNumCurrency=972,countryCode3='TJK' WHERE countryID=186 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='TZ',countryNumCurrency=834,countryCode3='TZA' WHERE countryID=187 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='TH',countryNumCurrency=764,countryCode3='THA' WHERE countryID=188 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='TG',countryNumCurrency=952,countryCode3='TGO' WHERE countryID=189 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='TK',countryNumCurrency=554,countryCode3='TKL' WHERE countryID=190 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='TO',countryNumCurrency=776,countryCode3='TON' WHERE countryID=191 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='TT',countryNumCurrency=780,countryCode3='TTO' WHERE countryID=192 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='TN',countryNumCurrency=788,countryCode3='TUN' WHERE countryID=193 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='TR',countryNumCurrency=949,countryCode3='TUR' WHERE countryID=194 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='TM',countryNumCurrency=0,countryCode3='TKM' WHERE countryID=195 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='TC',countryNumCurrency=840,countryCode3='TCA' WHERE countryID=196 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='TV',countryNumCurrency=0,countryCode3='TUV' WHERE countryID=197 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='UG',countryNumCurrency=800,countryCode3='UGA' WHERE countryID=198 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='UA',countryNumCurrency=980,countryCode3='UKR' WHERE countryID=199 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='AE',countryNumCurrency=784,countryCode3='ARE' WHERE countryID=200 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='US',countryNumCurrency=840,countryCode3='USA' WHERE countryID=1 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='UY',countryNumCurrency=858,countryCode3='URY' WHERE countryID=202 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='VI',countryNumCurrency=0,countryCode3='VIR' WHERE countryID=220 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='UZ',countryNumCurrency=860,countryCode3='UZB' WHERE countryID=203 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='VU',countryNumCurrency=548,countryCode3='VUT' WHERE countryID=204 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='VA',countryNumCurrency=978,countryCode3='VAT' WHERE countryID=205 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='VE',countryNumCurrency=937,countryCode3='VEN' WHERE countryID=206 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='VN',countryNumCurrency=704,countryCode3='VNM' WHERE countryID=207 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='WF',countryNumCurrency=953,countryCode3='WLF' WHERE countryID=209 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='WS',countryNumCurrency=882,countryCode3='WSM' WHERE countryID=162 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='YE',countryNumCurrency=886,countryCode3='YEM' WHERE countryID=210 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='ZM',countryNumCurrency=894,countryCode3='ZMB' WHERE countryID=211 AND countryCode IS NULL")
cnn.execute("UPDATE countries SET countryCode='ZW',countryNumCurrency=0,countryCode3='ZWE' WHERE countryID=212 AND countryCode IS NULL")

call checkaddcolumn("sections","sRecommend",TRUE,bitfield,"","")

on error resume next
printtickdiv("Checking for devicenotifications table")
err.number = 0
sSQL = "SELECT * FROM devicenotifications"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding devicenotifications table")
	cnn.execute("CREATE TABLE devicenotifications (dnID "&txtcl&"(127) PRIMARY KEY, dnLastUpdated "&datecl&")")
end if

printtickdiv("Checking for Shipping Insurance upgrade")
call checkaddcolumn("admin","insuranceDomPercent",FALSE,dblcl,"","")
call checkaddcolumn("admin","insuranceDomMin",FALSE,dblcl,"","")
call checkaddcolumn("admin","insuranceIntPercent",FALSE,dblcl,"","")
call checkaddcolumn("admin","insuranceIntMin",FALSE,dblcl,"","")
call checkaddcolumn("admin","shipInsuranceDom",FALSE,bytecl,"","")
call checkaddcolumn("admin","noCarrierDomIns",FALSE,bitfield,"","")
call checkaddcolumn("admin","noCarrierIntIns",FALSE,bitfield,"","")
if checkaddcolumn("admin","shipInsuranceInt",FALSE,bytecl,"","") then
	shipins=0
	insmin=0
	inspercent=0
	if is_numeric(addshippinginsurance) AND is_numeric(shipinsuranceamt) then
		shipins=abs(addshippinginsurance)
		if addshippinginsurance>0 then
			inspercent=shipinsuranceamt
		else
			insmin=shipinsuranceamt
		end if
	end if
	if forceinsuranceselection then shipins=3
	cnn.execute("UPDATE admin SET shipInsuranceDom=" & shipins & ",insuranceDomMin=" & insmin & ",insuranceDomPercent=" & inspercent & " WHERE adminID=1")
	cnn.execute("UPDATE admin SET shipInsuranceInt=shipInsuranceDom,insuranceIntMin=insuranceDomMin,insuranceIntPercent=insuranceDomPercent WHERE adminID=1")
end if

call checkaddcolumn("coupons","cpnInsurance",FALSE,bitfield,"","1")

call checkaddcolumn("products","pUpload",FALSE,bytecl,"","")
call checkaddcolumn("admin","uploadDir",FALSE,txtcl,"("&txtcollen&")","")

printtickdiv("Checking for image upload upgrade")
on error resume next
err.number = 0
sSQL="SELECT upID FROM imageuploads WHERE upID=0"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding image upload table")
	sSQL ="CREATE TABLE imageuploads (upID "&autoinc&",upOrderID INT DEFAULT 0,upComments "&txtcl&"("&txtcollen&"),upFilename "&txtcl&"(255))"
	cnn.execute(sSQL)
end if

call checkaddcolumn("admin","adminDeviceNotifAlert",FALSE,txtcl,"(255)","")

call checkaddcolumn("admin","blockMultiPurchase",FALSE,smallcl,"",IIfVr(blockmultipurchase<>"",blockmultipurchase,0))
call checkaddcolumn("admin","blockMaxCartAdds",FALSE,smallcl,"","")

call checkaddcolumn("products","pCustomCSS",FALSE,txtcl,"(255)","")

call checkaddcolumn("products","pSchemaType",FALSE,bytecl,"","")

call checkaddcolumn("sections","sCustomCSS",FALSE,txtcl,"(255)","")

function getcheckfromadmin(admincol)
	getcheckfromadmin=""
	on error resume next
	err.number = 0
	sSQL = "SELECT "&admincol&" FROM admin WHERE adminID=1"
	rs.open sSQL,cnn,0,1
	if err.number=0 then getcheckfromadmin=rs(admincol)
	rs.close
	on error goto 0
end function
on error resume next
printtickdiv("Checking for Admin Shipping upgrade")
err.number = 0
sSQL = "SELECT * FROM adminshipping"
rs.Open sSQL,cnn,0,1
errnum=err.number
rs.Close
on error goto 0
if errnum<>0 then
	printtick("Adding Admin Shipping table")
	sSQL="CREATE TABLE adminshipping (adminShipID INT PRIMARY KEY," & _
		"adminPacking "&smallcl&" DEFAULT 0," & _
		"AusPostAPI "&txtcl&"(100)," & _
		"adminCanPostUser "&txtcl&"(100)," & _
		"adminCanPostLogin "&txtcl&"(100)," & _
		"adminCanPostPass "&txtcl&"(100)," & _
		"adminUSPSUser "&txtcl&"(100)," & _
		"smartPostHub "&txtcl&"(15)," & _
		"adminUPSUser "&txtcl&"(100)," & _
		"adminUPSpw "&txtcl&"(100)," & _
		"adminUPSAccess "&txtcl&"(100)," & _
		"adminUPSAccount "&txtcl&"(100)," & _
		"adminUPSNegotiated "&smallcl&" DEFAULT 0," & _
		"FedexAccountNo "&txtcl&"(100)," & _
		"FedexMeter "&txtcl&"(100)," & _
		"FedexUserKey "&txtcl&"(100)," & _
		"FedexUserPwd "&txtcl&"(100)," & _
		"DHLSiteID "&txtcl&"(50)," & _
		"DHLSitePW "&txtcl&"(50)," & _
		"DHLAccountNo "&txtcl&"(50)," & _
		"shipStationUser "&txtcl&"(50)," & _
		"shipStationPass "&txtcl&"(50))"
	cnn.execute(sSQL)
	
	AusPostAPI=getcheckfromadmin("AusPostAPI")
	adminCanPostLogin=getcheckfromadmin("adminCanPostLogin")
	adminCanPostPass=getcheckfromadmin("adminCanPostPass")
	adminCanPostUser=getcheckfromadmin("adminCanPostUser")
	adminPacking=getcheckfromadmin("adminPacking")
	adminUSPSUser=getcheckfromadmin("adminUSPSUser")
	smartPostHub=getcheckfromadmin("smartPostHub")
	adminUPSUser=getcheckfromadmin("adminUPSUser")
	adminUPSpw=getcheckfromadmin("adminUPSpw")
	adminUPSAccess=getcheckfromadmin("adminUPSAccess")
	adminUPSAccount=getcheckfromadmin("adminUPSAccount")
	adminUPSNegotiated=getcheckfromadmin("adminUPSNegotiated")
	FedexAccountNo=getcheckfromadmin("FedexAccountNo")
	FedexMeter=getcheckfromadmin("FedexMeter")
	FedexUserKey=getcheckfromadmin("FedexUserKey")
	FedexUserPwd=getcheckfromadmin("FedexUserPwd")
	DHLSiteID=getcheckfromadmin("DHLSiteID")
	DHLSitePW=getcheckfromadmin("DHLSitePW")
	DHLAccountNo=getcheckfromadmin("DHLAccountNo")
	if adminPacking="" then adminPacking=0
	if adminUPSNegotiated="" then adminUPSNegotiated=0
	
	sSQL="INSERT INTO adminshipping (adminShipID,AusPostAPI,adminCanPostLogin,adminCanPostPass,adminCanPostUser,adminPacking,adminUSPSUser,smartPostHub,adminUPSUser,adminUPSpw,adminUPSAccess,adminUPSAccount,adminUPSNegotiated,FedexAccountNo,FedexMeter,FedexUserKey,FedexUserPwd,DHLSiteID,DHLSitePW,DHLAccountNo) VALUES (1,"
	sSQL=sSQL&"'"&escape_string(AusPostAPI)&"','"&escape_string(adminCanPostLogin)&"','"&escape_string(adminCanPostPass)&"','"&escape_string(adminCanPostUser)&"',"&adminPacking&",'"&escape_string(adminUSPSUser)&"','"&escape_string(smartPostHub)&"','"&escape_string(adminUPSUser)&"','"&escape_string(adminUPSpw)&"','"&escape_string(adminUPSAccess)&"','"&escape_string(adminUPSAccount)&"',"&IIfVr(adminUPSNegotiated,"1","0")&",'"&escape_string(FedexAccountNo)&"','"&escape_string(FedexMeter)&"','"&escape_string(FedexUserKey)&"','"&escape_string(FedexUserPwd)&"','"&escape_string(DHLSiteID)&"','"&escape_string(DHLSitePW)&"','"&escape_string(DHLAccountNo)&"')"
	cnn.execute(sSQL)

	call dropconstraintcolumn("admin","AusPostAPI")
	call dropconstraintcolumn("admin","adminCanPostLogin")
	call dropconstraintcolumn("admin","adminCanPostPass")
	call dropconstraintcolumn("admin","adminCanPostUser")
	call dropconstraintcolumn("admin","adminPacking")
	call dropconstraintcolumn("admin","adminUSPSUser")
	call dropconstraintcolumn("admin","adminUSPSpw")
	call dropconstraintcolumn("admin","smartPostHub")
	call dropconstraintcolumn("admin","adminUPSUser")
	call dropconstraintcolumn("admin","adminUPSpw")
	call dropconstraintcolumn("admin","adminUPSAccess")
	call dropconstraintcolumn("admin","adminUPSAccount")
	call dropconstraintcolumn("admin","adminUPSNegotiated")
	call dropconstraintcolumn("admin","FedexAccountNo")
	call dropconstraintcolumn("admin","FedexMeter")
	call dropconstraintcolumn("admin","FedexUserKey")
	call dropconstraintcolumn("admin","FedexUserPwd")
	call dropconstraintcolumn("admin","DHLSiteID")
	call dropconstraintcolumn("admin","DHLSitePW")
	call dropconstraintcolumn("admin","DHLAccountNo")
end if
response.flush

sSQL = "SELECT * FROM adminshipping WHERE adminShipID=1"
rs.Open sSQL,cnn,0,1
hasadminshipsettings=NOT rs.EOF
rs.Close
if NOT hasadminshipsettings then
	printtick("Adding Admin Shipping Settings")
	AusPostAPI=getcheckfromadmin("AusPostAPI")
	adminCanPostLogin=getcheckfromadmin("adminCanPostLogin")
	adminCanPostPass=getcheckfromadmin("adminCanPostPass")
	adminCanPostUser=getcheckfromadmin("adminCanPostUser")
	adminPacking=cint(getcheckfromadmin("adminPacking"))
	adminUSPSUser=getcheckfromadmin("adminUSPSUser")
	smartPostHub=getcheckfromadmin("smartPostHub")
	adminUPSUser=getcheckfromadmin("adminUPSUser")
	adminUPSpw=getcheckfromadmin("adminUPSpw")
	adminUPSAccess=getcheckfromadmin("adminUPSAccess")
	adminUPSAccount=getcheckfromadmin("adminUPSAccount")
	adminUPSNegotiated=cint(getcheckfromadmin("adminUPSNegotiated"))
	FedexAccountNo=getcheckfromadmin("FedexAccountNo")
	FedexMeter=getcheckfromadmin("FedexMeter")
	FedexUserKey=getcheckfromadmin("FedexUserKey")
	FedexUserPwd=getcheckfromadmin("FedexUserPwd")
	DHLSiteID=getcheckfromadmin("DHLSiteID")
	DHLSitePW=getcheckfromadmin("DHLSitePW")
	DHLAccountNo=getcheckfromadmin("DHLAccountNo")
	if adminPacking="" then adminPacking=0
	if adminUPSNegotiated="" then adminUPSNegotiated=0
	
	sSQL="INSERT INTO adminshipping (adminShipID,AusPostAPI,adminCanPostLogin,adminCanPostPass,adminCanPostUser,adminPacking,adminUSPSUser,smartPostHub,adminUPSUser,adminUPSpw,adminUPSAccess,adminUPSAccount,adminUPSNegotiated,FedexAccountNo,FedexMeter,FedexUserKey,FedexUserPwd,DHLSiteID,DHLSitePW,DHLAccountNo) VALUES (1,"
	sSQL=sSQL&"'"&escape_string(AusPostAPI)&"','"&escape_string(adminCanPostLogin)&"','"&escape_string(adminCanPostPass)&"','"&escape_string(adminCanPostUser)&"',"&adminPacking&",'"&escape_string(adminUSPSUser)&"','"&escape_string(smartPostHub)&"','"&escape_string(adminUPSUser)&"','"&escape_string(adminUPSpw)&"','"&escape_string(adminUPSAccess)&"','"&escape_string(adminUPSAccount)&"',"&adminUPSNegotiated&",'"&escape_string(FedexAccountNo)&"','"&escape_string(FedexMeter)&"','"&escape_string(FedexUserKey)&"','"&escape_string(FedexUserPwd)&"','"&escape_string(DHLSiteID)&"','"&escape_string(DHLSitePW)&"','"&escape_string(DHLAccountNo)&"')"
	cnn.execute(sSQL)
	
	call dropconstraintcolumn("admin","AusPostAPI")
	call dropconstraintcolumn("admin","adminCanPostLogin")
	call dropconstraintcolumn("admin","adminCanPostPass")
	call dropconstraintcolumn("admin","adminCanPostUser")
	call dropconstraintcolumn("admin","adminPacking")
	call dropconstraintcolumn("admin","adminUSPSUser")
	call dropconstraintcolumn("admin","adminUSPSpw")
	call dropconstraintcolumn("admin","smartPostHub")
	call dropconstraintcolumn("admin","adminUPSUser")
	call dropconstraintcolumn("admin","adminUPSpw")
	call dropconstraintcolumn("admin","adminUPSAccess")
	call dropconstraintcolumn("admin","adminUPSAccount")
	call dropconstraintcolumn("admin","adminUPSNegotiated")
	call dropconstraintcolumn("admin","FedexAccountNo")
	call dropconstraintcolumn("admin","FedexMeter")
	call dropconstraintcolumn("admin","FedexUserKey")
	call dropconstraintcolumn("admin","FedexUserPwd")
	call dropconstraintcolumn("admin","DHLSiteID")
	call dropconstraintcolumn("admin","DHLSitePW")
	call dropconstraintcolumn("admin","DHLAccountNo")
end if

call checkaddcolumn("products","pSiteID",FALSE,bytecl,"","")

call checkaddcolumn("products","pTitle2",FALSE,txtcl,"(255)","")
call checkaddcolumn("products","pMetaDesc2",FALSE,txtcl,"(255)","")
call checkaddcolumn("sections","sTitle2",FALSE,txtcl,"(255)","")
call checkaddcolumn("sections","sMetaDesc2",FALSE,txtcl,"(255)","")
call checkaddcolumn("products","pSearchParams2",FALSE,memocl,"","")

call checkaddcolumn("products","pTitle3",FALSE,txtcl,"(255)","")
call checkaddcolumn("products","pMetaDesc3",FALSE,txtcl,"(255)","")
call checkaddcolumn("sections","sTitle3",FALSE,txtcl,"(255)","")
call checkaddcolumn("sections","sMetaDesc3",FALSE,txtcl,"(255)","")
call checkaddcolumn("products","pSearchParams3",FALSE,memocl,"","")

call checkaddcolumn("orders","ordTransSession",FALSE,txtcl,"(255)","")

if checkaddcolumn("countries","countryTaxThreshold",FALSE,dblcl,"","") then
	if is_numeric(uktaxthreshold) then
		cnn.execute("UPDATE countries SET countryTaxThreshold=" & escape_string(uktaxthreshold) & " WHERE countryID IN (73,107,142,201,214,216)")
	end if
end if

on error resume next
cnn.execute("ALTER TABLE products "&altcl&" pCustomCSS VARCHAR(255) NULL")

cnn.execute("CREATE INDEX aceOrderID_Indx ON abandonedcartemail(aceOrderID)")

cnn.execute("CREATE INDEX cartClientId_Indx ON cart(cartClientId)")
cnn.execute("CREATE INDEX cartCompleted_Indx ON cart(cartCompleted)")
cnn.execute("CREATE INDEX cartDateAdded_Indx ON cart(cartDateAdded)")
cnn.execute("CREATE INDEX cartOrderID_Indx ON cart(cartOrderID)")
cnn.execute("CREATE INDEX cartProdID_Indx ON cart(cartProdID)")
cnn.execute("CREATE INDEX cartSessionID_Indx ON cart(cartSessionID)")

cnn.execute("CREATE INDEX coCartID_Indx ON cartoptions(coCartID)")
cnn.execute("CREATE INDEX coOptID_Indx ON cartoptions(coOptID)")

cnn.execute("CREATE INDEX countryName_Indx ON countries(countryName)")

cnn.execute("CREATE INDEX cpnStartDate_Indx ON coupons(cpnStartDate)")
cnn.execute("CREATE INDEX cpnEndDate_Indx ON coupons(cpnEndDate)")

cnn.execute("CREATE INDEX cpaCpnID_Indx ON cpnassign(cpaCpnID)")
cnn.execute("CREATE INDEX cpaAssignment_Indx ON cpnassign(cpaAssignment)")

cnn.execute("CREATE INDEX mSpID_Indx ON multisections(pID)")
cnn.execute("CREATE INDEX mSpSection_Indx ON multisections(pSection)")

cnn.execute("CREATE INDEX mSCscID_Indx ON multisearchcriteria(mSCscID)")
cnn.execute("CREATE INDEX mSCpID_Indx ON multisearchcriteria(mSCpID)")

cnn.execute("CREATE INDEX optGroup_Indx ON options(optGroup)")

cnn.execute("CREATE INDEX ordClientId_Indx ON orders(ordClientId)")
cnn.execute("CREATE INDEX ordDate_Indx ON orders(ordDate)")
cnn.execute("CREATE INDEX ordSessionID_Indx ON orders(ordSessionId)")
cnn.execute("CREATE INDEX ordStatus_Indx ON orders(ordStatus)")

cnn.execute("CREATE INDEX poProdID_Indx ON prodoptions(poProdID)")
cnn.execute("CREATE INDEX poOptionGroup_Indx ON prodoptions(poOptionGroup)")

cnn.execute("CREATE INDEX pDateAdded_Indx ON products(pDateAdded)")
cnn.execute("CREATE INDEX pDisplay_Indx ON products(pDisplay)")
cnn.execute("CREATE INDEX pManufacturer_Indx ON products(pManufacturer)")
cnn.execute("CREATE INDEX pName_Indx ON products(pName)")
cnn.execute("CREATE INDEX pNumRatings_Indx ON products(pNumRatings)")
cnn.execute("CREATE INDEX pNumSales_Indx ON products(pNumSales)")
cnn.execute("CREATE INDEX pOrder_Indx ON products(pOrder)")
cnn.execute("CREATE INDEX pPopularity_Indx ON products(pPopularity)")
cnn.execute("CREATE INDEX pPrice_Indx ON products(pPrice)")
cnn.execute("CREATE INDEX pSection_Indx ON products(pSection)")
cnn.execute("CREATE INDEX pSKU_Indx ON products(pSKU)")
cnn.execute("CREATE INDEX pTotRating_Indx ON products(pTotRating)")

cnn.execute("CREATE INDEX stateName_Indx ON states(stateName)")
cnn.execute("CREATE INDEX stateName2_Indx ON states(stateName2)")
cnn.execute("CREATE INDEX stateName3_Indx ON states(stateName3)")

cnn.execute("CREATE INDEX rvCustomerID_Indx ON recentlyviewed(rvCustomerID)")
cnn.execute("CREATE INDEX rvDate_Indx ON recentlyviewed(rvDate)")
cnn.execute("CREATE INDEX rvProdId_Indx ON recentlyviewed(rvProdId)")
cnn.execute("CREATE INDEX rvProdSection_Indx ON recentlyviewed(rvProdSection)")
cnn.execute("CREATE INDEX rvSessionID_Indx ON recentlyviewed(rvSessionID)")

cnn.execute("CREATE INDEX scGroup_Indx ON searchcriteria(scGroup)")
cnn.execute("CREATE INDEX scName_Indx ON searchcriteria(scName)")
cnn.execute("CREATE INDEX scName2_Indx ON searchcriteria(scName2)")
cnn.execute("CREATE INDEX scName3_Indx ON searchcriteria(scName3)")
cnn.execute("CREATE INDEX scOrder_Indx ON searchcriteria(scOrder)")

cnn.execute("CREATE INDEX scgOrder_Indx ON searchcriteriagroup(scgOrder)")

cnn.execute("CREATE INDEX sectionDisabled_Indx ON sections(sectionDisabled)")
cnn.execute("CREATE INDEX sectionName_Indx ON sections(sectionName)")
cnn.execute("CREATE INDEX sectionName2_Indx ON sections(sectionName2)")
cnn.execute("CREATE INDEX sectionName3_Indx ON sections(sectionName3)")
cnn.execute("CREATE INDEX sectionOrder_Indx ON sections(sectionOrder)")
cnn.execute("CREATE INDEX topSection_Indx ON sections(topSection)")
on error goto 0

if success then
	Application.Lock()
	Application("getadminsettings")=""
	Application.UnLock()
	cnn.execute("UPDATE admin SET updLastCheck="&datedelim&vsusdate(Date()-100)&datedelim&",updRecommended='',updSecurity=0,updShouldUpd=0")
	printtick("Updating version number to 'Ecommerce Plus "&sVersion&"'")
	cnn.execute("UPDATE admin SET adminVersion='Ecommerce Plus "&sVersion&"'")
	printtick("<strong>Everything updated successfully ! ! !</strong>")
	response.write "<script type=""text/javascript"">iqueue.push('C');</script>" & vbCrLf
	' response.write "<meta http-equiv=""Refresh"" content=""8; URL=updatestore.asp?posted=2"">"
else
	printtick("<font color='#FF0000'><b>Terminated but with errors</b></font>")
end if

sSQL = "INSERT INTO auditlog (userID,eventType,eventDate,eventSuccess,eventOrigin,areaAffected) VALUES ('UPDATE','UPDATESTORE',"&datedelim&vsusdate(Now())&datedelim&","&IIfVr(success,1,0)&",'UPDATER " & sVersion & "','DBVERSION')"
cnn.execute(sSQL)

elseif request.querystring("posted")="2" then
	rs.open "SELECT adminUSZones,adminCountry,adminShipping,adminIntShipping,adminAltRates FROM admin WHERE adminID=1",cnn,0,1
	splitUSZones=(int(rs("adminUSZones"))=1)
	countryID=rs("adminCountry")
	adminIntShipping=int(rs("adminIntShipping"))
	shipType=int(rs("adminShipping"))
	adminAltRates=rs("adminAltRates")
	rs.close
	alternateratesweightbased=FALSE
	if adminAltRates>0 then
		sSQL = "SELECT altrateid FROM alternaterates WHERE (altrateid=2 OR altrateid=5) AND (usealtmethod<>0 OR usealtmethodintl<>0)"
		rs.open sSQL,cnn,0,1
		alternateratesweightbased = NOT rs.EOF
		rs.close
	end if
	editzones = ((shipType=2 OR shipType=5 OR adminIntShipping=2 OR adminIntShipping=5 OR alternateratesweightbased) AND splitUSZones)
	if editzones then
		sSQL="SELECT stateID,stateName,pzName FROM states LEFT JOIN postalzones ON postalzones.pzID=states.stateZone WHERE stateCountryID="&countryID&" AND pzName IS NULL"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			response.write "<div style=""padding:24px;border:1px solid #e0e0e0;width:80%;margin:0 auto;background:#FFF;-moz-border-radius:10px;-webkit-border-radius:10px;margin-top:40px;"">"
			response.write "<div style=""color:#FF0000;font-weight:bold"">IMPORTANT, some of your states do not have a postal zone assigned.<br />You should correct this in the states admin page!!</div>"
			do while NOT rs.EOF
				response.write "<div style=""font-weight:bold"">The following state does not have a postal zone: " & rs("stateName") & "</div>"
				rs.movenext
			loop
			response.write "</div>"
		end if
		rs.close
	end if
%>
<div style="padding:24px;border:1px solid #e0e0e0;width:80%;margin:0 auto;background:#FFF;-moz-border-radius:10px;-webkit-border-radius:10px;margin-top:40px;">
<p style="font-size: 20px;font-family : Arial,sans-serif;font-weight : normal;padding-top: 6px;color : #2F3D6F;margin-top:0px;">The database upgrade script has completed.</p>
<p>After updating, please check our updater checklist / troubleshooting section on this page.</p>
<p><a style="color:#5377B4;text-decoration:none;" href="https://www.ecommercetemplates.com/updater_info.asp#checklist" target="_blank">https://www.ecommercetemplates.com/updater_info.asp#checklist</a></p>
<p>Please bookmark the above page so you can refer to it if you encounter any problems.</p>
<p>Please note that database script does not copy the updated scripts to your web. This must be done separately as detailed in the instructions.</p>
<p> </p>
<p>Please now delete this file from your web.</p>
</div>

<div style="padding:24px;border:1px solid #e0e0e0;width:80%;margin:0 auto;background:#FFF;-moz-border-radius:10px;-webkit-border-radius:10px;margin-top:40px;">
<p style="font-size: 20px;font-family : Arial,sans-serif;font-weight : normal;padding-top: 6px;color : #32427C;margin-top:0px;">ECT News</p>
<p>After you delete this file please take a look at our latest designs...</p>
<ul style="line-height:1.6;margin-left:10px;">
<li><a style="color:#5377B4;text-decoration:none;" href="https://www.ecommercetemplates.com/premium-responsive-design.asp" target="_blank">Premium Responsive Designs</a></li>
<li><a style="color:#5377B4;text-decoration:none;" href="https://www.ecommercetemplates.com/Generic-Version-Ecommerce-Plus" target="_blank">New Generic Version</a></li>
<li><a style="color:#5377B4;text-decoration:none;" href="https://www.ecommercetemplates.com/CSS-Premium-Layouts" target="_blank">Premium CSS Layouts</a></li>
</ul>
<p>We now also offer some related services...</p>
<ul style="line-height:1.6;margin-left:10px;">
<li>If you would like us to switch your site to one of our responsive or premium responsive designs we offer a <a style="color:#5377B4;text-decoration:none;" href="https://www.ecommercetemplates.com/Responsive-Design-Upgrade-Service" target="_blank">Responsive Design Upgrade Service</a>
with the price of the replacement software included in the package.</li>
<li>Of course you can do the integration yourself if you prefer in which case you would need the <a style="color:#5377B4;text-decoration:none;" href="https://www.ecommercetemplates.com/replacement-software.asp" target="_blank">Replacement Software</a></li>
<li>We can also add the Premium CSS Layouts to your existing site with the <a style="color:#5377B4;text-decoration:none;" href="https://www.ecommercetemplates.com/CSS-Layout-Service" target="_blank">CSS Layout Service</a></li>
<li>Check out all the tools and services on our <a style="color:#5377B4;text-decoration:none;" href="https://www.ecommercetemplates.com/ecommercetools.asp" target="_blank">Ecommerce Tools</a> page.</li>
</ul>
</div>
<%
else
	capturecardenabled=FALSE
	sSQL = "SELECT payProvEnabled FROM payprovider WHERE payProvID=10 AND payProvEnabled<>0"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then capturecardenabled=TRUE
	rs.close
	
	warncanadapostupdated=FALSE
	on error resume next
	err.number = 0
	sSQL = "SELECT adminID FROM admin WHERE adminShipping=6 OR adminIntShipping=6"
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then warncanadapostupdated = TRUE
	rs.Close
	on error goto 0
	if errnum<>0 then warncanadapostupdated = FALSE
	
	on error resume next
	err.number = 0
	sSQL = "SELECT altrateid FROM alternaterates WHERE altrateid=6 AND (usealtmethod<>0 OR usealtmethodintl<>0)"
	rs.Open sSQL,cnn,0,1
	errnum=err.number
	if NOT rs.EOF then warncanadapostupdated = TRUE
	rs.Close
	on error goto 0
	if errnum<>0 then warncanadapostupdated = FALSE
	
	on error resume next
	err.number = 0
	sSQL = "SELECT adminCanPostPass FROM admin"
	rs.Open sSQL,cnn,0,1
	errnum=err.number
	rs.Close
	on error goto 0
	if errnum=0 then warncanadapostupdated = FALSE
%>
<script type="text/javascript">
/* <![CDATA[ */
function checkform(frm){
<%	if capturecardenabled then %>
	if(!document.getElementById("capturecardenabled").checked){
		alert("Support for Capture Card has now been disabled due to PA-DSS and PCI requirements. Proceeding will remove Capture Card functionality from your website. Please check the box to proceed if you are in agreement.");
		return(false);
	}
<%	end if
	if warncanadapostupdated then %>
	if(!document.getElementById("warncanadapostupdated").checked){
		alert("The Canada Post registration system has changed. If you are using Canada Post shipping rates you must re-register with Canada Post in the Ecommerce Plus Shipping Methods admin page. Please check the box to indicate you understand this.");
		return(false);
	}
<%	end if %>
	if(!document.getElementById("csscheckout").checked){
		alert("Please check the box to indicate you have added the CSS Checkout file to your site.");
		return(false);
	}
	if(!document.getElementById("jscheckout").checked){
		alert("Please check the box to indicate you have added the JS Checkout file to your site.");
		return(false);
	}
	return(true);
}
/* ]]> */
</script>
<form action="updatestore.asp" method="post" onsubmit="return checkform(this)">
<input type="hidden" name="posted" value="1">
<table width="100%">
<tr><td width="100%">
<table width="100%">
<tr><td width="100%">
<p style="font-size: 20px;font-family : Arial,sans-serif;font-weight : normal;padding-top: 6px;color : #2F3D6F;margin-top:0px;">Version <%=sVersion%> Ecommerce Templates ASP Updater</p>
<p>Please note that clicking the button below will update your database to the current version. However it will not copy the updated scripts to your web. This must be done separately as detailed in the instructions.</p>
<p>Please make sure you have backed up your site and database before proceeding.</p>
<p>After performing the upgrade, please delete this file from your web.</p>
<%	sqlserversupported=TRUE
	if sqlserver=TRUE then
		on error resume next
		err.number = 0
		sSQL = "SELECT SERVERPROPERTY('productversion') as prodversion"
		rs.open sSQL,cnn,0,1
		errnum=err.number
		if NOT rs.EOF then prodversion = rs("prodversion")
		rs.close
		if errnum=0 then
		end if
		pverarr=split(prodversion,".")
		if int(pverarr(0))<9 then sqlserversupported=FALSE
		on error goto 0
	end if
	if sqlserversupported=FALSE then %>
		<p style="color:#FF0000">It seems that you are using SQL Server 2000. We are afraid that as this database version is no longer supported by Microsoft we are unable to support provide compatibility with this product. Please update your SQL Server database to a supported version.</p>
<%	else
%>
<div style="padding:24px;border:1px solid #e0e0e0;margin-top:10px;">
<p>The Ecommerce Templates cart checkout is now fully CSS based. If you have not already included the CSS file necessary for formatting it will be very difficult for your customers to navigate the checkout section and other parts of your site. More details are available here...<br />
<a href="https://www.ecommercetemplates.com/support/topic.asp?TOPIC_ID=107040" target="_blank">https://www.ecommercetemplates.com/support/topic.asp?TOPIC_ID=107040</a>
</p>
<p>The Ecommerce Templates cart checkout now uses an external javascript file, js/ectcart.js. If you have not already included the JS file necessary your customers will not be able to add products to cart or checkout. More details are available here...<br />
<a href="https://www.ecommercetemplates.com/support/topic.asp?TOPIC_ID=107040" target="_blank">https://www.ecommercetemplates.com/support/topic.asp?TOPIC_ID=107040</a>
</p>
<p><span style="color:red">IMPORTANT NOTE:</span> The ECT Cart CSS and JS Files change between versions to take account of the new classes for any new features. When updating, you should also update the css file included in this updater in the /css directory and the javascript file in the /js directory.</p>
<p><input type="checkbox" id="csscheckout" />Please check here to indicate you have added this CSS file to your site.</p>
<p><input type="checkbox" id="jscheckout" />Please check here to indicate you have added this JS file to your site.</p>
</div>
<%		if capturecardenabled then %>
<p><input type="checkbox" id="capturecardenabled" /> Support for Capture Card has now been disabled due to PA-DSS and PCI requirements. Proceeding will remove Capture Card functionality from your website. Please check the box to proceed if you are in agreement.</p>
<%		end if
		if warncanadapostupdated then %>
<p><input type="checkbox" id="warncanadapostupdated" /> The Canada Post registration system has changed. If you are using Canada Post shipping rates you must re-register with Canada Post in the Ecommerce Plus Shipping Methods admin page. Please check the box to indicate you understand this.</p>
<%		end if %>
<p>Please click below to start your upgrade.</p>
<p><input style="background:#0070ba;color:#fff;font-size:14px;cursor:pointer;padding:5px 10px;-moz-border-radius:10px;-webkit-border-radius:10px" type="submit" value="Upgrade to version <%=sVersion%>" /></p>
<%	end if %>
</td></tr>
</table>
</td></tr>
</table>
</form>
<%
end if

Set rs =nothing
Set rs2=nothing
Set rs3=nothing
cnn.close
Set cnn=nothing
%>
</div>
</body>
</html>
