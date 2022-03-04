<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
Dim outputvals()
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
success=true
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs2=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
thedate=date()
thedate=DateSerial(year(thedate),month(thedate),day(thedate))
fromdate=getpost("fromdate")
todate=getpost("todate")
hasfromdate=FALSE
hastodate=FALSE
if lcase(adminencoding)="iso-8859-1" then raquo="»" else raquo=">"
sub writemenulevel(id,itlevel)
	dim wmlindex
	if itlevel<10 then
		for wmlindex=0 TO ubound(alldata,2)
			if alldata(2,wmlindex)=id then
				print "<option value='"&alldata(0,wmlindex)&"'"
				if IsArray(selstatus) then
					for ii=0 to UBOUND(selstatus)
						if Int(selstatus(ii))=alldata(0,wmlindex) then print " selected=""selected"""
					next
				end if
				print ">"
				for index=0 to itlevel-2
					print raquo & " "
				next
				print alldata(1,wmlindex)&"</option>" & vbCrLf
				if alldata(3,wmlindex)=0 then call writemenulevel(alldata(0,wmlindex),itlevel+1)
			end if
		next
	end if
end sub
function getdatesql(datecol)
	getdatesql=""
	if NOT (hasfromdate OR hastodate) then
		' nothing
	elseif hasfromdate AND hastodate then
		getdatesql=" AND " & datecol & " BETWEEN " & vsusdate(thefromdate) & " AND " & vsusdatetime(dateadd("s",-1,thetodate))
	elseif hasfromdate then
		getdatesql=" AND " & datecol & " BETWEEN " & vsusdate(thefromdate) & " AND " & vsusdatetime(dateadd("s",-1,thefromdate+1))
	end if
end function
if fromdate<>"" then
	hasfromdate=TRUE
	if is_numeric(fromdate) then
		thefromdate=(thedate-fromdate)
	else
		err.number=0
		on error resume next
		thefromdate=DateValue(fromdate)
		if err.number <> 0 then
			thefromdate=thedate
			success=false
			errmsg=yyDatInv & " - " & fromdate
		end if
		on error goto 0
	end if
	hastodate=TRUE
	if todate="" then
		hastodate=FALSE
	elseif is_numeric(todate) then
		thetodate=(thedate-todate)
	else
		err.number=0
		on error resume next
		thetodate=DateValue(todate)
		if err.number <> 0 then
			thetodate=thedate
			success=false
			errmsg=yyDatInv & " - " & todate
		end if
		on error goto 0
	end if
	if hasfromdate AND hastodate then
		if thefromdate > thetodate then
			tmpdate=thetodate
			thetodate=thefromdate
			thefromdate=tmpdate

			tmpdate=hasfromdate
			hasfromdate=hastodate
			hastodate=tmpdate
		end if
		thetodate=thetodate + 1
	end if
else
	thefromdate=Date()-365
	thetodate=Date()
end if
sSQL="SELECT statID,statPrivate FROM orderstatus WHERE statPrivate<>'' ORDER BY statID"
rs.open sSQL,cnn,0,1
	allstatus=rs.GetRows
rs.close
themask=cStr(DateSerial(2003,12,11))
themask=replace(themask,"2003","yyyy")
themask=replace(themask,"12","mm")
themask=replace(themask,"11","dd")
'for each objItem in request.form
'	print objItem & ": " & request.form(objItem) & "<br>"
'next
ordstate=getpost("ordstate")
ordcountry=getpost("ordcountry")
ordstatus=getpost("ordstatus")
thecat=getpost("scat")
payprovider=getpost("payprovider")
stext=getpost("stext")
stsearch=getpost("stsearch")
%>
<script src="popcalendar.js"></script>
<script>try{languagetext('<%=adminlang%>');}catch(err){}</script>
<h2><%=yySalRep%></h2>
		  <form method="post" action="adminorders.asp" name="psearchform">
			<div class="orderssearch">
			  <div class="orderssearchrow">
                <div style="text-align:right;width:25%"><input type="button" onclick="popUpCalendar(this, document.forms.psearchform.fromdate, '<%=themask%>', 0)" value="<%=yyOrdFro%>" /></div>
				<div style="width:25%"><div style="position:relative;display:inline"><input type="text" class="orddatesel" size="14" name="fromdate" value="<%=fromdate%>" style="vertical-align:middle" /> <input type="button" onclick="document.forms.psearchform.fromdate.value='<%=thedate%>'" value="<%=yyToday%>" /></div></div>
				<div style="text-align:right;width:25%"><input type="button" onclick="popUpCalendar(this, document.forms.psearchform.todate, '<%=themask%>', -205)" value="<%=yyOrdTil%>" /></div>
				<div style="width:25%"><div style="position:relative;display:inline"><input type="text" class="orddatesel" size="14" name="todate" value="<%=todate%>" style="vertical-align:middle" /> <input type="button" onclick="document.forms.psearchform.todate.value='<%=thedate%>'" value="<%=yyToday%>" /></div></div>
			  </div>
			  <div class="orderssearchrow">
				<div style="text-align:center"><input type="checkbox" name="notsection" value="ON" title="<%=yyNot%>" style="vertical-align:middle" <% if getpost("notsection")="ON" then print "checked=""checked"" "%>/> <%=yySection%></div>
				<div style="text-align:center"><input type="checkbox" name="notstatus" value="ON" title="<%=yyNot%>" style="vertical-align:middle" <% if getpost("notstatus")="ON" then print "checked=""checked"" "%>/> <%=yyOrdSta%></div>
				<div style="text-align:center"><input type="checkbox" name="notstate" value="ON" title="<%=yyNot%>" style="vertical-align:middle" <% if getpost("notstate")="ON" then print "checked=""checked"" "%>/> <%=yyState%></div>
				<div style="text-align:center"><input type="checkbox" name="notcountry" value="ON" title="<%=yyNot%>" style="vertical-align:middle" <% if getpost("notcountry")="ON" then print "checked=""checked"" "%>/> <%=yyCountry%></div>
			  </div>
			  <div class="orderssearchrow">
				<div style="text-align:center"><select name="scat" size="5" multiple="multiple" style="max-width:200px"><%
						sSQL="SELECT sectionID,sectionWorkingName,topSection,rootSection FROM sections " & IIfVr(adminonlysubcats, "WHERE rootSection=1 ORDER BY sectionWorkingName", "ORDER BY sectionOrder")
						rs.open sSQL,cnn,0,1
						if rs.eof then
							success=false
						else
							alldata=rs.getrows
							success=true
						end if
						rs.close
						if thecat<>"" then selstatus=Split(thecat, ",") else selstatus=""
						if IsArray(alldata) then
							if adminonlysubcats=true then
								for rowcounter=0 to UBOUND(alldata,2)
									print "<option value='"&alldata(0,rowcounter)&"'"
									if IsArray(selstatus) then
										for ii=0 to UBOUND(selstatus)
											if Int(selstatus(ii))=alldata(0,rowcounter) then print " selected=""selected"""
										next
									end if
									print ">"&alldata(1,rowcounter)&"</option>" &vbCrLf
								next
							else
								call writemenulevel(0,1)
							end if
						end if %>
					  </select></div>
				<div style="text-align:center"><select name="ordstatus" size="5" multiple="multiple" style="max-width:200px"><%
						if ordstatus<>"" then selstatus=Split(ordstatus, ",") else selstatus=""
						for index=0 to UBOUND(allstatus,2)
							print "<option value=""" & allstatus(0,index) & """"
							if IsArray(selstatus) then
								for ii=0 to UBOUND(selstatus)
									if Int(selstatus(ii))=Int(allstatus(0,index)) then print " selected=""selected"""
								next
							end if
							print ">" & allstatus(1,index) & "</option>"
						next %></select></div>
				<div style="text-align:center"><select name="ordstate" size="5" multiple="multiple" style="max-width:200px"><%
						sSQL="SELECT stateID,stateName,stateAbbrev FROM states WHERE stateCountryID=" & origCountryID & " AND stateEnabled=1 ORDER BY stateName"
						rs.open sSQL,cnn,0,1
						if ordstate<>"" then selstatus=Split(ordstate, ",") else selstatus=""
						do while NOT rs.EOF
							print "<option value=""" & IIfVr(usestateabbrev,rs("stateAbbrev"),rs("stateName")) & """"
							if IsArray(selstatus) then
								for ii=0 to UBOUND(selstatus)
									if trim(selstatus(ii))=IIfVr(usestateabbrev,rs("stateAbbrev"),rs("stateName")) then print " selected=""selected"""
								next
							end if
							print ">" & rs("stateName") & "</option>"
							rs.MoveNext
						loop
						rs.close %></select></div>
				<div style="text-align:center"><select name="ordcountry" size="5" multiple="multiple" style="max-width:200px"><%
						sSQL="SELECT countryID,countryName FROM countries WHERE countryEnabled=1 ORDER BY countryOrder DESC, countryName"
						rs.open sSQL,cnn,0,1
						if ordcountry<>"" then selstatus=Split(ordcountry, ",") else selstatus=""
						do while NOT rs.EOF
							print "<option value=""" & rs("countryName") & """"
							if IsArray(selstatus) then
								for ii=0 to UBOUND(selstatus)
									if trim(selstatus(ii))=rs("countryName") then print " selected=""selected"""
								next
							end if
							print ">" & rs("countryName") & "</option>"
							rs.MoveNext
						loop
						rs.close %></select></div>
			  </div>
			  <div class="orderssearchrow">
				<div style="text-align:center">
						<div style="border-bottom:1px solid #D2DDEE;padding:2px">
							<div style="text-align:center"><input type="checkbox" name="notpayprov" value="ON" title="<%=yyNot%>" style="vertical-align:middle" <% if getpost("notpayprov")="ON" then print "checked "%>/> <%=yyPayMet%>:</div>
						</div>
						<div style="padding:2px">
							<div style="text-align:center"><select name="payprovider" size="5" multiple="multiple" style="max-width:200px"><%
							sSQL="SELECT payProvID,payProvName FROM payprovider WHERE payProvEnabled=1 ORDER BY payProvOrder"
							rs.open sSQL,cnn,0,1
							if payprovider<>"" then selstatus=Split(payprovider, ",") else selstatus=""
							do while NOT rs.EOF
								print "<option value=""" & rs("payProvID") & """"
								if IsArray(selstatus) then
									for ii=0 to UBOUND(selstatus)
										if int(selstatus(ii))=rs("payProvID") then print " selected=""selected"""
									next
								end if
								print ">" & rs("payProvName") & "</option>"
								rs.MoveNext
							loop
							rs.close %></select></div>
						</div>
				</div>
				<div style="text-align:center;vertical-align:middle">
					<div style="text-align:center;padding:5px"><input type="text" style="width:180px" name="stext" value="<%=htmlspecials(stext)%>" placeholder="<%=yySeaTxt%>" /></div>
					<div style="text-align:center;padding:5px"><select name="stype" size="1" style="width:190px">
						<option value=""><%=yySearch&" "&yySrchAl%></option>
						<option value="any" <% if request("stype")="any" then print "selected=""selected"""%>><%=yySearch&" "&yySrchAn%></option>
						<option value="exact" <% if request("stype")="exact" then print "selected=""selected"""%>><%=yySearch&" "&yySrchEx%></option>
						</select></div> 
					<div style="text-align:center;padding:5px">
				<input type="checkbox" name="stsearch" value="cartprodid" title="<%=yyPrId%>" <% if instr(stsearch, "cartprodid")>0 then print "checked=""checked"" "%>/> <%=yyID%> 
				<input type="checkbox" name="stsearch" value="cartprodname" title="<%=yyPrName%>" <% if instr(stsearch, "cartprodname")>0 then print "checked=""checked"" "%>/> <%=yyPrNam%> 
				<input type="checkbox" name="stsearch" value="ordaffiliate" title="<%=yyAffili%>" <% if instr(stsearch, "ordaffiliate")>0 then print "checked=""checked"" "%>/> <%=yyAffili%>
					</div>
				</div>
				<div style="text-align:center;vertical-align:middle">
					<div style="text-align:center;padding:5px">Display</div>
					<div style="text-align:center;padding:5px">
						<select name="numresults" id="numresults" size="1" style="vertical-align:middle">
						<option value="">Top 100 Sales</option>
						<option value="200" <% if getpost("numresults")="200" then print "selected=""selected"""%>>Top 200 Sales</option>
						<option value="300" <% if getpost("numresults")="300" then print "selected=""selected"""%>>Top 300 Sales</option>
						<option value="400" <% if getpost("numresults")="400" then print "selected=""selected"""%>>Top 400 Sales</option>
						<option value="500" <% if getpost("numresults")="500" then print "selected=""selected"""%>>Top 500 Sales</option>
						<option value="all" <% if getpost("numresults")="all" then print "selected=""selected"""%>>All Sales</option></select>
					</div>
				</div>
				<div style="text-align:center;vertical-align:middle">
					<div style="text-align:center;padding:5px">Generate Results</div>
					<div style="text-align:center;padding:5px">
						<select name="grouping" size="1" style="vertical-align:middle" onchange="document.getElementById('numresults').disabled=this.selectedIndex!=0">
						<option value="4">Totals</option>
						<option value="5" <% if getpost("grouping")="5" then print "selected=""selected"""%>><%="Graph By Day"%></option>
						<option value="1" <% if getpost("grouping")="1" then print "selected=""selected"""%>><%=yyGrByWk%></option>
						<option value="2" <% if getpost("grouping")="2" then print "selected=""selected"""%>><%=yyGrByMo%></option>
						<option value="3" <% if getpost("grouping")="3" then print "selected=""selected"""%>><%=yyGrByYr%></option></select>
						<input type="button" value="Stats" onclick="document.forms.psearchform.action='adminstats.asp';document.forms.psearchform.submit();" />
					</div>
				</div>
			  </div>
			</div>
		  </form>
<%
hascategory=FALSE
whereclause="WHERE cartCompleted=1 "
if ordstatus<>"" then whereclause=whereclause & "AND " & IIfVs(getpost("notstatus")="ON","NOT ") & "(ordStatus IN ("&ordstatus&")) " else whereclause=whereclause & "AND ordStatus<>0 AND ordStatus<>1 "
if payprovider<>"" then whereclause=whereclause & "AND " & IIfVs(getpost("notpayprov")="ON","NOT ") & "(ordPayProvider IN ("&payprovider&")) "
if ordstate<>"" then whereclause=whereclause & "AND " & IIfVs(getpost("notstate")="ON","NOT ") & "(ordState IN ('" & replace(replace(ordstate,", ",","),",","','") & "')) "
if ordcountry<>"" then whereclause=whereclause & "AND " & IIfVs(getpost("notcountry")="ON","NOT ") & "(ordCountry IN ('" & replace(replace(ordcountry,", ",","),",","','") & "')) "
orderclause=replace(whereclause, "cartCompleted=1 AND ", "")
if getpost("stext")<>"" AND getpost("stsearch")<>"" then
	sText=escape_string(stext)
	aText=Split(sText)
	aFields= split(getpost("stsearch"),",")
	if request("stype")="exact" then
		whereclause=whereclause & "AND ("
		for index=0 to UBOUND(afields)
			whereclause=whereclause & aFields(index) & " LIKE '%"&sText&"%' "
			if aFields(index)="ordaffiliate" then orderclause=orderclause & " AND " & aFields(index) & " LIKE '%"&sText&"%' "
			if index < UBOUND(afields) then whereclause=whereclause & "OR "
		next
		whereclause=whereclause & ") "
	else
		sJoin="AND "
		if request("stype")="any" then sJoin="OR "
		whereclause=whereclause & "AND ("
		whereand=" AND "
		for index=0 to UBOUND(afields)
			whereclause=whereclause & "("
			for rowcounter=0 to UBOUND(aText)
				whereclause=whereclause & aFields(index) & " LIKE '%"&aText(rowcounter)&"%' "
				if aFields(index)="ordaffiliate" then orderclause=orderclause & " AND " & aFields(index) & " LIKE '%"&sText&"%' "
				if rowcounter<UBOUND(aText) then whereclause=whereclause & sJoin
			next
			whereclause=whereclause & ") "
			if index < UBOUND(afields) then whereclause=whereclause & "OR "
		next
		whereclause=whereclause & ") "
	end if
end if
if thecat<>"" then
	sectionids=getsectionids(thecat, TRUE)
	if sectionids<>"" then whereclause=whereclause & "AND " & IIfVs(getpost("notsection")="ON","NOT ") & "(products.pSection IN (" & sectionids & ")) "
	hascategory=TRUE
end if
if getpost("grouping")="1" OR getpost("grouping")="2" OR getpost("grouping")="3" OR getpost("grouping")="5" then
	success=TRUE
	dateSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 1 ")&"ordDate "
	if mysqlserver then
		dateSQL=dateSQL & "FROM products INNER JOIN cart ON products.pID=cart.cartProdID INNER JOIN orders ON cart.cartOrderID=orders.ordID "
	else
		dateSQL=dateSQL & "FROM products INNER JOIN (cart INNER JOIN orders ON cart.cartOrderID=orders.ordID) ON products.pID=cart.cartProdID "
	end if
	dateSQL=dateSQL & whereclause
	rs.open dateSQL & "ORDER BY ordDate"&IIfVs(mysqlserver=TRUE," LIMIT 0,1"),cnn,0,1
	if NOT rs.EOF then
		minfromdate=rs("ordDate")
		if NOT hasfromdate OR minfromdate > thefromdate then
			thefromdate=minfromdate
			hasfromdate=TRUE
		end if
	else
		success=FALSE
	end if
	rs.close
	rs.open dateSQL & "ORDER BY ordDate DESC"&IIfVs(mysqlserver=TRUE," LIMIT 0,1"),cnn,0,1
	if NOT rs.EOF then
		maxtodate=rs("ordDate")
		if NOT hastodate OR maxtodate < thetodate then
			thetodate=maxtodate
			hastodate=true
		end if
	else
		success=FALSE
	end if
	rs.close
	' print "Dates: " & thefromdate & ", " & thetodate & "<br>"
	if success then
		if getpost("grouping")="1" then ' week
			thefromdate=dateserial(DatePart("yyyy",thefromdate), DatePart("m",thefromdate), DatePart("d",thefromdate))
			thefromdate=dateadd("d",0-(datepart("w",thefromdate,2)-1),thefromdate) ' round to beginning of week
		elseif getpost("grouping")="2" then ' month
			thefromdate=dateserial(DatePart("yyyy",thefromdate), DatePart("m",thefromdate), 1)
		elseif getpost("grouping")="5" then ' day
			thefromdate=dateserial(DatePart("yyyy",thefromdate), DatePart("m",thefromdate), DatePart("d",thefromdate))
		else
			thefromdate=dateserial(DatePart("yyyy",thefromdate), 1, 1)
		end if
		thetodate=dateadd("d", 1, dateserial(DatePart("yyyy",thetodate), DatePart("m",thetodate), DatePart("d",thetodate)))
	end if
	' print "Dates: " & thefromdate & ", " & thetodate & "<br>"
	if thefromdate > thetodate then success=FALSE

	print "<div style=""margin:30px""><h2>"&yySalGra&"</h2></div>"
	if NOT success then
		print "<div style=""margin:30px;text-align:center"">Empty Result Set</div>"
	else
		redim outputvals(2,100)
		maxtotal=0
		maxorders=0
		rowcounter=0
		' sSQL="SELECT SUM(cartQuantity) AS numorders,SUM(cartProdPrice*cartQuantity) AS theordtot,SUM(ordHandling) AS tothandling,SUM(ordStateTax) AS totstatetax,SUM(ordCountryTax) AS totcountrytax,SUM(ordHSTTax) AS tothsttax,SUM(ordDiscount) AS totdiscount, SUM(ordShipping) AS totshipping "
		sSQL="SELECT SUM(cartQuantity) AS numorders,SUM(cartProdPrice*cartQuantity) AS theordtot "
		if mysqlserver then
			sSQL=sSQL & "FROM products RIGHT JOIN cart ON products.pID=cart.cartProdID INNER JOIN orders ON cart.cartOrderID=orders.ordID "
		else
			sSQL=sSQL & "FROM products RIGHT JOIN (cart INNER JOIN orders ON cart.cartOrderID=orders.ordID) ON products.pID=cart.cartProdID "
		end if
		sSQL=sSQL & whereclause
		
		sSQLopts="SELECT SUM(coPriceDiff*cartQuantity) AS theordtot "
		if mysqlserver then
			sSQLopts=sSQLopts & "FROM cartoptions INNER JOIN (cart LEFT OUTER JOIN products ON cart.cartProdId=products.pID) ON cartoptions.coCartID=cart.cartID INNER JOIN orders ON cart.cartOrderID=orders.ordID "
		else
			sSQLopts=sSQLopts & "FROM cartoptions INNER JOIN ((cart LEFT OUTER JOIN products ON cart.cartProdId=products.pID) INNER JOIN orders ON cart.cartOrderID=orders.ordID) ON cartoptions.coCartID=cart.cartID "
		end if
		sSQLopts=sSQLopts & whereclause
		' print "<div>"&sSQL&"<br />"&sSQLopts&"</div>" : response.flush
		themaxdate=thetodate
		do while thefromdate<themaxdate
			if getpost("grouping")="1" then ' week
				thetodate=dateadd("ww", 1, thefromdate)
			elseif getpost("grouping")="2" then ' month
				thetodate=dateadd("m", 1, thefromdate)
			elseif getpost("grouping")="5" then ' day
				thetodate=dateadd("d", 1, thefromdate)
			else
				thetodate=dateadd("yyyy", 1, thefromdate)
			end if
			rs.open sSQL & getdatesql("cartDateAdded"),cnn,0,1
			if NOT rs.EOF then
				if isnull(rs("numorders")) then
					outputvals(0, rowcounter)=thefromdate
					outputvals(1, rowcounter)=0
					outputvals(2, rowcounter)=0
				elseif clng(rs("numorders"))=0 then
					outputvals(0, rowcounter)=thefromdate
					outputvals(1, rowcounter)=0
					outputvals(2, rowcounter)=0
				else
					'thetot=(rs("theordtot")+rs("totshipping")+rs("tothandling")+rs("totstatetax")+rs("totcountrytax")+rs("tothsttax"))-rs("totdiscount")
					thetot=rs("theordtot")
					outputvals(0, rowcounter)=thefromdate
					outputvals(1, rowcounter)=rs("numorders")
					outputvals(2, rowcounter)=IIfVr(isnull(thetot), 0, thetot)
				end if
			end if
			rs.close

			rs.open sSQLopts & getdatesql("cartDateAdded"),cnn,0,1
			if NOT rs.EOF then
				if isnull(rs("theordtot")) then
				' elseif cint(rs("theordtot"))=0 then
				else
					thetot=rs("theordtot")
					outputvals(2, rowcounter)=outputvals(2, rowcounter) + IIfVr(isnull(thetot), 0, thetot)
				end if
			end if
			rs.close

			if outputvals(1, rowcounter) > maxorders then maxorders=outputvals(1, rowcounter)
			if outputvals(2, rowcounter) > maxtotal then maxtotal=outputvals(2, rowcounter)
			thefromdate=thetodate
			rowcounter=rowcounter+1
			if rowcounter >= UBOUND(outputvals, 2) then redim preserve outputvals(2, UBOUND(outputvals, 2) + 100)
		loop

		if getpost("grouping")="2" then
			timediv="Monthly"
			for index=0 to rowcounter-1
				outputvals(0,index)=monthname(datepart("m",outputvals(0,index)),TRUE)&" "&datepart("yyyy",outputvals(0,index))
			next
		elseif getpost("grouping")="1" then
			timediv="Weekly"
		elseif getpost("grouping")="5" then
			timediv="Daily"
		else
			timediv="Yearly"
			for index=0 to rowcounter-1
				outputvals(0,index)=datepart("yyyy",outputvals(0,index))
			next
		end if %>
			<div style="padding-top:20px">
<%	saveLCID=session.LCID
	origmax=maxtotal
	divisor=1
	if maxtotal>0 then
		do while maxtotal/divisor>10
			divisor=divisor*10
		loop
	end if
	if maxtotal<>int(maxtotal/divisor)*divisor then maxtotal=int((maxtotal/divisor)+1)*divisor
	if maxtotal/divisor<3 then divisor=divisor/2
	do while maxtotal-divisor>origmax
		maxtotal=maxtotal-divisor
	loop
	if maxtotal/divisor<3 then divisor=divisor/2
%>
<svg id="svgelement" style="margin:auto;display:block;font-family:Arial,Helvetica,sans-serif;font-size:12px;width:90%;height:400px" viewBox="0 0 100 100" preserveAspectRatio="none">
	<text x="10" y="5" style="display:none;font-size:0.4em"><%=timediv&" Sales: " & outputvals(0,0) & " to " & outputvals(0,rowcounter-1)%></text>
	<text x="10" y="95" style="display:none;font-size:0.2em;text-anchor:middle"><%=outputvals(0,0)%></text>
	<text x="90" y="95" style="display:none;font-size:0.2em;text-anchor:middle"><%=outputvals(0,rowcounter-1)%></text>
<%
	tempdivisor=divisor
	do while tempdivisor<=maxtotal
		thisy=10+int((tempdivisor / maxtotal) * 80)
		print "<text x=""8.5"" y="""&(101-thisy)&""" style=""display:none;font-size:0.2em;text-anchor:end"">"&FormatCurrencyZeroDP(tempdivisor)&"</text>"&vbCrLf
		tempdivisor=tempdivisor+divisor
	loop
%>
	<g transform="translate(0,100) scale(1,-1)">
		<rect stroke="#c0c8d0" fill="none" x="10" y="10" width="80" height="80" vector-effect="non-scaling-stroke" />
<%		markers=""
		print "<polygon style=""fill:#506070"" points=""10,10 "
		onlyyears=rowcounter>100
		for index=0 to rowcounter-1
			pixelcolor="c0c8d0"
			isyear=FALSE
			if rowcounter=1 then thisx=10 else thisx=vsround(10+((index/(rowcounter-1))*80),2)
			if getpost("grouping")="1" then ' week
				if datepart("m", outputvals(0, index))=1 AND datepart("d", outputvals(0, index))<8 then pixelcolor="FF6070" : isyear=TRUE
			elseif getpost("grouping")="2" then ' month
				if datepart("m", outputvals(0, index))=1 then pixelcolor="FF6070" : isyear=TRUE
			end if
			session.LCID=1033
			print thisx&","&(10+int((outputvals(2, index) / maxtotal) * 80))&" "
			session.LCID=saveLCID
			if index<>0 AND index<>rowcounter-1 then
				if isyear OR onlyyears=FALSE then markers=markers&"<path style=""stroke:#"&pixelcolor&";vector-effect:non-scaling-stroke"" d=""M"&thisx&",8 L"&thisx&",10"" />"&vbCrLf
			end if
		next
		session.LCID=1033
		if rowcounter=1 then print "90,"&(10+int((outputvals(2,0) / maxtotal) * 80))&" 90" else print vsround(10+(((index-1)/(rowcounter-1))*80),2)
		session.LCID=saveLCID
		print ",10"" />" & vbCrLf
		print markers
		tempdivisor=divisor
		do while tempdivisor<=maxtotal
			thisy=10+int((tempdivisor / maxtotal) * 80)
			print "<path style=""stroke:#c0c8d0;vector-effect:non-scaling-stroke"" d=""M9,"&thisy&" L10,"&thisy&""" />"&vbCrLf
			tempdivisor=tempdivisor+divisor
		loop
%>
	</g>
</svg>
			</div>
			
			<div style="padding-top:20px">
<%	origmax=maxorders
	divisor=1
	if maxorders>0 then
		do while maxorders/divisor>10
			divisor=divisor*10
		loop
	end if
	if maxorders<>int(maxorders/divisor)*divisor then maxorders=int((maxorders/divisor)+1)*divisor
	if maxorders/divisor<3 then divisor=divisor/2
	do while maxorders-divisor>origmax
		maxorders=maxorders-divisor
	loop
	if maxorders/divisor<3 then divisor=divisor/2
%>
<svg id="svgelementorders" style="margin:auto;display:block;font-family:Arial,Helvetica,sans-serif;font-size:12px;width:90%;height:400px" viewBox="0 0 100 100" preserveAspectRatio="none">
	<text x="10" y="5" style="display:none;font-size:0.4em"><%=timediv&" Orders: " & outputvals(0,0) & " to " & outputvals(0,rowcounter-1)%></text>
	<text x="10" y="95" style="display:none;font-size:0.2em;text-anchor:middle"><%=outputvals(0,0)%></text>
	<text x="90" y="95" style="display:none;font-size:0.2em;text-anchor:middle"><%=outputvals(0,rowcounter-1)%></text>
<%
	tempdivisor=divisor
	do while tempdivisor<=maxorders
		thisy=10+int((tempdivisor / maxorders) * 80)
		print "<text x=""8.5"" y="""&(101-thisy)&""" style=""display:none;font-size:0.2em;text-anchor:end"">"&tempdivisor&"</text>"&vbCrLf
		tempdivisor=tempdivisor+divisor
	loop
%>
	<g transform="translate(0,100) scale(1,-1)">
		<rect stroke="#c0c8d0" fill="none" x="10" y="10" width="80" height="80" vector-effect="non-scaling-stroke" />
<%		markers=""
		print "<polygon style=""fill:#506070"" points=""10,10 "
		onlyyears=rowcounter>100
		for index=0 to rowcounter-1
			pixelcolor="c0c8d0"
			isyear=FALSE
			if rowcounter=1 then thisx=10 else thisx=vsround(10+((index/(rowcounter-1))*80),2)
			if getpost("grouping")="1" then ' week
				if datepart("m", outputvals(0, index))=1 AND datepart("d", outputvals(0, index))<8 then pixelcolor="FF6070" : isyear=TRUE
			elseif getpost("grouping")="2" then ' month
				if datepart("m", outputvals(0, index))=1 then pixelcolor="FF6070" : isyear=TRUE
			end if
			session.LCID=1033
			print thisx&","&(10+int((outputvals(1, index) / maxorders) * 80))&" "
			session.LCID=saveLCID
			if index<>0 AND index<>rowcounter-1 then
				if isyear OR onlyyears=FALSE then markers=markers&"<path style=""stroke:#"&pixelcolor&";vector-effect:non-scaling-stroke"" d=""M"&thisx&",8 L"&thisx&",10"" />"&vbCrLf
			end if
		next
		session.LCID=1033
		if rowcounter=1 then print "90,"&(10+int((outputvals(1,0) / maxorders) * 80))&" 90" else print vsround(10+(((index-1)/(rowcounter-1))*80),2)
		session.LCID=saveLCID
		print ",10"" />" & vbCrLf
		print markers
		tempdivisor=divisor
		do while tempdivisor<=maxorders
			thisy=10+int((tempdivisor / maxorders) * 80)
			print "<path style=""stroke:#c0c8d0;vector-effect:non-scaling-stroke"" d=""M9,"&thisy&" L10,"&thisy&""" />"&vbCrLf
			tempdivisor=tempdivisor+divisor
		loop
%>
	</g>
</svg>
			</div>
			
<script>
window.onresize = svgresized;
function svgresized(){
	var svgelem=document.getElementById("svgelement");
    var boundingwidth=svgelem.getBoundingClientRect().width;
	var boundingheight=svgelem.getBoundingClientRect().height;
	var xscale=400.0 / boundingwidth;
	var textelems=svgelem.getElementsByTagName("text");
	for (var svgi=0; svgi<textelems.length; svgi++) {
		var txtxpos=textelems[svgi].getAttribute('x');
		textelems[svgi].setAttribute('transform',"scale("+xscale+",1) translate("+(txtxpos*((1/xscale)-1))+",0)");
		textelems[svgi].style.display='';
	}
	
	var svgelem=document.getElementById("svgelementorders");
    var boundingwidth=svgelem.getBoundingClientRect().width;
	var boundingheight=svgelem.getBoundingClientRect().height;
	var xscale=400.0 / boundingwidth;
	var textelems=svgelem.getElementsByTagName("text");
	for (var svgi=0; svgi<textelems.length; svgi++) {
		var txtxpos=textelems[svgi].getAttribute('x');
		textelems[svgi].setAttribute('transform',"scale("+xscale+",1) translate("+(txtxpos*((1/xscale)-1))+",0)");
		textelems[svgi].style.display='';
	}
}
svgresized();
</script>
<%
	end if
elseif getpost("grouping")="4" then
	sub displayemptyresults()
		print "<div class=""stattblrow"">" & _
				"<div style=""margin:30px;text-align:center;padding:50px"">Empty Result Set</div>" & _
			"</div>"
		hasresults=FALSE
	end sub
	hasresults=TRUE
	whereclause=whereclause & getdatesql("cartDateAdded")
	orderclause=orderclause & getdatesql("ordDate")
%>
            <div style="padding-top:20px">
			  <h2>Order Results (Not limited by product / section)</h2>
<%
	sSQL="SELECT COUNT(ordID) AS numorders,SUM(ordTotal) AS theordtot,SUM(ordHandling) AS tothandling,SUM(ordStateTax) AS totstatetax,SUM(ordCountryTax) AS totcountrytax,SUM(ordHSTTax) AS tothsttax,SUM(ordDiscount) AS totdiscount, SUM(ordShipping) AS totshipping "
	sSQL=sSQL & "FROM orders "
	sSQL=sSQL & orderclause
	'print "<div>"&sSQL&"</div>"
	rs.open sSQL,cnn,0,1
	if rs.EOF then
		hasresults=FALSE
	else
		print "<div class=""stattbl"">"
		if rs("numorders")=0 then
			hasresults=FALSE
			call displayemptyresults()
		else
			print "<div class=""stattblrow"">" & _
					"<div>"&yyTotOrd&"</div><div>"&xxOrdTot&"</div><div>"&xxShippg&"</div><div>"&xxHndlg&"</div><div>"&xxDscnts&"</div><div>"&xxStaTax&"</div>" & IIfVs(origCountryID=2,"<div>"&xxHST&"</div>") & "<div>"&xxCntTax&"</div><div>"&xxGndTot&"</div>" & _
				"</div>"
			print "<div class=""stattblrow"">" & _
					"<div>" & rs("numorders") & "</div><div>" & FormatEuroCurrency(rs("theordtot")) & "</div><div>" & FormatEuroCurrency(rs("totshipping")) & "</div><div>" & FormatEuroCurrency(rs("tothandling")) & "</div><div>" & FormatEuroCurrency(rs("totdiscount")) & "</div><div>" & FormatEuroCurrency(rs("totstatetax")) & "</div>" & IIfVs(origCountryID=2,"<div>" & FormatEuroCurrency(rs("tothsttax")) & "</div>") & "<div>" & FormatEuroCurrency(rs("totcountrytax")) & "</div><div>" & FormatEuroCurrency((rs("theordtot")+rs("totshipping")+rs("tothandling")+rs("totstatetax")+rs("totcountrytax")+rs("tothsttax"))-rs("totdiscount")) & "</div>" & _
				"</div>"
		end if
		print "</div>"
%>
		<div style="margin-top:10px">
			<a download="ectorderresults.csv" href="data:application/octet-stream,<%
				print yyTotOrd&"%2C"&xxOrdTot&"%2C"&xxShippg&"%2C"&xxHndlg&"%2C"&xxDscnts&"%2C"&xxStaTax & IIfVs(origCountryID=2,"%2C"&xxHST) & "%2C"&xxCntTax&"%2C"&xxGndTot&"%0A"
				print rs("numorders") & "%2C" & vsround(rs("theordtot"),2) & "%2C" & vsround(rs("totshipping"),2) & "%2C" & vsround(rs("tothandling"),2) & "%2C" & vsround(rs("totdiscount"),2) & "%2C" & vsround(rs("totstatetax"),2) & IIfVs(origCountryID=2,"%2C" & vsround(rs("tothsttax"),2)) & "%2C" & vsround(rs("totcountrytax"),2) & "%2C" & vsround((rs("theordtot")+rs("totshipping")+rs("tothandling")+rs("totstatetax")+rs("totcountrytax")+rs("tothsttax"))-rs("totdiscount"),2)
			%>">Click Here For CSV Download</a>
		</div>
<%	end if
	rs.close
%>
			</div>
<%	response.flush 
	sSQL="SELECT SUM(cartQuantity) AS numorders,SUM(cartProdPrice*cartQuantity) AS theordtot "
	if hascategory then
		if mysqlserver then
			sSQL=sSQL & "FROM products INNER JOIN cart ON products.pID=cart.cartProdID INNER JOIN orders ON cart.cartOrderID=orders.ordID "
		else
			sSQL=sSQL & "FROM products INNER JOIN (cart INNER JOIN orders ON cart.cartOrderID=orders.ordID) ON products.pID=cart.cartProdID "
		end if
	else
		sSQL=sSQL & "FROM cart INNER JOIN orders ON cart.cartOrderID=orders.ordID "
	end if
	sSQL=sSQL & whereclause

	sSQLopts="SELECT SUM(coPriceDiff*cartQuantity) AS theordtot "
	if hascategory then
		if mysqlserver then
			sSQLopts=sSQLopts & "FROM cartoptions INNER JOIN cart ON cartoptions.coCartID=cart.cartID LEFT OUTER JOIN products ON cart.cartProdId=products.pID INNER JOIN orders ON cart.cartOrderID=orders.ordID "
		else
			sSQLopts=sSQLopts & "FROM cartoptions INNER JOIN ((cart LEFT OUTER JOIN products ON cart.cartProdId=products.pID) INNER JOIN orders ON cart.cartOrderID=orders.ordID) ON cartoptions.coCartID=cart.cartID "
		end if
	else
		sSQLopts=sSQLopts & "FROM (cartoptions INNER JOIN cart ON cartoptions.coCartID=cart.cartID) INNER JOIN orders ON cart.cartOrderID=orders.ordID "
	end if
	sSQLopts=sSQLopts & whereclause
	' print "<div>"&sSQL&"</div><div>"&sSQLopts&"</div>" : response.flush
	if hasresults then %>
			<div style="padding-top:20px">
			  <h2><%=yySalRes%></h2>
<%
		totopts=0
		rs.open sSQLopts,cnn,0,1
		if NOT rs.EOF then
			if isnull(rs("theordtot")) then
			' elseif cint(rs("theordtot"))=0 then
			else
				totopts=rs("theordtot")
			end if
		end if
		rs.close

		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			print "<div class=""stattbl"">"
			if isnull(rs("numorders")) then
				call displayemptyresults()
			elseif clng(rs("numorders"))=0 then
				call displayemptyresults()
			else
				print "<div class=""stattblrow""><div>"&yyTotItm&"</div><div>"&yyItmTot&"</div></div>"
				print "<div class=""stattblrow""><div>" & rs("numorders") & "</div><div>" & FormatEuroCurrency(rs("theordtot")+totopts) & "</div></div>"
			end if
			print "</div>"
%>
			<div style="margin-top:10px">
				<a download="ectsalesresults.csv" href="data:application/octet-stream,<%
					print yyTotItm&"%2C"&yyItmTot&"%0A"
					print rs("numorders") & "%2C" & vsround(rs("theordtot")+totopts,2)
				%>">Click Here For CSV Download</a>
			</div>
<%		end if
		rs.close
%>
			</div>
<%		response.flush
	end if
	if hasresults then
		print "<div style=""overflow:auto"">"
		for index2=1 to 2
			if mysqlserver then
				sSQL="SELECT SUM(cartQuantity) AS thecount,SUM((cartProdPrice+IFNULL((SELECT SUM(coPriceDiff) FROM cartoptions WHERE coCartID=cartID),0))*cartQuantity) AS theordtot,cartProdID "
				if hascategory then
					sSQL=sSQL&"FROM products INNER JOIN cart ON products.pID=cart.cartProdID INNER JOIN orders ON cart.cartOrderID=orders.ordID "
				else
					sSQL=sSQL&"FROM cart INNER JOIN orders ON cart.cartOrderID=orders.ordID "
				end if
				sSQL=sSQL&whereclause & " GROUP BY cartProdID "
			elseif sqlserver then
				sSQL="SELECT "&IIfVs(getpost("numresults")<>"all","TOP "&IIfVr(getpost("numresults")="",100,getpost("numresults")))&" cartProdID,SUM(cartQuantity) AS thecount,(SUM((cartProdPrice+COALESCE(t1.summedOptions,0))*cartQuantity)) AS theordtot FROM "&IIfVr(hascategory,"products INNER JOIN cart ON products.pID=cart.cartProdID","cart")&" LEFT JOIN (SELECT SUM(coPriceDiff) AS summedOptions,coCartID FROM cartoptions GROUP BY coCartID) t1 on cart.cartID=t1.coCartID INNER JOIN orders ON cart.cartOrderID=orders.ordID "&whereclause&" GROUP BY cartProdID "
			else
				sSQL="SELECT "&IIfVs(getpost("numresults")<>"all","TOP "&IIfVr(getpost("numresults")="",100,getpost("numresults")))&" SUM(cartQuantity) AS thecount,SUM((cartProdPrice+opts.optstot)*cartQuantity) AS theordtot,cartProdID "
				if hascategory then
					sSQL=sSQL & "FROM products INNER JOIN ((cart LEFT JOIN (SELECT coCartID,IIF(ISNULL(SUM(opts.coPriceDiff)),0,SUM(opts.coPriceDiff)) AS optstot FROM cartoptions opts GROUP BY opts.coCartID) opts ON cart.cartID=opts.coCartID) INNER JOIN orders ON cart.cartOrderID=orders.ordID) ON products.pID=cart.cartProdID "
				else
					sSQL=sSQL & "FROM (cart LEFT JOIN (SELECT coCartID,IIF(ISNULL(SUM(opts.coPriceDiff)),0,SUM(opts.coPriceDiff)) AS optstot FROM cartoptions opts GROUP BY opts.coCartID) opts ON cart.cartID=opts.coCartID) INNER JOIN orders ON cart.cartOrderID=orders.ordID "
				end if
				sSQL=sSQL & whereclause & " GROUP BY cartProdID "
			end if
			if index2=1 then
				sSQL=sSQL & "ORDER BY "&IIfVr(sqlserver OR mysqlserver, "thecount", "SUM(cartQuantity)")&" DESC"&IIfVs(mysqlserver AND getpost("numresults")<>"all"," LIMIT 0,"&IIfVr(getpost("numresults")="",100,getpost("numresults")))
			else
				sSQL=sSQL & "ORDER BY "&IIfVr(sqlserver OR mysqlserver, "theordtot", "SUM((cartProdPrice+opts.optstot)*cartQuantity)")&" DESC"&IIfVs(mysqlserver AND getpost("numresults")<>"all"," LIMIT 0,"&IIfVr(getpost("numresults")="",100,getpost("numresults")))
			end if
			' print "<div>"&sSQL&"</div>"
%>
			<div style="padding-top:20px;width:49%;min-width:800px;float:left;margin-right:1%">
			  <h2><%
					print IIfVr(getpost("numresults")<>"all",replace(yyTopSal,"100",getpost("numresults")),replace(yyTopSal,"100",""))
					if index2=1 then print " By Quantity" else print " By Amount" %></h2>
<%			rs.open sSQL,cnn,0,1
			if rs.EOF then
				call displayemptyresults()
			else
				print "<div class=""stattbl"">"
				print "<div class=""stattblrow""><div>"&yyPrId&"</div><div>"&yyPrName&"</div><div>"&replace(yyTotSal," ","&nbsp;")&"</div><div>"&yyAmount&"</div></div>"
				dldata=""
				do while NOT rs.EOF
					prodname=""
					sSQL="SELECT pName FROM products WHERE pID='" & escape_string(rs("cartProdID")) & "'"
					rs2.open sSQL,cnn,0,1
					if NOT rs2.EOF then prodname=rs2("pName")
					rs2.close
					if prodname="" then
						sSQL="SELECT "&IIfVs(NOT mysqlserver,"TOP 1 ")&"cartProdName FROM cart WHERE cartProdID='" & escape_string(rs("cartProdID")) & "' ORDER BY cartID DESC"&IIfVs(mysqlserver," LIMIT 0,1")
						rs2.open sSQL,cnn,0,1
						if NOT rs2.EOF then prodname=rs2("cartProdName")
						rs2.close
					end if
					theordtot=rs("theordtot")
					if isnull(theordtot) then theordtot=0
					print "<div class=""stattblrow""><div>" & rs("cartProdID") & "</div><div>" & prodname & "</div><div>" & rs("thecount") & "</div><div>" & formatnumber(theordtot,2,-1,0,0) & "</div></div>"
					dldata=dldata&rs("cartProdID") & "%2C" & replace(urlencode(prodname),"""","""""") & "%2C" & rs("thecount") & "%2C" & vsround(theordtot,2) & "%0A"
					rs.MoveNext
				loop
				print "</div>" %>
				<div style="margin-top:10px">
					<a download="ectstatsby<%=IIfVr(index2=1,"quantity","amount")%>.csv" href="data:application/octet-stream,<%
						print yyPrId&"%2C"&yyPrName&"%2C"&yyTotSal&"%2C"&yyAmount&"%0A"
						print dldata
					%>">Click Here For CSV Download</a>
				</div>
<%			end if
			rs.close
%>
			</div>
<%			response.flush
		next
		print "</div>" %>
			<div style="padding-top:20px">
			  <h2><%=yyTopCou%></h2>
<%		sSQLopts="SELECT SUM(coPriceDiff*cartQuantity) AS theordtot,ordCountry "
		if mysqlserver then
			sSQLopts=sSQLopts & "FROM cartoptions INNER JOIN cart ON cartoptions.coCartID=cart.cartID LEFT OUTER JOIN products ON cart.cartProdId=products.pID INNER JOIN orders ON cart.cartOrderID=orders.ordID "
		else
			sSQLopts=sSQLopts & "FROM cartoptions INNER JOIN ((cart LEFT OUTER JOIN products ON cart.cartProdId=products.pID) INNER JOIN orders ON cart.cartOrderID=orders.ordID) ON cartoptions.coCartID=cart.cartID "
		end if
		sSQLopts=sSQLopts & whereclause & " GROUP BY ordCountry"&IIfVs(mysqlserver=TRUE," LIMIT 0,100")
		' print "<div>"&sSQL&"<br />"&sSQLopts&"</div>" : response.flush
		
		totopts=0
		alloptions=""
		rs.open sSQLopts,cnn,0,1
		if NOT rs.EOF then alloptions=rs.getrows()
		rs.close

		sSQL="SELECT "&IIfVs(NOT mysqlserver,"TOP 100")&" SUM(cartQuantity) AS thecount,SUM(cartProdPrice*cartQuantity) AS theordtot, ordCountry "
		if mysqlserver then
			sSQL=sSQL & "FROM products INNER JOIN cart ON products.pID=cart.cartProdID INNER JOIN orders ON cart.cartOrderID=orders.ordID "
		else
			sSQL=sSQL & "FROM products INNER JOIN (cart INNER JOIN orders ON cart.cartOrderID=orders.ordID) ON products.pID=cart.cartProdID "
		end if
		sSQL=sSQL & whereclause & " GROUP BY ordCountry ORDER BY "&IIfVr(mysqlserver, "thecount", "SUM(cartQuantity)")&" DESC"&IIfVs(mysqlserver," LIMIT 0,100")
		' print "<div>"&sSQL&"</div>"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			print "<div class=""stattbl"">"
			print "<div class=""stattblrow""><div>"&yyCntNam&"</div><div>"&yyTotSal&"</div><div>"&yyAmount&"</div></div>"
			dldata=""
			do while NOT rs.EOF
				addoptions=0
				if isarray(alloptions) then
					for index=0 to UBOUND(alloptions, 2)
						if alloptions(1, index)=rs("ordCountry") then addoptions=alloptions(0, index) : exit for
					next
				end if
				print "<div class=""stattblrow""><div>" & rs("ordCountry") & "</div><div>" & rs("thecount") & "</div><div>" & formatnumber(rs("theordtot")+addoptions,2,-1,0,0) & "</div></div>"
				dldata=dldata&rs("ordCountry") & "%2C" & rs("thecount") & "%2C" & vsround(rs("theordtot")+addoptions,2) & "%0A"
				rs.MoveNext
			loop
			print "</div>" %>
			<div style="margin-top:10px">
				<a download="ectcountrystats.csv" href="data:application/octet-stream,<%
					print yyCntNam&"%2C"&yyTotSal&"%2C"&yyAmount&"%0A"
					print dldata
				%>">Click Here For CSV Download</a>
			</div>
<%		end if
		rs.close
%>
            </div>
<%	end if
else %>
			<div style="text-align:center;margin:80px">Please select a results format and press the &quot;Stats&quot; button.</div>
<%
end if ' grouping
%>
			<div style="text-align:center;margin:20px"><a href="admin.asp"><strong><%=yyAdmHom%></strong></a></div>
<%
cnn.Close
set rs=nothing
set cnn=nothing
%>