<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
'ANSI to Unicode conversion function by Lewis E. Moten III
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
addsuccess=TRUE
success=TRUE
successlines=0
faillines=0
pidnotfoundlines=0
pidnotfoundlineszerostock=0
zerostocklines=0
pidnotfoundpids=""
pidnotfoundpidszerostock=""
zerostockpids=""
stoppedonerror=FALSE
showaccount=TRUE
dorefresh=FALSE
isstockupdate=FALSE
isimagesupdate=FALSE
ismailinglistupdate=FALSE
iscouponupdate=FALSE
hasworkingname=FALSE
hasstartdate=FALSE
hasenddate=FALSE
CrLf=Chr(13) & Chr(10)
csvcurrpos=1
csvlen=0
Server.ScriptTimeout=360
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs2=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN

disableupdatechecker=TRUE
function getadminsettings()
	getadminsettings=TRUE
end function

'**************************************
' Name: ANSI to Unicode
' Description:Converts from ANSI to Unic
'     ode very fast. Inspired by code found in
'     UltraFastAspUpload by Cakkie (on PSC). T
'     his should work slightly faster then Cak
'     kies due to how some of the code has bee
'     n arranged.
' By: Lewis E. Moten III
'
' This code is copyrighted and has
' limited warranties. Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=7266&lngWId=4
' for details.
'**************************************
function ANSIToUnicode(ByRef pbinBinaryData)
	Dim lbinData	' Binary Data (ANSI)
	Dim llngLength	' Length of binary data (byte count)
	Dim lobjRs		' Recordset
	Dim lstrData	' Unicode Data
	' VarType Reference
	'8=Integer (this is expected var type)
	'17=Byte Subtype
	' 8192=Array
	' 8209=Byte Subtype + Array
	Set lobjRs=Server.CreateObject("ADODB.Recordset")
	if VarType(pbinBinaryData)=8 then
		' Convert integers(4 bytes) To Byte Subtype Array (1 byte)
		llngLength=LenB(pbinBinaryData)
		if llngLength=0 then
			lbinData=ChrB(0)
		else
			Call lobjRs.Fields.Append("BinaryData", adLongVarBinary, llngLength)
			Call lobjRs.Open()
			Call lobjRs.AddNew()
			Call lobjRs.Fields("BinaryData").AppendChunk(pbinBinaryData & ChrB(0)) ' + Null terminator
			Call lobjRs.Update()
			lbinData=lobjRs.Fields("BinaryData").GetChunk(llngLength)
			Call lobjRs.Close()
		end if
	else
		lbinData=pbinBinaryData
	end if
	' Do REAL conversion now!	
	llngLength=LenB(lbinData)
	if llngLength=0 then
		lstrData=""
	else
		Call lobjRs.Fields.Append("BinaryData", 201, llngLength)
		Call lobjRs.Open()
		Call lobjRs.AddNew()
		Call lobjRs.Fields("BinaryData").AppendChunk(lbinData)
		Call lobjRs.Update()
		lstrData=lobjRs.Fields("BinaryData").Value
		Call lobjRs.Close()
	end if
				
	Set lobjRs=nothing
	ANSIToUnicode=lstrData
end function
function IIfVr(theExp,theTrue,theFalse)
if theExp then IIfVr=theTrue else IIfVr=theFalse
end function
function IIfVs(theExp,theTrue)
if theExp then IIfVs=theTrue else IIfVs=""
end function
function vsusdate(thedate)
	if mysqlserver=true then
		vsusdate="'" & DatePart("yyyy",thedate) & "-" & DatePart("m",thedate) & "-" & DatePart("d",thedate) & "'"
	elseif sqlserver=true then
		vsusdate="CAST('" & DatePart("yyyy",thedate) & "-" & IIfVs(DatePart("m",thedate)<10,"0") & DatePart("m",thedate) & "-" & IIfVs(DatePart("d",thedate)<10,"0") & DatePart("d",thedate) & "' AS DATETIME)"
	else
		vsusdate="#" & DatePart("m",thedate) & "/" & DatePart("d",thedate) & "/" & DatePart("yyyy",thedate) & "#"
	end if
end function
function is_numeric(tstr)
	is_numeric=isnumeric(trim(tstr&""))
end function
function escape_string(str)
	escape_string=trim(replace(str&"","'","''"))
	if mysqlserver=TRUE then escape_string=replace(escape_string,"\","\\")
end function
function getcsvline()
	getcsvline=""
	do while csvcurrpos <= csvlen
		tmpchar=mid(csvfile, csvcurrpos, 1)
		csvcurrpos=csvcurrpos+1
		if tmpchar=vbCr OR tmpchar=vbLf then exit do else getcsvline=getcsvline&tmpchar
	loop
	do while csvcurrpos <= csvlen
		tmpchar=mid(csvfile, csvcurrpos, 1)
		if tmpchar=vbCr OR tmpchar=vbLf then csvcurrpos=csvcurrpos+1 else exit do
	loop
end function
function GetFieldName(infoStr)
	sPos=InStr(infoStr, "name=")
	endPos=InStr(sPos + 6, infoStr, Chr(34) & ";")
	if endPos=0 then
		endPos=inStr(sPos + 6, infoStr, Chr(34))
	end if
	GetFieldName=mid(infoStr, sPos + 6, endPos - (sPos + 6))
end function
'This function retreives a file field's filename
function GetFileName(infoStr)
	sPos=InStr(infoStr, "filename=")
	endPos=InStr(infoStr, Chr(34) & CrLf)
	GetFileName=mid(infoStr, sPos + 10, endPos - (sPos + 10))
end function
'This function retreives a file field's MIME type
function GetFileType(infoStr)
	sPos=InStr(infoStr, "Content-Type: ")
	GetFileType=mid(infoStr, sPos + 14)
end function
function htmlspecials(thestr)
	htmlspecials=trim(replace(replace(replace(replace(thestr&"","&","&amp;"),">","&gt;"),"<","&lt;"),"""","&quot;"))
end function
if ectdemostore then
	biData=""
else
	biData=Request.BinaryRead(Request.TotalBytes)
end if
bidatalen=LenB(biData)
isposted=(bidatalen>0)
PostData=""
if isposted then
	PostData=ANSIToUnicode(biData)
	ContentType=Request.ServerVariables("HTTP_CONTENT_TYPE")
	ctArray=Split(ContentType, ";")
	if trim(ctArray(0))="multipart/form-data" then
		ErrMsg=""
		' grab the form boundary...
		bArray=Split(trim(ctArray(1)), "=")
		Boundary=trim(bArray(1))
		'Now use that to split up all the variables!
		formData=Split(PostData, Boundary)
		'Extract the information for each variable and its data
		FileCount=0
		for x=0 to UBound(formData)
			'Two CrLfs mark the end of the information about this field; everything after that is the value
			Infoend=InStr(formData(x), CrLf & CrLf)
			if Infoend > 0 then
				'Get info for this field, minus stuff at the end
				varInfo=mid(formData(x), 3, Infoend - 3)
				'Get value for this field, being sure to skip CrLf pairs at the start and the CrLf at the end
				if (InStr(varInfo, "filename=") > 0) then ' Is this a file?
					if GetFieldName(varInfo)="csvfile" then
						csvfile=mid(formData(x), Infoend + 4, Len(formData(x)) - Infoend - 7) & vbCrLf ' add a "known elephant"
						csvlen=len(csvfile)
					end if
					' GetFileName(varInfo) : GetFileType(varInfo)
					FileCount=FileCount + 1
				else ' It's a regular field
					varValue=mid(formData(x), Infoend + 4, Len(formData(x)) - Infoend - 7)
					fieldname=GetFieldName(varInfo)
					select case fieldname
					case "show_errors"
						show_errors=(varValue="ON")
					case "stop_errors"
						stop_errors=(varValue="ON")
					case "theaction"
						isupdate=(varValue="update")
					end select
				end if
			end if
		next
	else
		ErrMsg="Wrong encoding type!"
	end if
end if
progressevery=500
function csv_database_error()
	if show_errors then response.write "Line " & line_num & ", " & mysql_error & "<br />"
	csvsuccess=FALSE
	faillines=faillines+1
	successlines=successlines-1
end function
function getsection(swn)
	sSQL="SELECT sectionID FROM sections WHERE rootSection=1 AND sectionWorkingName='" & escape_string(swn) & "'"
	rs2.Open sSQL,cnn,0,1
	if NOT rs2.EOF then getsection=rs2("sectionID") else getsection=0
	rs2.Close
end function
function getmanufacturer(swn)
	sSQL="SELECT scID FROM manufacturer WHERE scGroup=0 AND scName='" & escape_string(swn) & "'"
	rs2.Open sSQL,cnn,0,1
	if NOT rs2.EOF then getmanufacturer=rs2("scID") else getmanufacturer=0
	rs2.Close
end function
function checkmanufacturer(pmanufacturer)
	checkmanufacturer=0
	if isarray(manufacturerarray) then
		for index=0 to UBOUND(manufacturerarray,2)
			if int(manufacturerarray(0,index))=int(pmanufacturer) AND manufacturerarray(1,index)<>"" then checkmanufacturer=pmanufacturer : exit for
		next
	end if
end function
function isnumcol(thecol)
	isnumcol=(thecol="pprice" OR thecol="pwholesaleprice" OR thecol="plistprice" OR thecol="pshipping" OR thecol="pshipping2" OR thecol="pweight" OR thecol="pdisplay" OR thecol="psell" OR thecol="pexemptions" OR thecol="pinstock" OR thecol="ptax" OR thecol="pdropship" OR thecol="porder" OR thecol="pmanufacturer" OR thecol="ptotrating" OR thecol="pnumratings")
end function
sub updatemanufacturer(pid,pmanufacturer)
	if allmanufacturers<>"" then
		cnn.execute("DELETE FROM multisearchcriteria WHERE mSCscID IN ("&allmanufacturers&") AND mSCpID='"&escape_string(pid)&"'")
	end if
	if pmanufacturer<>0 then
		cnn.execute("INSERT INTO multisearchcriteria (mSCscID,mSCpID) VALUES ("&pmanufacturer&",'"&escape_string(pid)&"')")
	end if
end sub
function execute_sql()
	if isimagesupdate then
		rs.open "SELECT * FROM products WHERE pID='" & replace(valuesarray(0),"'","''") & "'",cnn,1,3,&H0001
		if rs.EOF then
			pidnotfoundlines=pidnotfoundlines+1
			pidnotfoundpids=pidnotfoundpids&htmlspecials(valuesarray(0))&"<br />"
			successlines=successlines-1
		else
			if isupdate then
				if valuesarray(1)="" then
					sSQL="DELETE FROM productimages WHERE imageproduct='" & replace(valuesarray(0),"'","''") & "' AND imagetype=" & valuesarray(2) & " AND imagenumber=" & valuesarray(3)
				else
					sSQL="UPDATE productimages SET imagesrc='" & replace(valuesarray(1),"'","''") & "' WHERE imageproduct='" & replace(valuesarray(0),"'","''") & "' AND imagetype=" & valuesarray(2) & " AND imagenumber=" & valuesarray(3)
				end if
				cnn.execute(sSQL)
			else
				sSQL="DELETE FROM productimages WHERE imageproduct='" & replace(valuesarray(0),"'","''") & "' AND imagetype=" & valuesarray(2) & " AND imagenumber>=" & valuesarray(3)
				cnn.execute(sSQL)
				sSQL="INSERT INTO productimages (imageproduct,imagesrc,imagetype,imagenumber) VALUES ("
				sSQL=sSQL & "'" & replace(valuesarray(0),"'","''") & "','" & replace(valuesarray(1),"'","''") & "'," & valuesarray(2) & "," & valuesarray(3) & ")"
				cnn.execute(sSQL)
			end if
		end if
		rs.close
	elseif isstockupdate then
		' on error resume next
		if trim(valuesarray(4))<>"" then
			sSQL="UPDATE options SET optStock=" & valuesarray(3) & " WHERE optID=" & valuesarray(4)
		else
			sSQL="UPDATE products SET pInStock=" & valuesarray(3) & " WHERE pID='" & trim(valuesarray(0)) & "'"
		end if
		err.number=0
		cnn.execute(sSQL)
		on error goto 0
	elseif ismailinglistupdate then
		sSQL="UPDATE mailinglist SET "
		gotemail=FALSE
		emailaddress=""
		fullname=""
		for i=0 to columncount-1
			if columnarray(i)="email" then
				gotemail=TRUE
				emailaddress=trim(valuesarray(i))
				if emailaddress="" then gotemail=FALSE
			elseif columnarray(i)="full name" then
				fullname=trim(valuesarray(i))
			end if
		next
		if gotemail then
			sSQL="SELECT email FROM mailinglist WHERE email='" & escape_string(emailaddress) & "'"
			rs.open sSQL,cnn,0,1
			emailexists=NOT rs.EOF
			rs.close
			if emailexists then
				sSQL="UPDATE mailinglist SET isconfirmed=1"
				if fullname<>"" then sSQL=sSQL&",mlName='" & escape_string(fullname) & "'"
				sSQL=sSQL&" WHERE email='" & escape_string(emailaddress) & "'"
			else
				sSQL="INSERT INTO mailinglist (email,mlName,isconfirmed,mlConfirmDate) VALUES ('" & escape_string(emailaddress) & "','" & escape_string(fullname) & "',1," & vsusdate(date()) & ")"
			end if
			cnn.execute(sSQL)
		end if
	elseif isupdate then
		pid=""
		pimage=""
		plargeimage=""
		pgiantimage=""
		addcomma=""
		hasdimensions=FALSE
		dimspattern="PLENxPWIDxPHEI"
		pmanufacturer=0
		if mysqlserver=TRUE then
			sSQL="UPDATE products SET "
			for i=0 to columncount-1
				if i <> keycolumn then
					if columnarray(i)="plength" then
						dimspattern=replace(dimspattern,"PLEN",valuesarray(i))
						hasdimensions=TRUE
					elseif columnarray(i)="pwidth" then
						dimspattern=replace(dimspattern,"PWID",valuesarray(i))
						hasdimensions=TRUE
					elseif columnarray(i)="pheight" then
						dimspattern=replace(dimspattern,"PHEI",valuesarray(i))
						hasdimensions=TRUE
					elseif columnarray(i)="pimage" then
						pimage=valuesarray(i)
					elseif columnarray(i)="plargeimage" then
						plargeimage=valuesarray(i)
					elseif columnarray(i)="pgiantimage" then
						pgiantimage=valuesarray(i)
					else
						if columnarray(i)="psection" then
							if NOT is_numeric(valuesarray(i)) then valuesarray(i)=getsection(valuesarray(i))
						elseif columnarray(i)="pmanufacturer" then
							if NOT is_numeric(valuesarray(i)) then valuesarray(i)=getmanufacturer(valuesarray(i))
							valuesarray(i)=checkmanufacturer(valuesarray(i))
							pmanufacturer=valuesarray(i)
						elseif isnumcol(columnarray(i)) then
							if valuesarray(i)="" then valuesarray(i)=0
						end if
						
						sSQL=sSQL & addcomma & columnarray(i) & "='" & escape_string(valuesarray(i)) & "'"
						addcomma=","
					end if
				end if
			next
			if hasdimensions then
				dimspattern=replace(dimspattern,"PLEN",0)
				dimspattern=replace(dimspattern,"PWID",0)
				dimspattern=replace(dimspattern,"PHEI",0)
				sSQL=sSQL & addcomma & "pDims" & "='" & dimspattern & "'"
			end if
			sSQL=sSQL & " WHERE pID='" & replace(valuesarray(keycolumn),"'","''") & "'"
			' response.write "<b>" & sSQL & "</b><br />"
			cnn.execute(sSQL)
		else
			if mysqlserver then rs.CursorLocation=3
			rs.open "SELECT * FROM products WHERE pID='" & replace(valuesarray(keycolumn),"'","''") & "'",cnn,1,3,&H0001
			if rs.EOF then
				pidnotfoundlines=pidnotfoundlines+1
				pidnotfoundpids=pidnotfoundpids&htmlspecials(valuesarray(keycolumn))&"<br />"
				successlines=successlines-1
			else
				for i=0 to columncount-1
					if i <> keycolumn then
						if (rs.Fields(columnarray(i)).Type=3 OR rs.Fields(columnarray(i)).Type=5 OR rs.Fields(columnarray(i)).Type=11 OR rs.Fields(columnarray(i)).Type=17) AND trim(valuesarray(i)&"")="" then valuesarray(i)=0
						' response.write "Upd col: " & columnarray(i) & " - " & valuesarray(i) & " : " & rs.Fields(columnarray(i)).Type & "<br>" : response.flush
						if columnarray(i)="plength" then
							dimspattern=replace(dimspattern,"PLEN",valuesarray(i))
							hasdimensions=TRUE
						elseif columnarray(i)="pwidth" then
							dimspattern=replace(dimspattern,"PWID",valuesarray(i))
							hasdimensions=TRUE
						elseif columnarray(i)="pheight" then
							dimspattern=replace(dimspattern,"PHEI",valuesarray(i))
							hasdimensions=TRUE
						elseif columnarray(i)="pimage" then
							pimage=valuesarray(i)
						elseif columnarray(i)="plargeimage" then
							plargeimage=valuesarray(i)
						elseif columnarray(i)="pgiantimage" then
							pgiantimage=valuesarray(i)
						else
							on error resume next
							err.number=0
							if columnarray(i)="psection" then
								if NOT is_numeric(valuesarray(i)) then secval=getsection(valuesarray(i)) else secval=valuesarray(i)
								rs.Fields(columnarray(i))=secval
							elseif columnarray(i)="pmanufacturer" then
								if NOT is_numeric(valuesarray(i)) then secval=getmanufacturer(valuesarray(i)) else secval=valuesarray(i)
								secval=checkmanufacturer(secval)
								pmanufacturer=secval
								rs.Fields(columnarray(i))=secval
							elseif rs.Fields(columnarray(i)).Type=5 then
								rs.Fields(columnarray(i))=cdbl(valuesarray(i))
							else
								rs.Fields(columnarray(i))=valuesarray(i)
							end if
							errnum=err.number
							on error goto 0
							if errnum<>0 then
								if show_errors then
									faillines=faillines+1
									successlines=successlines-1
									response.write "Data type mismatch adding " & valuesarray(i) & " to column " & columnarray(i) & "<br>"
								end if
								if stop_errors then
									csvcurrpos=csvlen+1
									stoppedonerror=TRUE
								end if
							end if
						end if
					end if
				next
				if hasdimensions then
					dimspattern=replace(dimspattern,"PLEN",0)
					dimspattern=replace(dimspattern,"PWID",0)
					dimspattern=replace(dimspattern,"PHEI",0)
					rs.Fields("pDims")=dimspattern
				end if
				rs.Update
			end if
			rs.close
		end if
		if pimage<>"" then cnn.execute("UPDATE productimages SET imageSrc='" & escape_string(pimage) & "' WHERE imageProduct='" & escape_string(valuesarray(keycolumn)) & "' AND imageNumber=0 AND imageType=0")
		if plargeimage<>"" then cnn.execute("UPDATE productimages SET imageSrc='" & escape_string(plargeimage) & "' WHERE imageProduct='" & escape_string(valuesarray(keycolumn)) & "' AND imageNumber=0 AND imageType=1")
		if pgiantimage<>"" then cnn.execute("UPDATE productimages SET imageSrc='" & escape_string(pgiantimage) & "' WHERE imageProduct='" & escape_string(valuesarray(keycolumn)) & "' AND imageNumber=0 AND imageType=2")
		if hasmanufacturer then call updatemanufacturer(valuesarray(keycolumn),pmanufacturer)
	else
		pid=""
		pimage=""
		plargeimage=""
		pgiantimage=""
		addcomma=""
		hasdimensions=FALSE
		dimspattern="PLENxPWIDxPHEI"
		pcolumns="("
		pvalues=") VALUES ("
		pmanufacturer=0
		cpnname=""
		if mysqlserver=TRUE then
			sSQL="INSERT INTO products ("
			addcomma=""
			for i=0 to columncount-1
				sSQL=sSQL & addcomma & columnarray(i)
				addcomma=","
			next
			sSQL=sSQL & ") VALUES ("
			addcomma=""
			for i=0 to columncount-1
				if columnarray(i)="plength" then
					dimspattern=replace(dimspattern,"PLEN",valuesarray(i))
					hasdimensions=TRUE
				elseif columnarray(i)="pwidth" then
					dimspattern=replace(dimspattern,"PWID",valuesarray(i))
					hasdimensions=TRUE
				elseif columnarray(i)="pheight" then
					dimspattern=replace(dimspattern,"PHEI",valuesarray(i))
					hasdimensions=TRUE
				elseif columnarray(i)="pimage" then
					pimage=valuesarray(i)
				elseif columnarray(i)="plargeimage" then
					plargeimage=valuesarray(i)
				elseif columnarray(i)="pgiantimage" then
					pgiantimage=valuesarray(i)
				else
					if columnarray(i)="cpnname" then cpnname=valuesarray(i)
					if columnarray(i)="pid" then
						pid=valuesarray(i)
					elseif columnarray(i)="psection" then
						if NOT is_numeric(valuesarray(i)) then valuesarray(i)=getsection(valuesarray(i))
					elseif columnarray(i)="pmanufacturer" then
						if NOT is_numeric(valuesarray(i)) then valuesarray(i)=getmanufacturer(valuesarray(i))
						valuesarray(i)=checkmanufacturer(valuesarray(i))
						pmanufacturer=valuesarray(i)
					elseif isnumcol(columnarray(i)) then
						if valuesarray(i)="" then valuesarray(i)=0
					end if
					pcolumns=pcolumns & addcomma & columnarray(i)
					pvalues=pvalues & addcomma & "'" & escape_string(valuesarray(i)) & "'"
					addcomma=","
				end if
			next
			if hasdimensions then
				dimspattern=replace(dimspattern,"PLEN",0)
				dimspattern=replace(dimspattern,"PWID",0)
				dimspattern=replace(dimspattern,"PHEI",0)
				pcolumns=pcolumns & addcomma & "pDims"
				pvalues=pvalues & addcomma & "'" & escape_string(dimspattern) & "'"
			end if
			if iscouponupdate then
				sSQL="INSERT INTO coupons " & pcolumns & IIfVs(NOT hasworkingname,"cpnworkingname") & IIfVs(NOT hasstartdate,"cpnstartdate") & IIfVs(NOT hasenddate,"cpnenddate") & pvalues & IIfVs(NOT hasworkingname,"'" & escape_string(cpnname) & "'") & IIfVs(NOT hasstartdate,vsusdate(dateserial(2000,1,1))) & IIfVs(NOT hasenddate,vsusdate(dateserial(3000,1,1))) & ")"
			else
				sSQL="INSERT INTO products " & pcolumns & pvalues & ")"
			end if
			' response.write "<b>" & sSQL & "</b><br />"
			on error resume next
			err.number=0
			cnn.execute(sSQL)
			errnum=err.number
			errdesc=err.description
			on error goto 0
		else
			on error resume next
			rs.open IIfVr(iscouponupdate,"coupons","products"),cnn,1,3,&H0002
			rs.AddNew
			for i=0 to columncount-1
				' response.write "Add col: " & columnarray(i) & " - " & valuesarray(i) & "<br>"
				if columnarray(i)="plength" then
					dimspattern=replace(dimspattern,"PLEN",valuesarray(i))
					hasdimensions=TRUE
				elseif columnarray(i)="pwidth" then
					dimspattern=replace(dimspattern,"PWID",valuesarray(i))
					hasdimensions=TRUE
				elseif columnarray(i)="pheight" then
					dimspattern=replace(dimspattern,"PHEI",valuesarray(i))
					hasdimensions=TRUE
				elseif columnarray(i)="pimage" then
					pimage=valuesarray(i)
				elseif columnarray(i)="plargeimage" then
					plargeimage=valuesarray(i)
				elseif columnarray(i)="pgiantimage" then
					pgiantimage=valuesarray(i)
				else
					if columnarray(i)="cpnname" then cpnname=valuesarray(i)
					if columnarray(i)="psection" then
						if NOT is_numeric(valuesarray(i)) then secval=getsection(valuesarray(i)) else secval=valuesarray(i)
						rs.Fields(columnarray(i))=secval
					elseif columnarray(i)="pmanufacturer" then
						if NOT is_numeric(valuesarray(i)) then secval=getmanufacturer(valuesarray(i)) else secval=valuesarray(i)
						secval=checkmanufacturer(secval)
						pmanufacturer=secval
						rs.Fields(columnarray(i))=secval
					elseif rs.Fields(columnarray(i)).Type=5 then
						if is_numeric(valuesarray(i)) then rs.Fields(columnarray(i))=cdbl(valuesarray(i)) else rs.Fields(columnarray(i))=0
					else
						if columnarray(i)="pid" then pid=valuesarray(i)
						rs.Fields(columnarray(i))=valuesarray(i)
					end if
				end if
			next
			if hasdimensions then
				dimspattern=replace(dimspattern,"PLEN",0)
				dimspattern=replace(dimspattern,"PWID",0)
				dimspattern=replace(dimspattern,"PHEI",0)
				rs.Fields("pDims")=dimspattern
			end if
			err.number=0
			if iscouponupdate then
				if NOT hasworkingname then rs.Fields("cpnworkingname")=cpnname
				if NOT hasstartdate then rs.Fields("cpnstartdate")=dateserial(2000,1,1)
				if NOT hasenddate then rs.Fields("cpnenddate")=dateserial(3000,1,1)
			end if
			rs.Update
			errnum=err.number
			errdesc=err.description
			rs.close
			on error goto 0
		end if
		if errnum<>0 then
			if errnum=-2147217887 OR errnum=-2147467259 then
				errdesc="Error, duplicate ID column.<br />" & errdesc
				pidnotfoundlines=pidnotfoundlines+1
			else
				faillines=faillines+1
			end if
			if show_errors then response.write "Adding pID: &quot;" & valuesarray(keycolumn) & "&quot; - " & errdesc & " (" &  errnum & ")<br>"
			if stop_errors then
				csvcurrpos=csvlen+1
				stoppedonerror=TRUE
			end if
			csvsuccess=FALSE
			successlines=successlines-1
		end if
		on error resume next
			if pimage<>"" then cnn.execute("INSERT INTO productimages (imageProduct,imageSrc,imageNumber,imageType) VALUES ('" & escape_string(pid) & "','" & escape_string(pimage) & "',0,0)")
			if plargeimage<>"" then cnn.execute("INSERT INTO productimages (imageProduct,imageSrc,imageNumber,imageType) VALUES ('" & escape_string(pid) & "','" & escape_string(plargeimage) & "',0,1)")
			if pgiantimage<>"" then cnn.execute("INSERT INTO productimages (imageProduct,imageSrc,imageNumber,imageType) VALUES ('" & escape_string(pid) & "','" & escape_string(pgiantimage) & "',0,2)")
		on error goto 0
		if hasmanufacturer then call updatemanufacturer(pid,pmanufacturer)
	end if
end function
if isposted then
	' response.write '<meta http-equiv="refresh" content="2; url=admincsv.asp">';
	time_start=timer()
	column_list=lcase(replace(getcsvline(),"""",""))
	hasmanufacturer=FALSE
	allmanufacturers=""
	if column_list="imageproduct,imagesrc,imagetype,imagenumber" then
		isimagesupdate=TRUE
	elseif column_list="pid,pname,pprice,pinstock,optid,optiongroup,option" then
		isstockupdate=TRUE
	elseif left(column_list,15)="email,full name" then
		ismailinglistupdate=TRUE
	elseif instr(","&column_list&",","cpnnumber")>0 AND instr(","&column_list&",","cpnname")>0 AND instr(","&column_list&",","cpndiscount")>0 then
		iscouponupdate=TRUE
	else
		column_list_copy=replace(column_list," ","")
		column_list_copy=replace(column_list_copy,"plength","")
		column_list_copy=replace(column_list_copy,"pwidth","")
		column_list_copy=replace(column_list_copy,"pheight","")
		column_list_copy=replace(column_list_copy,"pimage","")
		column_list_copy=replace(column_list_copy,"plargeimage","")
		column_list_copy=replace(column_list_copy,"pgiantimage","")
		do while(column_list_copy <> replace(column_list_copy,",,",","))
			column_list_copy=replace(column_list_copy,",,",",")
		loop
		if right(column_list_copy,1)="," then column_list_copy=left(column_list_copy,len(column_list_copy)-1)
		on error resume next
		err.number=0
		cnn.execute("SELECT " & column_list_copy & " FROM products WHERE pID='abcwxyz'")
		errnum=err.number
		errdesc=err.description
		on error goto 0
		if errnum<>0 then
			errmsg=errdesc
			success=FALSE
		end if
	end if
	columnarray=split(lcase(column_list), ",")
	valuesarray=columnarray
	columncount=UBOUND(columnarray)+1
	columnnum=0
	keycolumn=""
	stockcolumn=""
	for i=0 to columncount-1
		columnarray(i)=trim(columnarray(i))
		if iscouponupdate then
			if columnarray(i)="cpnnumber" then keycolumn=i
		else
			if columnarray(i)="pid" then keycolumn=i
		end if
		if columnarray(i)="pinstock" then stockcolumn=i
		if columnarray(i)="pmanufacturer" then hasmanufacturer=TRUE
	next
	if keycolumn="" AND isimagesupdate=FALSE AND ismailinglistupdate=FALSE AND NOT iscouponupdate then
		success=FALSE
		errmsg="There was no pID column specified."
	end if
	if success then
		if isupdate then
			response.write "&nbsp;Updating row: "
		else
			response.write "&nbsp;Adding row: "
		end if
		if hasmanufacturer then
			sSQL="SELECT scID,scGroup,scName FROM searchcriteria WHERE scGroup=0"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then manufacturerarray=rs.getrows()
			rs.close
			if isarray(manufacturerarray) then
				for index=0 to UBOUND(manufacturerarray,2)
					allmanufacturers=allmanufacturers&manufacturerarray(0,index)&" "
				next
				allmanufacturers=replace(trim(allmanufacturers)," ",",")
			end if
		end if
		line_num=1
		totallines=20
		do while csvcurrpos < csvlen
			thechar=mid(csvfile, csvcurrpos, 1)
			' response.write "&lt;"&thechar&">"
			if NOT needquote then
				if thiscol="" AND thechar="""" then
					needquote=TRUE
				elseif thechar <> "," AND thechar <> vbCr AND thechar <> vbLf then
					thiscol=thiscol&thechar
				else
					valuesarray(columnnum)=thiscol
					columnnum=columnnum+1
					' response.write "<b>Adding col:</b>" & columnnum & ": " & thiscol & "<br>" : response.flush
					if columnnum=columncount OR thechar=vbCr OR thechar=vbLf then
						do while columnnum<columncount
							valuesarray(columnnum)=null
							columnnum=columnnum+1
						loop
						successlines=successlines+1
						columnnum=0
						execute_sql()
						if (line_num MOD progressevery)=0 then
							response.write line_num & ", "
							response.flush()
						end if
						needquote=FALSE
						do while csvcurrpos<csvlen
							tmpchar=mid(csvfile, csvcurrpos+1, 1)
							if tmpchar=vbCr OR tmpchar=vbLf then csvcurrpos=csvcurrpos+1 else exit do
						loop
						line_num=line_num + 1
					end if
					thiscol=""
				end if
			elseif thechar="""" then
				if mid(csvfile, csvcurrpos+1, 1)="""" then
					thiscol=thiscol & """"
					csvcurrpos=csvcurrpos + 1
				else
					needquote=FALSE
				end if
			else
				pos=instr(csvcurrpos, csvfile, """")
				if pos=0 then
					thiscol=thiscol & mid(csvfile, csvcurrpos, (csvlen+1) - csvcurrpos)
					csvcurrpos=csvlen
				else
					thiscol=thiscol & mid(csvfile, csvcurrpos, pos - csvcurrpos)
					' response.write "<br>ADDING THIS CHUNK: " & mid(csvfile, csvcurrpos, pos - csvcurrpos) & "<br>"
					csvcurrpos=pos-1
				end if
			end if
			csvcurrpos=csvcurrpos+1
		loop
		response.write line_num-1 & "</p>"
	end if
	time_end=timer()
%>
      <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr>
          <td width="100%">
			<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
                <td width="100%" colspan="2" align="center"><%
					if NOT success then response.write "<p>ERROR: " & errmsg & "</p>"
					if isupdate OR isimagesupdate then
						response.write "<p>Rows successfully updated " & successlines & "</p>"
						if faillines > 0 then response.write "<p>Error rows " & faillines & "</p>"
						response.write "<p>Rows where pID not found " & pidnotfoundlines & "</p>"
					else
						response.write "<p>Rows successfully added " & successlines & "</p>"
						if faillines > 0 then response.write "<p>Error rows " & faillines & "</p>"
						if pidnotfoundlines > 0 then response.write "<p>Rows with duplicate product id (pID) " & pidnotfoundlines & "</p>"
					end if
					response.write "<p>This page took: " & round(time_end - time_start,4) & " seconds</p>"
					if successlines + faillines > 0 then response.write "<p>That is " & round((time_end - time_start) / (successlines + faillines), 4) & " seconds per row</p>"
					if pidnotfoundpids<>"" then
						response.write "<div style=""display:inline-block;width:250px;text-align:left;margin:15px;padding:15px;border:1px solid black"">"
						response.write "<div style=""margin-bottom:10px"">Rows where pID not found " & pidnotfoundlines & "</div>"
						response.write pidnotfoundpids & "</div>"
					end if
                %></td>
			  </tr>
			  <tr> 
                <td width="100%" colspan="2" align="center"><br /><strong><% if stoppedonerror then response.write "<span style=""color:#FF0000"">" & yyOpFai & "</span>" else response.write yyUpdSuc%></strong><br /><br /><br /><br />
                        Please <a href="admin.asp"><strong><%=yyClkHer%></strong></a> for the admin home page<% if stoppedonerror then response.write " or <a href=""javascript:history.go(-1)""><strong>" & yyClkHer & "</strong></a> to go back and try again"%>.<br />
                        <br /><br />&nbsp;
                </td>
			  </tr>
			</table></td>
        </tr>
      </table>
<%
else
%>
		  <form name="mainform" method="post" action="admincsv.asp" enctype="multipart/form-data">
		  <input type="hidden" name="posted" value="1">
			<table width="100%" border="0" cellspacing="0" cellpadding="2">
			  <tr>
				<td width="100%" align="center" colspan="2"><strong><%=yyCSVUpl%></strong><br />&nbsp;
<%			if ectdemostore then %>
					<p align="center" style="color:#FF0000;font-weight:bold">Please note, CSV upload features are disabled for the demo store.<br />&nbsp;</p>
<%			end if %>
				</td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyCSVFlN%>:</strong></td>
				<td><input type="file" name="csvfile" /></td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyAct%>:</strong></td>
				<td><select name="theaction" size="1">
					<option value="add"><%=yyAddDB%></option>
					<option value="update"><%=yyUpdDB%></option>
					</select></td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyShwErr%>:</strong></td>
				<td><input type="checkbox" name="show_errors" value="ON" checked /></td>
			  </tr>
			  <tr>
				<td align="right"><strong><%=yyStpErr%>:</strong></td>
				<td><input type="checkbox" name="stop_errors" value="ON" checked /></td>
			  </tr>
			  <tr>
				<td width="100%" align="center" colspan="2">&nbsp;<br /><input type="submit" value="<%=yySubmit%>" /><br />&nbsp;</td>
			  </tr>
			  <tr> 
				<td width="100%" align="center" colspan="2"><br />
					  <a href="admin.asp"><strong><%=yyAdmHom%></strong></a><br />&nbsp;</td>
			  </tr>
			</table>
		  </form>
<%
end if
%>