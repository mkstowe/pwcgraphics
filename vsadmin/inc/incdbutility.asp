<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if storesessionvalue="" then storesessionvalue="virtualstore"
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE OR sDSN="" then response.end
set rs = Server.CreateObject("ADODB.RecordSet")
set rs2= Server.CreateObject("ADODB.RecordSet")
set cnn=Server.CreateObject("ADODB.Connection")
function dbuencodeimage(byval imurl)
	imurl=replace(imurl&"","\","/")
	imurl=replace(replace(replace(replace(imurl,"|","%7C"),"<","%3C"),"?","%3F"),">","%3E")
	imurl=replace(replace(replace(replace(imurl,"prodimages/","|"),".gif","<"),".png","?"),".jpg",">")
	dbuencodeimage=replace(imurl,"'","\'")
end function
function in_arraydbu(element,arr)
	in_arraydbu=FALSE
	if isarray(arr) then
		for xxi=0 to UBOUND(arr,2)
			if trim(arr(0,xxi))=trim(element) then
				in_arraydbu=TRUE
				exit function      
			end If
		next
	end if
end function
if getget("act")="delimages" then
	response.clear
	cnn.open sDSN
	sSQL="DELETE FROM productimages WHERE imageSrc='"&escape_string(getget("iu"))&"'"
	cnn.execute(sSQL)
	print getget("iname")
	response.end
	cnn.close
elseif getget("getprodfromimage")<>"" then
	response.clear
	cnn.open sDSN
	sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 1 ")&"imageProduct FROM productimages WHERE imageSrc='"&escape_string(getget("getprodfromimage"))&"'"&IIfVs(mysqlserver=TRUE," LIMIT 0,1")
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then response.write rs("imageProduct")
	rs.close
	response.end
	cnn.close
end if
sub gobackhome()
	print "<div style=""margin:20px;text-align:center""><p align=""center""><input type=""button"" value=""Please click here to go back to the DB Utility Home Screen"" onclick=""document.location='admindbutility.asp'"" /></div>"
end sub
%>
<script>
function editproduct(pid){
	document.getElementById('postform').action="adminprods.asp";
	document.getElementById('postform').target="dbedit";
	document.getElementById('control1').name="id";
	document.getElementById('control1').value=pid;
	document.getElementById('control2').name="act";
	document.getElementById('control2').value="modify";
	document.getElementById('control3').name="posted";
	document.getElementById('control3').value="1";
	document.getElementById('postform').submit();
}
function fixsection(tact){
	document.getElementById('postform').action="admindbutility.asp";
	document.getElementById('postform').target="";
	document.getElementById('control1').name="mainaction";
	document.getElementById('control1').value="checkdatabase";
	document.getElementById('control2').name="subaction";
	document.getElementById('control2').value=tact;
	document.getElementById('postform').submit();
}
</script>
<form id="postform" method="post" action="">
<input type="hidden" id="control1" value="" />
<input type="hidden" id="control2" value="" />
<input type="hidden" id="control3" value="" />
</form>
<%
tablenames="abandonedcartemail,address,admin,adminlogin,affiliates,ajaxfloodcontrol,alternaterates,auditlog,cart,cartoptions,contentregions,countries,coupons,cpnassign,customerlists,customerlogin,dropshipper,emailmessages,giftcertificate,giftcertsapplied,installedmods,ipblocking,mailinglist,multibuyblock,multisearchcriteria,multisections,notifyinstock,optiongroup,options,orders,orderstatus,passwordhistory,payprovider,postalzones,pricebreaks,prodoptions,productimages,productpackages,products,ratings,recentlyviewed,relatedprods,searchcriteria,searchcriteriagroup,sections,shipoptions,states,tmplogin,uspsmethods,zonecharges"
cnn.open sDSN
server.scripttimeout=1800
if getpost("mainaction")="checkdatabase" then
	' Fix Actions
	print "<div class=""ectred ectdbsection"">"
		print "<div style=""padding:10px 0"">Before fixing any database problems, please make sure you have made a backup of your database.</div>"
	print "</div>"
	if getpost("subaction")="remmultisearch" then
		print "<div class=""ectdbsection"">"
		print "<div class=""half_bottom"">Fixing entries where the manufacturer attributes are not the same as that set for the product</div>"
		sSQL="SELECT mSCpID,mSCscID FROM multisearchcriteria INNER JOIN products ON multisearchcriteria.mSCpID=products.pID AND products.pManufacturer<>multisearchcriteria.mSCscID INNER JOIN searchcriteria ON multisearchcriteria.mSCscID=searchcriteria.scID WHERE scGroup=0"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			do while NOT rs.EOF
				sSQL="DELETE FROM multisearchcriteria WHERE mSCpID='" & escape_string(rs("mSCpID")) & "' AND mSCscID='" & escape_string(rs("mSCscID")) & "'"
				ect_query(sSQL)
				rs.movenext
			loop
			print "<div class=""half_bottom"" style=""color:#369105"">Fixed - entries where the manufacturer attributes are not the same as that set for the product.</div>"
		else
			print "<div class=""half_bottom"">No entries where the manufacturer attributes are not the same as that set for the product.</div>"
		end if
		rs.close
		print "</div>"
	elseif getpost("subaction")="addmanuf" OR getpost("subaction")="remmanuf" then
		print "<div class=""ectdbsection"">"
		print "<div class=""half_bottom"">Inserting missing manufacturer entries in the Product Attributes table</div>"
		manufacturers=""
		if getpost("subaction")<>"remmanuf" then
			sSQL="SELECT scID FROM searchcriteria WHERE scGroup=0"
			rs.open sSQL,cnn,0,1
			manufacturers=rs.getrows
			rs.close
		end if
		sSQL="SELECT pID,pManufacturer FROM products LEFT JOIN multisearchcriteria ON products.pID=multisearchcriteria.mSCpID AND products.pManufacturer=multisearchcriteria.mSCscID WHERE pManufacturer<>0 AND mSCscID IS NULL"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			do while NOT rs.EOF
				if in_arraydbu(rs("pManufacturer"),manufacturers) then
					sSQL="INSERT INTO multisearchcriteria (mSCpID,mSCscID) VALUES ('" & escape_string(rs("pID")) & "','" & escape_string(rs("pManufacturer")) & "')"
				else
					sSQL="UPDATE products SET pManufacturer=0 WHERE pID='" & escape_string(rs("pID")) & "'"
				end if
				ect_query(sSQL)
				rs.movenext
			loop
			print "<div class=""half_bottom"" style=""color:#369105"">Fixed - missing manufacturer entries in the Product Attributes table.</div>"
		else
			print "<div class=""half_bottom"">No missing manufacturer entries in the Product Attributes table.</div>"
		end if
		rs.close
		print "</div>"
	elseif getpost("subaction")="missmscp" then
		print "<div class=""ectdbsection"">"
		print "<div class=""half_bottom"">Fixing entries in the multisearchcriteria table where the product does not exist</div>"
		sSQL="SELECT mSCpID,mSCscID FROM multisearchcriteria LEFT JOIN products ON multisearchcriteria.mSCpID=products.pID WHERE products.pID IS NULL"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			do while NOT rs.EOF
				sSQL="DELETE FROM multisearchcriteria WHERE mSCpID='" & escape_string(rs("mSCpID")) & "' AND mSCscID='" & escape_string(rs("mSCscID")) & "'"
				ect_query(sSQL)
				rs.movenext
			loop
			print "<div class=""half_bottom"" style=""color:#369105"">Fixed - entries with missing products in multisearchcriteria table.</div>"
		else
			print "<div class=""half_bottom"">No entries with missing products in multisearchcriteria table.</div>"
		end if
		rs.close
		print "</div>"
	elseif getpost("subaction")="missmscc" then
		print "<div class=""ectdbsection"">"
		print "<div class=""half_bottom"">Fixing entries in the multisearchcriteria table where the Attribute does not exist</div>"
		sSQL="SELECT mSCpID,mSCscID FROM multisearchcriteria LEFT JOIN searchcriteria ON multisearchcriteria.mSCscID=searchcriteria.scID WHERE searchcriteria.scID IS NULL"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			do while NOT rs.EOF
				sSQL="DELETE FROM multisearchcriteria WHERE mSCpID='" & escape_string(rs("mSCpID")) & "' AND mSCscID='" & escape_string(rs("mSCscID")) & "'"
				ect_query(sSQL)
				rs.movenext
			loop
			print "<div class=""half_bottom"" style=""color:#369105"">Fixed - entries with missing Attributes in multisearchcriteria table.</div>"
		else
			print "<div class=""half_bottom"">No entries with missing Attributes in multisearchcriteria table.</div>"
		end if
		rs.close
		print "</div>"
	elseif getpost("subaction")="misspimg" then
		print "<div class=""ectdbsection"">"
		print "<div class=""half_bottom"">Fixing entries in the productimages table where the Product does not exist</div>"
		sSQL="SELECT imageProduct FROM productimages LEFT JOIN products ON productimages.imageProduct=products.pID WHERE products.pID IS NULL"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			do while NOT rs.EOF
				sSQL="DELETE FROM productimages WHERE imageProduct='" & escape_string(rs("imageProduct")) & "'"
				ect_query(sSQL)
				rs.movenext
			loop
			print "<div class=""half_bottom"" style=""color:#369105"">Fixed - entries with missing Product in productimages table.</div>"
		else
			print "<div class=""half_bottom"">No entries with missing Product in productimages table.</div>"
		end if
		rs.close
		print "</div>"
	end if

	' Missing category / section
	print "<div class=""ectdbsection"">"
	print "<div class=""half_bottom"">Checking for products whose category / section is missing</div>"
	sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 100 ")&"pID,pSection,sectionID FROM products LEFT JOIN sections ON products.pSection=sections.sectionID WHERE sectionID IS NULL"&IIfVs(mysqlserver=TRUE," LIMIT 0,100")
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		print "<div class=""ectred half_bottom"">The following products link to a category which is not defined.</div>"
		print "<div class=""ectdbtable""><div class=""ectdbhead""><div class=""ectdbheadrow"">Product ID</div><div class=""ectdbheadrow"">Section ID</div><div class=""ectdbheadrow"">&nbsp;</div></div>"
		do while NOT rs.EOF
			print "<div class=""ectdbrow_cnt""><div class=""ectdbrow"">" & rs("pID") & "</div><div class=""ectdbrow"">" & rs("pSection") & "</div><div class=""ectdbrow""><input type=""button"" value=""Edit Product"" onclick=""editproduct('" & jscheck(rs("pID")) & "')"" /></div></div>"
			rs.movenext
		loop
		print "</div>"
	else
		print "<div class=""half_bottom"" style=""color:#369105"">All ok!</div>"
	end if
	rs.close
	print "</div>"
	
	response.flush
	
	' Category / section not a products section
	print "<div class=""ectdbsection"">"
	print "<div class=""half_bottom"">Checking for products whose category / section is not for products</div>"
	sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 100 ")&"pID,pSection,sectionID FROM products INNER JOIN sections ON products.pSection=sections.sectionID WHERE rootSection<>1"&IIfVs(mysqlserver=TRUE," LIMIT 0,100")
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		print "<div class=""ectred half_bottom"">The following products link to a category which is not for products.</div>"
		print "<div class=""ectdbtable""><div class=""ectdbhead""><div class=""ectdbheadrow"">Product ID</div><div class=""ectdbheadrow"">Section ID</div><div class=""ectdbheadrow"">&nbsp;</div></div>"
		do while NOT rs.EOF
			print "<div class=""ectdbrow_cnt""><div class=""ectdbrow"">" & rs("pID") & "</div><div class=""ectdbrow"">" & rs("pSection") & "</div><div class=""ectdbrow""><input type=""button"" value=""Edit Product"" onclick=""editproduct('" & jscheck(rs("pID")) & "')"" /></div></div>"
			rs.movenext
		loop
		print "</div>"
	else
		print "<div class=""half_bottom"" style=""color:#369105"">All ok!</div>"
	end if
	rs.close
	print "</div>"
	
	response.flush
	
	' Manufacturer Not In Multisearchcriteria Table.
	print "<div class=""ectdbsection"">"
	print "<div class=""half_bottom"">Checking for products where the manufacturer set has no entry in the Product Attributes table</div>"
	sSQL="SELECT COUNT(pID) AS thecount FROM products LEFT JOIN multisearchcriteria ON products.pID=multisearchcriteria.mSCpID AND products.pManufacturer=multisearchcriteria.mSCscID WHERE pManufacturer<>0 AND mSCscID IS NULL"
	rs.open sSQL,cnn,0,1
	thecount=0
	if NOT rs.EOF then thecount=rs("thecount")
	if thecount>0 then
		print "<div class=""ectred half_bottom"">There are " & thecount & " entries where the manufacturer set has no entry in the Product Attributes table.</div>"
		print "<div class=""half_bottom""><input type=""button"" value=""Remove manufacturer setting from these products."" onclick=""fixsection('remmanuf')"" /></div>"
		print "<div class=""half_bottom""><input type=""button"" value=""Add entry to Product Attributes table for these products."" onclick=""fixsection('addmanuf')"" /></div>"
	else
		print "<div class=""half_bottom"" style=""color:#369105"">All ok!</div>"
	end if
	rs.close
	print "</div>"
	
	response.flush
	
	if sqlserver OR mysqlserver then
		' Multisearchcriteria Table Entry to Manufacturer not the same as products.
		print "<div class=""ectdbsection"">"
		print "<div class=""half_bottom"">Checking for manufacturer attributes which are not the same as that set for the product</div>"
		sSQL="SELECT COUNT(mSCpID) AS thecount FROM multisearchcriteria INNER JOIN products ON multisearchcriteria.mSCpID=products.pID AND products.pManufacturer<>multisearchcriteria.mSCscID INNER JOIN searchcriteria ON multisearchcriteria.mSCscID=searchcriteria.scID WHERE scGroup=0"
		rs.open sSQL,cnn,0,1
		thecount=0
		if NOT rs.EOF then thecount=rs("thecount")
		if thecount>0 then
			print "<div class=""ectred half_bottom"">There are " & thecount & " entries where the manufacturer attributes are not the same as that set for the product.</div>"
			print "<div class=""half_bottom""><input type=""button"" value=""Remove these attribute entries."" onclick=""fixsection('remmultisearch')"" /></div>"
		else
			print "<div class=""half_bottom"" style=""color:#369105"">All ok!</div>"
		end if
		rs.close
		print "</div>"
		
		response.flush
	end if

	' Multisearchcriteria Product Missing.
	print "<div class=""ectdbsection"">"
	print "<div class=""half_bottom"">Checking for entries in the Product Attributes table where the product does not exist</div>"
	sSQL="SELECT COUNT(mSCpID) AS thecount FROM multisearchcriteria LEFT JOIN products ON multisearchcriteria.mSCpID=products.pID WHERE products.pID IS NULL"
	rs.open sSQL,cnn,0,1
	thecount=0
	if NOT rs.EOF then thecount=rs("thecount")
	if thecount>0 then
		print "<div class=""ectred half_bottom"">There are " & thecount & " entries in the Product Attributes table where the product does not exist.</div>"
		print "<div class=""half_bottom""><input type=""button"" value=""Delete entries with missing products in Product Attributes table."" onclick=""fixsection('missmscp')"" /></div>"
	else
		print "<div class=""half_bottom"" style=""color:#369105"">All ok!</div>"
	end if
	rs.close
	print "</div>"
	
	response.flush
	
	' Multisearchcriteria Attribute Missing.
	print "<div class=""ectdbsection"">"
	print "<div class=""half_bottom"">Checking for entries in the Product Attributes table where the attribute does not exist</div>"
	sSQL="SELECT COUNT(mSCscID) AS thecount FROM multisearchcriteria LEFT JOIN searchcriteria ON multisearchcriteria.mSCscID=searchcriteria.scID WHERE searchcriteria.scID IS NULL"
	rs.open sSQL,cnn,0,1
	thecount=0
	if NOT rs.EOF then thecount=rs("thecount")
	if thecount>0 then
		print "<div class=""ectred half_bottom"">There are " & thecount & " entries in the Product Attributes table where the attribute does not exist.</div>"
		print "<div class=""half_bottom""><input type=""button"" value=""Delete entries with missing sections in Product Attributes table."" onclick=""fixsection('missmscc')"" /></div>"
	else
		print "<div class=""half_bottom"" style=""color:#369105"">All ok!</div>"
	end if
	rs.close
	print "</div>"
	
	' Product Images entry has missing product
	print "<div class=""ectdbsection"">"
	print "<div class=""half_bottom"">Checking for entries in the Product Images table where the product does not exist</div>"
	sSQL="SELECT COUNT(imageProduct) AS thecount FROM productimages LEFT JOIN products ON productimages.imageProduct=products.pID WHERE products.pID IS NULL"
	rs.open sSQL,cnn,0,1
	thecount=0
	if NOT rs.EOF then thecount=rs("thecount")
	if thecount>0 then
		print "<div class=""ectred half_bottom"">There are " & thecount & " entries in the Product Images table where the product does not exist.</div>"
		print "<div class=""half_bottom""><input type=""button"" value=""Delete entries with missing products in Product Attributes table."" onclick=""fixsection('misspimg')"" /></div>"
	else
		print "<div class=""half_bottom"" style=""color:#369105"">All ok!</div>"
	end if
	rs.close
	print "</div>"

	gobackhome()
elseif getpost("mainaction")="buildcache" then
	sSQL="UPDATE searchcriteria SET scGroupOrder=(SELECT scgOrder FROM searchcriteriagroup WHERE searchcriteriagroup.scgID=searchcriteria.scGroup),scGroupTitle=(SELECT scgTitle FROM searchcriteriagroup WHERE searchcriteriagroup.scgID=searchcriteria.scGroup),scGroupTitle2=(SELECT scgTitle2 FROM searchcriteriagroup WHERE searchcriteriagroup.scgID=searchcriteria.scGroup),scGroupTitle3=(SELECT scgTitle3 FROM searchcriteriagroup WHERE searchcriteriagroup.scgID=searchcriteria.scGroup) WHERE EXISTS (SELECT * FROM searchcriteriagroup WHERE searchcriteriagroup.scgID=searchcriteria.scGroup)"
	ect_query(sSQL)
	
	sSQL="UPDATE multisearchcriteria SET mscDisplay=(SELECT pDisplay FROM products WHERE products.pID=multisearchcriteria.mSCpID) WHERE EXISTS (SELECT * FROM products WHERE products.pID=multisearchcriteria.mSCpID)"
	ect_query(sSQL)

	print "<div style=""margin:20px;text-align:center"">All done!!.</div>"
	gobackhome()
elseif getpost("mainaction")="nvarchar" AND getpost("subaction")="go" then
	Dim pktablenames(100)
	Dim pkcolumns(100)
	pkmaxindex=0
	print "<p style=""margin:40px 0px"">This script can take quite some time so please don't use the back button.</p>"
	response.flush
	sSQL="SELECT table_name,column_name,max_length,is_nullable FROM (SELECT t.name AS table_name,c.name AS column_name,tp.name AS data_type,c.max_length,c.is_nullable FROM sys.tables AS t " & _
		"INNER JOIN sys.columns c ON t.OBJECT_ID = c.OBJECT_ID " & _
		"INNER JOIN sys.types tp ON tp.user_type_id = c.user_type_id " & _
		"WHERE tp.name='nvarchar' AND t.name IN ('" & replace(tablenames,",","','") & "') )t"
	' print sSQL&"<br>"
	cnn.CommandTimeout=0
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		defaulttext=""
		defaultcname=""
		primarykeytext=""
		primarycname=""
		hasindex=FALSE
		isunique=FALSE
		indexname=""
		sSQL2="SELECT s.name AS SchemaName,t.name AS TableName,c.name AS ColumnName,d.name AS DefaultConstraintName,d.definition AS DefaultDefinition FROM sys.default_constraints d " & _
			"INNER JOIN sys.columns c ON d.parent_object_id = c.object_id AND d.parent_column_id = c.column_id " & _
			"INNER JOIN sys.tables t ON t.object_id = c.object_id " & _
			"INNER JOIN sys.schemas s ON s.schema_id = t.schema_id " & _
			"WHERE t.name='" & escape_string(rs("table_name")) & "' and c.name='" & escape_string(rs("column_name")) & "'"
		rs2.open sSQL2,cnn,0,1
		if NOT rs2.EOF then
			print rs2("TableName")&"->"&rs2("ColumnName")& "@ " & rs2("DefaultDefinition") & " : " & rs2("DefaultConstraintName") & "<br>"
			defaulttext=replace(replace(rs2("DefaultDefinition"),"(",""),")","")
			defaultcname=rs2("DefaultConstraintName")
			cnn.execute("ALTER TABLE " & rs2("TableName") & " DROP CONSTRAINT " & rs2("DefaultConstraintName"))
		end if
		rs2.close

		sSQL2="SELECT KU.table_name as TABLENAME,column_name as PRIMARYKEYCOLUMN,TC.CONSTRAINT_TYPE,TC.CONSTRAINT_NAME FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS AS TC " & _
			"INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE AS KU ON TC.CONSTRAINT_NAME = KU.CONSTRAINT_NAME AND KU.table_name='" & escape_string(rs("table_name")) & "' " & _
			"WHERE TC.CONSTRAINT_TYPE='PRIMARY KEY'"
		' "WHERE TC.CONSTRAINT_TYPE='PRIMARY KEY' AND column_name='" & escape_string(rs("column_name")) & "'"
		rs2.open sSQL2,cnn,0,1
		addcomma=""
		if NOT rs2.EOF then
			do while NOT rs2.EOF
				' print rs2("TableName")&"->"&rs2("ColumnName")& "@ " & rs2("DefaultDefinition") & " : " & rs2("DefaultConstraintName") & "<br>"
				if instr(","&primarykeytext&",",","&rs2("PRIMARYKEYCOLUMN")&",")=0 then primarykeytext=primarykeytext&addcomma&rs2("PRIMARYKEYCOLUMN")
				on error resume next
				cnn.execute("ALTER TABLE [" & rs2("TableName") & "] DROP CONSTRAINT " & rs2("CONSTRAINT_NAME"))
				on error goto 0
				primarycname=rs2("CONSTRAINT_NAME")
				addcomma=","
				rs2.movenext
			loop
			pktablenames(pkmaxindex)=rs("table_name")
			pkcolumns(pkmaxindex)=primarykeytext
			pkmaxindex=pkmaxindex+1
		end if
		rs2.close
		
		sSQL2="SELECT  i.name AS indexname,i.type_desc,c.name AS columnname,i.is_unique FROM sys.index_columns AS ic " & _
            "INNER JOIN sys.indexes AS i ON ic.object_id = i.object_id AND ic.index_id = i.index_id " & _
            "INNER JOIN sys.columns AS c ON ic.object_id = c.object_id AND ic.column_id = c.column_id " & _
			"WHERE ic.object_id=OBJECT_ID('" & rs("table_name") & "') AND c.name='" & escape_string(rs("column_name")) & "' AND i.type!=0 AND i.is_primary_key=0"
			' "WHERE ic.object_id=OBJECT_ID('" & rs("table_name") & "') AND c.name='" & escape_string(rs("column_name")) & "' AND i.type!=0 AND ic.is_included_column=0 AND i.is_primary_key=0"
		rs2.open sSQL2,cnn,0,1
		do while NOT rs2.EOF
			hasindex=TRUE
			indexname=rs2("indexname")
			isunique=rs2("is_unique")=1
			cnn.execute("DROP INDEX " & rs("table_name") & "." & rs2("indexname"))
			rs2.movenext
		loop
		rs2.close

		sSQL="ALTER TABLE [" & rs("table_name") & "] ALTER COLUMN " & rs("column_name") & " VARCHAR(" & IIfVr(rs("max_length")=-1,"MAX",rs("max_length")/2) & ")" & IIfVs(rs("is_nullable")=0," NOT")&" NULL"
		print sSQL&"<br>"
		response.flush()
		cnn.execute(sSQL)
		if trim(defaulttext&"")<>"" then
			print "defaulttext is : " & defaulttext & "<br>"
			sSQL="ALTER TABLE " & rs("table_name") & " ADD CONSTRAINT "&defaultcname&" DEFAULT "&defaulttext&" FOR "&rs("column_name")
			on error resume next
			err.number=0
			cnn.execute(sSQL)
			errnum=err.number
			errdesc=err.description
			on error goto 0
			if errnum<>0 then print "<div style=""color:#ff0000"">" & sSQL & "<br />" & errdesc & "</div>" : response.flush : response.end
		end if
		if hasindex then cnn.execute("CREATE "&IIfVs(isunique,"UNIQUE ")&"INDEX "&indexname&" ON " & rs("table_name") & "(" & rs("column_name") & ")")

		rs.movenext
	loop
	for index=0 to pkmaxindex-1
		sSQL="ALTER TABLE " & pktablenames(index) & " ADD PRIMARY KEY ("&pkcolumns(index)&")"
		print sSQL&"<br>"
		on error resume next
		cnn.execute(sSQL)
		on error goto 0
	next
	print "<p style=""margin:40px 0px"">All Done.</p>"
	print "<input type=""button"" value=""Back to Admin DB Utility Home"" onclick=""document.location='admindbutility.asp'"" />"
	print "<p>&nbsp;</p>"
elseif getpost("mainaction")="varcharmax" AND getpost("subaction")="go" then
	print "<p style=""margin:40px 0px"">This script can take quite some time so please don't use the back button.</p>"
	response.flush
	if getpost("tablename")<>"" AND getpost("columnname")<>"" then
		coldatatype=getpost("coldatatype")=getpost("tablename")
		tablenames=getpost("tablename")
	end if
	sSQL="SELECT table_name,column_name,max_length,is_nullable FROM (SELECT t.name AS table_name,c.name AS column_name,tp.name AS data_type,c.max_length,c.is_nullable FROM sys.tables AS t " & _
		"INNER JOIN sys.columns c ON t.OBJECT_ID = c.OBJECT_ID " & _
		"INNER JOIN sys.types tp ON tp.user_type_id = c.user_type_id " & _
		"WHERE ((tp.name IN ('nvarchar','varchar') AND c.max_length IN (8000,-1)) OR tp.name IN ('text','ntext')) AND t.name IN ('" & replace(tablenames,",","','") & "')" & IIfVs(getpost("columnname")<>""," AND c.name IN ('" & getpost("columnname") & "')") & " )t"
	' print sSQL&"<br>"
	coldatatype=getpost("datatype")
	if getpost("tablename")<>"" AND getpost("columnname")<>"" then
		coldatatype=getpost("coldatatype")
	end if
	if coldatatype=1 then datatype="VARCHAR(MAX)"
	if coldatatype=2 then datatype="VARCHAR(8000)"
	if coldatatype=3 then datatype="TEXT"
	cnn.CommandTimeout=0
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		' print rs("table_name") & " : " & rs("column_name") & " : " & rs("max_length") & "<br>"
		sSQL="ALTER TABLE [" & rs("table_name") & "] ALTER COLUMN " & rs("column_name") & " " & datatype & " NULL"
		print sSQL&"<br>"
		response.flush()
		on error resume next
		err.number=0
		cnn.execute(sSQL)
		errnum=err.number
		errdesc=err.description
		on error goto 0
		if errnum<>0 then
			print "<div style=""color:#FF0000"">CONVERSION ERROR: " & errdesc & "</div>"
		end if
		rs.movenext
	loop
	rs.close %>
	<p style="margin:40px 0px">All Done.</p>
	<form method="post" action="admindbutility.asp" id="theform">
		<input type="hidden" name="mainaction" value="varcharmax">
		<input type="submit" value="Back to Column List" />
	</form>
	<p style="margin:40px 0px">&nbsp;</p>
<%
elseif getpost("mainaction")="varcharmax" then %>
<script>
/* <![CDATA[ */
function checkform(){
	if(document.getElementById('datatype').selectedIndex==0){
		alert("Please select a datatype to convert to.");
		document.getElementById('datatype').focus();
		return false;
	}
	if(confirm("Have you taken a backup of your database and are ready to proceed with the conversion?"))
		return true;
	return false;
}
function altercolumnsize(telem,tblname,colname){
	if(confirm('Are you sure you want to do this?')){
		document.getElementById("tablename").value=tblname;
		document.getElementById("columnname").value=colname;
		document.getElementById("coldatatype").value=telem[telem.selectedIndex].value;
		document.getElementById("colsizeform").submit();
	}
	telem.selectedIndex=0;
}
/* ]]> */
</script>
<%	if NOT sqlserver OR mysqlserver then
		print "<p style=""margin:40px"">This script is only designed for use with a MicroSoft SQL Server database</p>"
		print "<input type=""button"" value=""Cancel"" onclick=""document.location='admindbutility.asp'"" />"
		print "<p>&nbsp;</p>"
	else
		thecount=0
		sSQL="SELECT table_name,column_name,max_length,data_type FROM (SELECT t.name AS table_name,c.name AS column_name,tp.name AS data_type,c.max_length,c.is_nullable FROM sys.tables AS t " & _
		"INNER JOIN sys.columns c ON t.OBJECT_ID = c.OBJECT_ID " & _
		"INNER JOIN sys.types tp ON tp.user_type_id = c.user_type_id " & _
		"WHERE ((tp.name IN ('nvarchar','varchar') AND c.max_length IN (8000,-1)) OR tp.name IN ('text','ntext')) AND t.name IN ('" & replace(tablenames,",","','") & "') )t"
		cnn.CommandTimeout=0
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			print "<div style=""max-height:300px;overflow-y:scroll;overflow-x:hidden;display:inline-block""><div class=""ecttable"" style=""width:auto"">"
			do while NOT rs.EOF
				currval=3
				print "<div class=""ecttablerow""><div>" &  rs("table_name") & "</div><div>" & rs("column_name") & "</div><div>" & ucase(rs("data_type"))
				if rs("data_type")="varchar" then
					print "(" & IIfVr(rs("max_length")=-1,"MAX",rs("max_length")) & ")"
					if rs("max_length")=-1 then currval=1 else currval=2
				end if
				print "</div>"
				
				print "<select onchange=""altercolumnsize(this,'"&rs("table_name")&"','"&rs("column_name")&"')"" size=""1"">"
					print "<option value="""">Alter this column to...</option>"
					print "<option value=""1""" & IIfVs(currval=1," selected=""selected""") & ">VARCHAR(MAX)</option>"
					print "<option value=""2""" & IIfVs(currval=2," selected=""selected""") & ">VARCHAR(8000)</option>"
					print "<option value=""3""" & IIfVs(currval=3," selected=""selected""") & ">TEXT</option>"
				print "</select>"
				print "</div>" & vbLf
				thecount=thecount+1
				rs.movenext
			loop
			print "</div></div>"
		end if
		rs.close

		print "<form method=""post"" id=""colsizeform"" action=""admindbutility.asp"" onsubmit=""return checkform()"">" & whv("mainaction","varcharmax") & whv("subaction","go")
		print "<p style=""margin-top:40px""><h4>You can set the size of individual columns above, or clicking below will change ALL the VARCHAR(MAX), VARCHAR(8000) and TEXT columns, (" & thecount & " total columns), in your database to use the datatype.</h4></p>"
		print "<p>Changel ALL columns to: <select size=""1"" id=""datatype"" name=""datatype""><option value="""">Please Select...</option><option value=""1"">VARCHAR(MAX)</option><option value=""2"">VARCHAR(8000)</option><option value=""3"">TEXT</option></select></p>"
		print "<p>This can reduce the amount of space your database uses and greatly increase the speed of queries.</p>"
		print "<p class=""ectred"">It is imperative that you take a backup of the database before proceeding as it is possible that errors can occur, (for instance if the changes cause your database to run out of disk space.)</p>"
		print "<p>After using the script please be sure to fully test your store.</p>"
		print "<p>&nbsp;</p>"
		call writehiddenidvar("tablename","")
		call writehiddenidvar("columnname","")
		call writehiddenidvar("coldatatype","")
		print "<input type=""submit"" value=""Convert VARCHAR(MAX), VARCHAR(8000) AND TEXT columns"" /> <input type=""button"" value=""Cancel"" onclick=""document.location='admindbutility.asp'"" />"
		print "</form>"
		print "<p>&nbsp;</p>"
	end if
elseif getpost("mainaction")="nvarchar" then %>
<script>
/* <![CDATA[ */
function checkform(){
	if(confirm("Have you taken a backup of your database and are ready to proceed with the conversion?"))
		return true;
	return false;
}
/* ]]> */
</script>
<%	if NOT sqlserver OR mysqlserver then
		print "<p style=""margin:40px"">This script is only designed for use with a MicroSoft SQL Server database</p>"
		print "<input type=""button"" value=""Cancel"" onclick=""document.location='admindbutility.asp'"" />"
		print "<p>&nbsp;</p>"
	else
		thecount=0
		sSQL="SELECT COUNT(*) AS tcnt FROM (SELECT t.name AS table_name,c.name AS column_name,tp.name AS data_type,c.max_length,c.is_nullable FROM sys.tables AS t " & _
		"INNER JOIN sys.columns c ON t.OBJECT_ID = c.OBJECT_ID " & _
		"INNER JOIN sys.types tp ON tp.user_type_id = c.user_type_id " & _
		"WHERE tp.name='nvarchar' AND t.name IN ('" & replace(tablenames,",","','") & "') )t"
		cnn.CommandTimeout=0
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then thecount=rs("tcnt")
		rs.close

		print "<p style=""margin-top:40px"">This script will change the NVARCHAR columns, (" & thecount & " total columns), in your database to VARCHAR columns of the same length and properties.</p>"
		print "<p>This can reduce the amount of space your database uses and greatly increase the speed of queries.</p>"
		print "<p class=""ectred"">It is imperative that you take a backup of the database before proceeding as it is possible that errors can occur, (for instance if the changes cause your database to run out of disk space.)</p>"
		print "<p>After using the script please be sure to fully test your store.</p>"
		print "<form method=""post"" action=""admindbutility.asp"" onsubmit=""return checkform()"">" & whv("mainaction","nvarchar") & whv("subaction","go")
		print "<p>&nbsp;</p>"
		print "<input type=""submit"" value=""Convert NVARCHAR columns to VARCHAR"" /> <input type=""button"" value=""Cancel"" onclick=""document.location='admindbutility.asp'"" />"
		print "</form>"
		print "<p>&nbsp;</p>"
	end if
elseif getpost("mainaction")="reorderimages" then
	if getpost("subaction")="go" then %>
<table style="margin:0 auto">
<tr><td width="100%">
<%		sSQL="SELECT imageProduct,imageNumber,imageType FROM productimages ORDER BY imageProduct,imageType,imageNumber"
		currImageType=-1
		currImageNumber=-1
		currImageProduct=""
		expectedimagenumber=0
		numberreordered=0
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			if currImageProduct<>rs("imageProduct") OR currImageType<>rs("imageType") then
				expectedimagenumber=0
				currImageProduct=rs("imageProduct")
				currImageType=rs("imageType")
			end if
			if rs("imageNumber")<>expectedimagenumber then
				sSQL="UPDATE productimages SET imageNumber=" & expectedimagenumber & " WHERE imageProduct='"&escape_string(currImageProduct)&"' AND imageType="&currImageType&" AND imageNumber="&rs("imageNumber")
				cnn.execute(sSQL)
				numberreordered=numberreordered+1
			end if
			expectedimagenumber=expectedimagenumber+1
			rs.movenext
		loop
		print "<p>&nbsp;</p>"
		print "<p>&nbsp;</p>"
		if numberreordered=0 then
			print "<p>No images needed reordering.</p>"
		else
			print "<p>A total of " & numberreordered & " images were reordered.</p>"
		end if
		print "<p>&nbsp;</p>"
		print "<p>&nbsp;</p>"
		call gobackhome()
		rs.close
%>
</td></tr>
</table>
<%	else
		numimages=0
		sSQL="SELECT COUNT(*) AS numimages FROM productimages"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			if NOT isnull(rs("numimages")) then numimages=rs("numimages")
		end if %>
		<table style="margin:0 auto">
<tr><td width="100%">
<form action="admindbutility.asp" method="post">
<input type="hidden" name="mainaction" value="reorderimages">
<input type="hidden" name="subaction" value="go">
<input type="hidden" name="posted" value="1">
<p>This script will check the order of the images in your Ecommerce Templates database ensuring they are in sequential order.</p>
<p>Please make sure you have backed up your site and database before proceeding.</p>
<p>&nbsp;</p>
<p>Please click below to start checking<% if numimages>0 then response.write " <strong>(" & numimages & " images)</strong>"%>.</p>
<p>&nbsp;</p>
<p><input style="background:#399908; color:#fff;cursor:pointer;padding:5px 10px;-moz-border-radius:10px;-webkit-border-radius:10px" type="submit" value="Reorder Images" /></p>
<p>&nbsp;</p>
</form>
<% call gobackhome() %>
</td></tr>
</table>
<%	end if
elseif getpost("mainaction")="checkimagerefs" then %>
<script>
function dodeleteimages(){
	if(confirm('Are you sure you want to delete these images?\n\nPlease note, if the directory you select contains non-product images, these will be deleted.\n\nYou should make sure you have a backup before you continue.\n')){
		document.getElementById('subaction').value='dodeleteimages';
		document.getElementById('dodeleteimages').submit();
	}
}
</script>
<form method="post" id="dodeleteimages" action="admindbutility.asp">
<input type="hidden" name="mainaction" value="checkimagerefs" />
<input type="hidden" name="subaction" id="subaction" value="dodeleteimages" />
<input type="hidden" name="productimagesfolder" id="productimagesfolder" value="<%=getpost("productimagesfolder")%>" />
</form>
<%	function isimagefile(fname)
		lfname=lcase(fs.getextensionname(fname))
		isimagefile=lfname="gif" OR lfname="jpg" OR lfname="png" OR lfname="jpeg" OR lfname="tif" OR lfname="gif" OR lfname="bmp"
	end function

	if defaultprodimages="" then defaultprodimages="prodimages/"
	hasfso=TRUE
	err.number=0
	on error resume next
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	if err.number<>0 then hasfso=FALSE
	on error goto 0
	if hasfso then
		if getpost("productimagesfolder")<>"" then
			filepath=getpost("productimagesfolder")
		else
			filepath=replace(request.servervariables("URL"),"/admindbutility.asp","")
			slashpos=instrrev(filepath, "/")
			if slashpos>0 then filepath=left(filepath, slashpos)
			filepath=filepath&defaultprodimages
		end if
		print "<strong>Current Product Images Path:</strong> " & filepath & "<br />"
		mappedfilepath=server.mappath(filepath)
		print "<div style=""font-size:0.8em"">If this location is incorrect then please set the location below (relative to the site root) of the product images directory.</div>"
		print "<div style=""padding:20px"">"
		print "<input type=""text"" size=""46"" id=""imagelocation"" style=""vertical-align:middle"" value=""" & htmlspecials(filepath) & """ placeholder=""Eg: /myproductimages/"" />"
		print " <input type=""button"" value=""Change Location"" style=""vertical-align:middle"" onclick=""document.getElementById('productimagesfolder').value=document.getElementById('imagelocation').value;document.getElementById('subaction').value='';document.getElementById('dodeleteimages').submit()"" />"
		print "</div>"
		err.number=0
		
		direxists=TRUE
		on error resume next
		set fo=fs.GetFolder(mappedfilepath)
		if err.number<>0 then direxists=FALSE
		on error goto 0
		
		if NOT direxists then
			print "That directory does not exist<br />"
		else
			if getpost("subaction")="dodeleteimages" then
				for each afile in fo.files
					if isimagefile(afile.name) then
						sSQL="SELECT imageSrc FROM productimages WHERE imageSrc LIKE '" & escape_string(afile.name) & "' OR imageSrc LIKE '%/" & escape_string(afile.name) & "'"
						rs.open sSQL,cnn,0,1
						isfound=NOT rs.EOF
						rs.close
						if NOT isfound then
							print "Deleting: " & afile.name & "<br />"
							afile.delete()
						end if
					end if
				next
			end if
		
			print "Number of files in folder: " & fo.files.count & "<br /><br />"
			numreferenced=0
			notreferenced=0
			
			print "<div style=""font-weight:bold"">Unreferenced Images</div>"
			time_start=timer()
			for each afile in fo.files
				if isimagefile(afile.name) then
					sSQL="SELECT imageSrc FROM productimages WHERE imageSrc LIKE '" & escape_string(afile.name) & "' OR imageSrc LIKE '%/" & escape_string(afile.name) & "'"
					rs.open sSQL,cnn,0,1
					isfound=NOT rs.EOF
					rs.close
					if NOT isfound then
						print afile.name & "<br />"
						notreferenced=notreferenced+1
					else
						numreferenced=numreferenced+1
					end if
				end if
			next
			if notreferenced=0 then
				print "There are no unreferenced images<br />"
			else
				if allowdeleteimages then
					print "<br /><input type=""button"" value=""Delete Unreferenced Images"" onclick=""dodeleteimages()"" /><br />"
				else
					print "<br /><input type=""button"" value=""Delete Unreferenced Images: DISABLED"" disabled=""disabled"" /><br />"
					print "<div style=""font-size:0.8em"">To allow the automatic deletion of unreferenced images, please add the parameter ""allowdeleteimages=TRUE"" to your vsadmin/includes.asp file</div>"
				end if
			end if
			
			print "<br /><div style=""font-size:0.9em"">Images referenced: " & numreferenced & "<br />"
			print "Images NOT referenced: " & notreferenced & "<br />"
			print "Time taken: " & vsround(timer()-time_start,2) & " seconds.</div><br /><br />"
		end if
		set fo=nothing
		set fs=nothing
	else
		print "It seems that the Scripting.FileSystemObject has been disabled on your server so this check cannot be carried out.<br />"
	end if
	gobackhome()
elseif getpost("mainaction")="checkimages" then
	if getpost("subaction")="go" then %>
<form id="viewform" method="post" action="adminprods.asp" target="viewprodwindow">
<input type="hidden" name="posted" value="1" />
<input type="hidden" name="act" value="modify" />
<input type="hidden" id="viewprodid" name="id" value="" />
</form>
<script>
/* <![CDATA[ */
var foundimages=0, missingimages=0, imageindex=0, imagearray=[];
function displayproductcallback(){
	if(ajaxobj.readyState==4){
		var retvals=ajaxobj.responseText;
		if(retvals=='')
			alert("Product id could not be found.");
		else{
			document.getElementById('viewprodid').value=retvals;
			document.getElementById('viewform').submit();
		}
	}
}
function displayproduct(imind){
	ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
	ajaxobj.onreadystatechange=displayproductcallback;
	ajaxobj.open("GET", "admindbutility.asp?getprodfromimage="+encodeURIComponent(imagearray[imind]),true);
	ajaxobj.send(null);
}
function chkExists(imageurl){
	ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
	if(imageurl.substr(0,7).toLowerCase()!='http://'&&imageurl.substr(0,8).toLowerCase()!='https://'&&imageurl.substr(0,1)!='/') imageurl='../'+imageurl;
	ajaxobj.open('HEAD',imageurl,false);
	try{
		ajaxobj.send(null);
		return ajaxobj.status!=404;
	}catch(err){
		return false;
	}
}
function addtotable(imageurl,isfound){
	imagearray[imageindex]=imageurl;
	imgtable = document.getElementById('imagestable');
	newrow=imgtable.insertRow(-1);
	newrow.id='trX'+imageindex;
	newcell=newrow.insertCell(-1);
	newcell.style.border='1px solid';
	newcell.align='center';
	newcell.innerHTML='<input type="checkbox" title="Select" name="selectimage" value="'+imageindex+'" id="selectimage'+imageindex+'" />';
	newcell=newrow.insertCell(-1);
	newcell.style.border='1px solid';
	newcell.innerHTML=imageurl
	newcell=newrow.insertCell(-1);
	newcell.style.border='1px solid';
	newcell.align='center';
	newcell.innerHTML='<input type="button" value="Display Product" onclick="displayproduct(\''+imageindex+'\')" />';
	imageindex++;
}
function updatefound(){
	document.getElementById('foundspan').innerHTML=foundimages;
	document.getElementById('missingspan').innerHTML=missingimages;
	document.getElementById('checkedspan').innerHTML=foundimages+missingimages;
}
function vsdecimg(timg){
	return decodeURIComponent(timg.replace('|','prodimages/').replace('<','.gif').replace('>','.jpg').replace('?','.png'));
}
function ChK(imageurl){
	imageurl=vsdecimg(imageurl);
	if(chkExists(imageurl)){
		foundimages++;
	}else{
		missingimages++;
		addtotable(imageurl,false);
	}
	updatefound();
}
var ajaxobj;
function confirmdelimg(){
	if(ajaxobj.readyState==4){
		var retvals=ajaxobj.responseText;
		var row=document.getElementById('trX'+retvals);
		row.parentNode.removeChild(row);
	}
}
function deleteimages(deleteall){
	var selectimg=document.getElementById('imagesform').getElementsByTagName("input");
	var numselected=0;
	var idstack=[];
	for(var myobjind in selectimg){
		var myobj=selectimg[myobjind];
		if(myobj.id&&myobj.id.substr(0,11)=='selectimage'){
			if(myobj.checked||deleteall){
				var theid=myobj.id.substr(11);
				idstack.push(myobj.id.substr(11));
				numselected++;
			}
		}
	}
	if(!deleteall&&numselected==0){
		alert('No Images Selected.');
		return;
	}else{
		for(var myobjind in idstack){
			var theid=idstack[myobjind];
			ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
			ajaxobj.onreadystatechange=confirmdelimg;
			ajaxobj.open("GET", "admindbutility.asp?act=delimages&iname="+theid+"&iu="+encodeURIComponent(imagearray[theid]),false);
			ajaxobj.send(null);
		}
	}
}
function finishup(shownextpage,currpage,numpages){
	imgtable = document.getElementById('imagestable');
	newrow=imgtable.insertRow(-1);
	newcell=newrow.insertCell(-1);
	newcell.style.border='1px solid';
	newcell.height=35;
	newcell.align='center';
	newcell.colSpan=3;
	if(foundimages+missingimages>0&&eCTtotalTime>0.01)
		document.getElementById('totaltimespan').innerHTML="("+(Math.round(eCTtotalTime/10)/100)+" seconds / "+((Math.round(((foundimages+missingimages)/eCTtotalTime)*100000))/100)+" per second)";
	if(imageindex==0){
		newcell.innerHTML='No broken image links detected';
	}else
		newcell.innerHTML='<input type="button" value="Delete Selected" onclick="deleteimages(false)" />&nbsp;&nbsp;<input type="button" value="Delete All" onclick="deleteimages(true)" />';
	if(shownextpage) newcell.innerHTML+='&nbsp;&nbsp;<input type="button" value="Next Batch ('+(currpage+1)+' of '+numpages+')" onclick="document.getElementById(\'nextbatchform\').submit()" />';
}
function checkall(tobj){
	var selectimg=document.getElementById('imagesform').getElementsByTagName("input");
	for(var myobjind in selectimg){
		var myobj=selectimg[myobjind];
		if(myobj.id&&myobj.id.substr(0,11)=='selectimage') myobj.checked=tobj.checked
	}
}
/* ]]> */
</script>
<form method="post" id="nextbatchform" action="admindbutility.asp">
<input type="hidden" name="pg" value="<% if NOT is_numeric(getpost("pg")) then print 2 else print int(getpost("pg"))+1%>" />
<input type="hidden" name="batchesof" value="<%=getpost("batchesof")%>" />
<input type="hidden" name="mainaction" value="checkimages">
<input type="hidden" name="subaction" value="go">
<input type="hidden" name="posted" value="1">
</form>
<form method="post" name="imagesform" id="imagesform" action="">
<table id="imagestable" name="imagestable" class="imagestable" width="600" cellspacing="2" cellpadding="2" style="border-collapse:collapse;border:1px solid black;margin:0 auto">
<tr height="35"><td colspan="3" align="center" style="border:1px solid"><strong>Checked:</strong> <span id="checkedspan">0</span>, <strong>Found:</strong> <span id="foundspan">0</span>, <strong>Not Found:</strong> <span id="missingspan">0</span> <span id="totaltimespan"></span></td></tr>
<tr><td align="center" style="border:1px solid"><input type="checkbox" title="Select All" onchange="checkall(this)" /></td><td style="border:1px solid"><strong>Missing Image URL</strong></td><td align="center" style="border:1px solid"><strong>Display Product</strong></td></tr>
</table>
</form>
<%		fijs="<script>/* <![CDATA[ */"&vbCrLf
		fijs=fijs&"var eCTtimer=new Date().getTime();"&vbCrLf
		sSQL="SELECT DISTINCT imageSrc FROM productimages ORDER BY imageSrc"
		if is_numeric(getpost("batchesof")) then batchesof=int(getpost("batchesof")) else batchesof=""
		iNumOfPages=1
		CurPage=1
		success=TRUE
		tpagesize=0
		if batchesof<>"" then
			rs.CursorLocation=3 ' adUseClient
			rs.CacheSize=batchesof
			rs.open sSQL,cnn
			if rs.eof or rs.bof then
				success=FALSE
				iNumOfPages=0
			else
				success=TRUE
				rs.MoveFirst
				rs.PageSize=batchesof
				tpagesize=batchesof
				CurPage=1
				if is_numeric(getpost("pg")) then CurPage=int(getpost("pg"))
				iNumOfPages=int((rs.RecordCount + (batchesof-1)) / batchesof)
				rs.AbsolutePage=CurPage
			end if
		else
			rs.open sSQL,cnn,0,1
		end if
		Count=0
		do while NOT rs.EOF AND (Count< tpagesize OR batchesof="")
			fijs=fijs&"ChK('"&dbuencodeimage(rs("imageSrc"))&"');"&vbCrLf
			Count=Count+1
			rs.movenext
		loop
		fijs=fijs&"var eCTtotalTime=new Date().getTime()-eCTtimer;"&vbCrLf
		fijs=fijs&"finishup("&IIfVr(CurPage<iNumOfPages,"true","false")&","&CurPage&","&iNumOfPages&");"&vbCrLf
		fijs=fijs&"/* ]]> */</script>"&vbCrLf
		response.write fijs
		call gobackhome()
	else
		numimages=0
		if sqlserver then sSQL="SELECT COUNT(DISTINCT imageSrc) AS numimages FROM productimages" else sSQL="SELECT COUNT(*) AS numimages FROM (SELECT DISTINCT imageSrc FROM productimages)"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			if NOT isnull(rs("numimages")) then numimages=rs("numimages")
		end if %>
<table style="margin:0 auto">
<tr><td width="100%">
<%		if instr(request.servervariables("HTTP_USER_AGENT"),"Firefox")>0 then %>
<form action="admindbutility.asp" method="post">
<input type="hidden" name="mainaction" value="checkimages">
<input type="hidden" name="subaction" value="go">
<input type="hidden" name="posted" value="1">
<p>This script will check your Ecommerce Plus templates database for broken image links.</p>
<p>Please make sure you have backed up your site and database before proceeding.</p>
<p>&nbsp;</p>
<%			if numimages>1500 then %>
<p>
	<select name="batchesof">
	<option value="">Check all</option>
	<option value="1000">Check in batches of 1000</option>
	<option value="5000">Check in batches of 5000</option>
	<option value="10000" selected="selected">Check in batches of 10000</option>
	<option value="20000">Check in batches of 20000</option>
	<option value="30000">Check in batches of 30000</option>
	<option value="40000">Check in batches of 40000</option>
	</select>
</p><p>&nbsp;</p><p>
	<select name="pg">
	<option value="1">Start at beginning</option>
<%				for index=1 to 20 %>
	<option value="<%=index+1%>">Skip first <%=index%> batch(s)</option>
<%				next %>
	</select>
</p><p>&nbsp;</p>
<%			end if %>
<p>Please click below to start checking<% if numimages>0 then response.write " <strong>(" & numimages & " images)</strong>"%>.</p>
<p>&nbsp;</p>
<p><input style="background:#399908; color:#fff;cursor:pointer;padding:5px 10px;-moz-border-radius:10px;-webkit-border-radius:10px" type="submit" value="Check for Broken Image Links" /></p>
<p>&nbsp;</p>
<p>Please note, images can only be checked at the rate of about 5 a second. If you have<br />many images this process can take a long time so please be patient.</p>
</form>
<%		else %>
<p>We're sorry but this script is current only designed to work in the Firefox web browser.</p>
<p>&nbsp;</p>
<p><input style="background:#399908; color:#fff;cursor:pointer;padding:5px 10px;-moz-border-radius:10px;-webkit-border-radius:10px" type="button" value="Return to Database Utilities" onclick="document.location='admindbutility.asp'" /></p>
<%		end if
		call gobackhome() %>
</td></tr>
</table>
<%	end if
else
%>
<form id="mainactionform" action="admindbutility.asp" method="post">
<input type="hidden" name="mainaction" id="mainaction" value="">
<div style="margin:30px">
	<table style="margin:0 auto">
	<tr><td colspan="2"><strong>Please choose from the following options</strong></td></tr>
<%	if ectdemostore then %>
	<tr><td colspan="2"><div style="padding:30px 0px">This function has been disabled for the demo store.</div></td></tr>
<%	else %>
	<tr><td>Check database integrity</td><td><input type="button" value="Go" onclick="document.getElementById('mainaction').value='checkdatabase';document.getElementById('mainactionform').submit()" /></td></tr>
<%		if FALSE then %>
	<tr><td>Recreate database caches</td><td><input type="button" value="Go" onclick="document.getElementById('mainaction').value='buildcache';document.getElementById('mainactionform').submit()" /></td></tr>
<%		end if %>
	<tr><td>Check for broken image links</td><td><input type="button" value="Go" onclick="document.getElementById('mainaction').value='checkimages';document.getElementById('mainactionform').submit()" /></td></tr>
	<tr><td>Check for unreferenced product images</td><td><input type="button" value="Go" onclick="document.getElementById('mainaction').value='checkimagerefs';document.getElementById('mainactionform').submit()" /></td></tr>
	<tr><td>Reorder images in database</td><td><input type="button" value="Go" onclick="document.getElementById('mainaction').value='reorderimages';document.getElementById('mainactionform').submit()" /></td></tr>
	<tr><td>NVARCHAR to VARCHAR</td><td><input type="button" value="Go" onclick="document.getElementById('mainaction').value='nvarchar';document.getElementById('mainactionform').submit()" /></td></tr>
	<tr><td>VARCHAR(MAX) / TEXT / VARCHAR(8000)</td><td><input type="button" value="Go" onclick="document.getElementById('mainaction').value='varcharmax';document.getElementById('mainactionform').submit()" /></td></tr>
<%	end if %>
	</table>
</div>
</form>
<%
end if
cnn.close
set rs = nothing
set rs2= nothing
set cnn = nothing
%>