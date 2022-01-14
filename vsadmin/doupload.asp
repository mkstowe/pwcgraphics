<%
Response.Buffer = True
'=========================================
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protect under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
isadmincsv=TRUE
%>
<!--#include file="db_conn_open.asp"-->
<!--#include file="includes.asp"-->
<!--#include file="inc/languageadmin.asp"-->
<%
if storesessionvalue="" then storesessionvalue="virtualstore"
if NOT disallowlogin then
<!--#include file="inc/incloginfunctions.asp"-->
end if
if SESSION("loggedon")<>storesessionvalue OR disallowlogin=TRUE then response.redirect "login.asp"
isprinter=false
CrLf = Chr(13) & Chr(10)
Session.LCID = 1033
%>
<!doctype html>
<html>
<head>
<title>Admin Upload</title>
<link rel="stylesheet" type="text/css" href="adminstyle.css"/>
</head>
<body>

<%
'***************************************
' File:	  Upload.asp
' Author: Jacob "Beezle" Gilley
' Email:  avis7@airmail.net
' Date:   12/07/2000
' Comments: The code for the Upload, CByteString, 
'			CWideString	subroutines was originally 
'			written by Philippe Collignon...or so 
'			he claims. Also, I am not responsible
'			for any ill effects this script may
'			cause and provide this script "AS IS".
'			Enjoy!
'****************************************

Class FileUploader
	Public  Files
	Private mcolFormElem

	Private Sub Class_Initialize()
		Set Files = Server.CreateObject("Scripting.Dictionary")
		Set mcolFormElem = Server.CreateObject("Scripting.Dictionary")
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(Files) Then
			Files.RemoveAll()
			Set Files = Nothing
		End If
		If IsObject(mcolFormElem) Then
			mcolFormElem.RemoveAll()
			Set mcolFormElem = Nothing
		End If
	End Sub

	Public Property Get Form(sIndex)
		Form = ""
		If mcolFormElem.Exists(LCase(sIndex)) Then Form = mcolFormElem.Item(LCase(sIndex))
	End Property

	Public Default Sub Upload()
		Dim biData, sInputName
		Dim nPosBegin, nPosEnd, nPos, vDataBounds, nDataBoundPos
		Dim nPosFile, nPosBound

		if ectdemostore=TRUE then biData="" else biData=Request.BinaryRead(Request.TotalBytes)
		nPosBegin = 1
		nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))
		
		If (nPosEnd-nPosBegin) <= 0 Then Exit Sub
		 
		vDataBounds = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
		nDataBoundPos = InstrB(1, biData, vDataBounds)
		
		Do Until nDataBoundPos = InstrB(biData, vDataBounds & CByteString("--"))
			nPos = InstrB(nDataBoundPos, biData, CByteString("Content-Disposition"))
			nPos = InstrB(nPos, biData, CByteString("name="))
			nPosBegin = nPos + 6
			nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(34)))
			sInputName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
			nPosFile = InstrB(nDataBoundPos, biData, CByteString("filename="))
			nPosBound = InstrB(nPosEnd, biData, vDataBounds)
			
			If nPosFile <> 0 And  nPosFile < nPosBound Then
				Dim oUploadFile, sFileName
				Set oUploadFile = New UploadedFile
				
				nPosBegin = nPosFile + 10
				nPosEnd =  InstrB(nPosBegin, biData, CByteString(Chr(34)))
				sFileName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
				oUploadFile.FileName = Right(sFileName, Len(sFileName)-InStrRev(sFileName, "\"))

				nPos = InstrB(nPosEnd, biData, CByteString("Content-Type:"))
				nPosBegin = nPos + 14
				nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))
				
				oUploadFile.ContentType = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
				
				nPosBegin = nPosEnd+4
				nPosEnd = InstrB(nPosBegin, biData, vDataBounds) - 2
				oUploadFile.FileData = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
				
				If oUploadFile.FileSize > 0 Then Files.Add LCase(sInputName), oUploadFile
			Else
				nPos = InstrB(nPos, biData, CByteString(Chr(13)))
				nPosBegin = nPos + 4
				nPosEnd = InstrB(nPosBegin, biData, vDataBounds) - 2
				If Not mcolFormElem.Exists(LCase(sInputName)) Then mcolFormElem.Add LCase(sInputName), CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
			End If

			nDataBoundPos = InstrB(nDataBoundPos + LenB(vDataBounds), biData, vDataBounds)
		Loop
	End Sub

	'String to byte string conversion
	Private Function CByteString(sString)
		Dim nIndex
		For nIndex = 1 to Len(sString)
		   CByteString = CByteString & ChrB(AscB(Mid(sString,nIndex,1)))
		Next
	End Function

	'Byte string to string conversion
	Private Function CWideString(bsString)
		Dim nIndex
		CWideString =""
		For nIndex = 1 to LenB(bsString)
		   CWideString = CWideString & Chr(AscB(MidB(bsString,nIndex,1))) 
		Next
	End Function
End Class

Class UploadedFile
	Public ContentType
	Public FileName
	Public FileData
	Public FileSuccess
	
	Public Property Get FileSize()
		FileSize = LenB(FileData)
	End Property

	Public Sub SaveToDisk(sPath)
		Dim oFS, oFile
		Dim nIndex
	
		If sPath = "" Or FileName = "" Then Exit Sub
		If Mid(sPath, Len(sPath)) <> "\" Then sPath = sPath & "\"
	
		Set oFS = Server.CreateObject("Scripting.FileSystemObject")
		If Not oFS.FolderExists(sPath) Then Exit Sub
		
		on error resume next
		err.number = 0
		Set oFile = oFS.CreateTextFile(sPath & FileName, True)
		errnum=err.number
		on error goto 0
		if errnum=0 then
			FileSuccess = TRUE
			For nIndex = 1 to LenB(FileData)
			    oFile.Write Chr(AscB(MidB(FileData,nIndex,1)))
			Next
			oFile.Close
		else
			FileSuccess = FALSE
		end if
	End Sub
	
	Public Sub SaveToDatabase(ByRef oField)
		If LenB(FileData) = 0 Then Exit Sub
		
		If IsObject(oField) Then
			oField.AppendChunk FileData
		End If
	End Sub
End Class
function getfullpath(impath)
	if left(impath, 1)="/" then
		impath = server.mappath(impath)
	else
		thispath = server.mappath("dummydir")
		for index=0 to 1
			pos = instrrev(thispath, "\")
			if pos>0 then thispath = left(thispath, pos-(1-index))
		next
		if impath="" then impath="prodimages/"
		impath=replace(impath, "/", "\")
		if left(impath, 1)="\" then impath = right(impath, 2)
		impath = thispath & impath
	end if
	getfullpath = impath
end function
function getabsolutepath(impath)
	impath = replace(impath, "\", "/")
	if left(impath, 1)="/" then
		' Nothing
	else
		thispath = request.servervariables("URL")
		for index=0 to 1
			pos = instrrev(thispath, "/")
			if pos>0 then thispath = left(thispath, pos-(1-index))
		next
		impath = thispath & impath
	end if
	getabsolutepath = impath
end function
function validextension(lfn)
	if right(lfn, 4)=".gif" OR right(lfn, 4)=".jpg" OR right(lfn, 5)=".jpeg" OR right(lfn, 4)=".png" then ' OR right(lfn, 4)=".bmp" OR right(lfn, 4)=".art" OR right(lfn, 4)=".wmf" OR right(lfn, 4)=".emf" OR right(lfn, 4)=".mov" OR right(lfn, 4)=".xbm" OR right(lfn, 4)=".avi" OR right(lfn, 4)=".mpg" OR right(lfn, 5)=".mpeg") then
		validextension = TRUE
	else
		validextension = FALSE
	end if
end function
function validimagecontent()
	validimagecontent = FALSE
	if ascb(midb(File.FileData, 1, 1))=&HFF AND ascb(midb(File.FileData, 2, 1))=&HD8 then validimagecontent=TRUE ' JPEG
	if ascb(midb(File.FileData, 1, 1))=&H89 AND ascb(midb(File.FileData, 2, 1))=&H50 AND ascb(midb(File.FileData, 3, 1))=&H4E AND ascb(midb(File.FileData, 4, 1))=&H47 then validimagecontent=TRUE ' PNG
	if ascb(midb(File.FileData, 1, 1))=&H47 AND ascb(midb(File.FileData, 2, 1))=&H49 AND ascb(midb(File.FileData, 3, 1))=&H46 AND ascb(midb(File.FileData, 4, 1))=&H38 AND ascb(midb(File.FileData, 5, 1))=&H39 then validimagecontent=TRUE ' GIF
end function
sub writeimagejs(impath, fdname)
	impath = replace(impath, "\", "/")
	if right(impath, 1)<>"/" then impath = impath & "/"
	impath = impath & fdname
	response.write "<script>"
	if Uploader.Form("imagefield")="pImage" then response.write "window.opener.document.getElementById('smim0').value='"&impath&"';"
	if Uploader.Form("imagefield")="pLargeImage" then response.write "window.opener.document.getElementById('lgim0').value='"&impath&"';"
	if Uploader.Form("imagefield")="pGiantImage" then response.write "window.opener.document.getElementById('gtim0').value='"&impath&"';"
	response.write "window.opener.document.getElementById('"&Uploader.Form("imagefield")&"').value='"&impath&"';"
	response.write "</script>"
end sub
sub showfiledetails(fdsuccess, fdname, fdsize)
	if fdsuccess then
		response.write "<p align=""center"">&nbsp;<br /><strong>"&yyFileUp&"</strong><br />&nbsp;<br />"
		response.write yyDetai & ": " & fdname & " (" & fdsize & " bytes)</p>"
	else
		writeerror(yyNoWrFl&"<br /><br />"&yyChkFP)
	end if
end sub
sub writeerror(theerr)
	response.write "<p align=""center"">&nbsp;<br /><strong>ERROR! "&theerr&"</strong></p>"
end sub
if Mid(SESSION("loggedonpermissions"),6,1)<>"X" then
	response.write "<table width=""100%"" border=""0"" bgcolor=""""><tr><td width=""100%"" colspan=""4"" align=""center""><p>&nbsp;</p><p>&nbsp;</p><p><strong>"&yyOpFai&"</strong></p><p>&nbsp;</p><p>"&yyNoPer&" <br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;<br /><a href=""admin.asp""><strong>"&yyAdmHom&"</strong></a>.</p><p>&nbsp;</p></td></tr></table>"
elseif ectdemostore=TRUE then
	response.write "<table width=""100%"" border=""0"" bgcolor=""""><tr><td width=""100%"" colspan=""4"" align=""center""><p>&nbsp;</p><p><strong>"&yyOpFai&"</strong></p><p>&nbsp;</p><p>This function has been disabled for the demo store for security reasons.</p><p><a href=""javascript:window.close()""><strong>Close Window</strong></a>.</p><p>&nbsp;</p></td></tr></table>"
else
	response.write "<p>&nbsp;</p>"
	if useaspuploadforimages then
		Set Uploader = Server.CreateObject("Persits.Upload")
		Uploader.SetMaxSize 800000, True
		abspath = getabsolutepath(request.querystring("defimagepath"))
		on error resume next
		err.number = 0
		Uploader.SaveToMemory
		errnum = err.number
		errdesc = err.description
		on error goto 0
		if errnum=0 then
			set File = Uploader.Files("imagefile")
			if NOT File is nothing then
				lcasefn = lcase(File.FileName)
				if instr(File.ContentType, "image/") > 0 AND validextension(lcasefn) AND (File.ImageType="GIF" OR File.ImageType="JPG" OR File.ImageType="PNG") then
					File.SaveAsVirtual abspath & File.FileName
					call writeimagejs(request.querystring("defimagepath"), File.FileName)
					call showfiledetails(TRUE, File.FileName, File.Size)
				else
					writeerror(yyIlFlT)
				end if
			end if
		else
			writeerror(yyNoWrFl&"<br /><br />"&yyChkFP)
			response.write "<p align=""center"">ASPUpload Error Code ("&errnum&")<br/>"&errdesc&"</p>"
		end if
	else
		Dim Uploader, File
		Set Uploader = New FileUploader
		Uploader.Upload()
		if Uploader.Files.Count = 0 then
			writeerror(yyFlNtFn)
		else
			for each File In Uploader.Files.Items
				imagepath = getfullpath(Uploader.Form("defimagepath"))
				lcasefn = lcase(File.FileName)
				if instr(File.ContentType, "image/") > 0 AND validextension(lcasefn) then
					if validimagecontent() then
						File.SaveToDisk imagepath
						call writeimagejs(Uploader.Form("defimagepath"), File.FileName)
						call showfiledetails(File.FileSuccess, File.FileName, File.FileSize)
					else
						writeerror(yyIlFlT)
					end if
				else
					writeerror(yyIlFlT)
				end if
			next
		end if
	end if
	response.write "<p align=""center""><a href=""javascript:window.close()""><strong>"&yyClsWin&"</strong></a></p>"
end if %>
</body>
</html>
