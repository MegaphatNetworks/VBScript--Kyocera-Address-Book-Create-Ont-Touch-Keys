Dim sPath : sPath = getPath()
Dim oFS : Set oFS = CreateObject("Scripting.FileSystemObject")

If isConsole = 0 Then
	sMsg = "Kyocera/Copystar Address Book ""One-Touch"" Addition." & vbCrlf
	sMsg = sMsg & "(c)" & year(now) & " Megaphat Networks, All Rights Reserved" & vblf
	sMsg = sMsg & vblf
	sMsg = sMsg & "Cannot execute script!" & vblf
	sMsg = sMsg & "This script MUST run from CScript (Command Prompt) such as:" & vblf & vblf
	sMsg = sMsg & Left(WScript.ScriptFullName,3) & ">CScript """ & WScript.ScriptFullName & """"
	Msgbox sMsg, vbError
	Die
End If

If wscript.arguments.count = 0 then 
	sMsg = "Kyocera/Copystar Address Book ""One-Touch"" Addition." & vbCrlf
	sMsg = sMsg & "(c)" & year(now) & " Megaphat Networks, All Rights Reserved" & vblf
	sMsg = sMsg & vblf
	sMsg = sMsg & "Cannot execute script!" & vblf
	sMsg = sMsg & "you will need to provide the Path\File of the XML Address Book.  For example:" & vblf & vblf
	sMsg = sMsg & Left(WScript.ScriptFullName,3) & ">CScript """ & WScript.ScriptFullName & """" & " ""C:\Username\Desktop\Kyocera-Copystar-Address-Book.XML""" & vbcrlf & vbCrlf
	sMsg = sMsg & "And yes, you WILL need to add the quotations!"
	doLog sMsg
	Die
End If	

sArg = wscript.arguments.item(0)
sMsg = "Kyocera/Copystar Address Book ""One-Touch"" Addition." & vbCrlf
sMsg = sMsg & "(c)" & year(now) & " Megaphat Networks, All Rights Reserved" & vblf
sMsg = sMsg & vblf

doLog "Deleting any existing log file..."
If oFS.FileExists(sArg & ".log") Then oFS.DeleteFile (sArg & ".log")

doLog "Attempting to process Address Book: " & sArg

If Not oFS.FileExists(sArg) Then 
	sMsg = "File: " & sArg
	sMsg = sMsg & "Does not exist!"
	doLog sMsg
	Die
End If

sLogFile = GetPath & fmtDateTime("{0:yyyyMMdd-HHmmSS}", Array(now)) & ".log"
Set oAF = oFS.OpenTextFile(sLogFile, 8, True)
Set oTS = oFS.OpenTextFile(sArg, 1,True)

doLog "Setting headers and footers..."
sHead1 = "<?xml version=""1.0""?>"
sHead2 = "<DeviceAddressBook_v5_2>"
sFoot1 = "</DeviceAddressBook_v5_2>"
sFoot2 = "<!--"


doLog "Opening " & sArg
Dim saOT: Redim saOT(0)
Dim saAB: Redim saAB(0)
Dim iOT, iAB: iOT = 0: iAB = 0
Do While Not oTS.AtEndOfStream 
	sLine = trim(oTS.ReadLine)
	doLog "READ: " & sLine
	if sLine = sHead1 Then doLog "FOUND Header 1"
	if sLine = sHead2 Then doLog "FOUND Header 2"
	if sLine = sFoot1 Then doLog "FOUND Footer 1"
	if Instr(sLine,sFoot2) > 0 Then doLog "FOUND Footer 2"
	If (sLine <> sHead1) and (sLine <> sHead2) and (sLine <> sFoot1) and (Instr(sLine, sFoot2) = 0) Then
		iAB = iAB +1
		Redim Preserve saAB(iAB) 
		saAB(iAB) = sLine
		doLog "LINE: " & iAB & vbcrlf & sLine & vbcrlf & "- Not a header or footer., adding to AB entry." 
		If Instr(sLine, "<Item Type=""Contact""") > 0 Then 
			doLog "- Retrieving OTK values..."
			iOT = iOT +1
			Redim Preserve saOT(iOT)
			saOT(iOT) = CreateOTK(sLine)
			doLog "OTK: " & saOT(iOT)
		End If
	End If
Loop	
doLog "Finished reading file " & sArg
oTS.Close
oAF.Close
Set oAF = Nothing

doLog ""
doLog ""
doLog "Writing to " & sArg & ".NEW.txt"
fileWrite sArg & ".NEW.txt", sHead1
doLog sHead1
fileAppend sArg & ".NEW.txt", sHead2
doLog sHead2
For j = 1 to iAB
	fileAppend sArg & ".NEW.txt", "  " & saAB(j)
	doLog saAB(j)
Next
For j =  1 to iOT
	fileAppend sArg & ".NEW.txt", saOT(j)
	doLog saOT(j)
Next
fileAppend sArg & ".NEW.txt", sFoot1
doLog sFoot1

doLog ""
doLog "COMPLETE."
Die

Function CreateOTK(sStr)
	sStr = Replace(sStr,"<Item ","")
	sStr = Replace(sStr," />","")
	Dim sItem: Redim sItem(0)
	Dim sA: sA = Split(sStr, Chr(34) & " ")
	Dim oItems : Set oItems = CreateObject("Scripting.Dictionary")
	Dim oItem
	'<Item Type="Contact" Id="35" DisplayName="Daniel W" SendKeisyou="0" DisplayNameKana="Daniel W" MailAddress="" SendCorpName="" SendPostName="" SendAddrName="" SmbHostName="swcfile" SmbPath="\company files\scans\daniel w" SmbLoginName="scanner" SmbLoginPasswd="ED1CE55E30053916" SmbPort="139" FtpPath="" FtpHostName="" FtpLoginName="" FtpLoginPasswd="" FtpPort="21" FaxNumber="" FaxSubaddress="" FaxPassword="" FaxCommSpeed="BPS_9600" FaxECM="Off" FaxEncryptKeyNumber="0" FaxEncryption="Off" FaxEncryptBoxEnabled="Off" FaxEncryptBoxID="0000" InetFAXAddr="" InetFAXMode="Full" InetFAXResolution="3" InetFAXFileType="TIFF_MH" IFaxSendModeType="IFAX" InetFAXDataSize="1" InetFAXPaperSize="1" InetFAXResolutionEnum="Default" InetFAXPaperSizeEnum="Default" />
	'<Item Type="OneTouchKey" Id="1" DisplayName="Adam" AddressId="1" AddressType="SMB" />
  
	For i = 0 to uBound(sA)
		sA(i) = Replace(sA(i),chr(34),"")
		sItem = Split(sa(i),"=")
		oItems.Add sItem(0),sItem(1)
		Erase sItem
	Next
	For Each sKey in oItems
		sItem = sKey
		sVal = oItems.Item(sKey)
		If sItem = "MailAddress" and sVal <>"" Then 
			'Create OTK for Mail Address
			sOTK = "  <Item Type=""OneTouchKey"" Id=""" & oItems.Item("Id") & """ DisplayName=""" 
			sOTK = sOTK & oItems.Item("DisplayName") & """ AddressId=""" & oItems.Item("Id") & """ AddressType=""EMAIL"" />"
		Else
			'Create OTK for SMB
			sOTK = "  <Item Type=""OneTouchKey"" Id=""" & oItems.Item("Id") & """ DisplayName=""" 
			sOTK = sOTK & oItems.Item("DisplayName") & """ AddressId=""" & oItems.Item("Id") & """ AddressType=""SMB"" />"
		End If
	Next
	CreateOTK = sOTK
End Function


Function isConsole()
	If Instr(Wscript.FullName,"cscript") > 0 Then isConsole = 1 Else isConsole = 0
End Function

Function getPath()
	Set oShell = CreateObject("WScript.Shell")
	getPath = oShell.CurrentDirectory & "\"
	Set oShell = Nothing
End Function

Sub doLog(sStr)
    'Set oFile = oFS.OpenTextFile(sPath & "_log.log", 8,True)
	wscript.echo sStr
	fileAppend sArg & ".log", sStr
	'oFile.WriteLine time & " " & sStr
    'oFile.Close
    'Set oFSO = Nothing
End Sub

Sub fileWrite(szFileName, szData)
	Set oThisFile = oFS.OpenTextFile(szFileName, 2,True)

	oThisFile.WriteLine szData
	oThisFile.Close
	Set oThisFile = Nothing
End Sub


Sub fileAppend(szFileName, szData)
	Set oThisFile = oFS.OpenTextFile(szFileName, 8,True)
	oThisFile.WriteLine szData
	oThisFile.Close
	Set oThisFile = Nothing
End Sub

Function fmtDateTime(sFmt, aData)
	''USAGE: fmtDateTime("{0:yyyyMMdd}", Array(now))
	Dim g_oSB : Set g_oSB = CreateObject("System.Text.StringBuilder")
   g_oSB.AppendFormat_4 sFmt, (aData)
   fmtDateTime = g_oSB.ToString()
   g_oSB.Length = 0
End Function


Sub Die()
	WScript.Quit
End Sub