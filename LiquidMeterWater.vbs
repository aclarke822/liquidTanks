'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Get the current values from OIL_PAD, OIL_WELLS & SWD_WELLS
'
'Liquid Meter Water Script to FTP Upload for ProdView Format 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'Liquid Meter - Water Meters
'Facility Types: OIL_PAD , OIL_WELLS, SWD_WELLS

'On Error Resume Next
'SiteService is FRANK
Dim siteService, siteService2, LogLevel
siteService = "FRANK.UIS" 
siteService2 = "FRANK.PNT"

'Set Log Level 
'| 0 = No Log
'| 1 = Important
'| 2 = Everything
LogLevel = 2

'Now - Date get
Dim NowDateGet
NowDateGet = Month(now()) &"-"& Day(now()) &"-"& Year(now()) & " " & Hour(now()) &"-"& Minute(now()) &"-"& Second(now())

'Making Objects
Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
Dim objFac : Set objFac = CreateObject("CxScript.Facilities")
Dim PntClient : Set PntClient = CreateObject("CxPnt.PntClient")
Dim objPoints : Set objPoints = CreateObject("CxScript.Points")	
Dim objXML : Set objXML = CreateObject("Msxml2.DOMDocument.6.0")
Dim fileOut : Set fileOut = objFso.CreateTextFile("C:\CygNet\Scripts\TempGood\LiquidMeterWater " & NowDateGet & ".csv")
Dim logFile : Set logFile = objFso.CreateTextFile("C:\CygNet\Scripts\TempLog\LiquidMeterWaterLog.txt")
Dim BadFile : Set BadFile = objFso.CreateTextFile("C:\CygNet\Scripts\Bad\LiquidMeterWaterBad.txt")
Dim objGlobFunc : Set objGlobFunc = CreateObject("CxScript.GlobalFunctions")
Dim currPointObj : Set currPointObj = CreateObject("CxScript.Points")

'Connect
PntClient.Connect(siteService2)

objFac.UpdateNow()

'Log 
Call LogHeader(LogLevel)
Call WriteLogSucc("Successfully created all objects",LogLevel)

'Global Function Objects
Dim f_site : f_site = "FRANK"
Dim f_status : f_status = f_site & ".UIS:BATCH_STATUS_LIQWAT"
objGlobFunc.EnableLiveMode True
objGlobFunc.setpoint f_status, "In-Progress", now

'Create the Facility TagLists 
Dim TagListPAD, TagListWELLS, TagListSWD
objFac.GetFacilityTagList siteService, "facility_is_active=Y;facility_type=OIL_PAD",TagListPAD
objFac.GetFacilityTagList siteService, "facility_is_active=Y;facility_type=OIL_WELLS",TagListWELLS
objFac.GetFacilityTagList siteService, "facility_is_active=Y;facility_type=SWD_WELLS",TagListSWD

'Log 
Call WriteLogSucc("Successfully grabbed each Facility Type's Facility TagLists",LogLevel)

'Print the header of the ProdView 
fileOut.Writeline "prodview scada import"
fileOut.Writeline "1.0"
fileOut.Writeline "imperial"

BadFile.Writeline "Facility ID & Tank Description | Reason for Failure"  
BadFile.Writeline "---------------------------------------------------"

'Call Functions Here | Input: Facility TagLists
GetXMLCurrentValues(TagListPAD)
GetXMLCurrentValues(TagListWELLS)
GetXMLCurrentValues(TagListSWD)

'Log 
Call WriteLogSucc("Successful Point Retrieval",LogLevel)

'FTP Copy over
Call Copy("C:\CygNet\Scripts\TempGood", "\\wstr.com\ftp\CYGNET\PRD", "LiquidMeterWater " & NowDateGet & ".csv")
fileOut.close
logFile.close
BadFile.close

'Archive 
Call Archive("C:\CygNet\Scripts\TempGood\LiquidMeterWater " & NowDateGet & ".csv", "C:\CygNet\Scripts\Archive\LiquidMeterWater " & NowDateGet & ".csv")
Call Archive("C:\CygNet\Scripts\TempLog\LiquidMeterWaterLog.txt", "C:\CygNet\Scripts\Archive\LiquidMeterWaterLog " & NowDateGet & ".txt")

'Status report
if err.number = 0 then 
	objGlobFunc.setpoint f_status, "Complete", now
elseif err.number <> 0 then 
	objGlobFunc.setpoint f_status, "Complete w Errors", now
end if

'===============================
'=== Functions & Subroutines ===
'===============================

Function GetXMLCurrentValues(TagListArray)
	'The input is an entire TagList from the GetFacilityTagList function 
	
	'Declare 6 element array for UDC strings, XML String
	Dim arrStrUDC(), strXML, arrPoints(), elementCounter, UDCcounter, ArrayElementTool, tag, strNodes, maxCount
	Dim pntArrayCnt, tagpart1, strPntXML, facType, tempUDC
	pntArrayCnt = 0
	ArrayElementTool = UBound(TagListArray)
	
	'This arrStrUDC array is hard coded based on the specific UDCs you want to use
	Redim  arrStrUDC(0)
	
	Dim strFacID
	'An array of the UDCs as strings [0 to 5]
	arrStrUDC(0) = "VWY"
	
	'Variables used for first double for loop counters
	maxCount = (UBound(arrStrUDC)+1) * (ArrayElementTool + 1)
	Redim arrPoints(maxCount)
	
	'Setting up good and bad file dictioinaries
	Dim D1 : Set D1 = CreateObject("Scripting.Dictionary")
	Dim D1Bad : Set D1Bad = CreateObject("Scripting.Dictionary")
	Dim D1TimeBad : Set D1TimeBad = CreateObject("Scripting.Dictionary") 
	
	'Log
	Call WriteLogInfo("Getting " & maxCount & " data points",LogLevel)
	Call WriteLogInfo("Total # of Facilities: " & ArrayElementTool + 1,LogLevel)
	
	'Loop to create a long XML string
	strXML = "<cygPtInfo><Parameters><Value /><timestamp /><activestatus /></Parameters><Points>"
	for elementCounter=0 to ArrayElementTool
		strFacID = TagListArray(elementCounter)
		D1.Add strFacID, CreateObject("Scripting.Dictionary")
		D1.Item(strFacID).Add "Desc", objFac.GetFacilityAttribute(strFacID, "FACILITY_DESC")
		D1.Item(strFacID).Add "FacID", Replace(strFacID,"FRANK.UIS::","")
		tagpart1 = Replace(TagListArray(elementCounter),"::",":")
		facType = objFac.GetFacilityAttribute(strFacID, "FACILITY_TYPE")
		for UDCcounter=0 to UBound(arrStrUDC)
			tag = tagpart1 & "_" & arrStrUDC(UDCcounter)
			if facType = "OIL_WELLS" Then
				tempUDC = "VWY0"
				tag = tagpart1 & "_" & tempUDC
			elseif facType = "SWD_WELLS" Then
				tempUDC = "INJVOLPD"
				tag = tagpart1 & "_" & tempUDC
			end if 
			strXML = strXML & "<node cygTag=" & chr(34) & tag & chr(34) & " />"
			If pntArrayCnt =< maxCount then 
				arrPoints(pntArrayCnt) = tag
				pntArrayCnt = pntArrayCnt + 1
			End If
		next
	next
	strXML = strXML & "</Points></cygPtInfo>"
	
	'Log
	Call WriteLogSucc("String XML created",LogLevel)
	
	'Creating the XML object with an array of the points
	objPoints.AddPointsArray arrPoints, False
	objPoints.ResolveNow 2
	objPoints.UpdateNow 2
	
	strPntXML = objPoints.GetPointsXML(strXML)
	
	objXML.async = False
	objXML.LoadXML strPntXML
	
	'This makes strNodes = all of the Points in the XML string from last Nested For Loop
	Set strNodes = objXML.documentElement.SelectSingleNode("//cygPtInfo/Points").childNodes
	
	'Log
	Call WriteLogSucc("Ready to get attributes from the string XML.",LogLevel)
	
	'Variables for dictionary and child nodes || Must go thru the child nodes (they will be in a random order); so to go thru this list we must get the attributes from each child node
	Dim child, strValue, strCygTag, strFacTag, strUdc, strActiveStatus, strTimeStamp, strPointID
	For Each child in strNodes
		strValue = CheckValue(child.getAttribute("Value"))
		strCygTag = child.getAttribute("cygTag")
		strFacTag = GetFacTag(strCygTag)
		strUdc = GetUDC(strCygTag)
		strActiveStatus = child.getAttribute("activestatus")
		
		
		'Add CygTag as the Key to the Dictionary; then add the Value as the dictionary's value
		If strActiveStatus = "1" Then
			D1.Item(strFacTag).Add "Value", strValue
			D1.Item(strFacTag).Add "UDC", strUdc
			Call WriteLogInfo("Writing " & strCygTag & " to Good File", LogLevel)
			strPointID = currPointObj.Point(strFacTag &"."& strUdc).GetAttribute("pointid")
			D1.Item(strFacTag).Add "PointID", strPointID
			'No need to add an extra representation of the UDC, we already have the UDC
				'if strUDC = "VWY" Then 
				'	D1.Item(strFacTag).Add strUdc, strValue
				'	D1.Item(strFacTag).Add "UDC",1
				'	Call WriteLogInfo("Writing " & strCygTag & " to Good File", LogLevel)
				'	strPointID = currPointObj.Point(strFacTag &"."& strUdc).GetAttribute("pointid")
				'	D1.Item(strFacTag).Add "PointID", strPointID
				'Elseif strUDC = "VWY0" Then
				'	D1.Item(strFacTag).Add strUdc, strValue
				'	D1.Item(strFacTag).Add "UDC",2
				'	Call WriteLogInfo("Writing " & strCygTag & " to Good File", LogLevel)
				'	strPointID = currPointObj.Point(strFacTag &"."& strUdc).GetAttribute("pointid")
				'	D1.Item(strFacTag).Add "PointID", strPointID
				'Elseif strUDC = "INJVOLPD" Then 
				'	D1.Item(strFacTag).Add strUdc, strValue
				'	D1.Item(strFacTag).Add "UDC",3
				'	Call WriteLogInfo("Writing " & strCygTag & " to Good File", LogLevel)
				'	strPointID = currPointObj.Point(strFacTag &"."& strUdc).GetAttribute("pointid")
				'	D1.Item(strFacTag).Add "PointID", strPointID
				'Else
				'	D1.Item(strFacTag).Add "UDC",4
				'End If 
		ElseIf strActiveStatus = "0" Then
			D1Bad.Add strFacTag, CreateObject("Scripting.Dictionary")
			D1Bad.Item(strFacTag).Add strUdc, "Inactive"
			Call WriteLogInfo("Writing " & strCygTag & " to Bad File", LogLevel)
		' Elseif strActiveStatus = "Null" AND strUDC = "VWY" Then
			' if strValue <> "" Then
				' D1.Item(strFacTag).Add "Value", strValue
			    ' D1.Item(strFacTag).Add "UDC", strUdc
				' Call WriteLogInfo("Writing " & strCygTag & " to Bad File", LogLevel)
			' elseif strValue = "" Then
				' D1Bad.Add strCygTag, "Null Active Status and not printed b/c no value"
				' Call WriteLogInfo("Writing " & strCygTag & " to Bad File", LogLevel)
			' end if
		Else 
			D1Bad.Add strFacTag, CreateObject("Scripting.Dictionary")
			D1Bad.Item(strFacTag).Add strUdc, ""
			Call WriteLogInfo("Writing " & strCygTag & " to Bad File", LogLevel)
		End If
		'Site.Service:FacID_UDC
	Next
	
	'Variables 
	Dim i, j, k, arrD1Keys, printDate, TimeStampVal, BSandW
	printDate = Date() - 1
	TimeStampVal = CheckTimeStamp(printDate)
	BSandW = 100
	
	'Now we print out our Dictionary | Log
	arrD1Keys = D1.Keys
	Call WriteLogSucc("Dictionary of Created",LogLevel)
	
	For i = 0 to UBound(arrD1Keys)
		If D1.Item(arrD1Keys(i)).Item("Value") >= 0 Then
			fileOut.Writeline "LIQUID METER," &_
			D1.Item(arrD1Keys(i)).Item("Desc") &" Water" &","&_ 
									TrimLZ(D1.Item(arrD1Keys(i)).Item("PointID")) &","&_
									TimeStampVal &","&_ 
									D1.Item(arrD1Keys(i)).Item("Value") &","&_
									BSandW
		Else
			D1Bad.Add arrD1Keys(i), "Negative Value"
		End If
	Next
	
	'Log
	Call WriteLogSucc("Facility Type " & facType & " has finished getting values. ",LogLevel)
	
	'Variables for Bad dictionary
	Dim arrBadPoints, arrBadReason
	arrBadPoints = D1Bad.Keys 
	arrBadReason = D1Bad.Items
	
	'Print bad dictionary
	For j = 0 to UBound(arrBadPoints)
		BadFile.Writeline "Bad Point: " & arrBadPoints(j)
	Next
	
	'Log
	Call WriteLogSucc("Facility Type finished getting values. ",LogLevel)
End Function

Sub WriteLogSucc(str,level)
	If level => 2 Then 
		logFile.Writeline now &" - "& str
	End If 
End Sub

Sub WriteLogInfo(str,level)
	If level => 1 Then 
		logFile.Writeline now &" - "& str
	End If 
End Sub

Sub LogHeader(level)
	If level => 1 Then 
		
		logFile.Writeline now &" - Time Log Starts"
		logFile.Writeline ""
		logFile.Writeline "Begin log: "
	End If 
End Sub

Function CheckValue(value)
	'check
	If not IsNull(value) Then
		If Len(value) > 0 Then
			On Error Resume Next
				CheckValue = CInt(Replace(value," ",""))
				If err.Number > 0 Then
					Wscript.Echo Err.Description
					Err.clear
				End If
			On Error Goto 0
		Else
			CheckValue = -9999
		End If
	Else
		CheckValue = -9999
	End If
End Function

Function GetFacTag(PntTag)
	Dim strUdcFunct
	strUdcFunct = Split(PntTag, "_")(UBound(Split(PntTag, "_")))
	GetFacTag = Replace(Replace(PntTag, "_" & strUdcFunct,""),":","::")
End Function

Function GetUDC(PntTag)
	GetUDC = Split(PntTag, "_")(UBound(Split(PntTag, "_")))
End Function

Function CheckTimeStamp(TimeStamp)
	Dim day, month
	if not IsNull(TimeStamp) Then
		month = DatePart("m", TimeStamp)
		day = DatePart("d", TimeStamp)
		if month < 10 Then month = "0" & month
		if day < 10 Then day = "0" & day
		
		CheckTimeStamp = Year(TimeStamp) & month & day
	elseif IsNull(TimeStamp) Then
		CheckTimeStamp = ""
		
	End If 
End Function

Function FTPUpload(sSite, sUsername, sPassword, sLocalFile, sRemotePath)
	Const OpenAsDefault = -2
	Const FailIfNotExist = 0
	Const ForReading = 1
	Const ForWriting = 2
	Dim oFTPScriptFSO, oFTPScriptShell, sFTPScript, sFTPTemp, sFTPTempFile, sFTPResults, fFTPScript, fFTPResults, sResults
	Set oFTPScriptFSO = CreateObject("Scripting.FileSystemObject")
	Set oFTPScriptShell = CreateObject("WScript.Shell")

	sRemotePath = Trim(sRemotePath)
	sLocalFile = Trim(sLocalFile)
	  
	  '----------Path Checks---------
	  'Here we willcheck the path, if it contains
	  'spaces then we need to add quotes to ensure
	  'it parses correctly.
	If InStr(sRemotePath, " ") > 0 Then
		If Left(sRemotePath, 1) <> """" And Right(sRemotePath, 1) <> """" Then
			sRemotePath = chr(34) & sRemotePath & chr(34)
		End If
	End If
	  
	If InStr(sLocalFile, " ") > 0 Then
		If Left(sLocalFile, 1) <> """" And Right(sLocalFile, 1) <> """" Then
			sLocalFile = chr(34) & sLocalFile & chr(34)
		End If
	End If

	'Check to ensure that a remote path was
	'passed. If it's blank then pass a "\"
	If Len(sRemotePath) = 0 Then
	'Please note that no premptive checking of the
	'remote path is done. If it does not exist for some
	'reason. Unexpected results may occur.
		sRemotePath = "\"
	End If
	
	'Check the local path and file to ensure
	'that either the a file that exists was
	'passed or a wildcard was passed.
	If InStr(sLocalFile, "*") Then
		If InStr(sLocalFile, " ") Then
			FTPUpload = "Error: Wildcard uploads do not work if the path contains a space." & vbCRLF
			FTPUpload = FTPUpload & "This is a limitation of the Microsoft FTP client."
			Exit Function
		End If
	ElseIf Len(sLocalFile) = 0 Or Not oFTPScriptFSO.FileExists(sLocalFile) Then
	'nothing to upload
		FTPUpload = "Error: File Not Found."
		Exit Function
	End If
	'--------END Path Checks---------
	  
	  'build input file for ftp command
	sFTPScript = sFTPScript & "USER " & sUsername & vbCRLF
	sFTPScript = sFTPScript & sPassword & vbCRLF
	sFTPScript = sFTPScript & "cd " & sRemotePath & vbCRLF
	sFTPScript = sFTPScript & "binary" & vbCRLF
	sFTPScript = sFTPScript & "prompt n" & vbCRLF
	sFTPScript = sFTPScript & "put " & sLocalFile & vbCRLF
	sFTPScript = sFTPScript & "quit" & vbCRLF & "quit" & vbCRLF & "quit" & vbCRLF
	
	sFTPTemp = oFTPScriptShell.ExpandEnvironmentStrings("%TEMP%")
	sFTPTempFile = sFTPTemp & "\" & oFTPScriptFSO.GetTempName
	sFTPResults = sFTPTemp & "\" & oFTPScriptFSO.GetTempName
	
	'Write the input file for the ftp command
	'to a temporary file.
	Set fFTPScript = oFTPScriptFSO.CreateTextFile(sFTPTempFile, True)
	fFTPScript.WriteLine(sFTPScript)
	
	fFTPScript.Close
	Set fFTPScript = Nothing  
	
	oFTPScriptShell.Run "%comspec% /c FTP -n -s:" & sFTPTempFile & " " & sSite & " > " & sFTPResults, 0, TRUE
	
	Wscript.Sleep 1000
	
	'Check results of transfer.
	Set fFTPResults = oFTPScriptFSO.OpenTextFile(sFTPResults, ForRading, FailIfNotExist, OpenAsDefault)
	sResults = fFTPResults.ReadAll
	fFTPResults.Close
	
	oFTPScriptFSO.DeleteFile(sFTPTempFile)
	oFTPScriptFSO.DeleteFile (sFTPResults)
	
	If InStr(sResults, "226 Successfully transferred") > 0 Then
		FTPUpload = True
	ElseIf InStr(sResults, "File not found") > 0 Then
		FTPUpload = "Error: File Not Found"
	ElseIf InStr(sResults, "cannot log in.") > 0 Then
		FTPUpload = "Error: Login Failed."
	Else
		FTPUpload = "Error: Unknown."
	End If

	MsgBox FTPUpload

	Set oFTPScriptFSO = Nothing
	Set oFTPScriptShell = Nothing
End Function

Function Copy(source, destination, file)
	Call WriteLogInfo("Copying files to remote location...",LogLevel)
	Dim WshShellScriptExec, WSHShell, strLogFile, strCmd
	Set WSHShell = CreateObject("Wscript.Shell")
	strLogFile = source & "\CopyProcess.Log"
	strCmd = "robocopy """ & source & """ """ & destination & """ """ & file & """ /XO /NFL /NDL /NP /R:0 /W:1 /LOG+:""" & strLogFile &""""
	
	Call WriteLogInfo("Cmd: " & strCmd,LogLevel)
	WshShellScriptExec = WshShell.Run(strCmd, 0, True)
	
	Call WriteLogInfo("End of Copy. File copy status: " & WshShellScriptExec,LogLevel) 
End Function

Function Archive(source, archivedFile)
	Dim myFSO : Set myFSO = CreateObject("Scripting.FileSystemObject")
	myFSO.MoveFile source, archivedFile 
End Function

Function TrimLZ(str)
	If Left(str, 1) = 0 Then
		TrimLZ = TrimLZ(Mid(str, 2, Len(str)))
	Else 
		TrimLZ = str
	End If
End Function