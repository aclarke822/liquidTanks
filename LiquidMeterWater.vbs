'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Get the current values from OIL_PAD, OIL_WELLS & SWD_WELLS
'
'Liquid Meter Water Script to FTP Upload For ProdView Format 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'Liquid Meter - Water Meters
'Facility Types: OIL_PAD , OIL_WELLS, SWD_WELLS

'On Error Resume Next
'siteUIS is FRANK
Dim siteUIS, sitePNT, LogLevel
siteUIS = "FRANK.UIS" 
sitePNT = "FRANK.PNT"

'Set Log Level 
'| 0 = No Log
'| 1 = Important
'| 2 = Everything
LogLevel = 2

'Now - Date get
Dim NowDateGet : NowDateGet = Month(now()) &"-"& Day(now()) &"-"& Year(now()) & " " & Hour(now()) &"-"& Minute(now()) &"-"& Second(now())

'Making Objects
Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
Dim PntClient : Set PntClient = CreateObject("CxPnt.PntClient")
Dim fileOut : Set fileOut = objFso.CreateTextFile("C:\CygNet\Scripts\TempGood\LiquidMeterWater " & NowDateGet & ".csv")
Dim logFile : Set logFile = objFso.CreateTextFile("C:\CygNet\Scripts\TempLog\LiquidMeterWaterLog.txt")
Dim BadFile : Set BadFile = objFso.CreateTextFile("C:\CygNet\Scripts\Bad\LiquidMeterWaterBad.txt")
Dim objGlobFunc : Set objGlobFunc = CreateObject("CxScript.GlobalFunctions")

'Print the header of the ProdView 
fileOut.Writeline "prodview scada import"
fileOut.Writeline "1.0"
fileOut.Writeline "imperial"

BadFile.Writeline "Facility ID & Tank Description | Reason For Failure"  
BadFile.Writeline "---------------------------------------------------"

'Connect
PntClient.Connect(sitePNT)

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
'Needs error checking and testing
Dim arrFacTypes : arrFacTypes = Array("OIL_PAD", "OIL_WELLS", "SWD_WELLS")
For i = 0 to UBound(arrFacTypes)
    Call WriteToFile(PrepareDictionary(GetXMLCurrentValues(siteUIS, arrFacTypes(i)), arrFacTypes(i)))
    Call WriteLogSucc("Successfully processed " & arrFacTypes(i), LogLevel)
Next

'FTP Copy over
Call Copy("C:\CygNet\Scripts\TempGood", "\\wstr.com\ftp\CYGNET\PRD", "LiquidMeterWater " & NowDateGet & ".csv")
fileOut.close
logFile.close
BadFile.close

'Archive 
Call Archive("C:\CygNet\Scripts\TempGood\LiquidMeterWater " & NowDateGet & ".csv", "C:\CygNet\Scripts\Archive\LiquidMeterWater " & NowDateGet & ".csv")
Call Archive("C:\CygNet\Scripts\TempLog\LiquidMeterWaterLog.txt", "C:\CygNet\Scripts\Archive\LiquidMeterWaterLog " & NowDateGet & ".txt")

'Status report
If err.number = 0 then 
	objGlobFunc.setpoint f_status, "Complete", now
ElseIf err.number <> 0 then 
	objGlobFunc.setpoint f_status, "Complete w Errors", now
end If

'===============================
'=== Functions & Subroutines ===
'===============================

Function GetXMLCurrentValues(strSiteServ, strFacType)
    'The input is an entire TagList from the GetFacilityTagList function
    Dim objFac : Set objFac = CreateObject("CxScript.Facilities")
    Dim objPoints : Set objPoints = CreateObject("CxScript.Points")	
    Dim strXML, arrPoints(), i, j, tag, maxCount, arrTagList, pntArrayCnt, strFacType, strFacTag

    objFac.GetFacilityTagList strSiteServ, "facility_is_active=Y;facility_type=" & strFacType, arrTagList
    Redim arrPoints(UBound(arrTagList) + 1)

	'Log
	Call WriteLogInfo(strFacType & "|Getting " & UBound(arrTagList) + 1 & " data points from " & UBound(arrTagList) + 1 & "facilities", LogLevel)
	
	'Loop to create a long XML string
	strXML = "<cygPtInfo><Parameters><Value /><timestamp /><activestatus /></Parameters><Points>"
	For i = 0 to UBound(arrTagList)
		strFacTag = arrTagList(i)
		strFacType = objFac.GetFacilityAttribute(strFacTag, "FACILITY_TYPE")
		For j = 0 to UBound(arrStrUDC)
			If strFacType = "OIL_WELLS" Then
				tag = Replace(strFacTag,"::",":") & "_VWY0"
			ElseIf strFacType = "SWD_WELLS" Then
                tag = Replace(strFacTag,"::",":") & "_INJVOLPD"
            ElseIf strFacType = "OIL_PAD" Then
                tag = Replace(strFacTag,"::",":") & "_VWY"
            Else
                Wscript.Echo "This should never be displayed. Ever. If so, something is wrong. 04/24/20."
			end If 
			strXML = strXML & "<node cygTag=" & chr(34) & tag & chr(34) & " />"
			arrPoints(i) = tag
		Next
	Next
	strXML = strXML & "</Points></cygPtInfo>"
	
	'Log
	Call WriteLogSucc("String XML created For " & strFacType, LogLevel)
	
	'Creating the XML object with an array of the points
	objPoints.AddPointsArray arrPoints, False
	objPoints.ResolveNow 2
	objPoints.UpdateNow 2
	
    GetXMLCurrentValues = objPoints.GetPointsXML(strXML)
    Call WriteLogSucc("Successful Point Retrieval For " & strFacType, LogLevel)
End Function

Function PrepareDictionary(strPntXML, strFacType)
    'Setting up good and bad file dictioinaries
    Dim currPointObj : Set currPointObj = CreateObject("CxScript.Points")
    Dim objXML : Set objXML = CreateObject("Msxml2.DOMDocument.6.0")
	objXML.async = False
	objXML.LoadXML strPntXML
	
	'This makes strNodes = all of the Points in the XML string from last Nested For Loop
	Dim strNodes : Set strNodes = objXML.documentElement.SelectSingleNode("//cygPtInfo/Points").childNodes
	
	'Log
	Call WriteLogSucc("Ready to get attributes from the string XML.",LogLevel)
	
	'Variables For dictionary and child nodes || Must go thru the child nodes (they will be in a random order); so to go thru this list we must get the attributes from each child node
    Dim child, strValue, strCygTag, strFacTag, strUdc, strActiveStatus, strTimeStamp, strPointID
    Dim dictionary : Set dictionary = CreateObject("Scripting.Dictionary")
    dictionary.Add "Type", strFacType

    For Each child in strNodes
		strValue = CheckValue(child.getAttribute("Value"))
		strCygTag = child.getAttribute("cygTag")
		strFacTag = GetFacTag(strCygTag)
		strUdc = GetUDC(strCygTag)
        strActiveStatus = child.getAttribute("activestatus")
        strPointID = currPointObj.Point(strFacTag &"."& strUdc).GetAttribute("pointid")

        If NOT dictionary.Exists(strFacTag) Then dictionary.Add strFacTag, CreateObject("Scripting.Dictionary")
        dictionary.Item(strFacTag).Add strCygTag, CreateObject("Scripting.Dictionary")
		dictionary.Item(strFacTag).Item(strCygTag).Add "Desc", objFac.GetFacilityAttribute(strFacTag, "FACILITY_DESC")
        dictionary.Item(strFacTag).Item(strCygTag).Add "Value", strValue
        dictionary.Item(strFacTag).Item(strCygTag).Add "UDC", strUdc
        dictionary.Item(strFacTag).Item(strCygTag).Add "PointID", TrimLZ(strPointID)
            
		'Add CygTag as the Key to the Dictionary; then add the Value as the dictionary's value
        If strActiveStatus = "1" Then
            dictionary.Item(strFacTag).Item(strCygTag).Add "Quality", "Good"
			Call WriteLogInfo(strCygTag & " is good", LogLevel)
		ElseIf strActiveStatus = "0" Then
			dictionary.Item(strFacTag).Item(strCygTag).Add "Quality", "Inactive"
			Call WriteLogInfo(strCygTag & " is bad (inactive)", LogLevel)
		ElseIf strActiveStatus = "Null" AND strUDC = "VWY" Then 'Why did you make this distinction here?
			If strValue <> "" Then 'Why did you make this distinction here?
                dictionary.Item(strFacTag).Item(strCygTag).Add "Quality", "VWY/Null/Not Blank"
                Call WriteLogInfo(strCygTag & " is bad (VWY/Null/Not Blank)", LogLevel)
			ElseIf strValue = "" Then'Why did you make this distinction here?
                dictionary.Item(strFacTag).Item(strCygTag).Add "Quality", "VWY/Null/Blank"
                Call WriteLogInfo(strCygTag & " is bad (VWY/Null/Blank)", LogLevel)
            End If
		Else 
			dictionary.Item(strFacTag).Item(strCygTag).Add "Quality", "Other"
			Call WriteLogInfo(strCygTag & " is bad (other)", LogLevel)
		End If
	Next
    PrepareDictionary = dictionary
    Call WriteLogSucc("Successful Dictionary Preparation For " & strFacType, LogLevel)
End Function

Sub WriteToFile(D1)
	'Variables 
	Dim i, j, k, arrD1Keys, printDate, TimeStampVal, BSandW, D2, D3
	printDate = Date() - 1
	TimeStampVal = CheckTimeStamp(printDate)
	BSandW = 100
	
	'Now we print out our Dictionary | Log
	arrD1Keys = D1.Keys 'Top level dictionary with FacIDs as keys
	Call WriteLogSucc("Dictionary of Created",LogLevel)
    
    dictionary.Item(strFacTag).Item(strCygTag).Add "UDC", strUdc

    For i = 0 to UBound(arrD1Keys) 
        Set D2 = D1.Item(arrD1Keys(i)) 'Facility dictionary with PointTags as keys
        For j = 0 to UBound(arrD2Keys)
            Set D3 = D2.Item(arrD1Keys(i)) 'Point dictionary filled with point info
            If D3.Item("Quality") = "Good" Then
                If CInt(D3.Item("Value")) >= 0 Then
                    fileOut.Writeline "LIQUID METER," & D3.Item("Desc") & " Water" & "," & D3.Item("PointID")) & "," & TimeStampVal & "," & D3.Item("Value") & "," & BSandW
                End If
            Else
                BadFile.Writeline arrD2Keys(i) & "|" & D3.Item("Quality")
            End If
        Next
	Next
	
	Call WriteLogSucc("Facility Type finished getting values. ",LogLevel)
End Sub

Sub WriteLogSucc(str, level)
	If level => 2 Then 
		logFile.Writeline now &" - "& str
	End If 
End Sub

Sub WriteLogInfo(str, level)
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
	If not IsNull(value) Then
		If Len(value) > 0 Then
			On Error Resume Next
				CheckValue = CInt(Replace(value," ",""))
                If err.Number > 0 Then
                    CheckValue = -9999
					Wscript.Echo Err.Description 'Maybe write this to log as well
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
	If not IsNull(TimeStamp) Then
		month = DatePart("m", TimeStamp)
		day = DatePart("d", TimeStamp)
		If month < 10 Then month = "0" & month
		If day < 10 Then day = "0" & day
		
		CheckTimeStamp = Year(TimeStamp) & month & day
	ElseIf IsNull(TimeStamp) Then
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
	  'Here we willcheck the path, If it contains
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
	'remote path is done. If it does not exist For some
	'reason. Unexpected results may occur.
		sRemotePath = "\"
	End If
	
	'Check the local path and file to ensure
	'that either the a file that exists was
	'passed or a wildcard was passed.
	If InStr(sLocalFile, "*") Then
		If InStr(sLocalFile, " ") Then
			FTPUpload = "Error: Wildcard uploads do not work If the path contains a space." & vbCRLF
			FTPUpload = FTPUpload & "This is a limitation of the Microsoft FTP client."
			Exit Function
		End If
	ElseIf Len(sLocalFile) = 0 Or Not oFTPScriptFSO.FileExists(sLocalFile) Then
	'nothing to upload
		FTPUpload = "Error: File Not Found."
		Exit Function
	End If
	'--------END Path Checks---------
	  
	  'build input file For ftp command
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
	
	'Write the input file For the ftp command
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