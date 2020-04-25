'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Get the current values from OIL_PAD, OIL_WELLS & SWD_WELLS
'
'Liquid Meter Water Script to FTP Upload For ProdView Format 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'Liquid Meter - Water Meters
'Facility Types: OIL_PAD , OIL_WELLS, SWD_WELLS
'In testing, test with On Error Resume Next off
'In production, it's good to have On Error Resume Next on to make sure all processes and files close out instead of accumulating
'On Error Resume Next
Dim strSite : strSite = "FRANK"
Dim strSiteUIS : strSiteUIS = strSite & ".UIS" 
Dim strSitePNT : strSitePNT = strSite & ".PNT"
Dim strStatusPointTag : strStatusPointTag = strSite & ".UIS:BATCH_STATUS_LIQWAT"

'Set Log Level 
'| 0 = No Log
'| 1 = Important
'| 2 = Everything
Dim LogLevel : LogLevel = 2

'Now - Date get
Dim strFileTS : strFileTS = Month(now()) &"-"& Day(now()) &"-"& Year(now()) & " " & Hour(now()) &"-"& Minute(now()) &"-"& Second(now())

'Making Objects
Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
Dim PntClient : Set PntClient = CreateObject("CxPnt.PntClient")
Dim fileGood : Set fileGood = objFso.CreateTextFile("C:\CygNet\Scripts\TempGood\LiquidMeterWater " & strFileTS & ".csv")
Dim fileLog : Set fileLog = objFso.CreateTextFile("C:\CygNet\Scripts\TempLog\LiquidMeterWaterLog.txt")
Dim fileBad : Set fileBad = objFso.CreateTextFile("C:\CygNet\Scripts\Bad\LiquidMeterWaterBad.txt")
Dim objGlobFunc : Set objGlobFunc = CreateObject("CxScript.GlobalFunctions")

'Log 
Call LogHeader(LogLevel)
Call WriteLog("Successfully created all objects", 2)

'Print the header of the ProdView 
fileGood.Writeline "prodview scada import"
fileGood.Writeline "1.0"
fileGood.Writeline "imperial"

fileBad.Writeline "Facility ID & Tank Description | Reason For Failure"  
fileBad.Writeline "---------------------------------------------------"

'Connect
PntClient.Connect(strSitePNT)
objFac.UpdateNow()

'Global Function Objects
objGlobFunc.EnableLiveMode True
objGlobFunc.setpoint strStatusPointTag, "In-Progress", now

'Create the Facility TagLists 
'Needs error checking and testing
Dim arrFacTypes : arrFacTypes = Array("OIL_PAD", "OIL_WELLS", "SWD_WELLS")
For i = 0 to UBound(arrFacTypes)
    Call WriteToFile(PrepareDictionary(GetXMLCurrentValues(strSiteUIS, arrFacTypes(i)), arrFacTypes(i)))
    Call WriteLog("Successfully processed " & arrFacTypes(i), 2)
Next

'FTP Copy over
Call Copy("C:\CygNet\Scripts\TempGood", "\\wstr.com\ftp\CYGNET\PRD", "LiquidMeterWater " & strFileTS & ".csv")
fileGood.close
fileLog.close
fileBad.close

'Archive 
Call Archive("C:\CygNet\Scripts\TempGood\LiquidMeterWater " & strFileTS & ".csv", "C:\CygNet\Scripts\Archive\LiquidMeterWater " & strFileTS & ".csv")
Call Archive("C:\CygNet\Scripts\TempLog\LiquidMeterWaterLog.txt", "C:\CygNet\Scripts\Archive\LiquidMeterWaterLog " & strFileTS & ".txt")

'Status report
If Err.number = 0 then 
	objGlobFunc.setpoint strStatusPointTag, "Complete", now
ElseIf Err.number <> 0 then 
	objGlobFunc.setpoint strStatusPointTag, "Complete w Errors", now
end If

'===============================
'=== Functions & Subroutines ===
'===============================

Function GetXMLCurrentValues(strSiteServ, strFacType)
	Call WriteLog("Get XML Current values for " & strFacType, 2)
    Dim objFac : Set objFac = CreateObject("CxScript.Facilities")
    Dim objPoints : Set objPoints = CreateObject("CxScript.Points")	
    Dim strXML, arrPoints(), i, j, strPointTag, arrTagList, strFacTag

    objFac.GetFacilityTagList strSiteServ, "facility_is_active=Y;facility_type=" & strFacType, arrTagList
    Redim arrPoints(UBound(arrTagList) + 1)

	Call WriteLog(strFacType & "|" & UBound(arrTagList) + 1 & "facilities to process", 2)

	strXML = "<cygPtInfo><Parameters><Value /><timestamp /><activestatus /></Parameters><Points>"
	For i = 0 to UBound(arrTagList)
		strFacTag = arrTagList(i)
		strFacType = objFac.GetFacilityAttribute(strFacTag, "FACILITY_TYPE")
		For j = 0 to UBound(arrStrUDC)
			If strFacType = "OIL_WELLS" Then 'Could do a Switch here but this is fine too. Other options as well
				strPointTag = Replace(strFacTag,"::",":") & "_VWY0"
			ElseIf strFacType = "SWD_WELLS" Then
                strPointTag = Replace(strFacTag,"::",":") & "_INJVOLPD"
            ElseIf strFacType = "OIL_PAD" Then
                strPointTag = Replace(strFacTag,"::",":") & "_VWY"
            Else
                Wscript.Echo "This should never be displayed. Ever. If so, something is wrong. Fix it."
			end If 
			strXML = strXML & "<node cygTag=" & chr(34) & strPointTag & chr(34) & " />"
			arrPoints(i) = strPointTag
		Next
	Next
	strXML = strXML & "</Points></cygPtInfo>"
	
	Call WriteLog("XML current values request created For " & strFacType, 2)
	
	objPoints.AddPointsArray arrPoints, False
	objPoints.ResolveNow 2
	objPoints.UpdateNow 2
	
    GetXMLCurrentValues = objPoints.GetPointsXML(strXML)
    Call WriteLog("Finish get XML Current values for " & strFacType, 2)
End Function

Function PrepareDictionary(strPntXML, strFacType)
	Call WriteLog("Begin prepare dictionary for " & strFacType, 2)
	Dim pntChild, strValue, strCygTag, strFacTag, strUdc, strActiveStatus, strTimeStamp, strPointID, strQuality
    Dim dictionary : Set dictionary = CreateObject("Scripting.Dictionary")
    Dim currPointObj : Set currPointObj = CreateObject("CxScript.Points")
    Dim objXML : Set objXML = CreateObject("Msxml2.DOMDocument.6.0")
	
	objXML.async = False
	objXML.LoadXML strPntXML
	dictionary.Add "Type", strFacType

	'This makes arrNodes = all of the Points in the XML string from last Nested For Loop
	Dim arrNodes : Set arrNodes = objXML.documentElement.SelectSingleNode("//cygPtInfo/Points").childNodes
	Call WriteLog(strFacType & "|" & UBound(arrNodes) + 1 & " nodes to process", 2)   
    
    For Each pntChild in arrNodes
        strValue = CheckValue(pntChild.getAttribute("Value"))
        strCygTag = pntChild.getAttribute("cygTag")
        strFacTag = GetFacTag(strCygTag)
        strUdc = GetUDC(strCygTag)
        strActiveStatus = pntChild.getAttribute("activestatus")
        strPointID = currPointObj.Point(strFacTag & "." & strUdc).GetAttribute("pointid")

        If NOT dictionary.Exists(strFacTag) Then dictionary.Add strFacTag, CreateObject("Scripting.Dictionary")
        dictionary.Item(strFacTag).Add strCygTag, CreateObject("Scripting.Dictionary")
        dictionary.Item(strFacTag).Item(strCygTag).Add "Desc", objFac.GetFacilityAttribute(strFacTag, "FACILITY_DESC")
        dictionary.Item(strFacTag).Item(strCygTag).Add "Value", strValue
        dictionary.Item(strFacTag).Item(strCygTag).Add "UDC", strUdc
        dictionary.Item(strFacTag).Item(strCygTag).Add "PointID", TrimLZ(strPointID)
            
        If strActiveStatus = "1" Then
			If CInt(strValue) >= 0 Then
				strQuality = "Good"
			Else
				strQuality = "Status=Active/Value=(Negative OR Null OR Error Occurred in CheckValue)"
			End If
		ElseIf strActiveStatus = "0" Then
			strQuality = "Inactive Status"
		ElseIf strActiveStatus = "Null" AND strUDC = "VWY" Then
			If strValue <> "" Then
                strQuality = "UDC=VWY/Status=Null/Value=Not Blank"
			ElseIf strValue = "" Then
                strQuality = "UDC=VWY/Status=Null/Value=Blank"
            End If
		Else
			strQuality = "Other"
		End If
		
		Call WriteLog(strCygTag & " is " & strQuality, 2)
		dictionary.Item(strFacTag).Item(strCygTag).Add "Quality", strQuality
	Next
    PrepareDictionary = dictionary
    Call WriteLog("Finish prepare dictionary for " & strFacType, 2)
End Function

Sub WriteToFile(D1)
	Call WriteLog("Begin write file for " & D1.Item("Type"), 2)
	Dim i, j, arrD1Keys, arrD2Keys, strTimeStamp, intBSandW, D2, D3
	strTimeStamp = CheckTimeStamp(Date() - 1)
	intBSandW = 100
	
	'Now we print out our Dictionary | Log
	arrD1Keys = D1.Keys 'D1 is Top level dictionary with FacIDs as keys    
	Call WriteLog(strFacType & "|" & UBound(arrD1Keys) + 1 & "facilities in dictionary to process", 2)

	For i = 0 to UBound(arrD1Keys) 
		Set D2 = D1.Item(arrD1Keys(i)) 'D2 is the Facility dictionary with PointTags as keys
		arrD2Keys = D2.Keys
		For j = 0 to UBound(arrD2Keys)
			Set D3 = D2.Item(arrD2Keys(j)) 'D3 is the Point dictionary filled with point info
			If D3.Item("Quality") = "Good" Then
				fileGood.Writeline "LIQUID METER," & D3.Item("Desc") & " Water" & "," &_
				D3.Item("PointID") & "," & strTimeStamp & "," & D3.Item("Value") & "," & intBSandW
			Else
				fileBad.Writeline arrD2Keys(j) & "|" & D3.Item("Quality")
			End If
		Next
	Next
	
	Call WriteLog("Finish write file for " & D1.Item("Type"), 2)
End Sub

Sub WriteLog(str, level)
	If level => LogLevel Then 
		fileLog.Writeline now & " - "& str
	End If 
End Sub

Sub LogHeader(level)
	If level => LogLevel Then
		fileLog.Writeline "LiquidMeterWater.vbs - Start"
		fileLog.Writeline now
		fileLog.Writeline "Begin log: "
	End If 
End Sub

Function CheckValue(strPointTag, strValue)
	If NOT IsNull(strValue) Then
		If Len(strValue) > 0 Then
			On Error Resume Next
				CheckValue = CInt(Replace(strValue," ",""))
                If Err.Number > 0 Then
                    CheckValue = -9999
					Call WriteLog(strPointTag & "|" Err.Description, 2)
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
	GetFacTag = Replace(Replace(PntTag, "_" & GetUDC(PntTag),""),":","::")
End Function

Function GetUDC(PntTag)
	GetUDC = Split(PntTag, "_")(UBound(Split(PntTag, "_")))
End Function

Function CheckTimeStamp(TimeStamp)
	Dim day, month
	If NOT IsNull(TimeStamp) Then
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
	ElseIf Len(sLocalFile) = 0 OR NOT oFTPScriptFSO.FileExists(sLocalFile) Then
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
	Call WriteLog("Copying files to remote location...", 2)
	Dim WshShellScriptExec, WSHShell, strfileLog, strCmd
	Set WSHShell = CreateObject("Wscript.Shell")
	strfileLog = source & "\CopyProcess.Log"
	strCmd = "robocopy """ & source & """ """ & destination & """ """ & file & """ /XO /NFL /NDL /NP /R:0 /W:1 /LOG+:""" & strfileLog &""""
	
	Call WriteLog("Cmd: " & strCmd, 2)
	WshShellScriptExec = WshShell.Run(strCmd, 0, True)
	
	Call WriteLog("End of Copy. File copy status: " & WshShellScriptExec, 2) 
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