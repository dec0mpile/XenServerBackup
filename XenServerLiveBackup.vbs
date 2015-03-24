'User configurable variables
'Server names must be seperated by a comma.  Names are case sensitive so you have to get the
'name of your server PERFECT!
strServers = "ServerName,ServerName2"
strUser = "root"
strpw = "mypassword"
strXenServer = "XenServer IP Address or hostname"
strBackupPath = "d:\xsbackups"
'Make sure to use the short file name format
strXenCenterPath = "C:\Progra~1\Citrix\XenCenter\"
'Set to TRUE if you want to send a status email
binSendEmail = TRUE
strSMTPFrom = "XSLiveBackup@mycompany.com"
strSMTPTo = "spiceworkshelpdesk@mycompany.com"
strSMTPRelay = "smtp relay IP address or host name"
'Number of days to keep the logs and backup files
numKeepLogs = 10
numKeepBackups = 5



'************************************************************************************************
'Do not edit beyond this point
'************************************************************************************************
Dim errStatus, strLogName
Dim fs, logFile
If Right(strBackupPath, 1) = "\" Then
	strBackupPath = Left(strBackupPath, Len(strBackupPath) - 1)
End If

If Right(strXenCenterPath, 1) <> "\" Then
	strXenCenterPath = strXenCenterPath & "\"
End If

strRunXE = strXenCenterPath & "xe.exe -s " & strXenServer & " -u " & strUser & " -pw " & strpw & " "
arrBackups = Split(strServers, ",")
const ForAppending = 8
errStatus = 0

'Set some global objects
Set WshShell = CreateObject ("Wscript.shell")

'Create the log file and introduction information into log file
Call logSetup
'Check if XenCenter Path is correct
If fs.FileExists(strXenCenterPath & "xe.exe") Then
	'Loop through all selected servers and back them up
	For iServers = 0 to UBound(arrBackups)
		Call backupVM(Trim(arrBackups(iServers)))
	Next
Else
	writeLog("XenCenter not found, aborting backup!")
	errStatus = UBound(arrBackups) + 2
End If

'Finish up the log file
Call logClose
'Send status message via email
Call sendMsg

'Script completed -- REM line below out if you want to run this in a scheduled task
'wscript.echo "Done!"



Sub backupVM(strServer)
	'First, check that the VM exists, and get UUID and name-label information.  It is possible to 
	'name a snapshot as a backup source, but this script does not allow it, so checking for that too
	Set objExec = WSHshell.Exec(strRunXE & "vm-list params=uuid,name-label,is-a-snapshot name-label=" & strServer)
	strStatus = "Not Found"
	Do While Not objExec.StdOut.AtEndOfStream
		strStatus = "Found"
		strUUID = stripValue(objExec.StdOut.ReadLine())
		strVM = stripValue(objExec.StdOut.ReadLine())
		strSnapShot = stripValue(objExec.StdOut.ReadLine())
		strTemp = objExec.StdOut.ReadLine()
		strTemp = objExec.StdOut.ReadLine()
	
		If strSnapShot = "false" and strVM <> "Control domain on host" Then
			strStatus = "Good"
		End If
	Loop
	
	If strStatus = "Not Found" Then
		strResult = SetErrorStatus("Add")
		writeLog("No VM by that name: " & strServer)
		Exit Sub
	ElseIf strStatus = "Found" Then
		strResult = SetErrorStatus("Add")
		writeLog("VM is a Snapshot or Template, no backup performed: " & strServer)
		Exit Sub
	Else
		writeLog("VM Found: " & strServer)
	End If

	'Remove old backups
	writeLog("Cleaning up old backup files")
	Set oFolder = fs.GetFolder(strBackupPath)
	Set oAllFiles = oFolder.Files
	numDM = 0
	For Each oFile in oAllFiles
		If Left(oFile.Name, 7) = "Backup-" and DateDiff("d", NOW, oFile.DateLastModified) > numKeepBackups Then
			numDM = numDM + 1
			ReDim Preserve arrFileName(numDM)
			arrFileName(numDM) = oFile.Name
			fs.DeleteFile oFile.Path
		End If
	Next
	writeLog(numDM & " files deleted")
	For x = 1 to numDM
		writeLog("     " & arrFileName(x))
	Next
	
	'Snapshot the VM
	writeLog("Snapshoting server: " & strServer)
	Set objExec = WSHshell.Exec(strRunXE & "vm-snapshot new-name-label=" & strServer & "-XenServer-Live-Backup uuid=" & strUUID)
	strSSID = objExec.StdOut.ReadLine()
	strResult = strSSID
	Do While Not objExec.StdOut.AtEndOfStream
		strTemp = objExec.StdOut.ReadLine()
		writeLog(strTemp)
		strResult = strResult & ":" & strTemp
	Loop
	If InStr(UCase(strResult), "ERROR") Then
		writeLog("Error creating snapshot, see above")
		Exit Sub
	End If
	
	'Set snapshot to NOT be a template
	writeLog("Setting snapshot status...")
	Set objExec = WSHshell.Exec(strRunXE & "template-param-set is-a-template=false uuid=" & strSSID)
	Do While Not objExec.StdOut.AtEndOfStream
		writeLog(objExec.StdOut.ReadLine())
	Loop
		
	'Export to backup destination
	strTime = Replace(Now(), "/", "-")
	strTime = Replace(strTime, " ", "-")
	strTime = Replace(strTime, ":", "-")
	strName = "Backup-" & strServer & "-" & strTime & ".xva"
	writeLog("Backup filename: " & strName)
	Set objExec = WSHshell.Exec(strRunXE & "vm-export vm=" & strSSID & " filename=" & strBackupPath & "\" & strName)
	strResult = ""
	Do While Not objExec.StdOut.AtEndOfStream
		strTemp = objExec.StdOut.ReadLine()
		writeLog(strTemp)
		strResult = strResult & ":" & strTemp
	Loop
	If InStr(UCase(strResult), "SUCCEEDED") = 0 Then
		strResult = SetErrorStatus("Add")
		writeLog("**************   Error during backup of " & strServer & " **************")
	End If
	
	'Remove the snapshot
	Set objExec = WSHshell.Exec(strRunXE & "vm-uninstall uuid=" & strSSID & " force=true")
	strResult = ""
	Do While Not objExec.StdOut.AtEndOfStream
		strTemp = objExec.StdOut.ReadLine()
		writeLog(strTemp)
		strResult = strResult & " " & strTemp
	Loop
	If InStr(strResult, "All objects destroyed") = 0 Then
		strResult = SetErrorStatus("Add")
		writeLog("**************Error deleting snapshot for " & strServer & " **************")
	End If
End Sub


Function stripValue(strValue)
	arrStrip = Split(strValue, ":")
	stripValue = Trim(arrStrip(1))
End Function


Sub writeLog(strText)
	logFile.WriteLine Now() & ":  " & strText
End Sub


Function SetErrorStatus(strTask)
	If strTask = "Add" Then
		errStatus = errStatus + 1
		SetErrorStatus = errStatus
	Else
		numServers = UBound(arrBackups) + 1
		If errStatus = 0 Then
			SetErrorStatus = "Success"
		ElseIf errStatus >= numServers Then
			SetErrorStatus = "Failed"
		Else
			SetErrorStatus = "Partial Failure"
		End If
	End If
End Function


Sub logSetup
	Set fs = CreateObject("Scripting.FileSystemObject")

	If Not fs.FolderExists(strBackupPath) Then
		fs.CreateFolder(strBackupPath)
	End If
	
	Set oFolder = fs.GetFolder(strBackupPath)
	Set oAllFiles = oFolder.Files
	For Each oFile in oAllFiles
		If Left(oFile.Name, 16) = "XenServerBackups" and DateDiff("d", NOW, oFile.DateLastModified) > numKeepLogs Then
			fs.DeleteFile oFile.Path
		End If
	Next

	strLogName = "XenServerBackups-" & Replace(Date, "/", "-") & ".log"
	Set logFile = fs.OpenTextFile (strBackupPath & "\" & strLogName, ForAppending, True)	
	
	logFile.WriteLine "=========================================================================================================="
	logFile.WriteLine "Backup for:  " & Now()
	logFile.WriteLine "=========================================================================================================="
	For x = 0 to UBound(arrBackups)
		If x = 0 Then
			logFile.WriteLine "Servers targeted for backup: " & arrBackups(x)
		Else
			logFile.WriteLine "                             " & arrBackups(x)
		End If
	Next
	logFile.WriteLine "Backup User                : " & strUser 
	logFile.WriteLine "Password                   : *********" 
	logFile.WriteLine "Xen Server                 : " & strXenServer
	logFile.WriteLine "Backup Destination         : " & strBackupPath
	logFile.WriteLine
End Sub


Sub logClose
	'Finish up the log file
	logFile.WriteLine
	strStatus = SetErrorStatus("Read")
	logFile.WriteLine "Backup completed on " & Now()
	logFile.WriteLine "Backup Status: " & strStatus
	logFile.WriteLine
	logFile.WriteLine
	logFile.Close
End Sub


Sub sendMsg
	'Send status message via email
	If binSendEmail = False Then
		Exit Sub
	End If
	
	Set oMessage = CreateObject("CDO.Message")
	oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
	oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPRelay
	oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
	oMessage.Configuration.Fields.Update
	
	strStatus = SetErrorStatus("Read")
	oMessage.Subject = "XenServer Backup Completed. Status: " & strStatus & ".  Date: " & Now()
	oMessage.From = strSMTPFrom
	oMessage.To = strSMTPTo
	strText = vbCRLF & "XenServer Backup Completed. " & vbCRLF & vbCRLF
	strText = strText & "Backup User                : " & strUser & vbCRLF
	strText = strText & "Password                   : *********" & vbCRLF
	strText = strText & "Xen Server                 : " & strXenServer & vbCRLF
	strText = strText & "Backup Destination         : " & strBackupPath & vbCRLF
	strText = strText & vbCRLF
		strText = strText & "Backup Status              : " & strStatus
	oMessage.TextBody = strText
	oMessage.AddAttachment strBackupPath & "\" & strLogName
	oMessage.Send
End Sub
