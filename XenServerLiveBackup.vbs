'XenServerBackup script version 2.0
'This is a Visual Basic Script to perform live backups from XenServer hosts.
'Project page https://github.com/dec0mpile/XenServerBackup

'Use of this software is "as-is". I take no responsibility for the results of making
'use of this or related programs and any data directly or indirectly affected.
'Please test it in your environment before using it on production servers.

'This script is licensed under the GPL v3 license.
'A copy of the license is found on the project page and
'https://www.gnu.org/licenses/gpl-3.0.html

'############################### User configurable variables ####################################

'### XenServer VM settings ###
	'VM names are stored in the strServers variable and must be separated by a comma.
	'Names are case sensitive so you have to get the name of your server PERFECT!
	strServers = "ServerName,ServerName2,ServerName3"

'### XenServer and log-on details ###
	'XenServer root user
	strUser = "root"
	'Password of the root user
	strpw = "mypassword"
	'XenServer IP address or hostname
	strXenServer = "X.X.X.X"

'### XenServer backup settings ###
	'Backup path location. This script supports mapped network shares
	strBackupPath = "D:\XSBackups"
	'Path of XenCenter installation. Make sure to use the short path format
	strXenCenterPath = "C:\PROGRA~2\Citrix\XenCenter\"
	'Compress image .xva file?
	'++++++++++++++++++++++++++++++WARNING:+++++++++++++++++++++++++++++++++++++++++++++
	'Image file compression is performed on the XenServer host.
	'This means that the XenServer host will have increased CPU usage during backup.
	'The backup times with compression are significantly longer.
	'Test this very carefully to ensure CPU usage does not affect guest machines.
	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'Set to TRUE if you want the images to be compressed.
	bolCompressImage = FALSE
	'Server identification string. This is added to the file name, and it is used to
	'identify the server that is being backed up. This setting is useful if you have
	'multiple host backups going to the same backup location.
	'++++++++++++++++++++++++++++++++++++ NOTE +++++++++++++++++++++++++++++++++++++++++
	'This string is also used to identify which files to delete in the backup cleanup
	'function. So make sure that it is a meaningful name.
	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	strIndentify = "ProductionXen1"

'### Email Options ###
	'Set to TRUE if you want to send a status email
	bolSendEmail = FALSE
	'The log is always saved in the backup directory. Set this to TRUE if you want
	'the email to also include the log file as an attachment.
	bolIncludeAttachment = FALSE
	'This is the "From" email address that will be used. Eg: "XSLiveBackup@mycompany.com"
	strSMTPFrom = "Your_XenServer_Backup@example.com"
	'This is the "To" email address that will be used. Eg: "sysadmin@mycompany.com"
	strSMTPTo = "sysadmin@example.com"
	'The IP or hostname of the SMTP relay server
	strSMTPRelay = "X.X.X.X"

'### Backup Days ###
	'Number of days to keep the logs and backup files.
	'Anything older will be permanently deleted
	numKeepBackups = 5

'### Show Pop-Up message on script completion ###
	'This flag should be set to FALSE if you want to run this script as a scheduled task
	bolShowCompletionMsg = FALSE

'########################### END User configurable variables ####################################

'************************************************************************************************
'Do not edit beyond this point
'************************************************************************************************

Dim errStatus, strLogName
Dim fs, logFile

Set fs = CreateObject("Scripting.FileSystemObject")

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

'Cleans up the old backups and log files.
Call DelFiles(strBackupPath,strIndentify,numKeepBackups)
'Finish up the log file
Call logClose
'Send status message via email
Call sendMsg

If bolShowCompletionMsg = TRUE Then
	WScript.Echo "XenServerBackup script execution is complete!"
End If


Sub backupVM(strServer)

	'Fix the VM name to allow for possibility of spaces
	strQuote = """"
	strServer = strQuote & strServer & strQuote

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
	strResult = ""

	If bolCompressImage = TRUE Then
		strName = "Backup-" & strServer & "-" & strTime & "-" & strIndentify & ".xva.gz"
		writeLog("Backup filename: " & strName)

		Set objExec = WSHshell.Exec(strRunXE & "vm-export vm=" & strSSID & " compress=true" & " filename=" & strBackupPath & "\" & strName)
	Else
		strName = "Backup-" & strServer & "-" & strTime & "-" & strIndentify & ".xva"
		writeLog("Backup filename: " & strName)

		Set objExec = WSHshell.Exec(strRunXE & "vm-export vm=" & strSSID & " filename=" & strBackupPath & "\" & strName)
	End if

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

	If Not fs.FolderExists(strBackupPath) Then
		fs.CreateFolder(strBackupPath)
	End If

	strLogName = "XenServerBackups-" & Replace(Date, "/", "-") & "-" & strIndentify & ".log"
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


'Finish the log file.
Sub logClose
	
	logFile.WriteLine
	strStatus = SetErrorStatus("Read")
	logFile.WriteLine "Backup completed on " & Now()
	logFile.WriteLine "Backup Status: " & strStatus
	logFile.WriteLine
	logFile.WriteLine
	logFile.Close
	
End Sub


'Cleans up the old backups and log files.
Sub DelFiles (ByVal strDelPath, ByVal strIdentifier, ByVal intNumDaysToKeep)

	Set oFolder = fs.GetFolder(strDelPath)
	Set oAllFiles = oFolder.Files
	
	numDM = 0
	
	writeLog("Cleaning up old backup files")
	
	For Each oFile in oAllFiles
		If Abs(InStr(oFile.Name, strIdentifier) <> 0 and DateDiff("d", NOW, oFile.DateLastModified)) > intNumDaysToKeep Then
			numDM = numDM + 1
			ReDim Preserve arrFileName(numDM)
			
			arrFileName(numDM) = oFile.Name
			
			writeLog("Deleting file: " & oFile.Path)	
			fs.DeleteFile oFile.Path
			
		End If
	Next
	
	writeLog(numDM & " files deleted")
	
End Sub


Sub sendMsg

	'Send status message via email
	If bolSendEmail = FALSE Then
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
	strText = vbCRLF & "XenServer Backup completed for server: " & strIndentify & vbCRLF & vbCRLF
	strText = strText & "Backup User                : " & strUser & vbCRLF
	strText = strText & "Password                   : *********" & vbCRLF
	strText = strText & "Xen Server                 : " & strXenServer & vbCRLF
	strText = strText & "Backup Destination         : " & strBackupPath & vbCRLF
	strText = strText & vbCRLF
	strText = strText & "Backup Status              : " & strStatus
	oMessage.TextBody = strText

	If bolIncludeAttachment = TRUE Then
		oMessage.AddAttachment strBackupPath & "\" & strLogName
	End If

	oMessage.Send
	
End Sub
