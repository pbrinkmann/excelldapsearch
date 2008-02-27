

'
' Installs the Excel LDAP Query COM dll and Excel macro.
'
' ** THIS SCRIPT SHOULD NOT BE USED **
' ** USE THE NSIS INSTALLER NOW **
'

Dim r			' holds function return value (should be true/false to indicate success/failure)
Dim errMsg		' global error message string to hold error message from failed functions
errMsg = ""

Dim installDir	' hardcoded installation directory
installDir = "C:\Program Files\Excel LDAP Search"

	MsgBox errMsg, vbCritical, "Please use the installer instead"
	Wscript.Quit(1)

'
' Ask the user if it's really OK to proceed
'
r =  DisplayInstallPrompt()

if not r then
	MsgBox errMsg, vbCritical, "Excel LDAP Query Install Failed"
	Wscript.Quit(1)
end if


'
' Copy the files to the install directory
'
r = CopyFiles()

if not r then
	MsgBox errMsg, vbCritical, "Excel LDAP Query Install Failed"
	Wscript.Quit(1)
end if


'
' Have the user customize the LDAP connection parameters
'
r = PromptUserToCustomize()
if not r then
	MsgBox errMsg, vbInformation, "LDAP Settings Customization Failed"
	' Wscript.Quit(1)  treat this as a non-fatal error (user can always edit the file later)
end if


'
' Create a registry entry pointing to the config file
'
r = CreateRegistryEntry()
if not r then
	MsgBox errMsg, vbCritical, "Excel LDAP Query Install Failed"
	Wscript.Quit(1)
end if


'
' Register the COM dll
'
r = RegisterDLL()

if not r then
	MsgBox errMsg, vbCritical, "Excel LDAP Query Install Failed"
	Wscript.Quit(1)
end if


'
' Add the menu item to Tool's menu in Excel
'
r = AddToolsMenuItem()

if not r then
	MsgBox errMsg, vbCritical, "Excel LDAP Query Install Failed"
	Wscript.Quit(1)
end if

MsgBox "Done!", vbOKOnly, "Excel LDAP Query Install Complete"

'
' ------------------------------------------
' Subroutines
' ------------------------------------------
'

'
' DispalyInstallPrompt
'
' -displays installtion prompt
'
Function DisplayInstallPrompt()
	Dim res
	res = MsgBox("Ok to add the Excel LDAP Search macro to Excel?" & vbCrLf & vbCrLf & "NOTE: please make sure Excel is closed before continuing!!", vbOKCancel, "Excel LDAP Search Install Confirmation")
	
	if res <> vbOk then
		errMsg = "Excel LDAP Search install aborted by user"
		DisplayInstallPrompt = false
		exit Function
	end if
	
	DisplayInstallPrompt = true
		
End Function


'
' Copyfiles
'
' -copies files to installation directory
'
Function CopyFiles()
	Dim fs
	Set fs = CreateObject("Scripting.FileSystemObject")
	

	'
	' See if the installation directory exists and if the user wants to overwrite it.
	'
	if fs.FolderExists(installdir) then
		if MsgBox("It appears the installation folder already exists.  Is it OK to overwrite the contents?", vbOKCancel, "Installation Folder Exists") = vbCancel then
			errMsg = "Excel LDAP Search install cancelled by user"
			CopyFiles = False
			exit Function
		end if
	else	
		'
		' Make directory since it doesn't exist
		'
		On Error Resume Next
		Err.Clear
		
		fs.CreateFolder(installDir)
		
		if Err.number <> 0 Then
			errMsg = "Failed to create installation folder: " & Err.Description
			CopyFiles = False
			exit Function
		end if
	end if
	
	'
	' Copy the files (assuming they're in the CWD)
	'
	Dim filesToCopy
	filesToCopy = Array("LdapQuery.dll", "msvcp71.dll","msvcr71.dll", "nsldap32v50.dll", "macro_template.xlt","ldap_params.ini")
	
	Dim file

	On Error Resume Next
	Err.Clear
	for each file in filesToCopy
		if not fs.FileExists(file) then
			errMsg = "ERROR: installation file " & file & " is missing!"
			CopyFiles = false
			exit Function
		end if

		fs.CopyFile file, installDir & "\", true
		
		if Err <> 0 Then
			errMsg = "ERROR: file copy for " & file & " failed (is Excel still running?)" & vbCrLf & vbCrLf & Err.Description
			CopyFiles = false
			exit Function
		end if

	next
	
	CopyFiles = True
	
End Function


Function PromptUserToCustomize
	Dim WshShell, oExec
	Set WshShell = WScript.CreateObject("WScript.Shell")
	
	Dim iniFile
	iniFile = installDir & "\ldap_params.ini"
	
	MsgBox "Please customize these settings to match your set.  They can be edited later in " & iniFile, vbOK, "Customize LDAP Settings"
	
	set oExec = WshShell.Exec ("notepad """ & iniFile & """")
	
	Do While oExec.Status = 0
		WScript.Sleep 100
	Loop

	if oExec.ExitCode <> 0 Then
		errMsg = "WARNING: Please make sure the file " & iniFile & " is configured properly"
		PromptUserToCustomize = False
		exit Function
	end if
	
	PromptUserToCustomize = True
End Function


Function CreateRegistryEntry()
	Dim WshShell
	Set WshShell = WScript.CreateObject("WScript.Shell")
	
	Dim regKey, regValue
	
	regKey = "HKLM\Software\Excel LDAP Search\iniLocation"
	regValue = installDir & "\ldap_params.ini"

	WshShell.RegWrite regKey, regValue
	
	if WshShell.RegRead(regKey) <> regValue Then
		errMsg = "ERROR: Failed to write to the registry key " & regKey
		CreateRegistryEntry = false
		exit Function
	End If
	
	CreateRegistryEntry = true

End Function


Function RegisterDLL()
	Dim WshShell, oExec
	Set WshShell = WScript.CreateObject("WScript.Shell")
	
	set oExec = WshShell.Exec ("regsvr32 """ & installDir & "\LdapQuery.dll""")
	
	Do While oExec.Status = 0
		WScript.Sleep 100
	Loop

	if oExec.ExitCode <> 0 Then
		errMsg = "ERROR: LDAP Query DLL registration failed"
		RegisterDLL = False
		exit Function
	end if
	
	RegisterDLL = True
End Function


Function AddToolsMenuItem()

	On Error Resume Next
	Err.Clear
	
	Dim excel
	set excel = CreateObject("Excel.Application")
	
	if Err <> 0 Then
		errMsg = "ERROR: Unable to run Excel (is it installed correctly?)"
		AddToolsMenuItem = False
		exit Function
	End If
	
	On Error GoTo 0


    Dim LdapQueryFormTag
    LdapQueryFormTag = "LdapQueryShowFormMenuItem"
    
    Dim CmdBar
    Dim CmdBarMenu
    Dim CmdBarMenuItem
    
    '
    ' Point to the Worksheet Menu Bar
    '
    Set CmdBar = excel.Application.CommandBars("Worksheet Menu Bar")
    
    '
    ' Point to the Tools menu on the menu bar
    '
    Set CmdBarMenu = CmdBar.Controls("Tools")
    
    '
    ' See if we already have an item in there, and exit if we do
    '

	Dim toolsMenu
	Dim ldapMenuItem
	set toolsMenu = excel.Application.CommandBars("Tools")
	set ldapMenuItem = toolsMenu.FindControl(,,"LdapQueryShowFormMenuItem")

	if not ldapMenuItem is Nothing then
		MsgBox "The Tools menu item was already installed.", vbOKOnly, "Installation Note"
		excel.Quit		
		AddToolsMenuItem = True
		exit Function
	end if

    
    '
    ' Add a new menu item to the Tools menu
    '
    Set CmdBarMenuItem = CmdBarMenu.Controls.Add
    
    '
    ' Set the properties for the new control
    '
    With CmdBarMenuItem
        .Caption = "Run LDAP Search"
        '.OnAction = "'" & ThisWorkbook.Name & "'!ShowQueryForm"
        '.OnAction = "ShowQueryForm"
        .OnAction = "'" & installDir & "\macro_template.xlt'!ShowQueryForm"
        .Tag = LdapQueryFormTag
    End With
    
    excel.Quit
    
    
	AddToolsMenuItem = True    
	Exit Function 


End Function


