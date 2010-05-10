

'
' Installs the Excel "LDAP Search" menu item
'


Dim r			' holds function return value (should be true/false to indicate success/failure)
Dim errMsg		' global error message string to hold error message from failed functions
errMsg = ""

Dim installDir	' installation directory


'
' Fetch the installation directory from the registry (sets global installDir var)
'
r = GetInstallDir()

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

Wscript.Quit(0)  ' successfull completion

'
' ------------------------------------------
' Subroutines
' ------------------------------------------
'


Function GetInstallDir()
	Dim WshShell
	Set WshShell = WScript.CreateObject("WScript.Shell")
	
	Dim regKey, regValue
	
	On Error Resume Next
	Err.Clear
	
	regKey = "HKLM\SOFTWARE\Excel LDAP Search\installDir"
	regValue = WshShell.RegRead(regKey)
	
	if Err <> 0 Then
		errMsg = "ERROR: Failed to read the installation directory from the registry"
		GetInstallDir = false
		exit Function
	End If
	
	On Error Goto 0
	
	installDir = regValue
	
	GetInstallDir = true

End Function


Function AddToolsMenuItem()

	'
	' start off by getting the Excel application object
	'
	On Error Resume Next
	Err.Clear
	
	Dim excel
	set excel = CreateObject("Excel.Application")
	
	if Err <> 0 Then
		errMsg = "ERROR: Unable to run Excel (is it installed?)"
		AddToolsMenuItem = False
		exit Function
	End If
	
	'
	' get the Excel version, and make sure it's a valid one.
	'
	Dim excelVersion
	
	Select Case Int(excel.Version)
		Case 12
			excelVersion = "Excel 2007"
		Case 11
			excelVersion = "Excel 2003"
		Case 11
			excelVersion = "Excel 2003"
		Case 10
			excelVersion = "Excel XP/2002"
		Case 9
			excelVersion = "Excel 2000"
		Case 8
			excelVersion = "Excel 97"
		Case 13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30 ' that's a lot of future versions (why doesn't 13 To 100 work?)
			excelVersion = "Version " & excel.Version
			if MsgBox("It appears you have a new version of Excel(" & excelVersion & ") that hasn't been tested." & vbCr & "Do you with to continue with this install?", vbYesNo, "Unkown Excel Version") = vbNo then
				errMsg = "User aborted install due to unknown Excel version"
				AddToolsMenuItem = False
				exit Function
			end if
		Case 5
			MsgBox "Eeek, you have a really old version of Excel.  Please upgrade to Excel 97 or later", vbError, "Unsupported Excel Version"
			errMsg = "ERROR: Excel version is too old"
			AddToolsMenuItem = False
			exit Function
		Case Else
			excelVersion = "Version Unknown"
			MsgBox "Your Excel version was reported as """ & excel.Version & """, which is an unrecognized version", vbError, "Unsupported Excel Version"
			if MsgBox("Do you with to try and continue with this install?", vbYesNo, "Unkown Excel Version") = vbNo then
				errMsg = "User aborted install due to unknown Excel version"
				AddToolsMenuItem = False
				exit Function
			end if
	end select


    Dim LdapQueryFormTag
    LdapQueryFormTag = "LdapQueryShowFormMenuItem"
    
    Dim commandBar
	Dim toolsMenu
	Dim ldapMenuItem
    
    '
    ' Point to the Worksheet Menu Bar
    '
    Set commandBar = excel.Application.CommandBars("Worksheet Menu Bar")
    
	If Err.Number Then
        errMsg = "Unable to find the ""Worksheet Menu Bar""" _
        & vbCr & "Is this a non-English version of Excel?" _
        & vbCr & vbCr & "Please contact the Excel LDAP Search author for help" _
        & vbCr & "http://excelldapsearch.sourceforge.net"
        AddToolsMenuItem = False
		exit Function
    End If
	
    '
    ' Point to the Tools menu on the menu bar
    '
    Set toolsMenu = commandBar.Controls("Tools")
	
	If Err.Number <> 0 Then
        errMsg =  "Unable to find the ""Tools"" menu item" _
        & vbCr & "Is this a non-English version of Excel?" _
        & vbCr & vbCr & "Please contact the Excel LDAP Search author for help" _
        & vbCr & "http://excelldapsearch.sourceforge.net"
        AddToolsMenuItem = False
		exit Function
    End If
	
    
	' no more manual error checking
    On Error GoTo 0
	
    '
    ' See if we already have an item in there, and nuke it if it's there
    '
	set ldapMenuItem = commandBar.FindControl(,,"LdapQueryShowFormMenuItem")

	if not ldapMenuItem is Nothing then
		ldapMenuItem.delete
		Set ldapMenuItem = Nothing
	end if
		
    '
    ' Add a new menu item to the Tools menu
    '
    Set ldapMenuItem = toolsMenu.Controls.Add
    
    '
    ' Set the properties for the new control
    '
    With ldapMenuItem
        .Caption = "Run LDAP Search"
        .OnAction = "'" & installDir & "\ldapquery_addin.xla'!ShowQueryForm"
        .Tag = LdapQueryFormTag
    End With
    
    excel.Quit

	'
	' give the user the happy news
	'
	MsgBox "Excel menu item installed for " & excelVersion & "." & vbCr & "If you also have older versions of Excel installed, see the Readme.html for manual installation.", vbInformation, "Excel Tools Menu Item Added"
	
    
	AddToolsMenuItem = True    
	Exit Function 


End Function


