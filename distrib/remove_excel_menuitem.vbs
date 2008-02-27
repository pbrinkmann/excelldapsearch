

'
' Removes the Excel "Run LDAP Search" menu item
'


Dim r			' holds function return value (should be true/false to indicate success/failure)
Dim errMsg		' global error message string to hold error message from failed functions
errMsg = ""

Dim installDir	' hardcoded installation directory
installDir = "C:\Program Files\Excel LDAP Search"


'
' Remove the menu item to Tool's menu in Excel
'
r = RemoveToolsMenuItem()

if not r then
	MsgBox errMsg, vbCritical, "Excel LDAP Query UnInstall Failed"
	Wscript.Quit(1)
end if

'
' ------------------------------------------
' Subroutines
' ------------------------------------------
'


Function RemoveToolsMenuItem()

	On Error Resume Next
	Err.Clear
	
	Dim excel
	set excel = CreateObject("Excel.Application")
	
	if Err <> 0 Then
		errMsg = "ERROR: Unable to run Excel (is it installed?)"
		RemoveToolsMenuItem = False
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
    ' See if we already have an item in there, and exit if we don't (maybe the user removed it manually?)
    '

	Dim toolsMenu
	Dim ldapMenuItem
	set toolsMenu = excel.Application.CommandBars("Tools")
	set ldapMenuItem = toolsMenu.FindControl(,,"LdapQueryShowFormMenuItem")

	if ldapMenuItem is Nothing then
		MsgBox "The Tools menu item was not found.", vbOKOnly, "Installation Note"
		excel.Quit		
		RemoveToolsMenuItem = True
		exit Function
	end if


	'
	' remove the menu item
	'     
    ldapMenuItem.Delete
    
    excel.Quit
    
	RemoveToolsMenuItem = True    
	Exit Function 


End Function


