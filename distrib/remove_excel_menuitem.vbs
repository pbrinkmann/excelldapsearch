Option Explicit

'
' Removes the Excel "Run LDAP Search" menu item
'


Dim r			' holds function return value (should be true/false to indicate success/failure)
Dim errMsg		' global error message string to hold error message from failed functions
errMsg = ""

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
    
 
    '
    ' See if we already have an item in there, and exit if we don't (maybe the user removed it manually?)
    '

	Dim toolsCommandBar
	Dim ldapMenuItem
	Set toolsCommandBar = GetToolsCommandBar(excel)
	
	If toolsCommandBar Is Nothing Then
		MsgBox "The Tools Command Bar was not found, you may have to manually remove the menu item", vbOKOnly, "Installation Note"
		excel.Quit		
		RemoveToolsMenuItem = False
		Exit Function
	End IF
	
	Set ldapMenuItem = toolsCommandBar.FindControl(,,"LdapQueryShowFormMenuItem")

	if ldapMenuItem is Nothing then
		MsgBox "The Tools menu item was not found.", vbOKOnly, "Installation Note"
		excel.Quit		
		RemoveToolsMenuItem = True
		Exit Function
	end if


	'
	' remove the menu item
	'     
    ldapMenuItem.Delete
    
    excel.Quit
    
	RemoveToolsMenuItem = True    
	Exit Function 


End Function


Function GetToolsCommandBar(excel)
	On Error Resume Next
	Err.Clear
	
	Set GetToolsCommandBar =  excel.Application.CommandBars("Tools")
	
	If Err.Number Then
		Set GetToolsCommandBar = Nothing
	End If
	
	On Error Goto 0
	
End Function

