Attribute VB_Name = "ldapQueryFormMacro"
' Excel LDAP Search
' Copyright (C) 2008 Paul Brinkmann <paul@paulb.org>
'
' This program is free software; you can redistribute it and/or modify it under the terms
' of the GNU General Public License as published by the Free Software Foundation; either
' version 2 of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
' without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
' See the GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License along with this
' program; if not, write to the
' Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
'

Option Explicit

' $Id$


Dim arHumanAttributes_() As String
Dim arLdapAttributes_() As String
Dim cKnownAttributes As Integer

Dim arUniqueAttributes_() As String
Dim cUniqueAttributes As Integer

Dim DLLVersion_ As Integer
Dim AttributeValueSeparator_ As String

Dim PreviousQueryString_ As String
'''''''''''''''''''''''''''''''''''


Sub ShowQueryForm()
    '
    ' load up the values from the config file
    '
    If initAttributeNames() <> True Then
        Exit Sub
    End If
    
    frmLdapQuery.Show
End Sub

Sub AddToolsMenuItem()
    '
    ' load up the values from the config file
    '
    If initAttributeNames() <> True Then
        Exit Sub
    End If
    
    frmLdapQuery.AddToolsMenuItem
End Sub


Function NumAttributes()
    NumAttributes = cKnownAttributes
End Function

Function arHumanAttributes(i As Integer)
    arHumanAttributes = arHumanAttributes_(i)
End Function

Function arLdapAttributes(i As Integer)
    arLdapAttributes = arLdapAttributes_(i)
End Function

Function NumUniqueAttributes()
    NumUniqueAttributes = cUniqueAttributes
End Function

Function arUniqueAttributes(i As Integer)
    arUniqueAttributes = arUniqueAttributes_(i)
End Function

Function DLLVersion()
    DLLVersion = DLLVersion_
End Function

Function AttributeValueSeparator()
    AttributeValueSeparator = AttributeValueSeparator_
End Function

Function PreviousQueryString()
    PreviousQueryString = PreviousQueryString_
End Function

Sub SetPreviousQueryString(queryString As String)
    PreviousQueryString_ = queryString
End Sub

'
' Load the attributes from the config file.
' We'll also use this as an occasion to load the baseDN since we have an active config object
'
' returns True/False to indicate if everything loaded OK or not
'
Function initAttributeNames()
   
    '
    ' Get the KnownAttributes object from the config file
    '
    Dim oLdapConfig As Object
    
    Set oLdapConfig = CreateObject("LdapQuery.LdapConfig")

    If oLdapConfig.Load <> 1 Then
        MsgBox "Unable to load the ldap_params.ini file. Check for it in the install dir or reinstall."
        initAttributeNames = False
        Exit Function
    End If

    
    Dim oKnownAttributes As Object
    Set oKnownAttributes = oLdapConfig.GetKnownAttributes()
    
    '
    ' Do some error checking and resize the attribute arrays now that we know the needed lengths
    '
    cKnownAttributes = oKnownAttributes.Count
    cUniqueAttributes = oKnownAttributes.UniqueCount
    
    If cKnownAttributes < 1 Then
        MsgBox "You need to define at least one LDAP attribute in your ldap_params.ini file. Opening file...", vbCritical
        openConfigFile
        initAttributeNames = False
        Exit Function
    End If
    
    If cUniqueAttributes < 1 Then
        MsgBox "You need to mark at least one attribute as unique (using '*') in your ldap_params.ini file. Opening file...", vbCritical
        openConfigFile
        initAttributeNames = False
        Exit Function
    End If
    
    ReDim arHumanAttributes_(1 To cKnownAttributes)
    ReDim arLdapAttributes_(1 To cKnownAttributes)
    ReDim arUniqueAttributes_(1 To cUniqueAttributes)

    '
    ' Finally, populate all the attribute arrays
    '
    Dim i As Integer
    For i = 1 To cKnownAttributes
        arHumanAttributes_(i) = oKnownAttributes.HumanName(0 + i)
        arLdapAttributes_(i) = oKnownAttributes.LdapName(0 + i)
    Next
    
    For i = 1 To cUniqueAttributes
        arUniqueAttributes_(i) = oKnownAttributes.UniqueLdapName(0 + i)
    Next
    
    '
    ' Initialize other config variables here
    '
    DLLVersion_ = oLdapConfig.DLLVersion
    AttributeValueSeparator_ = oLdapConfig.AttributeValueSeparator

    ' beware that this is the first use of frmLdapQuery and causes its Initialize
    ' sub to be called  I think?  hmmm it was causing problems if DLLVersion_ wasn't set beforehand
    frmLdapQuery.lblBaseDN.Caption = oLdapConfig.BaseDN
    frmLdapQuery.lblBaseDN.ControlTipText = oLdapConfig.BaseDN
    
    frmLdapQuery.tbQuery.Text = PreviousQueryString
    
    initAttributeNames = True
    
End Function


'
' translates a human attribute name (ex, "User ID") into its LDAP equivalent (ex, "uid")
'
Function humanToLdapAttribute(attributeName As String)
    Dim strAttribute
    Dim iAttribute As Integer
    Dim i As Integer
    iAttribute = -1
    
    ' find the index of the attribute
    For i = 1 To cKnownAttributes
        If arHumanAttributes(i) = attributeName Then
            iAttribute = i
            Exit For
        End If
    Next
    
    ' then look it up in the other array, assuming we found it
    If iAttribute = -1 Then
        humanToLdapAttribute = "INVALID ATTRIBUTE"
    Else
        humanToLdapAttribute = arLdapAttributes(iAttribute)
    End If

End Function

'
' translates an LDAP attribute name (ex, "uid") into its human equivalent (ex, "User ID")
'
Function LdapToHumanAttribute(attributeName As String)
    Dim strAttribute
    Dim iAttribute As Integer
    Dim i As Integer
    iAttribute = -1
    
    ' find the index of the attribute
    For i = 1 To cKnownAttributes
        If arLdapAttributes(i) = attributeName Then
            iAttribute = i
            Exit For
        End If
    Next
    
    ' then look it up in the other array, assuming we found it
    If iAttribute = -1 Then
        LdapToHumanAttribute = "INVALID ATTRIBUTE"
    Else
        LdapToHumanAttribute = arHumanAttributes(iAttribute)
    End If

End Function

Function ReadRegValue(RegKey As String, ByRef regValue)
    Dim WshShell
    Set WshShell = CreateObject("WScript.Shell")
    
    On Error Resume Next
    Err.Clear
    
    regValue = WshShell.RegRead("HKLM\SOFTWARE\Excel LDAP Search\" & RegKey)
    
    If Err <> 0 Then
        ReadRegValue = False
    Else
        ReadRegValue = True
    End If
    
    On Error GoTo 0
End Function

'
' Tries to read a reg key, uses specified default value if the key
' doesn't exist or an error occurred
'
Sub ReadRegValueWithDefault(RegKey As String, ByRef regValue, defaultValue)
    If Not ReadRegValue(RegKey, regValue) Then
        regValue = defaultValue
    End If
End Sub

' Takes three arguements - the key name, the value of the key,
' and the type of key (ex, "REG_SZ", "REG_DWORD", "REG_BIN",...)
Function WriteRegValue(RegKey As String, regValue, RegType As String)
    Dim WshShell
    Set WshShell = CreateObject("WScript.Shell")
    
    On Error Resume Next
    Err.Clear
    
    WshShell.RegWrite "HKLM\SOFTWARE\Excel LDAP Search\" & RegKey, regValue, RegType
    
    If Err <> 0 Then
        WriteRegValue = False
    Else
        WriteRegValue = True
    End If
    
    On Error GoTo 0
End Function

' Reads the installation directory from the registry, or returns the default dir if
' a problem occurrs
Function GetInstallDir()
    ' the easiest way to get the install dir is with the scripting shell
    Dim oWSH
    Set oWSH = CreateObject("WScript.Shell")
    Dim installDir
    
    On Error Resume Next
    Err.Clear
    GetInstallDir = oWSH.RegRead("HKLM\SOFTWARE\Excel LDAP Search\installDir")
    
    ' fall back to the default location
    If Err <> 0 Then
        GetInstallDir = "C:\program files\excel ldap search"
    End If
    
End Function

Sub openConfigFile()

    Dim oWSH
    Dim oEnv
    
    Set oWSH = CreateObject("WScript.Shell")
    Set oEnv = oWSH.Environment("process")
    
    Shell oEnv("SYSTEMROOT") & "\notepad.exe """ & GetInstallDir() & "\ldap_params.ini""", vbNormalFocus
End Sub

