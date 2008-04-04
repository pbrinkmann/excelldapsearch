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

Sub ShowQueryForm()
    frmLdapQuery.Show
End Sub

Sub AddToolsMenuItem()
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



'
' Load the attributes from the config file.
' We'll also use this as an occasion to load the baseDN since we have an active config object
'
Sub initAttributeNames()
   
    '
    ' Get the KnownAttributes object from the config file
    '
    Dim oLdapConfig As Object
    
    Set oLdapConfig = CreateObject("LdapQuery.LdapConfig")

    If oLdapConfig.Load <> 1 Then
        MsgBox "Unable to load the ldap_params.ini file"
        Exit Sub
    End If

    
    Dim oKnownAttributes As Object
    Set oKnownAttributes = oLdapConfig.GetKnownAttributes()
    
    '
    ' Resize the attribute arrays now that we know the needed lengths
    '
    cKnownAttributes = oKnownAttributes.Count
    cUniqueAttributes = oKnownAttributes.UniqueCount
    
    If cUniqueAttributes < 1 Then
        MsgBox "You need to mark at least one attribute as unique (using '*') in your ldap_params.ini file", vbCritical
        Exit Sub
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
    ' sneak in the baseDN thingy here
    '
    frmLdapQuery.lblBaseDN.Caption = oLdapConfig.BaseDN
    frmLdapQuery.lblBaseDN.ControlTipText = oLdapConfig.BaseDN
    
    '
    ' and grab the DLL version while we're doing things unlrelated to initializing attribute names,
    ' as this method name would imply.
    ' Maybe I should rename it some day?
    '
    DLLVersion_ = oLdapConfig.DLLVersion
End Sub


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

