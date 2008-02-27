Attribute VB_Name = "ldapQueryFormMacro"
Option Explicit

' $Id$


Dim arHumanAttributes_() As String
Dim arLdapAttributes_() As String
Dim cKnownAttributes As Integer

Dim arUniqueAttributes_() As String
Dim cUniqueAttributes As Integer


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
    ' We can immediately resize the attribute name arrays to the proper size
    '
    cKnownAttributes = oKnownAttributes.Count
    
    ReDim arHumanAttributes_(1 To cKnownAttributes)
    ReDim arLdapAttributes_(1 To cKnownAttributes)
    
    '
    ' It's a little trickier to get the # of unique attributes,
    ' but doing a quick iteration through all attributes gets the job done
    '
    Dim i As Integer
    cUniqueAttributes = 0

    For i = 1 To cKnownAttributes
        If InStr(oKnownAttributes.LdapName(0 + i), "*") = 1 Then
            cUniqueAttributes = cUniqueAttributes + 1
        End If
    Next
    
    If cUniqueAttributes < 1 Then
        MsgBox "You need to mark at least one attribute as unique (using '*') in your ldap_params.ini file", vbCritical
        Exit Sub
    End If
    
    ReDim arUniqueAttributes_(1 To cUniqueAttributes)
    
    '
    ' Finally, populate all the attribute arrays
    '
    Dim uniqueIndex ' current index into unique attribute array
    uniqueIndex = 1
    For i = 1 To cKnownAttributes
    
        ' strip the leading * and add to unique attributes if needed
        Dim ldapAttribute
        ldapAttribute = oKnownAttributes.LdapName(0 + i)
        
        If InStr(ldapAttribute, "*") = 1 Then
            ldapAttribute = Right(ldapAttribute, Len(ldapAttribute) - 1)
            arUniqueAttributes_(uniqueIndex) = ldapAttribute
            uniqueIndex = uniqueIndex + 1
        End If
        
        
        arHumanAttributes_(i) = oKnownAttributes.HumanName(0 + i)
        arLdapAttributes_(i) = ldapAttribute
        
        'MsgBox "just set: " & arHumanAttributes_(i) & " = " & arLdapAttributes_(i)
    Next
    
    '
    ' sneak in the baseDN thingy here
    '
    frmLdapQuery.lblBaseDN.Caption = oLdapConfig.BaseDN
    
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




