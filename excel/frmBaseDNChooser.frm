VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBaseDNChooser 
   Caption         =   "Choose new BaseDN for searches"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   OleObjectBlob   =   "frmBaseDNChooser.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBaseDNChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Dim selectedDN As String

Public Function GetSelectedDN()
    GetSelectedDN = selectedDN
End Function

Private Sub bnCancel_Click()
    frmBaseDNChooser.Hide
    
    selectedDN = ""
End Sub

Private Sub bnOK_Click()
    Dim oTreeControl As Object
    Set oTreeControl = frmBaseDNChooser.Controls.Item("LdapTreeBrowser")
    selectedDN = oTreeControl.selectedDN
    
    If selectedDN = "invalid selection" Then
        selectedDN = ""
    End If
    
    frmBaseDNChooser.Hide
End Sub

Private Sub UserForm_Initialize()
    frmBaseDNChooser.Controls.Add "LdapQuery.LdapTreeBrowser", "LdapTreeBrowser"
    Dim oTreeControl As Object
    Set oTreeControl = frmBaseDNChooser.Controls.Item("LdapTreeBrowser")
    oTreeControl.Height = 240
    oTreeControl.Width = 238
    
    selectedDN = ""
End Sub
