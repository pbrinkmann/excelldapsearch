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
