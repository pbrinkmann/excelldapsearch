VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQueryBuilder 
   Caption         =   "Query Builder"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7080
   OleObjectBlob   =   "frmQueryBuilder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmQueryBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bnAddBottomRule_Click()
    lbBottomRules.AddItem (GetRuleBuilderRule)
End Sub

Private Sub bnAddTopRule_Click()
    lbTopRules.AddItem (GetRuleBuilderRule)
End Sub

' $Id$


Private Sub bnCancel_Click()
    Hide
    
    'Unload frmQueryBuilder    <-- don't do this, we want to save the last query
End Sub

Private Sub bnRemoveBottomRule_Click()
    If lbBottomRules.ListIndex >= 0 Then
        lbBottomRules.RemoveItem lbBottomRules.ListIndex
    End If
End Sub

Private Sub bnRemoveTopRule_Click()
    If lbTopRules.ListIndex >= 0 Then
        lbTopRules.RemoveItem lbTopRules.ListIndex
    End If
End Sub

Private Sub bnReset_Click()
    Hide
    Unload frmQueryBuilder
    frmQueryBuilder.Show
End Sub

Private Sub bnSaveQuery_Click()
    Dim query As String
    
    '
    ' we need at least one primary rule
    '
    If lbTopRules.ListCount = 0 Then
        MsgBox "Please add at least one rule to the primary search criteria", vbExclamation
        Exit Sub
    End If
    
    query = GetCriteriaQuery(lbTopRules, obnTopAny)
    
    ''
    '' now on to the secondary criteria
    ''
    
    If cbMatchSecondaryCriteria = True Then
        query = "(&" & query
        query = query & GetCriteriaQuery(lbBottomRules, obnBottomAny)
        query = query & ")"
    End If
    
    frmLdapQuery.tbQuery = query
    
    Hide
    'Unload frmQueryBuilder    <-- don't do this, we want to save the last query
    
End Sub

'
' builds LDAP query string for the rules in this query rule listbox
'
Private Function GetCriteriaQuery(listBox, matchAnyRule As Boolean)
    Dim query
    query = ""
    
    '
    ' if we have multiple rules, start with the OR or AND operator
    '
    If listBox.ListCount > 1 Then
        If matchAnyRule = True Then
            query = "(|"
        Else
            query = "(&"
        End If
    End If
    
    '
    ' append all the rules
    '
    Dim rule
    For Each rule In listBox.List
        query = query & rule
    Next
    
    '
    ' if we have multiple rules, close the initial AND/OR
    '
    If listBox.ListCount > 1 Then
        query = query & ")"
    End If
    
    GetCriteriaQuery = query
End Function

Private Sub cbMatchSecondaryCriteria_Click()
    EnableSecondaryCriteria cbMatchSecondaryCriteria.Value
End Sub

Private Sub UserForm_Initialize()
    
    initAttributeNames
    
    Dim i As Integer
    
    For i = 1 To NumAttributes()
        cbRuleAttribute.AddItem arHumanAttributes(i)
    Next
    cbRuleAttribute.ListIndex = 0

    ' if you change these, be sure to update GetRuleBuilderRule()
    cbRuleCondition.AddItem "is"
    cbRuleCondition.AddItem "is not"
    cbRuleCondition.AddItem "contains"
    cbRuleCondition.AddItem "doesn't contain"
    cbRuleCondition.AddItem "starts with"
    cbRuleCondition.AddItem "ends with"
    cbRuleCondition.ListIndex = 0
    
    EnableSecondaryCriteria False
    
End Sub

Private Function GetRuleBuilderRule()

    ' should the search term be preceeded by a wildcard (*)?
    Dim wildcardPrefix As String
    
    ' should the search term be followed by a wildcard (*)?
    Dim wildcardPostfix As String
    
    ' do we needed to negate the expression (ex, is NOT xyz)?
    Dim bNegate As Boolean
    
    wildcardPrefix = ""
    wildcardPostfix = ""
    
    Select Case cbRuleCondition
        Case "is"
            bNegate = False
        Case "contains"
            bNegate = False
            wildcardPrefix = "*"
            wildcardPostfix = "*"
        Case "is not"
            bNegate = True
        Case "doesn't contain"
            bNegate = True
            wildcardPrefix = "*"
            wildcardPostfix = "*"
        Case "starts with"
            bNegate = False
            wildcardPostfix = "*"
        Case "ends with"
            bNegate = False
            wildcardPrefix = "*"
        Case Else
            MsgBox "Yikes, the match type wasn't recognized!  This shouldn't happen"
    End Select
    
    ' Since there isn't a != operator, we just surround the expression with "(!" and ")"
    Dim negatePrefix As String
    Dim negatePostfix As String
    
    If bNegate Then
        negatePrefix = "(!"
        negatePostfix = ")"
    Else
        negatePrefix = ""
        negatePostfix = ""
    End If

    GetRuleBuilderRule = negatePrefix & "(" & humanToLdapAttribute(cbRuleAttribute) & "=" & wildcardPrefix & tbRuleValue & wildcardPostfix & ")" & negatePostfix
    
End Function

Private Sub EnableSecondaryCriteria(bEnable As Boolean)
    bnAddBottomRule.Enabled = bEnable
    obnBottomAny.Enabled = bEnable
    obnBottomAll.Enabled = bEnable
    lbBottomRules.Enabled = bEnable
    bnRemoveBottomRule.Enabled = bEnable
    
    If bEnable = True Then
        lbBottomRules.BackColor = &HFFFFFF
    Else
        lbBottomRules.BackColor = &HC0C0C0
    End If
    
End Sub
