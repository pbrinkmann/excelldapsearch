VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLdapQuery 
   Caption         =   "Ldap Query"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6885
   OleObjectBlob   =   "frmLdapQuery.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLdapQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' $Id$

'
' given a search result, a target cell for the results, and a listbox with a list of
' attributes to display selected, this will print out those attributes starting at the target cell
' and moving right one cell for each subsequent attribute
'
Private Sub DisplaySearchResult(oSearchResult, Target, lbReturnAttributes)
    Dim cAttributes
    Dim i
    Dim oAttribute
    
    cAttributes = 0
    '
    ' loop through all the attributes we've requested, and
    ' print them if this search result has it
    '
    For i = 0 To lbReturnAttributes.ListCount - 1
        If lbReturnAttributes.Selected(i) Then
            ' handle special case of the DN
            If lbReturnAttributes.List(i) = "LDAP DN" Then
                Target.Offset(0, cAttributes).Value = oSearchResult.DN
            Else
                ' TODO: calling humanToLdapAttribute each time here is not optimal
                ' the index (i) should be the same as the attribute arrays so we should just directly index ldapAttributes
                ' ...  I think ...
                Set oAttribute = oSearchResult.GetAttributeByName(humanToLdapAttribute(lbReturnAttributes.List(i)))
                If oAttribute.IsValid Then
                    Target.Offset(0, cAttributes).Value = oAttribute.Value
                Else
                    Target.Offset(0, cAttributes).Value = "NO VALUE"
                End If
            End If
        
        cAttributes = cAttributes + 1
    End If
    Next ' end of attribute loop
End Sub


Private Sub DoSearch(queryText, Target As Range)

    On Error GoTo DoSearchError
    
    Dim oLdap, oLdapConfig
    
    Set oLdap = CreateObject("LdapQuery.LdapSearch")
    Set oLdapConfig = CreateObject("LdapQuery.LdapConfig")
    
    If oLdapConfig.Load <> 1 Then
        Target.Value = "Unable to load the ldap_params.ini file " & oLdapConfig.ServerName
        GoTo PEnd
    End If
    
    If oLdap.Connect(oLdapConfig.ServerName, oLdapConfig.BindDN, oLdapConfig.BindPW) <> 1 Then
        Target.Value = oLdap.ErrorString
        GoTo PEnd
    End If
    
    Dim c
    c = oLdap.Search(frmLdapQuery.lblBaseDN.Caption, queryText)
    If c = -1 Then
        Target.Value = oLdap.ErrorString
        GoTo PEnd
    End If
    
    Target.Value = "search found " & c & " entries"
    Set Target = Target.Offset(1, 0)
    
    '
    ' print out the header columns for out returned attributes
    '
    Dim cHeaderColumn
    Dim i
    cHeaderColumn = 0
    For i = 0 To lbQueryAttributes.ListCount - 1
        If lbQueryAttributes.Selected(i) Then
            Target.Offset(0, cHeaderColumn).Font.Bold = True
            Target.Offset(0, cHeaderColumn) = lbQueryAttributes.List(i)
            cHeaderColumn = cHeaderColumn + 1
        End If
    Next
    Set Target = Target.Offset(1, 0)
   
    Dim oSearchResults
    Set oSearchResults = oLdap.GetSearchResults()
   
    Dim oSearchResult
    Set oSearchResult = oSearchResults.GetFirstResult()
   
    Do While oSearchResult.IsValid() = 1
        DisplaySearchResult oSearchResult, Target, lbQueryAttributes
        Set Target = Target.Offset(1, 0)
        Set oSearchResult = oSearchResults.GetNextResult()
    Loop
   
PEnd:

    Exit Sub
    
DoSearchError:

    MsgBox "An error occurred: " & Err.Description & " at " & Err.Source, vbExclamation
    Err.Clear

End Sub


Private Sub bnClose_Click()
    frmLdapQuery.Hide
    Unload frmLdapQuery
End Sub


Private Sub bnOpenQueryBuilder_Click()
    frmQueryBuilder.Show
End Sub


Private Sub bnRunQuery_Click()
    '
    ' do some error checking first before we try and run the query
    '
    If tbQuery = "" Then
        MsgBox "Please enter a value in the query area", vbExclamation
        Exit Sub
    End If
    
    Dim i
    Dim bFoundOne As Boolean
    bFoundOne = False
    For i = 0 To lbQueryAttributes.ListCount - 1
        If lbQueryAttributes.Selected(i) Then
            bFoundOne = True
            GoTo endOfSelectLoop
        End If
    Next
endOfSelectLoop:

    If bFoundOne = False Then
        MsgBox "Please select at least one attribute to return", vbExclamation
        Exit Sub
    End If
    
    ActiveSheet.Cells.Clear
    DoSearch tbQuery.Value, ActiveSheet.Range("A1")
End Sub

Private Sub bnRowLookup_Click()

    On Error GoTo RowLookupClickError

    '
    ' form error checking first
    '
    If reDataRange.Value = "" Then
        MsgBox "Please enter a data range containing the attribute to search for", vbExclamation
        Exit Sub
    End If
    
    Dim i
    Dim bFoundOne As Boolean
    bFoundOne = False
    For i = 0 To lbReturnAttributes.ListCount - 1
        If lbReturnAttributes.Selected(i) Then
            bFoundOne = True
            GoTo endOfSelectLoop
        End If
    Next
endOfSelectLoop:
    
    If bFoundOne = False Then
        MsgBox "Please select at least one attribute to return", vbExclamation
        Exit Sub
    End If
    
    
    
    
    Dim Item
    
    '
    ' the range containing the values to run the LDAP searches on
    '
    Dim lookupValues As Range
    Set lookupValues = ActiveSheet.Range(reDataRange.Value)
    
    '
    ' which attribute are we searching for?
    '
    Dim searchAttribute
    searchAttribute = humanToLdapAttribute(cbSearchAttribute.Value)
    
    Dim oLdap, oLdapConfig
    Set oLdap = CreateObject("LdapQuery.LdapSearch")
    Set oLdapConfig = CreateObject("LdapQuery.LdapConfig")
    
    If oLdapConfig.Load <> 1 Then
        lookupValues(1).Offset(0, lColOffset.Caption).Value = "Unable to load the ldap_params.ini file"
        GoTo PEnd
    End If
    
    If oLdap.Connect(oLdapConfig.ServerName, oLdapConfig.BindDN, oLdapConfig.BindPW) <> 1 Then
        lookupValues(1).Offset(0, lColOffset.Caption).Value = oLdap.ErrorString
        GoTo PEnd
    End If
    
    '
    ' loop through each row item we need to look up
    '
    For Each Item In lookupValues
        
        ' skip blank items
        If Item.Value = "" Then
            GoTo ItemLoopEnd
        End If
    
        ' set the target for the lookup results
        Dim Target
        Set Target = Item.Offset(0, lColOffset.Caption)
        
        ' run the search
        Dim c
        c = oLdap.Search(frmLdapQuery.lblBaseDN.Caption, "(" & searchAttribute & "=" & Item.Value & ")")
        If c = -1 Then
            Target.Value = oLdap.ErrorString
            GoTo PEnd
        End If
        
        Dim oSearchResults
        Set oSearchResults = oLdap.GetSearchResults()


        '
        ' There should be exactly be one search result
        '
        If oSearchResults.ResultCount > 1 Then
            Target.Value = "MULTIPLE ENTRIES FOUND FOR THIS ROW, PLEASE REFINE SEARCH ATTRIBUTE"
            GoTo ItemLoopEnd
        End If
        
        If oSearchResults.ResultCount = 0 Then
            Target.Value = "ENTRY NOT FOUND"
            GoTo ItemLoopEnd
        End If
        
        '
        ' get the first (and only) search result
        '
        Dim oSearchResult
        Set oSearchResult = oSearchResults.GetFirstResult()
        
        '
        ' this check should be redundant
        '
        If oSearchResult.IsValid() = 0 Then
            Target.Value = "ENTRY NOT FOUND"
            GoTo ItemLoopEnd
        End If
        
        DisplaySearchResult oSearchResult, Target, lbReturnAttributes
      

ItemLoopEnd:

    Next ' end of row loop
  
   
PEnd:

    Exit Sub

RowLookupClickError:

    MsgBox "An error occurred: " & Err.Description & " from " & Err.Source, vbExclamation
    Err.Clear

End Sub


'
' Adds an item to the Tools menu to open the search form
'
Sub AddToolsMenuItem()

    Dim LdapQueryFormTag
    LdapQueryFormTag = "LdapQueryShowFormMenuItem"
    
    Dim CmdBar As CommandBar
    Dim CmdBarMenu As CommandBarControl
    Dim CmdBarMenuItem As CommandBarControl
    
    '
    ' Point to the Worksheet Menu Bar
    '
    Set CmdBar = Application.CommandBars("Worksheet Menu Bar")
    
    '
    ' Point to the Tools menu on the menu bar
    '
    Set CmdBarMenu = CmdBar.Controls("Tools")
    
    '
    ' See if we already have an item in there, and exit if we do
    '
    On Error Resume Next
    Dim CmdCtrl As CommandBarControl
    For Each CmdCtrl In Application.CommandBars.FindControls(Tag:=LdapQueryFormTag)
        If Not CmdCtrl Is Nothing Then ' for some reason, we always get one iteration through the foreach with a "nothing" object
            Exit Sub
        End If
    Next CmdCtrl
    
    '
    ' Add a new menu item to the Tools menu
    '
    Set CmdBarMenuItem = CmdBarMenu.Controls.Add
    
    '
    ' Set the properties for the new control
    '
    With CmdBarMenuItem
        .Caption = "Run LDAP Query"
        '.OnAction = "'" & GetInstallDir() & "\macro_template.xlt'!ShowQueryForm"
        .OnAction = "'" & GetInstallDir() & "\ldapquery_addin.xla'!ShowQueryForm"
        .Tag = LdapQueryFormTag
    End With
    

End Sub

Private Function GetInstallDir()
    Dim WshShell
    Set WshShell = CreateObject("WScript.Shell")
    
    Dim regKey, regValue
    
    On Error Resume Next
    Err.Clear
    
    regKey = "HKLM\SOFTWARE\Excel LDAP Search\installDir"
    regValue = WshShell.RegRead(regKey)
    
    If Err <> 0 Then
        MsgBox "Failed to read the installation directory from the registry, so I'm going to use the default install directory"
        GetInstallDir = "C:\Program Files\Excel LDAP Search"
        Exit Function
    End If
    
    On Error GoTo 0
    
    GetInstallDir = regValue
    
End Function

Private Function IsSelectedAttribute(TestAttrib)
    Dim i
        For i = 0 To lbReturnAttributes.ListCount - 1
        If lbReturnAttributes.Selected(i) Then
            If lbReturnAttributes.List(i) = TestAttrib Then
                IsSelectedAttribute = True
                Exit Function
            End If
        End If
    Next
    IsSelectedAttribute = False
End Function






Private Sub Frame2_Click()

End Sub

Private Sub lblBaseDN_Click()
    Unload frmBaseDNChooser ' reset it
    frmBaseDNChooser.Show
    
    If (frmBaseDNChooser.GetSelectedDN <> "") Then
        lblBaseDN.Caption = frmBaseDNChooser.GetSelectedDN
    End If
End Sub

' Keep the tooltip updated (is there a better way?)
Private Sub lblBaseDN_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblBaseDN.ControlTipText = lblBaseDN.Caption
End Sub

Private Sub lblOpenConfigFile_Click()
    
    ' the easiest way to get the install dir is with the scripting shell
    Dim oWSH
    Set oWSH = CreateObject("WScript.Shell")
    Dim installDir
    
    On Error Resume Next
    Err.Clear
    installDir = oWSH.RegRead("HKLM\SOFTWARE\Excel LDAP Search\installDir")
    
    ' fall back to the default location
    If Err <> 0 Then
        installDir = "C:\program files\excel ldap search"
    End If
    
    Shell Environ("windir") & "\notepad.exe """ & installDir & "\ldap_params.ini""", vbNormalFocus
End Sub

Private Sub sbColOffset_SpinUp()
    lColOffset.Caption = lColOffset.Caption + 1
End Sub

Private Sub sbColOffset_SpinDown()
    lColOffset.Caption = lColOffset.Caption - 1
    
    If lColOffset.Caption < 1 Then
        lColOffset.Caption = 1
    End If
        
End Sub

Private Sub tsQueryType_Change()
    If tsQueryType.Value = 0 Then
        frmExistingDataQuery.Visible = True
        frmExistingDataQuery.Enabled = True
        frmBlankSheetQuery.Visible = False
        frmBlankSheetQuery.Enabled = False
    Else
        frmExistingDataQuery.Visible = False
        frmExistingDataQuery.Enabled = False
        frmBlankSheetQuery.Visible = True
        frmBlankSheetQuery.Enabled = True
    End If
        
End Sub

Private Sub UserForm_Initialize()
    
    reDataRange.Value = Selection.Address
    
    '
    ' load up the values from the config file
    '
    initAttributeNames
    
    '
    ' fill in the unique attributes
    '
    Dim i As Integer
    
    
    For i = 1 To NumUniqueAttributes()
        cbSearchAttribute.AddItem LdapToHumanAttribute(arUniqueAttributes(i))
    Next
    
    cbSearchAttribute.Value = LdapToHumanAttribute(arUniqueAttributes(1))
    
    '
    ' fill in the searchable attributes
    '
    
    For i = 1 To NumAttributes()
        lbReturnAttributes.AddItem arHumanAttributes(i)
        lbQueryAttributes.AddItem arHumanAttributes(i)
    Next
    
    lbReturnAttributes.AddItem "LDAP DN"
    lbQueryAttributes.AddItem "LDAP DN"
    
    
    '
    ' position the tab pages and turn off the secondary page
    '
    frmBlankSheetQuery.Top = 42
    frmBlankSheetQuery.Left = 18
    frmExistingDataQuery.Top = 42
    frmExistingDataQuery.Left = 18
    frmBlankSheetQuery.Visible = False
    frmBlankSheetQuery.Enabled = False
    tsQueryType.Tabs(0).Caption = "Add LDAP data to existing rows"
    tsQueryType.Tabs(1).Caption = "Add LDAP entires to a blank sheet"
   
End Sub


