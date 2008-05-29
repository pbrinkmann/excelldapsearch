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

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim LastQueryTime
Dim PBLastPercent

Dim PBBlankSearch As Integer
Dim PBAddSearch As Integer




' percent given as  0 -> 100
' please use "Done" as the last stage description to guarantee the bar will refresh
Private Sub SetProgressBarPercent(stageDescription As String, percent As Integer, whichProgressBar As Integer)

    Dim progressFrame
    Dim progressLabel
    
    If whichProgressBar = PBBlankSearch Then
        Set progressFrame = frmPBBlankSearch
        Set progressLabel = lblPBBlankSearch
    Else
        Set progressFrame = frmPBAddSearch
        Set progressLabel = lblPBAddSearch
    End If
    
    ' to prevent unnecessary drawing, only update when the percent changes
    If IsEmpty(PBLastPercent) Then
        PBLastPercent = -100
    End If
    If percent = PBLastPercent And stageDescription <> "Done" Then
        Exit Sub
    End If
    PBLastPercent = percent
    
    progressFrame.caption = stageDescription
    progressLabel.Width = progressFrame.Width * (percent / 100#)
    If progressLabel.Width = progressFrame.Width Then
        progressLabel.Width = progressLabel.Width - 4
    End If
    progressLabel.caption = percent & "%"

End Sub


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
    
    SetPreviousQueryString (queryText)
    
    Dim oLdap, oLdapConfig
  
    
    SetProgressBarPercent "Loading Config", 0, PBBlankSearch
    DoEvents
    

    
    Set oLdap = CreateObject("LdapQuery.LdapSearch")
    Set oLdapConfig = CreateObject("LdapQuery.LdapConfig")
    
    If oLdapConfig.Load <> 1 Then
        Target.Value = "Unable to load the ldap_params.ini file " & oLdapConfig.ServerName
        GoTo PEnd
    End If
    
    SetProgressBarPercent "Connecting to LDAP Server", 5, PBBlankSearch
    DoEvents
    
    If oLdap.Connect(oLdapConfig.ServerName, oLdapConfig.ServerPort, oLdapConfig.BindDN, oLdapConfig.BindPW) <> 1 Then
        Target.Value = oLdap.ErrorString
        GoTo PEnd
    End If
    
    SetProgressBarPercent "Performing Search", 10, PBBlankSearch
    DoEvents
    
    Dim cSearchResults
    cSearchResults = oLdap.Search(frmLdapQuery.lblBaseDN.caption, queryText)
    If cSearchResults = -1 Then
        Target.Value = oLdap.ErrorString
        GoTo PEnd
    End If
    
    SetProgressBarPercent "Displaying Results", 25, PBBlankSearch
    DoEvents
    
    Target.Value = "search found " & cSearchResults & " entries"
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
    Set oSearchResults = oLdap.GetSearchResults(AttributeValueSeparator)
   
    Dim oSearchResult
    Set oSearchResult = oSearchResults.GetFirstResult()
   
    i = 1
    
    Do While oSearchResult.IsValid() = 1
        DisplaySearchResult oSearchResult, Target, lbQueryAttributes
        Set Target = Target.Offset(1, 0)
        Set oSearchResult = oSearchResults.GetNextResult()
        SetProgressBarPercent "Displaying Results", (25 + 75 * i / cSearchResults), PBBlankSearch
        DoEvents
        i = i + 1
    Loop
    
    SetProgressBarPercent "Done", 100, PBBlankSearch
   
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
    
    '
    ' See if the config file has changed since the last query and if we should exit
    '
    If CheckConfigFileModified = True Then
        Exit Sub
    End If
    
    
    '
    ' See if the user needs a confirmation
    '
    If cbPromptOnOverwrite.Value = True Then
        If MsgBox("Are you sure you want to replace all contents of this sheet?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    
    ActiveSheet.Cells.Clear
    frmPBBlankSearch.Visible = True
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
    
    
    '
    ' See if the config file has been modified since our last query
    '
    If CheckConfigFileModified = True Then
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
    
    frmPBAddSearch.Visible = True
    SetProgressBarPercent "Loading Config", 5, PBAddSearch
    DoEvents
    
    Dim oLdap, oLdapConfig
    Set oLdap = CreateObject("LdapQuery.LdapSearch")
    Set oLdapConfig = CreateObject("LdapQuery.LdapConfig")
    
    If oLdapConfig.Load <> 1 Then
        lookupValues(1).Offset(0, lColOffset.caption).Value = "Unable to load the ldap_params.ini file"
        GoTo PEnd
    End If
    
    SetProgressBarPercent "Connecting to LDAP Server", 10, PBAddSearch
    DoEvents
    
    If oLdap.Connect(oLdapConfig.ServerName, oLdapConfig.ServerPort, oLdapConfig.BindDN, oLdapConfig.BindPW) <> 1 Then
        lookupValues(1).Offset(0, lColOffset.caption).Value = oLdap.ErrorString
        GoTo PEnd
    End If
    
    '
    ' loop through each row item we need to look up
    '
    i = 0
    
    For Each Item In lookupValues
        i = i + 1
        SetProgressBarPercent "Updating Rows", (10 + 90 * i / lookupValues.Count), PBAddSearch
        DoEvents
        
        ' skip blank items
        If Item.Value = "" Then
            GoTo ItemLoopEnd
        End If
    
        ' set the target for the lookup results
        Dim Target
        Set Target = Item.Offset(0, lColOffset.caption)
        
        ' run the search
        Dim c
        c = oLdap.Search(frmLdapQuery.lblBaseDN.caption, "(" & searchAttribute & "=" & Item.Value & ")")
        If c = -1 Then
            Target.Value = oLdap.ErrorString
            GoTo PEnd
        End If
        
        Dim oSearchResults
        Set oSearchResults = oLdap.GetSearchResults(AttributeValueSeparator)


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

    SetProgressBarPercent "Done", 100, PBAddSearch
    DoEvents

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
        .caption = "Run LDAP Query"
        '.OnAction = "'" & GetInstallDir() & "\macro_template.xlt'!ShowQueryForm"
        .OnAction = "'" & GetInstallDir() & "\ldapquery_addin.xla'!ShowQueryForm"
        .Tag = LdapQueryFormTag
    End With
    

End Sub

Private Function GetInstallDir()

    Dim installDir As String
    If ReadRegValue("installDir", installDir) Then
        GetInstallDir = installDir
    Else
        MsgBox "Failed to read the installation directory from the registry, so I'm going to use the default install directory"
        GetInstallDir = "C:\Program Files\Excel LDAP Search"
    End If

    
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

Private Sub cbPromptOnOverwrite_Change()
    Dim regValue As Integer
    
    If cbPromptOnOverwrite.Value = True Then
        regValue = 1
    Else
        regValue = 0
    End If
    
    WriteRegValue "promptOnOverwrite", regValue, "REG_DWORD"
End Sub




Private Sub frmExistingDataQuery_Click()

End Sub

Private Sub lblBaseDN_Click()
    Unload frmBaseDNChooser ' reset it
    frmBaseDNChooser.Show
    
    If (frmBaseDNChooser.GetSelectedDN <> "") Then
        lblBaseDN.caption = frmBaseDNChooser.GetSelectedDN
        lblBaseDN.ControlTipText = frmBaseDNChooser.GetSelectedDN
        
    End If
End Sub


Private Sub lblOpenConfigFile_Click()
    
    openConfigFile
    
End Sub

Private Sub lblViewReadme_Click()
    ShellExecute 0, "open", GetInstallDir() & "\readme.html", vbNullString, vbNullString, Empty
End Sub

Private Sub lblVersion_Click()
    ShellExecute 0, "open", "http://excelldapsearch.sourceforge.net/cgi-bin/version_check/?dll=" & DLLVersion, vbNullString, vbNullString, Empty
End Sub

Private Sub sbColOffset_SpinUp()
    lColOffset.caption = lColOffset.caption + 1
End Sub

Private Sub sbColOffset_SpinDown()
    lColOffset.caption = lColOffset.caption - 1
    
    If lColOffset.caption < 1 Then
        lColOffset.caption = 1
    End If
        
End Sub

Private Sub tsQueryType_Change()
    
    WriteRegValue "lastTabOpen", tsQueryType.Value, "REG_DWORD"
    
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
    ' set version string
    '
    lblVersion = "version 0." & DLLVersion  ' TODO: deal with 1.xx versions later
    
    '
    ' set some constants
    '
    PBBlankSearch = 1
    PBAddSearch = 2
    
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
    ' position the tab pages and set the active page
    '
    frmBlankSheetQuery.Top = 42
    frmBlankSheetQuery.Left = 18
    frmExistingDataQuery.Top = 42
    frmExistingDataQuery.Left = 18
    tsQueryType.Tabs(0).caption = "Add LDAP data to existing rows"
    tsQueryType.Tabs(1).caption = "Add LDAP entires to a blank sheet"
   
    Dim LastTabOpen As Integer
    
    If ReadRegValue("lastTabOpen", LastTabOpen) Then
        tsQueryType.Value = LastTabOpen
    Else
        tsQueryType.Value = 0
    End If
    
    '
    ' Set the overwrite checkbox value from the registry
    '
    Dim promptOrNot
    If ReadRegValue("promptOnOverwrite", promptOrNot) Then
        If promptOrNot = 1 Then
            cbPromptOnOverwrite.Value = True
        Else
            cbPromptOnOverwrite.Value = False
        End If
        
    End If

    
End Sub

'
' Returns true if the config file has been modified since the last call to this function,
' AND if the user wants to reload the settings
'
Function CheckConfigFileModified()

    CheckConfigFileModified = False

    Dim oLdapConfig
    Set oLdapConfig = CreateObject("LdapQuery.LdapConfig")
    If IsEmpty(LastQueryTime) Then ' first time through, no existing timestamp
        LastQueryTime = oLdapConfig.ConfigFileTimestamp
        'MsgBox ("config file modified on " & LastQueryTime)
    Else
        If LastQueryTime <> oLdapConfig.ConfigFileTimestamp Then
            If MsgBox("The config file has been modified since the settings were loaded, cancel query and reload?", vbYesNo) = vbYes Then
                frmLdapQuery.Hide
                Unload frmLdapQuery
                frmLdapQuery.Show
               CheckConfigFileModified = True
            End If
        End If
    End If
End Function



