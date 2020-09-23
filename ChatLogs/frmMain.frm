VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Chat Log"
   ClientHeight    =   8700
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   10830
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   360
      Left            =   0
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   722
      TabIndex        =   0
      Top             =   0
      Width           =   10830
      Begin VB.CommandButton cmdToolbar 
         Height          =   360
         Index           =   10
         Left            =   4050
         Picture         =   "frmMain.frx":2CFA
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Options"
         Top             =   15
         Width           =   375
      End
      Begin VB.CommandButton cmdToolbar 
         Height          =   360
         Index           =   9
         Left            =   3690
         Picture         =   "frmMain.frx":30BD
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Word count"
         Top             =   15
         Width           =   375
      End
      Begin VB.CommandButton cmdToolbar 
         Height          =   360
         Index           =   8
         Left            =   3330
         Picture         =   "frmMain.frx":347C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Time chart"
         Top             =   15
         Width           =   375
      End
      Begin VB.CommandButton cmdToolbar 
         Height          =   360
         Index           =   7
         Left            =   2820
         Picture         =   "frmMain.frx":380D
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Session list"
         Top             =   15
         Width           =   375
      End
      Begin VB.CommandButton cmdToolbar 
         Height          =   360
         Index           =   6
         Left            =   2460
         Picture         =   "frmMain.frx":3BA5
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "User list"
         Top             =   15
         Width           =   375
      End
      Begin VB.CommandButton cmdToolbar 
         Height          =   360
         Index           =   5
         Left            =   2100
         Picture         =   "frmMain.frx":3F3C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Nick list"
         Top             =   15
         Width           =   375
      End
      Begin VB.CommandButton cmdToolbar 
         Height          =   360
         Index           =   4
         Left            =   1590
         Picture         =   "frmMain.frx":42D4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Search"
         Top             =   15
         Width           =   375
      End
      Begin VB.CommandButton cmdToolbar 
         Height          =   360
         Index           =   3
         Left            =   1230
         Picture         =   "frmMain.frx":4662
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Clear all"
         Top             =   15
         Width           =   375
      End
      Begin VB.CommandButton cmdToolbar 
         Height          =   360
         Index           =   2
         Left            =   735
         Picture         =   "frmMain.frx":46EA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Export to file(s)"
         Top             =   15
         Width           =   375
      End
      Begin VB.CommandButton cmdToolbar 
         Height          =   360
         Index           =   1
         Left            =   375
         Picture         =   "frmMain.frx":4A68
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Open folder"
         Top             =   15
         Width           =   375
      End
      Begin VB.CommandButton cmdToolbar 
         Height          =   360
         Index           =   0
         Left            =   15
         Picture         =   "frmMain.frx":4DF0
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Open file"
         Top             =   15
         Width           =   375
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoadFile 
         Caption         =   "&Load file"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuLoadFolder 
         Caption         =   "&Load folder"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export ..."
         Begin VB.Menu mnuExportLogs 
            Caption         =   "Sessions"
            Shortcut        =   ^E
         End
         Begin VB.Menu mnuLine7 
            Caption         =   "-"
         End
         Begin VB.Menu mnuExportPlugin 
            Caption         =   "(Plugin)"
            Index           =   0
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear all"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuNickList 
         Caption         =   "&Nick list"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSessionList 
         Caption         =   "&Session list"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuUserList 
         Caption         =   "&User list"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuLine6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTimeChart 
         Caption         =   "&Time Chart"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuWordCount 
         Caption         =   "&Word Count"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options ..."
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnuListView 
      Caption         =   "&ListView"
      Visible         =   0   'False
      Begin VB.Menu mnuClone 
         Caption         =   "&Clone"
      End
      Begin VB.Menu mnuLine8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Copyright (C) 2005 Kristian S. Stangeland

' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.

' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.

' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

' The class module containing all the logs
Public WithEvents Logs As clsContainer
Attribute Logs.VB_VarHelpID = -1

' Used to save which control to be cloned or printed
Public MenuControl As Object

Private Sub cmdToolbar_Click(Index As Integer)

    ' Call the appropriate function to deal with this click
    Select Case Index
    Case 0: mnuLoadFile_Click
    Case 1: mnuLoadFolder_Click
    Case 2: mnuExportLogs_Click
    Case 3: mnuClear_Click
    Case 4: mnuSearch_Click
    Case 5: mnuNickList_Click
    Case 6: mnuUserList_Click
    Case 7: mnuSessionList_Click
    Case 8: mnuTimeChart_Click
    Case 9: mnuWordCount_Click
    Case 10: mnuOptions_Click
    End Select

End Sub

Private Sub Logs_FileFound(ByVal FileName As String, AddManually As Boolean)

    ' Inform about the file
    frmProgress.lblDescription.Caption = "Adding " & FileName & " ..."
    
    ' Increase file count
    frmProgress.progressBar.Value = frmProgress.progressBar.Value + 1

    ' Update so we can se the change on the form
    DoEvents

End Sub

Private Sub Logs_FilesListed(ByVal Path As String, ByVal Amout As Long)

    On Error Resume Next
    
    ' Firstly, show the form and set the properties of its controls
    frmProgress.Show
    frmProgress.lblDescription = "Adding files from folders ..."
    frmProgress.progressBar.Value = 0
    frmProgress.progressBar.Max = Amout
    
End Sub

Private Sub Logs_StatisticCalculation(ByVal UserIndex As Long, ByVal Max As Long)

    ' Show the progress-form and set its values
    frmProgress.Show
    frmProgress.lblDescription = "Calculating statistics for user " & UserIndex & " of total " & Max & " ..."
    frmProgress.progressBar.Value = UserIndex
    frmProgress.progressBar.Max = Max
    
    ' Show changes
    DoEvents

End Sub

Private Sub Logs_StatisticsCalculated()

    ' We are finished, hide the progress-window
    frmProgress.Hide

End Sub

Private Sub Logs_FolderImported(ByVal Path As String)

    ' We are finished, hide the progress-window
    frmProgress.Hide

End Sub

Private Sub MDIForm_Load()

    ' Create a new container
    Set Logs = New clsContainer

    ' Initialize this form
    ChatApp.InitializeForm Me

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    ' Exit the application
    ChatApp.Terminate

End Sub

Private Sub mnuAbout_Click()

    ' Now, show the form
    frmAbout.Show
    
    ' Then start the animation
    frmAbout.StartAnimation

End Sub

Private Sub mnuClear_Click()

    ' Clear everything (by renewing the logs-class)
    Set Logs = New clsContainer

End Sub

Private Sub mnuClone_Click()

    ' Clone the correct form
    MenuControl.Parent.Clone

End Sub

Private Sub mnuExit_Click()

    ' Unload this form
    Unload Me

End Sub

Private Sub mnuExportLogs_Click()

    ' Show the export dialog
    frmExport.Update
    frmExport.Show

End Sub

Private Sub mnuExportPlugin_Click(Index As Integer)

    ' Execute the plugin-procedure
    ChatApp.ExportModules(mnuExportPlugin(Index).Tag).ExportProc Logs, New clsCommonDialogs

End Sub

Private Sub mnuLoadFile_Click()

    Dim OpenSave As New clsCommonDialogs
    
    ' Firstly, initialize the filter
    OpenSave.Filter = ChatApp.FileInputFilters
    
    ' Then, show the dialog
    OpenSave.ShowOpen Me.hWnd, "Open log"
    
    ' And continue if open was selected
    If OpenSave.File <> "" Then
    
        ' Load this file
        LoadFile OpenSave.File
    
    End If

End Sub

Public Sub LoadFile(sFile As String)

    ' Inform the parser that we're about to begin
    ChatApp.Parsers(ChatApp.ParserByFile(sFile)).ParserInitialize

    ' The file type is determined by analyzing the file extension
    Logs.AddLog sFile, ChatApp.ParserByFile(sFile)
    
    ' Refresh values
    Logs.RefreshStatistic

End Sub

Private Sub mnuLoadFolder_Click()

    Dim FolderBrowse As New clsFolderBrowse

    ' Customize form
    FolderBrowse.Description = LanguageConst("FolderBrowseText")
    FolderBrowse.Filters = ChatApp.FileInputFilters

    ' Firstly, ask the user of what folder to get
    If FolderBrowse.Browse = Browse_OK Then
    
        ' Then import the folder selected by the user
        Logs.ImportFolder FolderBrowse.ResultPath, FolderBrowse.ResultFilter, FolderBrowse.SubFolders
        
    End If

End Sub

Private Sub mnuNickList_Click()

    ' Invert value
    mnuNickList.Checked = Not mnuNickList.Checked

    ' Hide or show the form depending of the current check value
    If mnuNickList.Checked Then
        frmNickList.Update
        frmNickList.Show
    Else
        frmNickList.Hide
    End If

End Sub

Private Sub mnuOptions_Click()

    ' Initialize settings
    frmSettings.LoadAll
    
    ' Show the form
    frmSettings.Show

End Sub

Private Sub mnuPrint_Click()

    Dim Dialogs As New clsCommonDialogs

    ' Show the print-dialog and go further if the user has pressed OK
    If Dialogs.ShowPrinter(Me) = 0 Then
        ' Print the saved control
        ChatApp.PrintClass.PrintControl MenuControl
    End If

End Sub

Private Sub mnuSearch_Click()

    ' Show the search dialog box
    frmSearch.Show

End Sub

Private Sub mnuSessionList_Click()

    ' Invert value
    mnuSessionList.Checked = Not mnuSessionList.Checked

    ' Hide or show the form depending of the current check value
    If mnuSessionList.Checked Then
        frmSessionList.Caption = LanguageConst("AllSessions")
        frmSessionList.Show
        frmSessionList.Update Logs.Sessions
    Else
        frmSessionList.Hide
    End If

End Sub

Private Sub mnuTimeChart_Click()

    ' Set the source
    Set frmTimeChart.Sessions = frmMain.Logs.Sessions

    ' Show the time-chart-form
    frmTimeChart.Show

End Sub

Private Sub mnuUserList_Click()

    ' Invert value
    mnuUserList.Checked = Not mnuUserList.Checked

    ' Hide or show the form depending of the current check value
    If mnuUserList.Checked Then
        frmUserList.Show
        frmUserList.Update
    Else
        frmUserList.Hide
    End If

End Sub

Private Sub mnuView_Click()

    ' Make sure that the check-value of all form-elements are correct
    mnuUserList.Checked = ChatApp.IsVisible(frmUserList)
    mnuSessionList.Checked = ChatApp.IsVisible(frmSessionList)
    mnuNickList.Checked = ChatApp.IsVisible(frmNickList)
    
    ' In practice, this redraws the menu elements (which the above code obviously don't)
    mnuUserList.Caption = mnuUserList.Caption
    mnuSessionList.Caption = mnuSessionList.Caption
    mnuNickList.Caption = mnuNickList.Caption

End Sub

Private Sub mnuWordCount_Click()

    ' Show the word-form
    frmWordCount.Show

End Sub

Public Sub Update()

    Dim Plugin As Object, Tell As Long, bNotFirst As Boolean

    ' Hide the first element
    mnuExportPlugin(0).Visible = False

    ' Unload all created menu item
    For Tell = mnuExportPlugin.LBound + 1 To mnuExportPlugin.UBound
        Unload mnuExportPlugin(Tell)
    Next
    
    ' Reset variable
    Tell = mnuExportPlugin.LBound

    ' Update the export plugins
    For Each Plugin In ChatApp.ExportModules
    
        ' Only add it if this module is going to handle all work by ifself
        If Plugin.ExportHandleAll Then
    
            ' Create a new menu item if necessary
            If Tell > mnuExportPlugin.UBound Then
            
                ' Create a new item
                Load mnuExportPlugin(Tell)
            
            End If
            
            ' Set the properties of this export module
            mnuExportPlugin(Tell).Visible = True
            mnuExportPlugin(Tell).Caption = Plugin.MenuCaption
            mnuExportPlugin(Tell).Tag = Plugin.ExportName
            
            ' Increase the counter
            Tell = Tell + 1
        
        End If
    
    Next

End Sub
