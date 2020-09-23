VERSION 5.00
Object = "{E910F8E1-8996-4EE9-90F1-3E7C64FA9829}#1.1#0"; "vbaListView6.ocx"
Begin VB.Form frmWordCount 
   Caption         =   "Word Count"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   710
   Begin vbalListViewLib6.vbalListViewCtl lstWords 
      Height          =   6375
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   11245
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   1
      MultiSelect     =   -1  'True
      LabelEdit       =   0   'False
      AutoArrange     =   0   'False
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5280
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   353
      TabIndex        =   0
      Top             =   6960
      Width           =   5295
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmdFilters 
         Caption         =   "&Filters"
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   495
         Left            =   3600
         TabIndex        =   1
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.Label lblUnique 
      Caption         =   "&Unique words: "
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmWordCount"
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

Public WithEvents Search As clsWordCount
Attribute Search.VB_VarHelpID = -1
Public Filters As New Collection

Public Function Clone() As frmWordCount

    ' Create a new form
    Set Clone = New frmWordCount
    
    ' and finally show it
    Clone.Show

End Function

Private Sub cmdClose_Click()
    
    ' Just remove this form
    Unload Me

End Sub

Private Sub cmdFilters_Click()

    Dim currFilter As New frmFilter
    
    ' Show the new filter-dialog
    currFilter.Show
    
    ' Set the filter-class to this one
    Set currFilter.currFilters = Filters
    
    ' Update the filter-list
    currFilter.UpdateFilters currFilter.lstFilters
    
    ' Wait until the user is finished
    currFilter.WaitForFinish
    
    ' Remove the form
    Set currFilter = Nothing
    
End Sub

Private Sub cmdSearch_Click()

    ' Clear the listview
    lstWords.ListItems.Clear

    ' Initialize the class
    Set Search.ListView = lstWords
    Set Search.Sessions = frmMain.Logs.Sessions
    Set Search.Filters = Filters
    
    ' Count all words
    Search.CountWords
    
    ' Show the amount of words
    lblUnique.Caption = "&Unique words: : " & lstWords.ListItems.Count

End Sub

Private Sub Form_Load()

    ' Initialize the class
    Set Search = New clsWordCount
    
    ' Initialize this form
    ChatApp.InitializeForm Me
    
End Sub

Private Sub lstWords_ColumnClick(Column As cColumn)

    ' "Invert" the sort-order
    Column.SortOrder = NewSortOrder(Column.SortOrder)

    ' Sort according to the column type:
    Select Case Column.Key
    Case "Word" ' Sort by text
        Column.SortType = eLVSortTag
        
    Case "Occurrences" ' Sort by number
        Column.SortType = eLVSortItemData

    End Select
    
    ' Sort everything
    lstWords.ListItems.SortItems
   
End Sub

Private Sub Form_Resize()

    ' Only resize when we can
    If Me.WindowState <> 1 Then

        ' Resize the listview
        lstWords.Width = Me.ScaleWidth - 16
        lstWords.Height = Me.ScaleHeight - picToolbar.Height - lstWords.Top - 16
        
        ' Then set the position of the toolbar
        picToolbar.Top = lstWords.Height + lstWords.Top + 8
        picToolbar.Left = Me.ScaleWidth - 8 - picToolbar.Width
        
    End If

End Sub

Private Sub lstWords_ItemDblClick(Item As vbalListViewLib6.cListItem)
    
    Dim SearchClass As New clsSearch, frmResult As frmMessageList, oFilters As New Collection, Filter As Object
    
    ' Firstly, add all ordinary filters
    For Each Filter In Filters
        oFilters.Add Filter
    Next

    ' Then add the text-filter
    oFilters.Add ChatApp.CreateFilter(Filter_TextMatch, True, "(^|\W)" & Item.Text & "($|\W)", , True, True)
    
    ' Initialize class
    With SearchClass
        
        ' Initialize source
        Set .Container = frmMain.Logs
    
        ' Set the filters
        Set .Filters = oFilters

    End With
    
    ' Execute search
    SearchClass.InitiateSearch
    
    ' Show the result by creating a new form
    Set frmResult = New frmMessageList
    
    ' And update it
    frmResult.Update SearchClass.Messages

End Sub

Private Sub Search_SavingResult()

    ' Inform about the operation
    frmProgress.lblDescription.Caption = "Writing result to list view ..."
    
    ' Allow other events
    DoEvents

End Sub

Private Sub Search_SessionParsed(ByVal Index As Long)

    ' See if we don't have exceeded the maximum amount
    If Index + 1 > frmProgress.progressBar.Max Then
        Exit Sub
    End If

    ' Set the description
    frmProgress.lblDescription.Caption = "Counting words in session " & Index + 1 & " of " & frmProgress.progressBar.Max & " ..."

    ' Set the progressbar
    frmProgress.progressBar.Value = Index + 1

    ' Allow other events
    DoEvents

End Sub

Private Sub Search_WordCounted()

    ' Hide the progress-form
    frmProgress.Hide

End Sub

Private Sub Search_WordCounting(ByVal SessionAmount As Long)

    ' Show the progress-form
    frmProgress.Show
    
    ' Set the description
    frmProgress.lblDescription.Caption = "Counting words ..."

    ' Set the progressbar
    frmProgress.progressBar.Value = 0
    frmProgress.progressBar.Max = SessionAmount

End Sub

Private Sub lstWords_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' When right-clicking, ...
    If Button = 2 Then
    
        ' Save the control the procedure should call/print
        Set frmMain.MenuControl = lstWords
    
        ' ... invoke the associated popup menu
        Me.PopupMenu frmMain.mnuListView
    
    End If

End Sub
