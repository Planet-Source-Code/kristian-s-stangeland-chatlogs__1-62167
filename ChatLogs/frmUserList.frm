VERSION 5.00
Object = "{E910F8E1-8996-4EE9-90F1-3E7C64FA9829}#1.1#0"; "vbaListView6.ocx"
Begin VB.Form frmUserList 
   Caption         =   "User list"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   402
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   427
   Begin vbalListViewLib6.vbalListViewCtl lstUsers 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   10186
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
End
Attribute VB_Name = "frmUserList"
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

Public Function Clone() As frmUserList

    ' Create a new form
    Set Clone = New frmUserList
    
    ' Initialize it ...
    Clone.Update
    
    ' and finally show it
    Clone.Show

End Function

Private Sub Form_Load()

    ' Initialize this form
    ChatApp.InitializeForm Me

End Sub

Private Sub lstUsers_ColumnClick(Column As cColumn)

    ' "Invert" the sort-order
    Column.SortOrder = NewSortOrder(Column.SortOrder)

    ' Sort according to the column type:
    Select Case Column.Key
    Case "NickName", "Email" ' Sort by text
        Column.SortType = eLVSortString
        
    Case "Sessions", "Initiator" ' Sort by number
        Column.SortType = eLVSortNumeric
        
    Case "ART" ' Sort by date/time
        Column.SortType = eLVSortDate

    End Select
    
    ' Sort everything
    lstUsers.ListItems.SortItems
   
End Sub

Private Sub Form_Resize()

    ' Only resize when we can
    If Me.WindowState <> 1 Then
    
        ' Resize the list-control
        lstUsers.Width = Me.ScaleWidth - 16
        lstUsers.Height = Me.ScaleHeight - 16
    
    End If
    
End Sub

Public Sub Update()

    Dim Users As Collection, ListItem As cListItem, Tell As Long

    ' Hide the listbox to optimze the speed
    lstUsers.Visible = False

    ' Firstly, clear the list
    lstUsers.ListItems.Clear
    
    ' Grab the user-list from the parent
    Set Users = frmMain.Logs.Users
    
    ' Then, add new elements form the list
    For Tell = 1 To Users.Count
    
        ' Add the user's nickname
        Set ListItem = lstUsers.ListItems.Add(, , Users(Tell).NickName)
    
        ' ... and email, together with the reference
        ListItem.SubItems(1).Caption = Users(Tell).Email
        ListItem.SubItems(2).Caption = Users(Tell).Sessions.Count
        ListItem.SubItems(3).Caption = Users(Tell).InitiatedSessions
        ListItem.SubItems(4).Caption = Users(Tell).AverageRespond
        
        ' Remember this item's reference
        ListItem.Tag = Tell
    
    Next

    ' Show the listbox again
    lstUsers.Visible = True

End Sub

Private Sub lstUsers_ItemDblClick(Item As vbalListViewLib6.cListItem)

    Dim Dialog As frmSessionList, User As clsUser

    ' Create a new form
    Set Dialog = New frmSessionList
    
    ' Get the user
    Set User = frmMain.Logs.Users(Val(Item.Tag))
    
    ' Update it for this session
    Dialog.Caption = LanguageConst("SessionsWith") & User.NickName
    Dialog.Update User.Sessions

End Sub

Private Sub lstUsers_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' When right-clicking, ...
    If Button = 2 Then
    
        ' Save the control the procedure should call/print
        Set frmMain.MenuControl = lstUsers
    
        ' ... invoke the associated popup menu
        Me.PopupMenu frmMain.mnuListView
    
    End If

End Sub
