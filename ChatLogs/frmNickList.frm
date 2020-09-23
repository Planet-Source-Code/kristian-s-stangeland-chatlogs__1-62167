VERSION 5.00
Object = "{E910F8E1-8996-4EE9-90F1-3E7C64FA9829}#1.1#0"; "vbaListView6.ocx"
Begin VB.Form frmNickList 
   Caption         =   "Nick List"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   468
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   566
   Begin vbalListViewLib6.vbalListViewCtl lstNicks 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   11880
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
Attribute VB_Name = "frmNickList"
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

Public Function Clone() As frmNickList

    ' Create a new form
    Set Clone = New frmNickList
    
    ' Initialize it ...
    Clone.Update
    
    ' and finally show it
    Clone.Show

End Function

Private Sub Form_Load()

    ' Initialize this form
    ChatApp.InitializeForm Me

End Sub

Public Sub Update()

    Dim Users As Collection, ListItem As cListItem, User As clsUser, Tell As Long, lngNick As Long

    ' Hide the listbox to optimze the speed
    lstNicks.Visible = False

    ' Firstly, clear the list
    lstNicks.ListItems.Clear
    
    ' Grab the user-list from the parent
    Set Users = frmMain.Logs.Users
    
    ' Then, add new elements form the list
    For Tell = 1 To Users.Count

        ' Retrive the current user
        Set User = Users(Tell)
    
        ' Go through all registered nicknames by this user
        For lngNick = 0 To User.NameAmount

            ' Add an listitem with the nickname
            Set ListItem = lstNicks.ListItems.Add(, , CStr(User.NameValue(lngNick)))
        
            ' ... and email, together with the reference
            ListItem.SubItems(1).Caption = User.Email
            ListItem.SubItems(2).Caption = User.NameCount(lngNick)
            
            ' Remember this item's reference
            ListItem.Tag = Tell
        
        Next
    
    Next

    ' Show the listbox again
    lstNicks.Visible = True

End Sub

Private Sub lstNicks_ItemDblClick(Item As vbalListViewLib6.cListItem)

    Dim Dialog As frmSessionList, User As clsUser

    ' Create a new form
    Set Dialog = New frmSessionList
    
    ' Get the user
    Set User = frmMain.Logs.Users(Val(Item.Tag))
    
    ' Update it for this session
    Dialog.Caption = LanguageConst("SessionsWith") & User.NickName
    Dialog.Update User.Sessions

End Sub

Private Sub lstNicks_ColumnClick(Column As cColumn)

    ' "Invert" the sort-order
    Column.SortOrder = NewSortOrder(Column.SortOrder)

    ' Sort according to the column type:
    Select Case Column.Key
    Case "NickName", "Email" ' Sort by text
        Column.SortType = eLVSortString

    Case "Used"
        Column.SortType = eLVSortNumeric
    
    End Select
    
    ' Sort everything
    lstNicks.ListItems.SortItems
   
End Sub

Private Sub Form_Resize()

    ' Only resize when we can
    If Me.WindowState <> 1 Then
    
        ' Resize the list-control
        lstNicks.Width = Me.ScaleWidth - 16
        lstNicks.Height = Me.ScaleHeight - 16
    
    End If
    
End Sub

Private Sub lstNicks_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' When right-clicking, ...
    If Button = 2 Then
    
        ' Save the control the procedure should call/print
        Set frmMain.MenuControl = lstNicks
    
        ' ... invoke the associated popup menu
        Me.PopupMenu frmMain.mnuListView
    
    End If

End Sub
