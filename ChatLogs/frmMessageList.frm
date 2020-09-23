VERSION 5.00
Object = "{E910F8E1-8996-4EE9-90F1-3E7C64FA9829}#1.1#0"; "vbaListView6.ocx"
Begin VB.Form frmMessageList 
   Caption         =   "Message list"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   484
   Begin vbalListViewLib6.vbalListViewCtl lstMessages 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   10821
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
Attribute VB_Name = "frmMessageList"
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

Private currMessages As Collection

Public Function Clone() As frmMessageList

    ' Create a new form
    Set Clone = New frmMessageList
    
    ' Initialize it ...
    Clone.Update currMessages
    
    ' and finally show it
    Clone.Show

End Function

Private Sub Form_Load()

    ' Initialize this form
    ChatApp.InitializeForm Me
    
End Sub

Private Sub Form_Resize()
    
    ' Only resize when we can
    If Me.WindowState <> 1 Then
    
        ' Resize the list-control
        lstMessages.Width = Me.ScaleWidth - 16
        lstMessages.Height = Me.ScaleHeight - 16
    
    End If
    
End Sub

Public Sub Update(Messages As Collection)

    On Error Resume Next
    Dim Log As clsLog, Message As Object, ListItem As cListItem

    ' Firstly, clear the list
    lstMessages.ListItems.Clear

    ' Then, add all messages in the collection
    For Each Message In Messages
    
        ' Only add if this is REALLY a message, not a event
        If TypeOf Message Is clsMessage Then
    
            ' Firstly, add a new element
            Set ListItem = lstMessages.ListItems.Add(, , Message.MessageDate)
            
            ' Then set its subelements
            ListItem.SubItems(1).Caption = Message.Sender.NickName & " (" & Message.Sender.Email & ")"
            ListItem.SubItems(2).Caption = ChatApp.JoinParticipants(Message.Receiver, ", ")
            ListItem.SubItems(3).Caption = Message.Text
            
            ' Remember the original class
            ListItem.Tag = Message.Session.Parent.Index & "." & Message.Session.SessionID & "." & Message.MessageID
        
        End If
    
    Next

    ' Save the messages used
    Set currMessages = Messages

End Sub

Private Sub lstMessages_ColumnClick(Column As cColumn)

    ' "Invert" the sort order
    Column.SortOrder = NewSortOrder(Column.SortOrder)

    ' Sort according to the column type:
    Select Case Column.Key
    Case "MessageDate" ' Sort by date
        Column.SortType = eLVSortDate
       
    Case "Sender", "Receiver", "Message" ' Sort by text
        Column.SortType = eLVSortString

    End Select
    
    ' Sort everything
    lstMessages.ListItems.SortItems
   
End Sub

Private Sub RetriveClasses(sIndex As String, Session As clsSession, Message As clsMessage)

    Dim vIndexes As Variant
    
    ' Split the string by the delimiter so that we can retrive the session
    vIndexes = Split(sIndex, ".")
    
    ' Then get the session and the message
    Set Session = frmMain.Logs.Logs(Val(vIndexes(0))).Sessions(Val(vIndexes(1)))
    Set Message = Session.Events(Val(vIndexes(2)))

End Sub

Private Sub lstMessages_ItemDblClick(Item As vbalListViewLib6.cListItem)
    
    Dim Dialog As frmSession, Session As clsSession, Message As clsMessage
    
    ' Create a new form
    Set Dialog = New frmSession
    
    ' Get the classes
    RetriveClasses CStr(Item.Tag), Session, Message
    
    ' Update it for this session
    Dialog.Update Session, Message

End Sub

Private Sub lstMessages_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' When right-clicking, ...
    If Button = 2 Then
    
        ' Save the control the procedure should call/print
        Set frmMain.MenuControl = lstMessages
    
        ' ... invoke the associated popup menu
        Me.PopupMenu frmMain.mnuListView
    
    End If

End Sub
