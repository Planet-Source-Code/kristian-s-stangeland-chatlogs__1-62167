VERSION 5.00
Object = "{E910F8E1-8996-4EE9-90F1-3E7C64FA9829}#1.1#0"; "vbaListView6.ocx"
Begin VB.Form frmSessionList 
   Caption         =   "Session list"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   385
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   534
   Begin vbalListViewLib6.vbalListViewCtl lstSessions 
      Height          =   5055
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8916
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
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   521
      TabIndex        =   0
      Top             =   5280
      Width           =   7815
      Begin VB.CommandButton cmdTimeChart 
         Caption         =   "&Time Chart"
         Height          =   255
         Left            =   6600
         TabIndex        =   7
         Top             =   15
         Width           =   1095
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "&Filter"
         Height          =   255
         Left            =   5520
         TabIndex        =   5
         Top             =   15
         Width           =   1095
      End
      Begin VB.TextBox txtToDate 
         Height          =   285
         Left            =   3600
         TabIndex        =   4
         Top             =   0
         Width           =   1815
      End
      Begin VB.TextBox txtFromDate 
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label lblDateTo 
         Caption         =   "Until: "
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   30
         Width           =   735
      End
      Begin VB.Label lblDateFrom 
         Caption         =   "From:"
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   30
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmSessionList"
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

Public currSessions As Collection

Public Function Clone() As frmSessionList

    ' Create a new form
    Set Clone = New frmSessionList
    
    ' Initialize it ...
    Clone.Update currSessions
    
    ' and finally show it
    Clone.Show

End Function

Private Sub cmdTimeChart_Click()

    Dim newTimeChart As New frmTimeChart

    ' Use all current sessions for the time chart
    Set newTimeChart.Sessions = currSessions
    
    ' Show the time chart
    newTimeChart.Show

End Sub

Private Sub Form_Load()
    
    ' Initialize this form
    ChatApp.InitializeForm Me
    
End Sub

Private Sub Form_Resize()
    
    ' Resize the list-control
    lstSessions.Width = Me.ScaleWidth - 16
    lstSessions.Height = Me.ScaleHeight - 24 - picFilter.Height
    
    ' Then set the new position of the filter-box and resize it
    picFilter.Top = lstSessions.Top + lstSessions.Height + 8
    picFilter.Width = lstSessions.Width
    
    ' Whereafter we need to move different controls around
    cmdFilter.Left = picFilter.ScaleWidth - cmdFilter.Width - cmdTimeChart.Width
    cmdTimeChart.Left = cmdFilter.Left + cmdFilter.Width
    txtFromDate.Width = (cmdFilter.Left - (lblDateTo.Width * 2) - 32) / 2
    txtToDate.Width = txtFromDate.Width
    lblDateTo.Left = ((cmdFilter.Left - txtFromDate.Left - txtFromDate.Width) / 2) - ((lblDateTo.Width + txtToDate.Width) / 2) + txtFromDate.Left + txtFromDate.Width
    txtToDate.Left = lblDateTo.Left + lblDateTo.Width
    
End Sub

Private Sub cmdFilter_Click()

    ' Here we'll filter the sessions
    Update currSessions, CDate(txtFromDate), CDate(txtToDate), True

End Sub

Public Sub Update(Sessions As Collection, Optional StartDate As Variant, Optional EndDate As Variant, Optional FilterDate As Boolean)

    On Error Resume Next
    Dim Log As clsLog, Session As clsSession, ListItem As cListItem, bIgnore As Boolean
    
    ' Save the sessions
    Set currSessions = Sessions
   
    ' Firstly, clear the list
    lstSessions.ListItems.Clear

    ' Then, add all sessions in the collection
    For Each Session In Sessions
    
        ' Now, should we filter by date?
        If FilterDate Then
        
            ' If so, don't go further if the date is outside
            bIgnore = CBool(Session.SessionDate < StartDate Or Session.SessionDate > EndDate)
        
        Else
    
            ' Use this session-date as default
            If IsMissing(StartDate) Then
                StartDate = Session.SessionDate
            End If
            
            ' As the above
            If IsMissing(EndDate) Then
                EndDate = Session.SessionDate
            End If
    
            ' Now, if this session is BEFORE the start-date, use that date insted
            If Session.SessionDate < StartDate Then
                StartDate = Session.SessionDate
            End If
            
            ' As the above, only here witht the end-date
            If Session.SessionDate > EndDate Then
                EndDate = Session.SessionDate
            End If
    
        End If
        
        ' Add if we shouldn't ignore the session
        If (Not bIgnore) And ((Session.SessionDate >= StartDate And Session.SessionDate <= EndDate) Or (Not FilterDate)) Then
    
            ' Firstly, add a new element
            Set ListItem = lstSessions.ListItems.Add(, , Session.SessionDate)
            
            ' Then set its subelements
            ListItem.SubItems(1).Caption = ChatApp.JoinParticipants(Session.Participants, ", ")
            ListItem.SubItems(2).Caption = Session.Events.Count
            
            ' Remember the original class
            ListItem.Tag = Session.Parent.Index & "." & Session.SessionID
    
        End If
    
    Next

    If Not FilterDate Then
    
        ' Just set the dates to the textboxes
        txtFromDate.Text = StartDate
        txtToDate.Text = EndDate
    
    End If

End Sub

Private Sub lstSessions_ColumnClick(Column As cColumn)

    ' "Invert" the sort-order
    Column.SortOrder = NewSortOrder(Column.SortOrder)

    ' Sort according to the column type:
    Select Case Column.Key
    Case "MessageDate" ' Sort by date
        Column.SortType = eLVSortDate
       
    Case "Participants" ' Sort by text
        Column.SortType = eLVSortString
        
    Case "Events" ' Sort by number
        Column.SortType = eLVSortNumeric

    End Select
    
    ' Sort everything
    lstSessions.ListItems.SortItems
   
End Sub

Private Sub lstSessions_ItemDblClick(Item As vbalListViewLib6.cListItem)

    Dim Dialog As frmSession

    ' Create a new form
    Set Dialog = New frmSession
    
    ' Update it for this session
    Dialog.Update RetriveSession(Item.Tag)

End Sub

Private Function RetriveSession(sIndex As String) As clsSession

    Dim vIndexes As Variant
    
    ' Split the string by the delimiter so that we can retrive the session
    vIndexes = Split(sIndex, ".")
    
    ' Return the session
    Set RetriveSession = frmMain.Logs.Logs(Val(vIndexes(0))).Sessions(Val(vIndexes(1)))

End Function

Private Sub lstSessions_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' When right-clicking, ...
    If Button = 2 Then
    
        ' Save the control the procedure should call/print
        Set frmMain.MenuControl = lstSessions
    
        ' ... invoke the associated popup menu
        Me.PopupMenu frmMain.mnuListView
    
    End If

End Sub
