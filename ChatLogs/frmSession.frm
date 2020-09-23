VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSession 
   Caption         =   "Session"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   366
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   509
   Begin RichTextLib.RichTextBox txtConversation 
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4683
      _Version        =   393217
      BorderStyle     =   0
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmSession.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstParticipants 
      Height          =   1230
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   7335
   End
   Begin VB.Label lblDate 
      Caption         =   "26.07.2004"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblConversation 
      Caption         =   "&Conversation:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label lblParticipants 
      Caption         =   "&Participants:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblSessionDate 
      Caption         =   "&Session date:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmSession"
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

Public Session As clsSession

Public Function Clone() As frmSession

    ' Create a new form
    Set Clone = New frmSession
    
    ' Initialize it ...
    Clone.Update Session
    
    ' and finally show it
    Clone.Show

End Function

Public Sub Update(newSession As clsSession, Optional HighlighMessage As clsMessage)

    ' Firstly, save this session
    Set Session = newSession
    
    ' Set the date
    lblDate.Caption = Session.SessionDate
    
    ' And update everything else
    AddParticipants
    AddEvents HighlighMessage
    
End Sub

Public Sub AddParticipants()
    
    Dim User As clsUser, Tell As Long
    
    ' Clear the list
    lstParticipants.Clear
    
    ' Add all participants
    For Tell = 1 To Session.Participants.Count

        ' Get the user
        Set User = Session.Participants(Tell)

        ' Combine the nickname and email
        lstParticipants.AddItem User.NickName & " (" & User.Email & ")"
        lstParticipants.ItemData(lstParticipants.ListCount - 1) = Tell

    Next

End Sub

Public Sub AddEvents(HighlightMessage As clsMessage)

    On Error Resume Next
    Dim sText As String, oEvent As Object, NotFirstLine As Boolean, sNickName As String, lngPos As Long, lngLenght As Long

    ' Default value
    lngPos = -1

    ' Clear the conversation-window
    txtConversation.Text = ""
    
    ' Hide it to imporve performance
    txtConversation.Visible = False
    
    ' Add all events
    For Each oEvent In Session.Events
    
        ' Check for the event-type
        If TypeOf oEvent Is clsMessage Then
        
            ' Try to get the sender (this is not always possible due to long nicknames)
            sNickName = oEvent.Sender.NickName
        
            ' Generate the text to add
            sText = IIf(NotFirstLine, vbCrLf, "") & "(" & RemoveDate(oEvent.MessageDate) & ") " & sNickName & ": "
        
            ' Add the header
            txtConversation.SelLength = 1
            txtConversation.SelColor = IIf(oEvent.Sender Is Session.Participants(1), &HC0&, &HC00000)
            txtConversation.SelText = sText
            
            ' Then add the message
            txtConversation.SelLength = 1
            txtConversation.SelColor = vbBlack
            txtConversation.SelText = IndentLines(oEvent.Text)

            ' If this is the message to highlight ...
            If oEvent Is HighlightMessage Then
            
                ' Remember the position and lenght
                lngLenght = Len(sText)
                lngPos = txtConversation.SelStart - lngLenght
            
            End If
        
            ' This is not the first line any more
            NotFirstLine = True
        
        ElseIf TypeOf oEvent Is clsEvent Then
        
            ' Add the special event
            txtConversation.SelLength = 1
            txtConversation.SelBold = True
            txtConversation.SelText = IIf(NotFirstLine, vbCrLf, "") & "(" & RemoveDate(oEvent.EventDate) & ") " & oEvent.Text
            txtConversation.SelBold = False
        
        End If
    
    Next
    
    ' Show the window
    txtConversation.Visible = True
    txtConversation.SetFocus
    
    ' Highlight a message if found
    If lngPos >= 0 Then
    
        ' Set the new selection
        txtConversation.SelStart = lngPos
        txtConversation.SelLength = lngLenght
        
    End If

End Sub

Private Function IndentLines(sText As String) As String

    Dim Lines As Variant, Tell As Long

    ' Get all lines from the string
    Lines = Split(sText, vbCrLf)
    
    ' Begin indenting lines
    For Tell = LBound(Lines) + 1 To UBound(Lines)
    
        ' Indent line
        Lines(Tell) = Space(11) & Lines(Tell)
    
    Next

    ' Return the modified string
    IndentLines = Join(Lines, vbCrLf)

End Function

Private Function RemoveDate(dDate As Date) As Date

    ' Return only the time
    RemoveDate = TimeSerial(Hour(dDate), Minute(dDate), Second(dDate))

End Function

Private Sub Form_Load()

    ' Initialize this form
    ChatApp.InitializeForm Me

End Sub

Private Sub Form_Resize()

    ' Only resize when we can
    If Me.WindowState <> 1 Then

        ' Only three values should be resized
        lstParticipants.Width = Me.ScaleWidth - 16
        txtConversation.Width = lstParticipants.Width
        txtConversation.Height = Me.ScaleHeight - txtConversation.Top - 8
        txtConversation.RightMargin = txtConversation.Width + 10000

    End If

End Sub

Private Sub lstParticipants_DblClick()

    Dim Dialog As frmSessionList, User As clsUser

    ' Only go further if there's really an element selected
    If lstParticipants.ListCount > 0 Then
    
        ' Create a new form
        Set Dialog = New frmSessionList
        
        ' Get the user
        Set User = Session.Participants(lstParticipants.ItemData(lstParticipants.ListIndex))
        
        ' Update it for this session
        Dialog.Caption = LanguageConst("SessionsWith") & User.NickName
        Dialog.Update User.Sessions
    
    End If

End Sub
