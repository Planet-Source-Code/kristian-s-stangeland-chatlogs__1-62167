VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

' The current index and session we'll read from
Public EventID As Long
Public Session As clsSession

' A pointer to the array
Private aElements() As Element

Public Property Get EventDate() As Date
    ' Return the date of this event
    EventDate = aElements(EventID).Date
End Property

Public Property Let EventDate(ByVal vNewValue As Date)
    ' Set the date of this event
    aElements(EventID).Date = vNewValue
End Property

Public Property Get EventType() As Long
    ' Return the type of this event
    EventType = aElements(EventID).Type
End Property

Public Property Let EventType(ByVal vNewValue As Long)
    ' Set the type of this event
    aElements(EventID).Type = vNewValue
End Property

Public Property Get Text() As String
    ' Return the text of this event
    Text = aElements(EventID).Text
End Property

Public Property Let Text(ByVal vNewValue As String)
    ' Save the text
    aElements(EventID).Text = vNewValue
End Property

Public Property Get Sender() As clsUser
    
    Dim lngID As Long
    
    ' Retrive the ID
    lngID = aElements(EventID).Sender
    
    ' See if the ID is valid
    If lngID > 0 Then
        ' Return the sender of this event by using the saved senderID
        Set Sender = Session.Parent.Parent.Users(aElements(EventID).Sender)
    End If
    
End Property

Public Property Let Sender(ByVal vNewValue As clsUser)
    ' Save the sender
    aElements(EventID).Sender = vNewValue.UserID
End Property

Public Property Get Tag() As Variant
    
    ' Return the tag of this event safely
    If VarType(aElements(EventID).Tag) = vbObject Then
        Set Tag = aElements(EventID).Tag
    Else
        Tag = aElements(EventID).Tag
    End If
    
End Property

Public Property Let Tag(ByVal vNewValue As Variant)

    ' Save the variant safely
    If VarType(vNewValue) = vbObject Then
        Set aElements(EventID).Tag = vNewValue
    Else
        aElements(EventID).Tag = vNewValue
    End If
    
End Property

Friend Sub InitializeClass(refElements() As Element, oSession As clsSession)

    ' Make a reference to the array filled with the elements
    CopyMemory ByVal VarPtrArray(aElements), ByVal VarPtrArray(refElements), 4

    ' Set the session-object
    Set Session = oSession

End Sub

Public Function IsEqual(refEvent As clsEvent) As Boolean

    ' See if the three variables correspond
    IsEqual = CBool((refEvent.Session.SessionDate = Session.SessionDate) And (refEvent.EventDate = EventDate) And (refEvent.Text = Text))

End Function

Public Property Get IsMessage() As Boolean

    ' Nope, this is NOT a message
    IsMessage = False

End Property

Private Sub Class_Terminate()

    ' Clear the pointer
    ZeroMemory ByVal VarPtrArray(aElements), 4

End Sub
