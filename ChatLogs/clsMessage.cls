VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMessage"
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

' A message send to one user from another
Public Session As clsSession
Public MessageID As Long

' A pointer to the array
Private aElements() As Element

Public Property Get MessageDate() As Date
    ' Return the date of this message
    MessageDate = aElements(MessageID).Date
End Property

Public Property Let MessageDate(ByVal vNewValue As Date)
    ' Set the date of this message
    aElements(MessageID).Date = vNewValue
End Property

Public Property Get Style() As String
    ' Return the style of this message
    Text = aElements(MessageID).Style
End Property

Public Property Let Style(ByVal vNewValue As String)
    ' Save the style
    aElements(MessageID).Style = vNewValue
End Property

Public Property Get Text() As String
    ' Return the text of this message
    Text = aElements(MessageID).Text
End Property

Public Property Let Text(ByVal vNewValue As String)
    ' Save the text
    aElements(MessageID).Text = vNewValue
End Property

Public Property Get Sender() As clsUser
    
    Dim lngID As Long
    
    ' Retrive the ID of the sender
    lngID = aElements(MessageID).Sender

    ' See if the ID is valid
    If lngID > 0 Then
    
        ' Return the sender of this event by using the saved senderID
        Set Sender = Session.Parent.Parent.Users(lngID)
        
    End If
    
End Property

Public Property Let Sender(ByVal vNewValue As clsUser)

    ' Save the sender
    aElements(MessageID).Sender = vNewValue.UserID
    
End Property

Public Property Get Tag() As Variant
    
    ' Return the tag of this event safely
    If VarType(aElements(MessageID).Tag) = vbObject Then
        Set Tag = aElements(MessageID).Tag
    Else
        Tag = aElements(MessageID).Tag
    End If
    
End Property

Public Property Let Tag(ByVal vNewValue As Variant)

    ' Save the variant safely
    If VarType(vNewValue) = vbObject Then
        Set aElements(MessageID).Tag = vNewValue
    Else
        aElements(MessageID).Tag = vNewValue
    End If
    
End Property

Public Property Get Receiver() As Collection

    Dim Tell As Long, Users As Collection

    ' Firstly, create a new collection
    Set Receiver = New Collection
    
    ' Get the users-collection
    Set Users = Session.Parent.Parent.Users
    
    ' Then, add each user in the array
    For Tell = 1 To aElements(MessageID).ReceiverCount
    
        ' Add the user
        Receiver.Add Users(aElements(MessageID).Receiver(Tell))
    
    Next

End Property

Public Property Let Receiver(ByVal vNewValue As Collection)

    Dim Tell As Long, lngCount As Long
    
    ' Get the amount of users
    lngCount = vNewValue.Count
    
    ' Reset the count
    aElements(MessageID).ReceiverCount = lngCount
    
    ' And then the array
    ReDim aElements(MessageID).Receiver(lngCount)
    
    ' Whereafter we'll add all users
    For Tell = 1 To lngCount
    
        ' Set the element to the user-id of each user in the collection
        aElements(MessageID).Receiver(Tell) = vNewValue(Tell).UserID
    
    Next

End Property

Friend Sub InitializeClass(refElements() As Element, oSession As clsSession)

    ' Make a reference to the array filled with the elements
    CopyMemory ByVal VarPtrArray(aElements), ByVal VarPtrArray(refElements), 4

    ' Make a reference to the parent session
    Set Session = oSession

End Sub

Public Property Get IsMessage() As Boolean

    ' Yup, this is a message
    IsMessage = True

End Property

Private Sub Class_Terminate()

    ' Clear the pointer
    ZeroMemory ByVal VarPtrArray(aElements), 4

End Sub
