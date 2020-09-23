VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSession"
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

' Different information about a conversation-session
Public SessionDate As Date
Public SessionID As Long
Public Participants As New Collection
Public NickNames As New Collection
Public Events As clsElements
Public Parent As clsLog
Public Tag As String

' The array holding all events and messages
Private aElements() As Element
Private ElementCount As Long
Private ArraySize As Long

Public Sub ClearSession()

    Dim User As clsUser, Tell As Long

    ' Remove the reference this user has to the parent object - the session
    For Each User In Participants
        
        ' Remove the item
        ChatApp.RemoveFromCollection User.Sessions, Me
        
        ' If no sessions is refered, remove the user from the list completely
        If User.Sessions.Count = 0 Then
            
            ' Remove the user from the master list
            ChatApp.RemoveFromCollection Parent.Parent.Users, User
        
        End If
        
    Next
    
    ' Now, erase the array
    Erase aElements
    
    ' Then erase the count
    ElementCount = 0

    ' Remove all references
    Set Events = Nothing
    Set Participants = Nothing
    Set Parent = Nothing

End Sub

Public Function FindUser(SearchString As String) As clsUser

    Dim User As clsUser
    
    ' Go through all users
    For Each User In Participants
    
        ' See if the search string is in the email or nickname
        If Correspond(SearchString, User.Email) Or Correspond(SearchString, User.NickName) Then
            
            ' Return this user
            Set FindUser = User
        
            ' We are finished
            Exit Function
            
        End If
    
    Next

End Function

Public Function AddParticipant(NickName As String, Email As String, Optional lngUserID As Long = -1) As clsUser

    Dim User As New clsUser, Tell As Long
        
    ' See if we're instructed to just add a user
    If lngUserID >= 0 Then
    
        ' Use this class insted
        Set User = Parent.Parent.Users(lngUserID)
            
        ' Add the nickname
        User.AppendNickName NickName
    
    Else
        
        ' Add the nickname, if specified
        If LenB(NickName) > 0 Then
            User.AppendNickName NickName
        End If
        
        ' Set the email adress
        User.Email = Email
        
        ' Check weather or not the user already has been added
        Tell = Parent.Parent.FindUser(User)
        
        ' If so, ...
        If Tell >= 0 Then
        
            ' Use this class insted
            Set User = Parent.Parent.Users(Tell)
            
            ' Add the nickname
            User.AppendNickName NickName
        
        Else
        
            ' Otherwise, we need to add it to the list
            Parent.Parent.Users.Add User
            
            ' And then save the index
            User.UserID = Parent.Parent.Users.Count
        
        End If
    
    End If
        
    ' Return the user
    Set AddParticipant = User
    
    ' And finally, add it to the lists
    Participants.Add User
    
    ' Only add the nickname if it's specified
    If LenB(NickName) > 0 Then
        NickNames.Add NickName
    End If
    
End Function

Public Sub UpdateUsers()

    Dim User As clsUser

    ' Loop through each user
    For Each User In Participants
    
        ' Then add which session this user belongs to
        User.Sessions.Add Me
    
    Next

End Sub

Public Function AddMessage(MessageDate As Date, Sender As clsUser, Receiver As Collection, Text As String, Optional Style As String = "", Optional Tag As Variant = "")

    ' Increase the element-count
    ElementCount = ElementCount + 1
    
    ' Then reallocate the array if it's necessary
    If ElementCount > ArraySize Then
        
        ' Increase the array szie
        ArraySize = ArraySize * 4
        
        ' Reallocate the array
        ReDim Preserve aElements(1 To ArraySize)
    
    End If
    
    ' We need to refresh the event-class
    Events.InitializeClass aElements, Me, ElementCount
    
    ' Set the properties of this element
    aElements(ElementCount).IsMessage = True
    aElements(ElementCount).Date = MessageDate
    aElements(ElementCount).Style = Style
    aElements(ElementCount).Text = Text
    
    ' Set the tag-variable (safely)
    If VarType(Tag) = vbObject Then
        Set aElements(ElementCount).Tag = Tag
    Else
        aElements(ElementCount).Tag = Tag
    End If
    
    ' Set the sender, if it exists
    If Not (Sender Is Nothing) Then
        aElements(ElementCount).Sender = Sender.UserID
    End If
    
    ' Return this object
    Set AddMessage = Events(ElementCount)
    
    ' Then set the recievers
    AddMessage.Receiver = Receiver
    
End Function

Public Function AddEvent(EventDate As Date, Text As String, Optional EventType As Long, Optional Sender As Object)

    ' Firstly, increase the element-count
    ElementCount = ElementCount + 1
    
    ' Then reallocate the array if it's necessary
    If ElementCount > ArraySize Then
        
        ' Increase the array szie
        ArraySize = ArraySize * 3
        
        ' Reallocate the array
        ReDim Preserve aElements(1 To ArraySize)
    
    End If
    
    ' We need to refresh the event-class
    Events.InitializeClass aElements, Me, ElementCount
    
    ' Set each property of this element
    aElements(ElementCount).IsMessage = False
    aElements(ElementCount).Date = EventDate
    aElements(ElementCount).Type = EventType
    aElements(ElementCount).Text = Text
    
    ' Set the sender, if it exists
    If Not (Sender Is Nothing) Then
        aElements(ElementCount).Sender = Sender.UserID
    End If
    
    ' Return this object
    Set AddEvent = Events(ElementCount)
    
End Function

Public Function IsEqual(refSession As clsSession) As Boolean

    Dim Tell As Long

    ' Firstly, check both the amout of events and participants
    If Participants.Count = refSession.Participants.Count And ElementCount = refSession.Events.Count Then
    
        ' See if all participants are the same
        For Tell = 1 To Participants.Count
        
            ' See if they are not equal
            If Not Participants(Tell) Is refSession.Participants(Tell) Then
            
                  ' Nope, the session is NOT equal
                  Exit Function
            
            End If
        
        Next
        
        
        ' Then see if all events are the same
        For Tell = 1 To ElementCount
        
            ' Firstly, see if the type corresponds
            If aElements(Tell).IsMessage = refSession.Events(Tell).IsMessage Then
                
                ' Then check the type of the event
                If aElements(Tell).IsMessage Then
                    ' If the date isn't alike, ...
                    If aElements(Tell).Date <> refSession.Events(Tell).MessageDate Then
                        ' ... the session is neither alike
                        Exit Function
                    End If
                Else
                    ' Same as above, if it dosen't correspond, ...
                    If aElements(Tell).Date <> refSession.Events(Tell).EventDate Then
                        ' ... the session is not alike
                        Exit Function
                    End If
                End If
                
            Else
                ' Nope, the session is not alike
                Exit Function
            End If
        
            ' See if they are not equal
            If Not (aElements(Tell).Text = refSession.Events(Tell).Text) Then
            
                  ' Nope, the session is NOT equal
                  Exit Function
            
            End If
        
        Next
        
        ' The session is in fact totaly alike
        IsEqual = True
    
    End If

End Function

Private Function Correspond(sFirst As String, sSecond As String) As Boolean

    ' Return whether or not they correspoond
    If sSecond <> "" Then
        Correspond = (InStr(1, sFirst, sSecond) <> 0) Or (InStr(1, sSecond, sFirst) <> 0)
    End If

End Function

Private Sub Class_Initialize()
    
    ' Set the REAL size of the array
    ArraySize = 10
    
    ' Allocate elements
    ReDim aElements(1 To ArraySize)

    ' Initialize the events-class
    Set Events = New clsElements
    
    ' Call the initialize-procedure
    Events.InitializeClass aElements, Me, ElementCount

End Sub
