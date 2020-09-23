VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUser"
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

' Our type holding information regarding a certain nickname
Private Type NickName
    Counter As Long
    Value As String
End Type

' This class simply consist of information regarding a user
Public Sessions As New Collection
Public Email As String
Public Tag As String
Public UserID As Long

' Different useful statistics
Public AverageRespond As Date
Public InitiatedSessions As Long

' Internal variables used mainly to record all nicknames used
Private aNickNames() As NickName
Private lngNameCount As Long
Private lngMostUsed As Long

Public Sub AppendNickName(sNickName As String, Optional bAdded As Boolean)

    Dim Tell As Long

    ' See if the nickname is empty
    If LenB(sNickName) = 0 Then
        ' If so, there's no point of adding it
        Exit Sub
    End If

    ' First, see if this nickname already has been added
    For Tell = 0 To lngNameCount - 1

        ' See if they correspond in times of size
        If LenB(aNickNames(Tell).Value) = LenB(sNickName) Then
        
            ' Then see if they really correspond (this method is actually faster)
            If aNickNames(Tell).Value = sNickName Then
                
                ' This nickname does not need to be added, count it insted
                aNickNames(Tell).Counter = aNickNames(Tell).Counter + 1
                
                ' If this counter is bigger than the most used, use this insted
                If aNickNames(Tell).Counter > aNickNames(lngMostUsed).Counter Then
                    ' Change the index
                    lngMostUsed = Tell
                End If
                
                ' Exit the procedure
                Exit Sub
                
            End If
        
        End If

    Next
    
    ' Tell that the nickname has been added
    bAdded = True

    ' Firstly. reallocate the array completely
    ReDim Preserve aNickNames(lngNameCount)

    ' Then add the nickname
    aNickNames(lngNameCount).Value = sNickName
    aNickNames(lngNameCount).Counter = 1

    ' And finally, increase the counter
    lngNameCount = lngNameCount + 1

End Sub

Public Property Get NameAmount() As Long

    ' Return the amount of registered nicknames of this user
    NameAmount = lngNameCount - 1

End Property

Public Property Get NameValue(ByVal Index As Long) As String

    ' Return the value of a certain nickname
    NameValue = aNickNames(Index).Value

End Property

Public Property Let NameValue(ByVal Index As Long, ByVal vNewValue As String)

    ' Set the value of an element
    aNickNames(Index).Value = vNewValue

End Property

Public Property Get NameCount(ByVal Index As Long) As Long

    ' Retrive the amount this nickname has been used (that is, registered)
    NameCount = aNickNames(Index).Counter

End Property

Public Property Get NickName() As String

    On Local Error Resume Next

    ' Return the latest nickname
    NickName = aNickNames(lngMostUsed).Value

End Property

Public Function IsEqual(refUser As clsUser) As Boolean

    ' See if the three variables correspond
    IsEqual = CBool((UCase(Mid(refUser.Email, 1, Len(Email))) = UCase(Email) And Not (Email = "")) Or ((Mid(refUser.NickName, 1, Len(NickName)) = NickName) And NickName <> ""))

End Function

Public Sub CalculateValues()

    ' Call both functions and calculate them
    CalculateInitiatedSessions
    CalculateAverageRespond
    
End Sub

Private Sub CalculateInitiatedSessions()

    Dim Session As clsSession, oEvent As Object

    ' Clear the value
    InitiatedSessions = 0

    ' Go through all sessions of this user
    For Each Session In Sessions

        ' Check the first user and se if the sender is equal to the current user
        For Each oEvent In Session.Events
        
            ' The event must of course be a message
            If TypeOf oEvent Is clsMessage Then
            
                ' If the sender is the current user, increase the count
                If oEvent.Sender Is Me Then
                    InitiatedSessions = InitiatedSessions + 1
                End If
            
                ' The first message has been checked
                Exit For
            
            End If
        
        Next

    Next

End Sub

Private Sub CalculateAverageRespond()

    Dim Session As clsSession, oEvent As Object, oCompare As Object, Tell As Long
    Dim dTotal As Date, lngMessages As Long, lngIndex As Long

    ' Clear the value
    AverageRespond = 0

    ' Go through all sessions with this user
    For Each Session In Sessions
    
        ' Only include sessions with two participants
        If Session.Participants.Count = 2 Then
    
            ' Go through all events in each session, minus the last one
            For Tell = 2 To Session.Events.Count
            
                ' Get the event/message
                Set oEvent = Session.Events(Tell)
            
                ' The event must of course be a message
                If TypeOf oEvent Is clsMessage Then
                
                    ' Now, if the sender is equal to this user
                    If oEvent.Sender Is Me Then
                    
                        ' Move backwards searching for a message not from our user
                        For lngIndex = Tell - 1 To 1 Step -1
                        
                            ' Get the object
                            Set oCompare = Session.Events(lngIndex)
                         
                            ' See if this is a message
                            If TypeOf oCompare Is clsMessage Then
                            
                                ' ... see if the message before was not made by this user
                                If Not (oCompare.Sender Is Me) Then
                    
                                    ' Calulate the difference in time and append it to the total time
                                    dTotal = dTotal + CDate(oEvent.MessageDate - oCompare.MessageDate)
                        
                                    ' The number of messages to take the avarage of
                                    lngMessages = lngMessages + 1
                                
                                End If
                            
                                ' We're finish searching anyway
                                Exit For
                            
                            End If
                        
                        Next
                    
                    End If
        
                End If
                
            Next
        
        End If
    
    Next
    
    ' Ignore errors in the following calculation
    On Local Error Resume Next
    
    ' Return the avarage time
    AverageRespond = CDate(dTotal / lngMessages)

End Sub

