VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLog"
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

' Is invoked when a file is to be read
Event FileRead(FileName As String, FileType As String, FileData As Variant, Ignore As Boolean)

' The index of this log
Public Index As Long

' All sessions in this log
Public Sessions As New Collection

' This class' parent
Public Parent As New clsContainer

' Used to determine if a similar log already has been processed
Public LogLenght As Long
Public LogName As String
Public LogIgnore As Boolean

Public Sub RemoveSession(Session As clsSession)

    Dim Tell As Long
    
    ' Go through all sessions in this log
    For Tell = 1 To Sessions.Count
    
        ' Remove the session if it's equal
        If Sessions(Tell) Is Session Then
                
            ' Remove it
            Sessions.Remove Tell
            
            ' We're finished
            Exit Sub
    
        End If

    Next

    ' Update session ID
    For Tell = 1 To Sessions.Count
        
        ' Set the session ID
        Sessions(Tell).SessionID = Tell
    
    Next

End Sub

Public Sub ClearLog()

    Dim Tell As Long
    
    For Tell = 1 To Sessions.Count
        
        ' Invoke the cleanup-procedure
        Sessions(Tell).ClearSession
        
        ' Remove the session from the above list
        ChatApp.RemoveFromCollection Parent.Sessions, Sessions(Tell)
    
    Next
    
    ' Remove own references
    Set Parent = Nothing
    Set Sessions = Nothing
    
End Sub

Public Function AddUser(NickNames As Variant, Email As String) As clsUser

    Dim User As clsUser, Tell As Long, NickName
    
    ' Create a new class
    Set User = New clsUser
    
    ' Add each element in the collection
    For Each NickName In NickNames
        User.AppendNickName CStr(NickName)
    Next
    
    ' Set the email
    User.Email = Email

    ' Add the user to the list
    Parent.Users.Add User
    
    ' Set the index of this user
    User.UserID = Parent.Users.Count

End Function

Public Function AddSession(SessionDate As Date, Participants As Variant) As clsSession

    Dim Session As New clsSession, Tell As Long, User As clsUser
    
    ' Set the parent
    Set Session.Parent = Me
    
    ' Set the properties of this session
    Session.SessionDate = SessionDate

    ' Make sure that this actually is an array
    If IsArray(Participants) Then

        ' Add all participants
        For Tell = LBound(Participants) To UBound(Participants)
    
            ' The information of each participant is split into an array containing the nickname and email address
            Session.AddParticipant CStr(Participants(Tell)(0)), CStr(Participants(Tell)(1))
    
        Next
    
    ElseIf IsObject(Participants) Then ' Make sure that this is a object
    
        ' Then it is most likely an collection of users
        For Each User In Participants
        
            ' Just add the user
            Session.Participants.Add User
        
        Next
    
    End If
    
    ' And finally add this session to the collection and parent collectin
    Parent.Sessions.Add Session
    Sessions.Add Session
    
    ' Set the session ID
    Session.SessionID = Sessions.Count
    
    ' Return the session added
    Set AddSession = Session

End Function

Public Sub ImportLog(sFile As String, LogType As String)

    Dim bIgnore As Boolean, sFileData As String, bFileData() As Byte, Free As Long
    
    ' Tell about this
    RaiseEvent FileRead(sFile, LogType, sFileData, bIgnore)

    ' Initialize the parser
    Set Parent.Parser.LogClass = Me
    Parent.Parser.ParseType = LogType

    ' Save the file name of this log
    LogName = sFile

    ' Whether or not the parser handles I/O-operation iself
    If Parent.Parser.HandleIO Then
    
        ' In this case, nothing more should be done
        If bIgnore Then
            Exit Sub
        End If
    
        ' Set the file to use
        Parent.Parser.LogData = sFile
        
    Else
    
        ' Read the file
        If Not bIgnore Then
        
            ' Get a free file
            Free = FreeFile
        
            ' This is done by using the internal I/O-functions
            Open sFile For Binary As #Free
            
                ' Get the lenght of the file
                LogLenght = LOF(Free)
                
                ' The data type we're going to deliver
                If Parent.Parser.DeliverString Then
                            
                    ' Allocate string
                    sFileData = Space(LogLenght)
                    
                    ' Get information
                    Get #Free, , sFileData
                
                    ' Set the data to use
                    Parent.Parser.LogData = sFileData
                    
                Else
                
                    ' Allocate byte-array
                    ReDim bFileData(1 To LogLenght)
                
                    ' Get the information to the array
                    Get #Free, , bFileData

                End If
            
            Close #Free
            
        Else
    
            ' Set the data to use
            Parent.Parser.LogData = sFileData
    
        End If
    
    End If
    
    ' Parse the log
    Parent.Parser.BeginParsing bFileData

    ' We are finished, clear the parser (by making a new, fresh one)
    Set Parent.Parser = New clsParser

End Sub

Public Function IsEqual(refLog As clsLog) As Boolean

    Dim Tell As Long

    ' Surely, the file lenght must be equal
    If LogLenght = refLog.LogLenght Then
    
        ' Check the amout of sessions
        If refLog.Sessions.Count = refLog.Sessions.Count Then
    
            ' Amazingly, this log is equal
            IsEqual = True
    
        End If
    
    End If

End Function

