VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInput"
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

Public MainObject As Object

Property Get FileFilter() As String

    ' The filter used to browse for such files
    FileFilter = "Access Database (*.mdb)" & Chr(0) & "*.mdb"
    
End Property

Property Get FileExtension() As String

    ' The file extension this parser uses
    FileExtension = "mdb"
    
End Property

Property Get ParserHandleIO() As Boolean

    ' Whether or not the parser handles I/O iself
    ParserHandleIO = True

End Property

Property Get ParserName() As String

    ' This parser is primary used to import Access Databases
    ParserName = "Access Database"

End Property

Property Get ParserDeliverString() As Boolean

    ' Whether or not we're awaiting a string or a byte array
    ParserDeliverString = True

End Property

Property Get ParserDescription() As String

    ' The description of this parser
    ParserDescription = "Used to import Access Databases"

End Property

Private Function GetUsers(aUsers As String, Users() As Long, Logs As Object) As Collection

    Dim aArray, Tell As Long
    
    ' Create a new collection
    Set GetUsers = New Collection
    
    ' Split all users
    aArray = Split(aUsers, Chr(13))

    ' Loop through the resulting array ...
    For Tell = LBound(aArray) To UBound(aArray)
    
        ' ... and add each user
        GetUsers.Add Logs.Users(Users(aArray(Tell)))
    
    Next

End Function
 
Private Function GetDateOnly(dDate As Date) As Date

    ' Return the date-only
    GetDateOnly = DateSerial(Year(dDate), Month(dDate), Day(dDate))

End Function

Public Sub ParserInitialize()

    ' We are about to begin

End Sub

Public Sub ParseDocument(FileName As String, LogClass As Object)

    ' The database-objects
    Dim Database As Database
    Dim tblSessions As Recordset
    Dim tblUsers As Recordset
    Dim tblMessages As Recordset
    
    Dim Tell As Long, Logs As Object, aNickNames, strEmail As String
    Dim Session As Object, Sender As Object, Receiver As Object
    Dim Users() As Long, Sessions() As Long
    
    ' Get the container
    Set Logs = LogClass.Parent

    ' Now, see first if the file dosen't exist
    If Dir(FileName) = "" Then
    
        ' If so, we cannot do much more
        Exit Sub

    End If
    
    ' If it does exist, open it
    Set Database = OpenDatabase(FileName)

    ' Then open all the tables
    Set tblSessions = Database.OpenRecordset("Sessions")
    Set tblUsers = Database.OpenRecordset("Users")
    Set tblMessages = Database.OpenRecordset("Messages")
    
    ' Reallocate arrays
    ReDim Users(1 To CountRecords(Database, "Users"))
    ReDim Sessions(1 To CountRecords(Database, "Sessions"))
    
    ' Move to the beginning
    tblUsers.MoveFirst
    
    ' Go through the user-list
    Do Until tblUsers.EOF
    
        ' Retrive the information
        aNickNames = Split(tblUsers!NickNames, Chr(13))
        strEmail = tblUsers!Email
    
        ' Check if the user already exist
        For Tell = 1 To Logs.Users.Count
        
            ' See if the user correspond
            If IsUser(Logs.Users(Tell), aNickNames, strEmail) Then
            
                ' If so, save the user-index
                Users(tblUsers!id) = Tell
                
                ' We're finish searching
                Exit For
            
            End If
        
        Next
    
        ' Move to the next record
        tblUsers.MoveNext
    Loop
    
    ' Move to the beginning
    tblUsers.MoveFirst
    
    ' Then add all users that hasn't been added yet
    Do Until tblUsers.EOF
    
        ' If the user-element is empty, this means that the user dosen't exist, and thus we'll add it
        If Users(tblUsers!id) = 0 Then
        
            ' Add the user
            LogClass.AddUser Split(tblUsers!NickNames, Chr(13)), tblUsers!Email
        
            ' Then set the reference
            Users(tblUsers!id) = Logs.Users.Count
        
        End If

        ' Move to the next record
        tblUsers.MoveNext

    Loop
    
    ' Move to the beginning
    tblSessions.MoveFirst
    
    ' Go through the session-list
    Do Until tblSessions.EOF
    
        ' Check if the session already exist
        For Tell = 1 To Logs.Sessions.Count
        
            ' See if the sessions correspond
            If GetSessionDate(Logs.Sessions(Tell)) = CDate(tblSessions!Date) Then
            
                ' If so, save the session-index
                Sessions(tblSessions!id) = Tell
                
                ' We're finish searching
                Exit For
            
            End If
        
        Next
    
        ' Move to the next record
        tblSessions.MoveNext
    Loop
    
    ' Now, move to the beginning
    tblSessions.MoveFirst
    
    ' Then add all users that hasn't been added yet
    Do Until tblSessions.EOF
    
        ' If the session-element is empty, this means that the session dosen't exist, and thus we'll add it
        If Sessions(tblSessions!id) = 0 Then
        
            ' Add the user
            LogClass.AddSession GetDateOnly(CDate(tblSessions!Date)), GetUsers(tblSessions!Participants, Users, Logs)
        
            ' Then set the reference
            Sessions(tblSessions!id) = Logs.Sessions.Count
        
        End If

        ' Move to the next record
        tblSessions.MoveNext

    Loop
    
    ' Move messages-table to the beginning
    tblMessages.MoveFirst
    
    ' Ignore all errors henceforth
    On Local Error Resume Next
    
    ' And finally, add all messages
    Do Until tblMessages.EOF
    
        ' Get the parent session, sender and receiver
        Set Session = Logs.Sessions(Sessions(tblMessages!SessionID))
    
        ' See if this is an event or a message
        If tblMessages!IsEvent = 0 Then
            
            ' Get the sender and the receivers
            Set Sender = Logs.Users(Users(tblMessages!SenderID))
            Set Receiver = GetUsers(tblMessages!ReceiverID, Users, Logs)
            
            ' It's a message
            Session.AddMessage CDate(tblMessages!Date), Sender, Receiver, tblMessages!Text, IIf(IsNull(tblMessages!Style), "", tblMessages!Style)
        
            ' Clean up
            Set Sender = Nothing
            Set Receiver = Nothing
        
        Else
            
            ' It's an event
            Session.AddEvent CDate(tblMessages!Date), tblMessages!Text
        
        End If

        ' Move to the next record
        tblMessages.MoveNext
    
    Loop

End Sub
