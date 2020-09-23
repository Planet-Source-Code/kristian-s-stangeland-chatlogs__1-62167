VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutput"
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

Property Get MenuCaption() As String

    ' This module dosen't require any caption
    MenuCaption = ""
    
End Property

Property Get FileFilter() As String

    ' The filter used to browse for such files
    FileFilter = "Access Database (*.mdb)" & Chr(0) & "*.mdb"
    
End Property

Property Get FileExtension() As String

    ' The file extension this writer uses
    FileExtension = "mdb"
    
End Property

Property Get ExportName() As String

    ' This parser is primary used to write database-files
    ExportName = "Access Database"

End Property

Property Get ExportHandleAll() As Boolean

    ' If this output modules will do all the work by itself
    ExportHandleAll = False

End Property

Property Get ExportHandleIO() As Boolean

    ' Whether or not the export-function does all the I/O-operations itself
    ExportHandleIO = True

End Property

Property Get ExportMayAppend() As Boolean

    ' Whether or not the export-function returns the ENTIRE file-content or simply just what needs to be appended
    ExportMayAppend = False

End Property

Property Get ExporterDescription() As String

    ' The description of this parser
    ExporterDescription = "Used to export to a Access Database"

End Property

Public Function AddTable(Database As Database, sName As String, ParamArray Fields() As Variant) As Recordset
 
    Dim tblDef As TableDef, Tell As Long, tblID As Field, IndexObject As Index, IndexField As Field

    ' Create the table
    Set tblDef = Database.CreateTableDef(sName)
    
    ' Add the ID-field
    Set tblID = tblDef.CreateField("ID", dbLong)
    
    ' Set properties and append
    tblID.Attributes = dbAutoIncrField
    tblDef.Fields.Append tblID
    
    ' Then all fields
    For Tell = LBound(Fields) To UBound(Fields) Step 2
    
        ' Add the new field
        tblDef.Fields.Append tblDef.CreateField(Fields(Tell), Fields(Tell + 1))
    
    Next
    
    ' Now, create the index-field
    Set IndexObject = tblDef.CreateIndex("PrimaryKey")
    Set IndexField = IndexObject.CreateField("ID", dbLong)
    
    ' And add the field
    IndexObject.Fields.Append IndexField
    
    ' Set it as primary
    IndexObject.Primary = True
    
    ' Append the index
    tblDef.Indexes.Append IndexObject
    
    ' Add the table to the database
    Database.TableDefs.Append tblDef
    Database.TableDefs.Refresh
    
    ' Then create the recordset
    Set AddTable = Database.OpenRecordset(sName, dbOpenTable)

End Function

Private Function CreateNewDatabase(sFile As String, tblSessions As Recordset, tblUsers As Recordset, tblMessages As Recordset) As Database

    ' Create a new database at the specified location
    Set CreateNewDatabase = CreateDatabase(sFile, dbLangGeneral)
    
    ' Then create the session-table
    Set tblSessions = AddTable(CreateNewDatabase, "Sessions", "Date", dbDate, "Participants", dbText)
    
    ' Whereupon we'll create the user-table
    Set tblUsers = AddTable(CreateNewDatabase, "Users", "Nicknames", dbMemo, "Email", dbText)

    ' And finally, we'll make the message-table
    Set tblMessages = AddTable(CreateNewDatabase, "Messages", "SessionID", dbLong, "SenderID", _
                       dbLong, "ReceiverID", dbText, "Date", dbDate, "Text", dbMemo, "Style", _
                        dbText, "EventType", dbLong, "IsEvent", dbBoolean)
                        
End Function

Private Sub ClearTag(Logs As Object)

    Dim User, Session

    ' Just reset the tag-property in all users
    For Each User In Logs.Users
        User.Tag = ""
    Next
    
    ' Then in all sessions
    For Each Session In Logs.Sessions
        Session.Tag = ""
    Next

End Sub

Private Function GetForm(Collection As Object, sName As String) As Object

    Dim Form As Object
    
    ' Go through each form in the collection
    For Each Form In Collection
    
        ' See if the name corresponds
        If Form.Name = sName Then
        
            ' Return this form
            Set GetForm = Form
        
            ' We're finish searching
            Exit For
        
        End If
    
    Next

End Function

Public Function RetriveName(Session As Object, userIndex As Long) As String
    
    ' Return the second email of this session
    RetriveName = Session.Participants(userIndex).Email
    
End Function

Public Function ExportLog(Logs As Object, sExistingFile As String, FileOperation As Object) As String

    ' The database-objects
    Dim Database As Database
    Dim tblSessions As Recordset
    Dim tblUsers As Recordset
    Dim tblMessages As Recordset
    
    ' Used in different part of the export-process
    Dim User As Object
    Dim Session As Object
    Dim oEvent As Object, SessionID As Long
    Dim frmProgress As Object, dDate As Date
    Dim aNickNames, strEmail As String

    ' Then, get the progress-form (we know it is active)
    Set frmProgress = GetForm(MainObject.ActiveForms, "frmProgress")

    ' Now, see first if the file dosen't exist
    If Dir(sExistingFile) = "" Then
    
        ' If so, create the database
        Set Database = CreateNewDatabase(sExistingFile, tblSessions, tblUsers, tblMessages)

    Else
    
        ' If it does exist, read from it
        Set Database = OpenDatabase(sExistingFile)
    
        ' Then open all the tables
        Set tblSessions = Database.OpenRecordset("Sessions")
        Set tblUsers = Database.OpenRecordset("Users")
        Set tblMessages = Database.OpenRecordset("Messages")
    
    End If
    
    ' Clear tag-properties
    ClearTag Logs
    
    ' Move to first element (if the table actually has records)
    If Not (tblUsers.BOF And tblUsers.EOF) Then
        tblUsers.MoveFirst
    End If
    
    ' First, check out all the already existing users
    Do Until tblUsers.EOF
    
        ' Get email and nicknames
        aNickNames = Split(tblUsers!NickNames, Chr(13))
        strEmail = tblUsers!Email
    
        ' Now, loop through all users ...
        For Each User In Logs.Users
        
            ' ... and see if it correspond to this user
            If IsUser(User, aNickNames, strEmail) Then
            
                ' Save the reference to this user in the table
                User.Tag = tblUsers!id

                ' We're finish searching
                Exit For

            End If
        
        Next
        
        ' Move to next user
        tblUsers.MoveNext
    
    Loop
    
    ' Now, add the remaing users
    For Each User In Logs.Users
    
        ' If the tag-property is empty, we haven't already found the user
        If User.Tag = "" Then
    
            ' Add a new element
            tblUsers.AddNew
            tblUsers!NickNames = JoinNames(User.NickNames, Chr(13))
            tblUsers!Email = User.Email
            
            ' Remember the ID of this user
            User.Tag = CStr(tblUsers!id)
            
            ' Update the record
            tblUsers.Update
    
        End If
    
    Next
    
    ' Move to first element, if the table has records
    If Not (tblSessions.BOF And tblSessions.EOF) Then
        tblSessions.MoveFirst
    End If
    
    ' Now, loop through all sessions
    Do Until tblSessions.EOF
    
        ' Save the date of this session
        dDate = tblSessions!Date
    
        ' Go through each session
        For Each Session In Logs.Sessions
            
            ' If the dates correspond, this session has already been added
            If dDate = GetSessionDate(Session) Then
                
                ' Save the ID
                Session.Tag = tblSessions!id
                
                ' We're finsih searching
                Exit For
                
            End If
        
        Next
    
        ' Move to next element
        tblSessions.MoveNext
    Loop
    
    ' Ignore errors henceforth
    On Local Error Resume Next
    
    ' Further on, we'll add all sessions that hasn't been added yet
    For Each Session In Logs.Sessions
    
        ' Set the description of the progress-form plus its progress
        frmProgress.progressbar.Value = frmProgress.progressbar.Value + 1
        frmProgress.lblDescription = "Adding session " & frmProgress.progressbar.Value & " of " & frmProgress.progressbar.Max
        
        ' Allow events to occur
        MainObject.AllowEvents
    
        ' If the tag-property is empty, this session hasn't been added yet
        If Session.Tag = "" Then
                
            ' Add the session
            tblSessions.AddNew
            tblSessions!Date = GetSessionDate(Session)
            tblSessions!Participants = JoinID(Session.Participants, Chr(13))
            
            ' Retrive the ID
            SessionID = tblSessions!id
            
            ' Update the record
            tblSessions.Update
            
            ' Then, add all events and messages
            For Each oEvent In Session.Events

                ' Add a new message
                tblMessages.AddNew
                tblMessages!SessionID = SessionID
                tblMessages!SenderID = oEvent.Sender.Tag
                tblMessages!ReceiverID = JoinID(oEvent.Receiver, Chr(13))
                tblMessages!Text = oEvent.Text
                
                ' The following properties depend on the type of the event
                If oEvent.IsMessage Then
                
                    ' Don't set it if it's empty
                    If oEvent.Style <> "" Then
                        tblMessages!Style = oEvent.Style
                    End If
                
                    ' This is a normal message
                    tblMessages!Date = oEvent.MessageDate
                    tblMessages!EventType = 0
                    tblMessages!IsEvent = False
                
                Else
                
                    ' This is an event
                    tblMessages!Date = oEvent.EventDate
                    tblMessages!EventType = oEvent.EventType
                    tblMessages!IsEvent = True
                
                End If
                
                ' Update the message
                tblMessages.Update

            Next
        
        End If
    
    Next
    
    ' Clean up everything
    Set tblMessages = Nothing
    Set tblUsers = Nothing
    Set tblSessions = Nothing
    Set Database = Nothing
    
    ' Clear tag-properties
    ClearTag Logs
    
End Function





