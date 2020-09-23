VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContainer"
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

' Events
Event FilesListed(ByVal Path As String, ByVal Amout As Long)
Event FileFound(ByVal FileName As String, AddManually As Boolean)
Event FolderImported(ByVal Path As String)
Event StatisticCalculation(ByVal UserIndex As Long, ByVal Max As Long)
Event StatisticsCalculated()

' The parser used to get the information
Public Parser As New clsParser

' The collection that contains all the logs
Public Logs As New Collection
Public Users As New Collection
Public Sessions As New Collection

Public Function AddLog(sFile As String, LogType As String) As clsLog
    
    Dim Log As New clsLog, Session As clsSession, Tell As Long
    
    ' Surely, the first alternative cannot be opened
    If LogType = "Unknown" Then
        Exit Function
    End If
        
    ' Set the parent
    Set Log.Parent = Me
    
    ' Open this new log
    Log.ImportLog sFile, LogType
    
    ' If we've been told to ignore this log, to do
    If Log.LogIgnore Then
        Exit Function
    End If
    
    ' Then, see if we have already added this log
    For Tell = 1 To Logs.Count
    
        ' See if the two logs are equal
        If Logs(Tell).IsEqual(Log) Then
        
            ' Just ignore this log
            Exit Function
        
        End If
    
    Next
    
    ' Update user-lists
    For Each Session In Log.Sessions
    
        ' Update user lists in this session
        Session.UpdateUsers
    
    Next
    
    ' Add the log to the log-list
    Logs.Add Log
    
    ' And set its index
    Log.Index = Logs.Count
    
    ' And finally return this log
    Set AddLog = Log

End Function

Public Sub RefreshStatistic()

    Dim Tell As Long
    
    ' Go through all users
    For Tell = 1 To Users.Count
    
        ' Inform about the calucation
        RaiseEvent StatisticCalculation(Tell, Users.Count)
    
        ' Update the values
        Users(Tell).CalculateValues
    
    Next
    
    ' Inform that the calulations are finished
    RaiseEvent StatisticsCalculated

End Sub

Public Sub ImportFolder(sPath As String, sFileExtension As String, Optional bSubFolders As Boolean, Optional Attributes As VbFileAttribute = vbDirectory)
    
    Dim Files As Collection, Tell As Long, bAdd As Boolean, sFileName As String
    
    ' Get the file list
    Set Files = FileOperation.RetriveFileList(sPath, sFileExtension, bSubFolders, Attributes)
    
    ' All files have been found
    RaiseEvent FilesListed(sPath, Files.Count)
    
    ' Inform the parser that we're about to begin
    ChatApp.Parsers(ChatApp.ParserByFile(sFileExtension)).ParserInitialize
    
    ' Search through all of them and perform the requested action
    For Tell = 1 To Files.Count
    
        ' Get the file name of this element
        sFileName = Files(Tell)
        
        ' Reset variable
        bAdd = False
        
        ' Inform about the file
        RaiseEvent FileFound(sFileName, bAdd)
        
        ' Add the file 'automatically' if nothing else is specified
        If Not bAdd Then
        
            ' Add the file to the list
            AddLog sFileName, ChatApp.ParserByFile(sFileName)
            
        End If

    Next
    
    ' And inform the top-class by invoking this event
    RaiseEvent FolderImported(sPath)

    ' Update statistics
    RefreshStatistic

End Sub

Public Function FindUser(User As clsUser) As Long

    Dim Tell As Long
    
    ' Go through all added users
    For Tell = 1 To Users.Count
    
        ' See if it corresponds to the search-element
        If User.IsEqual(Users(Tell)) Then
        
            ' If it does, return its index
            FindUser = Tell
            
            ' ... and exit
            Exit Function
    
        End If
    
    Next

    ' Nothing was found
    FindUser = -1

End Function

Public Sub RemoveLog(Index)

    On Local Error Resume Next
    Dim Log As clsLog

    ' Firstly, get the log
    Set Log = Logs(Index)
    
    ' Continue if we got something
    If Not Log Is Nothing Then
    
        ' Firstly, remove the log from the collection
        Logs.Remove Index
    
        ' Then remove the log and all its subclasses
        Log.ClearLog
    
    End If

End Sub
