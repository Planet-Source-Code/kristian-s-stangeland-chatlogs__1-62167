VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsApplication"
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

' All the different parsers registered
Public Parsers As New Collection

' Export-modules
Public ExportModules As New Collection

' All plugins added and initialized
Public Plugins As New Collection

' All the settings of this application
Public Settings As New clsSettings

' The class containing the current language
Public Language As New clsLanguage

' Different prosedures used when printing
Public PrintClass As New clsPrint

Public Sub InitializeForm(Form As Object)

    ' Just set the language
    Language.UseLanguage Form

End Sub

Public Sub AddColumns(ListView As Object, ParamArray Columns() As Variant)

    Dim Tell As Long

    ' If the amount of existing columns correspond to what we're about to add, rename them insted
    If (UBound(Columns) - LBound(Columns) + 1) / 2 = ListView.Columns.Count Then
    
        ' Just rename the columns
        For Tell = LBound(Columns) To UBound(Columns) Step 2
            SetColumnText ListView, Tell / 2, CStr(Columns(Tell + 1))
        Next
    
    Else
    
        ' Clear all columns
        ListView.Columns.Clear
        
        ' Then add then, item for item
        For Tell = LBound(Columns) To UBound(Columns) Step 2
            ListView.Columns.Add , CStr(Columns(Tell)), CStr(Columns(Tell + 1))
        Next
    
    End If
    
    ' Update the listbox
    ListView.Refresh

End Sub

Public Sub ReloadLanguage()

    Dim Form As Form

    ' Load the language file referenced in the registry
    Language.LoadFile FileOperation.ValidPath(App.Path) & "Data\Languages\" & _
                       ChatApp.Settings.Setting("Settings", "LanguagePack", "English.lpk")

    ' Then update all forms
    For Each Form In Forms
        Language.UseLanguage Form
    Next

End Sub

Public Sub AllowEvents()

    ' Allow events to occur
    DoEvents

End Sub

Public Property Get ActiveForms() As Object

    ' Return all open forms
    Set ActiveForms = Forms

End Property

Public Property Get FileInputFilters() As String

    Dim Tell As Long
    
    ' Loop through all registered parsers
    For Tell = 1 To Parsers.Count
        
        ' Add its filter-string
        FileInputFilters = FileInputFilters & Parsers(Tell).FileFilter & IIf(Tell <> Parsers.Count, Chr$(0), "")
    
    Next

End Property

Public Property Get FileSystem() As clsFileOperation

    ' Return this class
    Set FileSystem = FileOperation

End Property

Public Property Get Logs() As clsContainer

    ' Return the used container
    Set Logs = frmMain.Logs

End Property

Public Sub ExecuteCommands(sCommands As String)

    On Local Error Resume Next
    Dim Lines, Line, Command, Paramenters

    ' Split the command into different lines
    Lines = SplitX(sCommands, "/")
    
    ' Go through each line
    For Each Line In Lines

        ' Retrive the command
        Command = Split(Line, " ", 2)
        
        ' See if we have a remaining
        If UBound(Command) > 0 Then
            ' The remaing is the paramenters
            Paramenters = SplitX(CStr(Command(1)), ",")
        Else
            ' If not, the command has no paramenters
            Paramenters = Array()
        End If

        ' Then process the command
        ProcessCommand CStr(Command(0)), Paramenters

    Next

End Sub

Public Sub ProcessCommand(sCommand As String, Optional Paramenters As Variant)

    ' Execute the different commands
    Select Case LCase(CStr(Command(0)))
    Case "run"
    
        ' We cannot execute this command without at least one paramenter
        If UBound(Paramenters) >= 0 Then
            ' Load the file
            frmMain.LoadFile ExtractValue(CStr(Paramenters(0)))
        End If
    
    Case "terminate"
    
        ' Exit the application
        Terminate
    
    End Select

End Sub

Private Function ExtractValue(sText As String) As Variant

    Dim lngBegin As Long, lngEnd As Long
    
    ' Find two quotation marks
    lngBegin = InStr(1, sText, """")
    lngEnd = InStr(lngBegin + 1, sText, """")

    ' Now, see if the text contains quotation marks
    If lngBegin And lngEnd Then
    
        ' Return the text between those two
        ExtractValue = Mid(sText, lngBegin + 1, lngEnd - lngBegin - 2)
    
    Else
    
        ' If not, this is most likely a number
        ExtractValue = Val(sText)
    
    End If

End Function

Public Sub Terminate()

    Dim Form As Form
    
    ' Go through each form ...
    For Each Form In Forms
    
        ' ... and close it
        Unload Form
    
    Next

End Sub

' Returns whether or not a form is active, without accidentally activating it
Public Function IsVisible(refForm As Object) As Boolean

    Dim Form As Form

    ' Search through all active forms
    For Each Form In Forms
        
        ' See if this form correspond to the form we're searching for
        If refForm Is Form Then
        
            ' If so, return the current state of the visible-property
            IsVisible = Form.Visible
        
            ' We're finished
            Exit For
        
        End If
    
    Next

End Function

Public Function CreateFilter(FilterType As eFilterType, ResultValue As Boolean, Optional ValueOne, Optional ValueTwo, Optional UsePattern, Optional IgnoreCase) As clsFilter

    On Error Resume Next

    ' Firstly, create a new filter-class
    Set CreateFilter = New clsFilter
    
    ' Then set its settings
    CreateFilter.FilterType = FilterType
    CreateFilter.ResultValue = ResultValue
    CreateFilter.ValueOne = ValueOne
    CreateFilter.ValueTwo = ValueTwo
    CreateFilter.UsePattern = UsePattern
    CreateFilter.IgnoreCase = IgnoreCase

End Function

Public Function ParserByFile(sFile As String) As String

    Dim sExtension As String, Parser As Object
    
    ' Get the file extension of this file
    sExtension = FileOperation.GetFileExtension(sFile)

    ' Loop though all parsers registered
    For Each Parser In Parsers
    
        ' If this parser handle the file extension, use this parser
        If Parser.FileExtension = sExtension Then
        
            ' Return the parser name
            ParserByFile = Parser.ParserName
        
            ' We're finished searching
            Exit Function
        
        End If
    
    Next
    
    ' The default return-value if nothing is found
    ParserByFile = "Unknown"

End Function

Public Function JoinParticipants(partCollection As Collection, Delimiter As String, Optional ByVal StartIndex As Long = -1, Optional ByVal EndIndex As Long = -1) As String

    Dim Tell As Long
    
    ' Set the default value if the assigned value is invalid
    If StartIndex <= -1 Then
        StartIndex = 1
    End If
    
    ' As the code above
    If EndIndex <= -1 Then
        EndIndex = partCollection.Count
    End If
    
    ' Go through the elements as specified
    For Tell = StartIndex To EndIndex
        
        ' Add the participant's nickname and email, with the delimiter at the end of each participant
        JoinParticipants = JoinParticipants & partCollection(Tell).NickName & IIf(partCollection(Tell).Email <> "", " (" & partCollection(Tell).Email & ")", "") & IIf(Tell < EndIndex, Delimiter, "")
    
    Next

End Function

' Remove an element from a collection
Public Sub RemoveFromCollection(refCollection As Collection, refClass As Object)

    Dim Tell As Long
    
    ' Go through all object-elements
    For Tell = 1 To refCollection.Count
    
        ' See if they are alike
        If refClass Is refCollection(Tell) Then
        
            ' Remove the element
            refCollection.Remove Tell
            
            ' We are finished
            Exit Sub
        
        End If
    
    Next

End Sub
