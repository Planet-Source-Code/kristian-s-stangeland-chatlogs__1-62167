VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContacts"
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

Private DOMDocument As DOMDocument

Property Get MenuCaption() As String

    ' This module dosen't require any caption
    MenuCaption = "Contacts"
    
End Property

Property Get FileFilter() As String

    ' Just ignored.
    FileFilter = "MSN Contact List (*.ctt)" & Chr(0) & "*.ctt"
    
End Property

Property Get FileExtension() As String

    ' The file extension this parser uses
    FileExtension = "ctt"
    
End Property

Property Get ExportName() As String

    ' This parser is primary used to parse messenger plus
    ExportName = "Contacts"

End Property

Property Get ExportHandleAll() As Boolean

    ' If this output modules will do all the work by itself
    ExportHandleAll = True

End Property

Property Get ExportHandleIO() As Boolean

    ' Just ignored.
    ExportHandleIO = False

End Property

Property Get ExportMayAppend() As Boolean

    ' In this case, this property is ignored.
    ExportMayAppend = False

End Property

Property Get ExporterDescription() As String

    ' The description of this parser
    ExporterDescription = "Used to export all loaded contacts to a single file"

End Property

Public Function RetriveName(Session As Object, userIndex As Long) As String
    
    ' Not used.
    
End Function

Public Sub ExportProc(LogContainer As Object, SaveDialog As Object)

    Dim sFile As String, Messenger, Service, ContactList, Contact, User, Free As Long

    ' Set the filter
    SaveDialog.Filter = FileFilter

    ' Invoke the file-open dialog
    SaveDialog.ShowSave 0&, "Save contact"
    
    ' Don't continue if no file was selected open was selected
    If SaveDialog.File = "" Then
    
        ' Exit the procedure
        Exit Sub
    
    End If
    
    ' Save the file for latter use
    sFile = SaveDialog.File

    ' Create a new document
    Set DOMDocument = New DOMDocument
    
    ' Then create the three main elements
    Set Messenger = DOMDocument.createElement("messenger")
    Set Service = DOMDocument.createElement("service")
    Set ContactList = DOMDocument.createElement("contactlist")
    
    ' Set properties
    Service.setAttribute "name", ".NET Messenger Service"
    
    ' Insert the top-level element
    DOMDocument.insertBefore Messenger, Nothing
    
    ' Then insert its children
    Messenger.appendChild DOMDocument.createTextNode(vbCrLf & Space(2))
    Messenger.appendChild Service
    Service.appendChild DOMDocument.createTextNode(vbCrLf & Space(4))
    Service.appendChild ContactList
    
    ' Add each contact to the list
    For Each User In LogContainer.Users
    
        ' Create a new element
        Set Contact = DOMDocument.createElement("contact")
    
        ' Set its contents
        Contact.Text = User.Email
    
        ' Add the element
        ContactList.appendChild DOMDocument.createTextNode(vbCrLf & Space(6))
        ContactList.appendChild Contact
    
    Next
    
    ' Add the final text nodes
    ContactList.appendChild DOMDocument.createTextNode(vbCrLf & Space(4))
    Service.appendChild DOMDocument.createTextNode(vbCrLf & Space(2))
    Messenger.appendChild DOMDocument.createTextNode(vbCrLf)
    
    ' Create and insert the process instruction
    DOMDocument.insertBefore DOMDocument.createProcessingInstruction("xml", "version=" & Chr(34) & "1.0" & Chr(34)), Messenger

    ' Get a free handle
    Free = FreeFile

    ' Now, save the resulting XML
    Open sFile For Binary As #Free
        Put #Free, , DOMDocument.xml
    Close #Free

End Sub

