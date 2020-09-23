Attribute VB_Name = "modGlobal"
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

Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' A type holding information about an event or a message
Type Element
    Date As Date
    Type As eEventType
    Text As String
    Style As String ' Not available in all types of chat logs
    Tag As Variant
    IsMessage As Boolean
    Sender As Long
    Receiver() As Long
    ReceiverCount As Long
End Type

' A type used to change the text of a column
Type LVCOLUMN
    mask As Long
    fmt As Long
    cx As Long
    pszText As String
    cchTextMax As Long
    iSubItem As Long
    iImage As Long
    iOrder As Long
End Type

' Used when initializing common controls
Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

' The different event-types
Enum eEventType
    Event_Unknown
    Event_Invitation
    Event_InvitationResponse
    Event_Join
    Event_Leave
End Enum

' Used in ListView
Public Const LVM_FIRST = &H1000
Public Const LVM_GETCOLUMN = (LVM_FIRST + 25)
Public Const LVM_SETCOLUMN = (LVM_FIRST + 26)

Public Const LVCF_TEXT = &H4
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const ICC_USEREX_CLASSES = &H200

' Different application related operations
Public ChatApp As New clsApplication
Public FileOperation As New clsFileOperation

' This is just a simplifying of how to extract a language string
Public Function LanguageConst(sName As String) As String

    ' Return the variable
    LanguageConst = ChatApp.Language.ConstantValue(sName)

End Function

Public Sub Main()

    ' Try to initialize the common-controls-dll, making XP-style possible
    If InitCommonControlsVB = False Then
        MsgBox LanguageConst("ErrComCTL"), vbCritical, "Error"
        End
    End If
    
    ' Load all plugins
    LoadPlugins
    
    ' Load langage pack
    ChatApp.ReloadLanguage
    
    ' Retrive the command line and process it
    ChatApp.ExecuteCommands Command$
    
    ' Show the main form
    frmMain.Update
    frmMain.Show

End Sub

Public Sub LoadPlugins()

    On Error Resume Next
    Dim sFile As Variant, Plugin As Object, strClassName As String
    
    ' Go through all executable files in the plugin-folder
    For Each sFile In FileOperation.RetriveFileList(FileOperation.ValidPath(App.Path) & "Plugins\", "*.dll", False, vbNormal)
    
        Select Case FileOperation.GetFileExtension(CStr(sFile))
        Case "dll"

            ' Clear all errors
            Err.Clear
            
            ' How the class is registered in the registry
            strClassName = FileOperation.GetNoExtension(FileOperation.GetFileName(CStr(sFile))) & ".PluginMain"
            
            ' Try to create the object
            Set Plugin = CreateObject(strClassName)
    
            If Err = 429 Then ' ERROR: ActiveX component can't create object
                
                ' Clear the error
                Err.Clear
                
                ' Try to register the object
                Shell "regsvr32 /s " & Chr(34) & CStr(sFile) & Chr(34)
                
                ' Load the plugin again
                Set Plugin = CreateObject(strClassName)
            
            End If
    
            ' Now, if this didn't work either, it could be something wrong with the registration of the DLL, so we'll retry
            If Err = 429 Then
            
                ' Firstly, clear the error
                Err.Clear
                
                ' Then, unregister the DLL
                Shell "regsvr32 /u /s " & Chr(34) & CStr(sFile) & Chr(34)
                
                ' And register it again
                Shell "regsvr32 /s " & Chr(34) & CStr(sFile) & Chr(34)
            
                ' Try to load the plugin for the last time
                Set Plugin = CreateObject(strClassName)
            
            End If
            
            ' If even this dosen't work, we'll just ignore the hole thing
            If Err <> 429 Then
    
                ' See if this plugin actually is allowed
                If ChatApp.Settings.PluginEnabled(Plugin.Name) Then
        
                    ' Add plugin
                    ChatApp.Plugins.Add Plugin, Plugin.Name
                
                    ' Initialize plugin
                    Plugin.Initialize ChatApp
                
                End If

            End If

        End Select
        
    Next

End Sub

Public Function InitCommonControlsVB() As Boolean

   On Error Resume Next
   
   Dim iccex As tagInitCommonControlsEx
   
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   
   On Error GoTo 0
   
End Function

Public Function GetTimeOnly(dateTime As Date) As Date

    ' Return only the time-part of the date
    GetTimeOnly = TimeSerial(Hour(dateTime), Minute(dateTime), Second(dateTime))

End Function

Public Sub SetColumnText(ListView As Object, ColumnIndex As Long, ColumnText As String)

    Dim Column As LVCOLUMN

    ' Initialize the type
    Column.mask = LVCF_TEXT
    Column.cchTextMax = Len(ColumnText) + 1
    Column.pszText = ColumnText & Chr(0)
    
    ' Change the text of the column
    SendMessage ListView.hWndListView, LVM_SETCOLUMN, ColumnIndex, Column

End Sub

Public Function NewSortOrder(ByVal SortOrder As ESortOrderConstants) As ESortTypeConstants
   
   ' Find the opposite sort-order
   Select Case SortOrder
   Case eSortOrderNone, eSortOrderDescending
      NewSortOrder = eSortOrderAscending
   Case eSortOrderAscending
      NewSortOrder = eSortOrderDescending
   End Select
   
End Function
