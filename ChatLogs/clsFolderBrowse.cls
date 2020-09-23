VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFolderBrowse"
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

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

' Different flag-values
Public Enum BIF_FLAGS
    BIF_RETURNONLYFSDIRS = 1
    BIF_DONTGOBELOWDOMAIN = 2
    BIF_RETURNFSANCESTORS = 8
    BIF_EDITBOX = 16
    BIF_VALIDATE = 32
    BIF_NEWDIALOGSTYLE = 64
    BIF_BROWSEINCLUDEFILES = 16384
End Enum

Enum BrowseResult
    Browse_OK
    Browse_CANCEL
End Enum

' The default maximum lenght of a path
Const MAX_PATH = 260

' Different variables used to customize the browse-form
Public Caption As String
Public Description As String
Public Filters As String
Public Flags As BIF_FLAGS

' Variables holding the result
Public ResultPath As String
Public ResultFilter As String
Public SubFolders As Boolean

' The form to use
Private currentBrowse As frmBrowse

' The main function which invokes the form
Public Function Browse(Optional bCustomizedForm As Boolean = True) As BrowseResult

    Dim iNull As Integer, lpIDList As Long, lResult As Long, aFilters
    Dim sPath As String, udtBI As BrowseInfo, Tell As Long

    ' Choose between the different ways of selecting a folder
    Select Case bCustomizedForm
    Case True

        ' Allocate a new form
        Set currentBrowse = New frmBrowse
        
        ' Customize it
        currentBrowse.Caption = Caption
        currentBrowse.lblDescription = Description
        
        ' Clear the filter-control
        currentBrowse.cmdFileType.Clear
        
        ' Split the filter-string
        aFilters = Split(Filters, Chr(0))
        
        ' Add all filters
        For Tell = LBound(aFilters) To UBound(aFilters) Step 2
            ' Add the filter
            currentBrowse.cmdFileType.AddItem aFilters(Tell)
        Next
        
        ' Set the default index
        currentBrowse.cmdFileType.ListIndex = 0
        
        ' Show it
        currentBrowse.Show
        
        ' Start looping
        Do Until currentBrowse.Tag <> ""
        
            ' Allow events to be processed
            DoEvents
            
            ' Wait a bit to save CPU
            Sleep 10
            
        Loop
        
        ' Get the result(s)
        ResultPath = currentBrowse.lstFolders.Path
        SubFolders = (currentBrowse.chkAddSubfolders.Value = 1)
        ResultFilter = aFilters((currentBrowse.cmdFileType.ListIndex * 2) + 1)
        
        ' Check the end result and return accordingly
        If currentBrowse.Tag = "OK" Then
            Browse = Browse_OK
        Else
            Browse = Browse_CANCEL
        End If
        
        ' Hide and deallocte form
        currentBrowse.Hide
        Set currentBrowse = Nothing
    
    Case False
    
        ' Firstly assume that we didn't get a OK
        Browse = Browse_CANCEL
    
        ' Initialize user-defined type
        With udtBI
            .lpszTitle = lstrcat(Description, "")
            .ulFlags = Flags
        End With
    
        'Show the 'Browse for folder' dialog
        lpIDList = SHBrowseForFolder(udtBI)
        
        ' Then try to get the folder
        If lpIDList Then
            
            ' Allocate buffer
            sPath = String$(MAX_PATH, 0)
            
            'Get the path from the IDList
            SHGetPathFromIDList lpIDList, sPath
            
            'free the block of memory
            CoTaskMemFree lpIDList
            iNull = InStr(sPath, vbNullChar)
            
            If iNull Then
            
                ' Set the return path
                ResultPath = Left$(sPath, iNull - 1)
                
                ' Everything went ok
                Browse = Browse_OK
                
            End If
            
        End If
        
    End Select
    
End Function

Private Sub Class_Initialize()

    ' Set the default value
    Flags = BIF_RETURNONLYFSDIRS

End Sub

