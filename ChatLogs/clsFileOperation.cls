VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Copyright (C) 2004 Kristian. S.Stangeland

'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

' Creates directories
Private Declare Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal psa As Any) As Long

Public Function ValidPath(sFile As String) As String

    ' Add a slash if needed
    ValidPath = sFile & IIf(Right(sFile, 1) = "\", "", "\")

End Function

Public Function GetFileName(sFile As String) As String

    ' Remove the path
    GetFileName = Right(sFile, Len(sFile) - InStrRev(sFile, "\"))

End Function

Public Function GetFilePath(sFile As String) As String

    ' Get the path
    GetFilePath = Left(sFile, InStrRev(sFile, "\"))

End Function

Public Function GetNoExtension(sFile As String) As String

    ' "Remove" the extension from the file name
    GetNoExtension = Left(sFile, InStrRev(sFile, ".") - 1)

End Function

Public Function GetFileExtension(sFile As String) As String

    ' Retrie the extension
    GetFileExtension = Right(sFile, Len(sFile) - InStrRev(sFile, "."))
    
End Function

Public Function RetriveFileList(sPath As String, sFileExtension As String, bSubFolders As Boolean, Optional Attributes As VbFileAttribute = vbDirectory) As Collection

    Dim Folders As New Collection, Folder, File As String, vFile
    Dim sFileName As String, bAdd As Boolean
    
    ' Create a new file container
    Set RetriveFileList = New Collection
    
    ' Find all files and folders
    File = Dir(FileOperation.ValidPath(sPath), Attributes)

    ' Loop until we've found the last file/folder
    Do While File <> ""
    
        ' Firstly, see if this in fact IS a file or folder
        If File <> "." And File <> ".." Then
    
            ' Then find out whether or not this is a file
            If File Like "*.*" Then
            
                ' Further on, it needs to meat a certain pattern
                If File Like sFileExtension Then
                
                    ' Add the real file name
                    RetriveFileList.Add FileOperation.ValidPath(sPath) & File
                    
                End If
            
            Else ' If not, this must be a folder
            
                ' Add this folder to the list
                Folders.Add FileOperation.ValidPath(sPath) & File
            
            End If
    
        End If
        
        ' Find the next file/folder
        File = Dir
    
    Loop
    
    ' Look in subfolders if requested
    If bSubFolders Then
        
        ' Go through all folders found
        For Each Folder In Folders
        
            ' Search inside this folder as well
            For Each vFile In RetriveFileList(CStr(Folder), sFileExtension, True, Attributes)
            
                ' Add the file to our list
                RetriveFileList.Add vFile
            
            Next
        
        Next
    
    End If

End Function

Public Sub SaveFile(Path As String, Data As String, Append As Boolean)

    Dim Free As Long
    
    ' If we're not appending and the file exist, delete it
    If (Not Append) And Dir(Path) <> "" Then
        Kill Path
    End If
    
    ' Firstly, create the directory, in case it dosen't already exist
    SHCreateDirectoryEx 0&, GetFilePath(Path), ByVal 0&
    
    ' Get a free socket
    Free = FreeFile
    
    ' Write data to file
    Open Path For Append As #Free
        Print #Free, Data
    Close #Free

End Sub

Public Function LoadFile(Path As String, Optional Index As Long = 1, Optional ByVal Lenght As Long = -1) As String

    Dim Free As Long, Buff As String
    
    ' Check the existence of the file
    If Dir(Path) = "" Then
        ' If the file dosen't already exist, don't try to read from it
        Exit Function
    End If
    
    ' Get a free handle
    Free = FreeFile
    
    ' Open file
    Open Path For Binary Access Read Lock Write As #Free
        
        ' If the lenght isn't specified, ignore it
        If Lenght < 0 Then
        
            ' Check and see if the lenght isn't a bit to long
            If LOF(Free) - Index < Lenght Then
                ' Accommodate for the new lenght
                Lenght = LOF(Free) - Index + 1
                
                ' If the new lenght is below or equal to zero, then we really don't need to get any data
                If Lenght <= 0 Then
                    Exit Function
                End If
                
            End If
        
        End If
        
        ' Don't proceed if the amout of data to be read is zero
        If Lenght <> 0 Then
        
            ' Allocate buffer according to the variable lenght if its not negative, if else, use lenght of file
            Buff = Space(IIf(Lenght >= 0, Lenght, LOF(Free)))
            
            ' Load data into the buffer from the specified index
            Get #Free, Index, Buff
        
        End If
        
    Close #Free
    
    ' Return the buffer
    LoadFile = Buff

End Function

Public Function FindFile(Path As String, Optional sExtension As String = "*", Optional bSearchTree As Boolean) As String
           
    Dim sFile As String, sFolder As Variant, cFolders As New Collection
    
    ' Sett variabel
    sFile = Dir(ValidPath(Path), vbDirectory) ' Legg til AND vbHidden for å attpåtil søke etter skjulte filer
    
    ' Søk gjennom hele mappen
    Do Until sFile = ""
        
        ' Vi må ekskludere "." og "..", idet disse ikke er virkelige filer
        If sFile <> "." And sFile <> ".." Then
            
            ' Finner ut hvorvidt dette er en fil eller ei
            If sFile Like "*.*" Then
            
                ' Sjekk at filen etterfølger ett visst mønster - benytter RegEx
                If sFile Like sExtension Then
                    
                    ' Vi har funnet filen
                    FindFile = ValidPath(Path) & sFile
                    
                    ' Intet mer skal gjøres
                    Exit Function
                
                End If
            
            Else ' Vi har funnet en mappe
                
                ' Behold mappen for senere utforskning
                cFolders.Add sFile
            
            End If
        
        End If
        
        ' Gå videre til neste fil
        sFile = Dir
    Loop
    
    ' Søk gjennom alle undermapper, såfremt dette er spesifisert
    If bSearchTree Then
    
        For Each sFolder In cFolders
        
            ' Søk gjennom mappe
            FindFile ValidPath(Path) & sFolder, sExtension, bSearchTree
            
        Next
    
    End If
           
End Function
