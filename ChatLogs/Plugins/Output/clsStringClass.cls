VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStringClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

' Faster string-allocation
Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal olestr&, ByVal BLen As Long) As String

Private sBuffer As String
Private lngCurrPos As Long
Private lngSegmentSize As Long
Private lngActual As Long

Private Sub Class_Initialize()
    
    ' Set default values
    lngSegmentSize = 128
    lngCurrPos = 1
    
    ' Set the actual lenght of the buffer string
    lngActual = lngSegmentSize
    
    ' Preallocate the string
    AllocateString

End Sub

Private Sub AllocateString()

    ' Reallocate the string
    sBuffer = SysAllocStringByteLen(0&, lngActual)

End Sub

Public Sub AppendString(sString As String)

    Dim sTemp As String

    ' Firstly, see if a reallocating is needed
    If lngCurrPos + Len(sString) - 1 > lngActual Then
    
        ' Save the content of the string
        sTemp = ToString
    
        ' Increase by one segment
        lngActual = lngActual + lngSegmentSize
        
        ' Increse the size of the next segment
        lngSegmentSize = lngSegmentSize * 8
        
        ' Create a new buffer
        AllocateString
        
        ' Then set the newly created buffer
        Mid(sBuffer, 1, Len(sTemp)) = sTemp
    
    End If

    ' Insert string in the correct position
    Mid(sBuffer, lngCurrPos, Len(sString)) = sString

    ' Increase the write-position
    lngCurrPos = lngCurrPos + Len(sString)

End Sub

Public Property Get ToString() As String

    ' Retrive the string
    ToString = Mid(sBuffer, 1, lngCurrPos - 1)

End Property
