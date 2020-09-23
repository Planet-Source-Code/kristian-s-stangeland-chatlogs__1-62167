Attribute VB_Name = "modGeneral"
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

' Used to optimize string-reading
Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (pDst As Any, ByVal ByteLen As Long)
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

' Used in conjunction with the below
Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

' The safe-array structure used in the fast string access
Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0) As SAFEARRAYBOUND
End Type

Public Const MB_PRECOMPOSED = &H1       ' Use precomposed chars
Public Const MB_COMPOSITE = &H2         ' Use composite chars
Public Const MB_USEGLYPHCHARS = &H4     ' Use glyph chars, not ctrl chars

Public Const WC_DEFAULTCHECK = &H100    ' Check for default char
Public Const WC_COMPOSITECHECK = &H200  ' Convert composite to precomposed
Public Const WC_DISCARDNS = &H10        ' Discard non-spacing chars
Public Const WC_SEPCHARS = &H20         ' Generate separate chars
Public Const WC_DEFAULTCHAR = &H40      ' Replace with default Char

Public Const CP_ACP = 0
Public Const CP_OEMCP = 1
Public Const CP_MACCP = 2
Public Const CP_THREAD_ACP = 3
Public Const CP_SYMBOL = 42
Public Const CP_UTF7 = 65000
Public Const CP_UTF8 = 65001

' Character numbers used in the function below
Public Const QuotationMark As Integer = 34
Public Const Apostrophe As Integer = 39
Public Const LessThan As Integer = 60
Public Const GreaterThan As Integer = 62
Public Const Ampersand As Integer = 38

Public Sub ConvertString(sText As String, clsString As clsStringClass)

    On Error Resume Next
    Dim sData As String, sResult As String, intChar() As Integer, currChar As Integer
    Dim SafeArray As SAFEARRAY1D, Tell As Long, lngLast As Long, bAddedAprostophe As Boolean
    Dim bAprostophe As Boolean, bQuote As Boolean
        
    ' Let the array point to the string
    InitializeArray intChar, SafeArray, sText

    ' Then let us begin at the first position
    lngLast = 1

    ' Go trough the entire array
    For Tell = 1 To UBound(intChar)
    
        ' Get the current char
        currChar = intChar(Tell)
    
        ' See if we shoukd replace this character with an HTML-entity
        If currChar = LessThan Or currChar = GreaterThan Or currChar = Ampersand Or currChar = QuotationMark Or currChar = Apostrophe Or currChar > 127 Then
        
            ' Add the text before
            clsString.AppendString Mid(sText, lngLast, Tell - lngLast)
            
            ' Add this enitity
            clsString.AppendString "&#" & currChar & ";"

            ' The last position is now this plus one
            lngLast = Tell + 1

        End If
    
    Next
    
    ' See if we need to add a string
    If lngLast < Len(sText) Then
        ' Add the string
        clsString.AppendString Mid(sText, lngLast)
    End If
    
    ' Clear up
    ZeroMemory ByVal VarPtrArray(intChar), 4
    
End Sub

Public Sub InitializeArray(intChar() As Integer, SafeArray As SAFEARRAY1D, sRefString As String)

    ' Firstly, point the integer array to the string
    With SafeArray
        .cDims = 1
        .Bounds(0).lLbound = 1
        .Bounds(0).cElements = Len(sRefString)
        .pvData = StrPtr(sRefString)
    End With
    
    ' Set the array
    CopyMemory ByVal VarPtrArray(intChar), VarPtr(SafeArray), 4

End Sub

Public Function UniToMB(sSrc As String, lCP As Long) As Byte()

    ' Convert Unicode to multi-byte
    Dim byDst() As Byte
    Dim lNC As Long, lSize As Long, lRet As Long

    ' Caluclate sizes
    lNC = Len(sSrc)
    lSize = 2 * lNC
    
    ' Allocate the buffer
    ReDim byDst(0 To lSize - 1)
    
    ' Convert the string
    lRet = WideCharToMultiByte(lCP, 0, StrPtr(sSrc), lNC, VarPtr(byDst(0)), lSize, vbNullString, 0)
    
    ' Reallocate the buffer (so it fits the result)
    ReDim Preserve byDst(0 To lRet - 1)
    
    ' Return the result
    UniToMB = byDst
    
End Function
