VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLanguage"
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

' Our main type that will contain all the data
Private Type Constant
    Name As String
    Value As Variant
End Type

' The script engine
Private Script As Object

' Local arrays that holds the data
Private aConstants() As Constant
Private aForms() As Constant

' The amount of constants
Public ConstantCount As Long
Private FormCount As Long

Public Property Get ConstantValue(ByVal Index As Variant) As Variant

    ' First, get the index
    GetIndex Index, aConstants
    
    ' Then see if it's valid
    If Index >= 0 Then
        ' Return the value
        ConstantValue = aConstants(Index).Value
    End If
    
End Property

Public Property Let ConstantValue(ByVal Index As Variant, ByVal vNewValue As Variant)
    
    ' First, get the index
    GetIndex Index, aConstants
    
    ' Then see if it's valid
    If Index >= 0 Then
        ' Set the valye
        aConstants(Index).Value = vNewValue
    End If

End Property

Public Sub LoadFile(sFileName As String)

    Dim sBuffer As String, Free As Long, aLines, sName As String, Tell As Long
    
    ' Clear all variables and forms
    ClearAll
    
    ' Retrive a free file-number
    Free = FreeFile
    
    ' Firstly, load the file
    Open sFileName For Binary As #Free

        ' Allocate buffer
        sBuffer = Space(LOF(Free))
        
        ' Retrive the data
        Get #Free, , sBuffer

    Close #Free

    ' Now, split the file into several lines
    aLines = Split(sBuffer, vbCrLf)
    
    ' Go through all the lines
    For Tell = LBound(aLines) To UBound(aLines)
    
        ' See if this is the beginning of a new section
        If Mid(aLines(Tell), 1, 1) = "[" Then
    
            ' Get the name of this section
            sName = Mid(aLines(Tell), 2, Len(aLines(Tell)) - 2)
    
            ' See if this is the constant-section
            If LCase(sName) = "constants" Then
            
                ' Yup, add all constants
                Tell = AddConstants(aLines, Tell + 1) - 1
            
            Else
            
                ' If not, this is a form
                Tell = AddForm(aLines, sName, Tell + 1) - 1
                
            End If
    
        End If
    
    Next

End Sub

Public Sub ClearAll()

    ' Clears both arrays
    Erase aConstants
    Erase aForms
    
    ' And their respective counter
    ConstantCount = 0
    FormCount = 0

End Sub

Public Sub UseLanguage(Form As Object)

    Dim Index As Variant

    ' Firstly, set the index to use
    Index = Form.Name

    ' Then get the ID from the index
    GetIndex Index, aForms

    ' If the index is valid, execute the script
    If Index >= 0 Then
    
        ' Reinitialze script
        Script.Reset
        
        ' Add the code
        Script.AddCode aForms(Index).Value
        
        ' Execute the code
        Script.Run "UseLanguage", Form, ChatApp
    
    End If

End Sub

Private Function AddConstants(aLines, Position As Long) As Long
 
    Dim Tell As Long, aLine
     
    ' Add all constants until we reach the end of the section
    For Tell = Position To UBound(aLines)
    
        ' We cannot add a line that's empty
        If LenB(aLines(Tell)) <> 0 Then

            ' If the line begins with a certain character, it is time to end the search
            If Mid(aLines(Tell), 1, 1) = "[" Then
                Exit For
            End If

            ' Split the line into to segments
            aLine = Split(aLines(Tell), "=", 2)

            ' Reallocate the array
            ReDim Preserve aConstants(ConstantCount)

            ' Set the data of this element
            aConstants(ConstantCount).Name = Trim(aLine(0))
            aConstants(ConstantCount).Value = ParseVariable(CStr(aLine(1)))

            ' Increase the count
            ConstantCount = ConstantCount + 1

        End If
            
    Next
    
    ' Return the position of the search
    AddConstants = Tell

End Function

Private Function AddForm(aLines, sName As String, Position As Long) As Long

    Dim Tell As Long, sStart As String

    ' The string to start each line with
    sStart = "Form."

    ' Reallocate array
    ReDim Preserve aForms(FormCount)

    ' Set the name of this element
    aForms(FormCount).Name = sName

    ' Set the beginning value
    aForms(FormCount).Value = "Sub UseLanguage(Form, ChatApp)" & vbCrLf

    ' Add all lines until we reach the next section
    For Tell = Position To UBound(aLines)
    
        ' We cannot add a line that's empty
        If LenB(aLines(Tell)) <> 0 Then

            ' If the line begins with a certain character, it is time to end the search
            If Mid(aLines(Tell), 1, 1) = "[" Then
                Exit For
            End If
            
            ' Set the resulting string, depending on the beginning character
            If Mid(aLines(Tell), 1, 1) = "#" Then
                ' Add a 'real' line
                aForms(FormCount).Value = aForms(FormCount).Value & Mid(aLines(Tell), 2) & vbCrLf
            Else
                ' Add the shortcut-line
                aForms(FormCount).Value = aForms(FormCount).Value & sStart & aLines(Tell) & vbCrLf
            End If
            
        End If
        
    Next

    ' Append the ending
    aForms(FormCount).Value = aForms(FormCount).Value & "End Sub" & vbCrLf

    ' Incrase the form-count
    FormCount = FormCount + 1
    
    ' Return the position of the search
    AddForm = Tell

End Function

Private Function ParseVariable(sText As String) As Variant

    Dim Tmp As Long

    ' See if the text contains quotation marks
    Tmp = InStr(1, sText, Chr(34))

    ' If so, ...
    If Tmp <> 0 Then
        ' ... it's a string. Return the text between the quotation marks
        ParseVariable = Mid(sText, Tmp + 1, InStr(Tmp + 1, sText, Chr(34)) - Tmp - 1)
    Else
        ' If not, we've gotten a numeric value
        ParseVariable = Val(sText)
    End If

End Function

Private Sub GetIndex(Index As Variant, aArray() As Constant)

    On Error Resume Next
    Dim Tell As Long

    ' If the index is numeric, just return it, ...
    If IsNumeric(Index) Then
    
        ' ... after verifying that it's correct
        If Index >= LBound(aArray) Or Index <= UBound(aArray) Then
            Exit Sub
        End If
        
    Else
    
        ' Set the string as lowercase
        Index = LCase(Index)
    
        ' If not, search through it for the index-string
        For Tell = LBound(aArray) To UBound(aArray)
        
            ' Correspond without being case-sensitive
            If LCase(aArray(Tell).Name) = Index Then
            
                ' Return the index
                Index = Tell
                
                ' We're finish searching
                Exit Sub
            
            End If
        
        Next
    
    End If

    ' No element was found
    Index = -1

End Sub

Private Sub Class_Initialize()

    ' Create the script-engine
    Set Script = CreateObject("MSScriptControl.ScriptControl")
    
    ' Set the language
    Script.Language = "VBScript"

End Sub

