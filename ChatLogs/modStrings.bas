Attribute VB_Name = "modStrings"
Option Explicit

' Used to optimize string-reading
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

Public Const SmallA As Integer = 97
Public Const SmallZ As Integer = 122
Public Const LargeA As Integer = 65
Public Const LargeZ As Integer = 90

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

' Used to split a string with the exception of data enclosed by quotation marks
Public Function SplitX(Text As String, Delimiter As String) As Variant

    Dim Tell As Long, Last As Long, Arr() As String, Cnt As Long
    
    ' If no text is specified, don't proceed
    If Text = "" Then
        Exit Function
    End If
    
    ' Default value
    Last = 1
    
    ' Go on searching
    Do Until Tell >= Len(Text)
        
        ' Find the next delimiter, while watching for quotation marks
        Tell = InStrX(Last, Text, Delimiter)
        
        ' If nothing was found, assume we've reached the end
        If Tell = 0 Then
            Tell = Len(Text) + 1
        End If
        
        ' Reallocate the result-array
        ReDim Preserve Arr(Cnt)
        
        ' Set the latest element in the newly allocated array
        Arr(Cnt) = Mid(Text, Last, Tell - Last)
        
        ' Move forward
        Cnt = Cnt + 1
        Last = Tell + 1
    Loop
    
    ' Return what we've found
    SplitX = Arr

End Function

' Used to search for a character in a string, while avoiding occurrences that is within two quotation marks
Public Function InStrX(ByVal Begin As Integer, Str As Variant, Optional SearchFor As String = " ") As Integer

    Dim Tell As Long, Buff As String, OneChar As String, DontLook As Boolean
    
    ' Go through all characters in the string from a certain start-position
    For Tell = Begin To Len(Str)
    
        ' Retrive the characters from this point
        Buff = Mid(Str, Tell, Len(SearchFor))
        
        ' Get one character from the code above
        OneChar = Mid(Buff, 1, 1)
        
        ' Now, if this is a quotation mark, turn on/off searching
        If OneChar = """" Then
            DontLook = Not DontLook
        End If
        
        ' If we've found what we're looking for, and it's not enclosed by quotation marks ...
        If (Not DontLook) And Buff = SearchFor Then
            
            ' ... return the position ...
            InStrX = Tell
            
            ' ... and exit the prosedure.
            Exit Function
            
        End If
        
    Next

End Function
