Attribute VB_Name = "modGlobal"
Option Explicit

Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function UrlUnescape Lib "shlwapi" Alias "UrlUnescapeA" (ByVal pszURL As String, ByVal pszUnescaped As String, pcchUnescaped As Long, ByVal dwFlags As Long) As Long

'escape #'s in paths
Enum URL_FLAGS
    URL_INTERNAL_PATH = &H800000
    URL_DONT_ESCAPE_EXTRA_INFO = &H2000000
    URL_ESCAPE_SPACES_ONLY = &H4000000
    URL_DONT_SIMPLIFY = &H8000000
End Enum

Public Const MAX_PATH = 260
Public Const ERROR_SUCCESS = 0

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

' Our main object
Public MainObject As Object

' Used primary to convert UTF-8 to Unicode
Public Function MBToUni(bySrc() As Byte, lCP As Long) As String

    Dim lBytes As Long, lNC As Long, lRet As Long
    
    ' Calculate the amount of bytes in this array
    lBytes = UBound(bySrc) - LBound(bySrc) + 1
    
    ' Copy the amount of bytes
    lNC = lBytes
    
    ' Now, allocate the resulting string
    MBToUni = String$(lNC, Chr(0))
    
    ' Convert the byte array to a unicode-string
    lRet = MultiByteToWideChar(lCP, 0, VarPtr(bySrc(LBound(bySrc))), lBytes, StrPtr(MBToUni), lNC)
    
    ' Chop of the string, only returning what's necessary
    MBToUni = Left$(MBToUni, lRet)
    
End Function

' Used to decode an URL-string
Public Function DecodeUrl(ByVal sUrl As String, Optional Flags As URL_FLAGS = URL_DONT_SIMPLIFY) As String

    Dim buff As String
    Dim dwSize As Long
    
    ' Make sure the string isn't empty
    If Len(sUrl) > 0 Then
       
        ' Allocate a buffer and save the lenght of the string
        buff = Space$(MAX_PATH)
        dwSize = Len(buff)
        
        ' Decode the URL
        If UrlUnescape(sUrl, buff, dwSize, Flags) = ERROR_SUCCESS Then
        
            ' Return the resulting string
            DecodeUrl = Left$(buff, dwSize)
        
        End If
       
    End If

End Function

' Load a file
Public Function LoadFile(sFile As String) As String

    Dim Free As Long
    
    ' Get a free handle
    Free = FreeFile
    
    ' Open the file, ...
    Open sFile For Binary As #Free
    
        ' Allocate the buffer
        LoadFile = Space(LOF(Free))
        
        ' Retrive the datA
        Get #Free, , LoadFile
    
    Close #Free

End Function
