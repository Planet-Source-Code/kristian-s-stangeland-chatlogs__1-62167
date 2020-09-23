VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommonDialogs"
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

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pShowOpenname As OpenFileNAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pShowOpenname As OpenFileNAME) As Long
Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Function PAGESETUPDLG Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PAGESETUPDLG) As Long
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Const FW_NORMAL = 400
Const DEFAULT_CHARSET = 1
Const OUT_DEFAULT_PRECIS = 0
Const CLIP_DEFAULT_PRECIS = 0
Const DEFAULT_QUALITY = 0
Const DEFAULT_PITCH = 0
Const FF_ROMAN = 16
Const CF_PRINTERFONTS = &H2
Const CF_SCREENFONTS = &H1
Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Const CF_EFFECTS = &H100&
Const CF_FORCEFONTEXIST = &H10000
Const CF_INITTOLOGFONTSTRUCT = &H40&
Const CF_LIMITSIZE = &H2000&
Const REGULAR_FONTTYPE = &H400
Const LF_FACESIZE = 32
Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32
Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40
Const DM_DUPLEX = &H1000&
Const DM_ORIENTATION = &H1&
Const PD_PRINTSETUP = &H40
Const PD_DISABLEPRINTTOFILE = &H80000

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type OpenFileNAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type PAGESETUPDLG
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    Flags As Long
    ptPaperSize As POINTAPI
    rtMinMargin As RECT
    rtMargin As RECT
    hInstance As Long
    lCustData As Long
    lpfnPageSetupHook As Long
    lpfnPagePaintHook As Long
    lpPageSetupTemplateName As String
    hPageSetupTemplate As Long
End Type

Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 31
End Type

Private Type CHOOSEFONT
    lStructSize As Long
    hwndOwner As Long
    hDC As Long
    lpLogFont As Long
    iPointSize As Long
    Flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

Private Type PRINTDLG_TYPE
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    Flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Private Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type

Private Type DEVMODE_TYPE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Enum OFN_FLAGS
    OFN_ALLOWMULTISELECT = &H200
    OFN_CREATEPROMPT = &H2000
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_EXPLORER = &H80000
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_FILEMUSTEXIST = &H1000
    OFN_HIDEREADONLY = &H4
    OFN_LONGNAMES = &H200000
    OFN_NOCHANGEDIR = &H8
    OFN_NODEREFERENCELINKS = &H100000
    OFN_NOLONGNAMES = &H40000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NOVALIDATE = &H100
    OFN_OVERWRITEPROMPT = &H2
    OFN_READONLY = &H1
    OFN_SHAREAWARE = &H4000
    OFN_SHAREFALLTHROUGH = 2
    OFN_SHARENOWARN = 1
    OFN_SHAREWARN = 0
    OFN_SHOWHELP = &H10
End Enum

Private SaveFileDialog As OpenFileNAME
Private OpenFileDialog As OpenFileNAME
Private CustomColors() As Byte

Private sv As Long
Private lFilter As String
Private lMaxFile As Long
Private lFlags As Long
Private lFile As String

Public Sub ShowOpen(hWnd As Long, Caption As String)
   
    Dim Ret As Long
   
    ' Initialize the type
    With OpenFileDialog
        .lStructSize = Len(OpenFileDialog)
        .hwndOwner = hWnd
        .hInstance = App.hInstance
        .lpstrFilter = lFilter
        .lpstrFile = Space$(lMaxFile - 1)
        .nMaxFile = lMaxFile
        .lpstrFileTitle = Space$(lMaxFile - 1)
        .nMaxFileTitle = 255
        .lpstrInitialDir = App.Path
        .lpstrTitle = Caption
        .Flags = lFlags
    End With
    
    ' Show the dialog and wait for input
    Ret = GetOpenFileName(OpenFileDialog)
    
    ' See whether or not the user has pressed cancel
    If Ret Then
        lFile = Trim$(OpenFileDialog.lpstrFile)
        lFile = Mid(lFile, 1, Len(lFile) - 1)
    Else
        lFile = ""
    End If

End Sub

Public Sub ShowSave(hWnd As Long, Caption As String)
   
    Dim Ret As Long
   
    ' Initialize the type
    With SaveFileDialog
        .lStructSize = Len(SaveFileDialog)
        .hwndOwner = hWnd
        .hInstance = App.hInstance
        .lpstrFilter = lFilter
        .lpstrFile = Space$(lMaxFile - 1)
        .nMaxFile = lMaxFile
        .lpstrFileTitle = Space$(lMaxFile - 1)
        .nMaxFileTitle = 255
        .lpstrInitialDir = App.Path
        .lpstrTitle = Caption
        .Flags = lFlags
    End With
    
    ' Show the dialog and wait for input
    Ret = GetSaveFileName(SaveFileDialog)
    
    ' See whether or not the user has pressed cancel
    If Ret Then
        lFile = Trim$(SaveFileDialog.lpstrFile)
        lFile = Mid(lFile, 1, Len(lFile) - 1)
    Else
        lFile = ""
    End If

End Sub

Public Function ShowColor(Optional hWnd As Long) As Long

    Dim cc As CHOOSECOLOR
    Dim Custcolor(16) As Long
    Dim lReturn As Long

    'set the structure size
    cc.lStructSize = Len(cc)
    'Set the owner
    cc.hwndOwner = hWnd
    'set the application's instance
    cc.hInstance = App.hInstance
    'set the custom colors (converted to Unicode)
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    'no extra flags
    cc.Flags = 0

    'Show the 'Select Color'-dialog
    If CHOOSECOLOR(cc) <> 0 Then
        ShowColor = cc.rgbResult
        CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
    Else
        ShowColor = -1
    End If
    
End Function

Public Function ShowFont(Optional hWnd As Long) As String

    Dim cf As CHOOSEFONT, lfont As LOGFONT, hMem As Long, pMem As Long
    Dim fontname As String, retval As Long
    
    With lfont
        .lfHeight = 0  ' determine default height
        .lfWidth = 0  ' determine default width
        .lfEscapement = 0  ' angle between baseline and escapement vector
        .lfOrientation = 0  ' angle between baseline and orientation vector
        .lfWeight = FW_NORMAL  ' normal weight i.e. not bold
        .lfCharSet = DEFAULT_CHARSET  ' use default character set
        .lfOutPrecision = OUT_DEFAULT_PRECIS  ' default precision mapping
        .lfClipPrecision = CLIP_DEFAULT_PRECIS  ' default clipping precision
        .lfQuality = DEFAULT_QUALITY  ' default quality setting
        .lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN  ' default pitch, proportional with serifs
        .lfFaceName = "Times New Roman" & vbNullChar  ' string must be null-terminated
    End With
    
    ' Create the memory block which will act as the LOGFONT structure buffer.
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lfont))
    pMem = GlobalLock(hMem)  ' lock and get pointer
    CopyMemory ByVal pMem, lfont, Len(lfont)  ' copy structure's contents into block
    
    ' Initialize dialog box: Screen and printer fonts, point size between 10 and 72.
    With cf
        .lStructSize = Len(cf)  ' size of structure
        .hwndOwner = hWnd  ' The parent window
        .hDC = Printer.hDC  ' device context of default printer (using VB's mechanism)
        .lpLogFont = pMem   ' pointer to LOGFONT memory block buffer
        .iPointSize = 120  ' 12 point font (in units of 1/10 point)
        .Flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
        .rgbColors = RGB(0, 0, 0)  ' black
        .nFontType = REGULAR_FONTTYPE  ' regular font type i.e. not bold or anything
        .nSizeMin = 10  ' minimum point size
        .nSizeMax = 72  ' maximum point size
    End With
    
    ' Now, call the function.  If successful, copy the LOGFONT structure back into the structure
    ' and then print out the attributes we mentioned earlier that the user selected.
    retval = CHOOSEFONT(cf)  ' Open the dialog box
    
    If retval <> 0 Then  ' Success
        CopyMemory lfont, ByVal pMem, Len(lfont)  ' copy memory back
        ' Now make the fixed-length string holding the font name into a "normal" string.
        ShowFont = Left(lfont.lfFaceName, InStr(lfont.lfFaceName, vbNullChar) - 1)
    End If
    
    ' Deallocate the memory block we created earlier.  Note that this must
    ' be done whether the function succeeded or not.
    retval = GlobalUnlock(hMem)  ' destroy pointer, unlock block
    retval = GlobalFree(hMem)  ' free the allocated memory
    
End Function

Public Function ShowPageSetupDlg(Optional hWnd As Long) As Long

    Dim m_PSD As PAGESETUPDLG
    
    ' Set the structure size
    m_PSD.lStructSize = Len(m_PSD)
    ' Set the owner window
    m_PSD.hwndOwner = hWnd
    ' Set the application instance
    m_PSD.hInstance = App.hInstance
    ' No extra flags
    m_PSD.Flags = 0

    ' Show the pagesetup dialog
    If PAGESETUPDLG(m_PSD) Then
        ShowPageSetupDlg = 0
    Else
        ShowPageSetupDlg = -1
    End If
    
End Function

Public Function ShowPrinter(frmOwner As Object, Optional PrintFlags As Long) As Long

    '-> Code by Donald Grover
    Dim PrintDlg As PRINTDLG_TYPE
    Dim DevMode As DEVMODE_TYPE
    Dim DevName As DEVNAMES_TYPE

    Dim lpDevMode As Long, lpDevName As Long
    Dim bReturn As Integer
    Dim objPrinter As Printer, NewPrinterName As String

    ' Use PrintDialog to get the handle to a memory
    ' block with a DevMode and DevName structures
    PrintDlg.lStructSize = Len(PrintDlg)
    PrintDlg.hwndOwner = frmOwner.hWnd
    PrintDlg.Flags = PrintFlags
    
    On Error Resume Next
    'Set the current orientation and duplex setting
    DevMode.dmDeviceName = Printer.DeviceName
    DevMode.dmSize = Len(DevMode)
    DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
    DevMode.dmPaperWidth = Printer.Width
    DevMode.dmOrientation = Printer.Orientation
    DevMode.dmPaperSize = Printer.PaperSize
    DevMode.dmDuplex = Printer.Duplex
    On Error GoTo 0

    'Allocate memory for the initialization hDevMode structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
    End If

    'Set the current driver, device, and port name strings
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With

    With Printer
        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With

    ' Allocate memory for the initial hDevName structure
    ' and copy the settings gathered above into this memory
    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If

    ' Call the print dialog up and let the user make changes
    If PrintDialog(PrintDlg) <> 0 Then

        ' First get the DevName structure.
        lpDevName = GlobalLock(PrintDlg.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree PrintDlg.hDevNames

        'Next get the DevMode structure and set the printer properties appropriately
        lpDevMode = GlobalLock(PrintDlg.hDevMode)
        CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
        GlobalFree PrintDlg.hDevMode
        NewPrinterName = UCase$(Left(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
        
        If Printer.DeviceName <> NewPrinterName Then
            For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                    Set Printer = objPrinter
                    'set printer toolbar name at this point
                End If
            Next
        End If

        On Error Resume Next
        'Set printer object properties according to selections made
        'by user
        Printer.Copies = DevMode.dmCopies
        Printer.Duplex = DevMode.dmDuplex
        Printer.Orientation = DevMode.dmOrientation
        Printer.PaperSize = DevMode.dmPaperSize
        Printer.PrintQuality = DevMode.dmPrintQuality
        Printer.ColorMode = DevMode.dmColor
        Printer.PaperBin = DevMode.dmDefaultSource
        On Error GoTo 0
        
    Else
        
        ' The user canceled the printing
        ShowPrinter = -1
        
    End If
    
End Function

Public Property Get MaxFile() As Long

    ' Return the size
    MaxFile = lMaxFile

End Property

Public Property Let MaxFile(ByVal vNewValue As Long)

    ' The size of the buffer used to return the file-names.
    ' Increase this if you're planning on retriving a lot of files.
    lMaxFile = vNewValue

End Property

Public Property Get Filter() As String

    ' Return the filter
    Filter = lFilter

End Property

Public Property Let Filter(ByVal vNewValue As String)

    ' Set the filter (delimitered by null)
    lFilter = vNewValue

End Property

Public Property Get Flags() As OFN_FLAGS

    ' Return the flags
    Flags = lFlags

End Property

Public Property Let Flags(ByVal vNewValue As OFN_FLAGS)

    ' Set the flags
    lFlags = vNewValue

End Property

Public Property Get File() As String
    
    ' Return the file selected
    File = lFile
    
End Property

Private Sub Class_Initialize()

    ' Set default value
    lMaxFile = 255

    ' Redim the variables to store the cutstom colors
    ReDim CustomColors(0 To 16 * 4 - 1) As Byte

End Sub
