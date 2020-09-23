VERSION 5.00
Begin VB.UserControl usrProgressbar 
   ClientHeight    =   2625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6720
   ScaleHeight     =   2625
   ScaleWidth      =   6720
   ToolboxBitmap   =   "usrProgressbar.ctx":0000
End
Attribute VB_Name = "usrProgressbar"
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

' API-calls needed
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' The different progress styles
Private Enum ProgressStyle
    PBS_SMOOTH = &H1
    PBS_VERTICAL = &H4
    PBS_MARQUEE = &H8
End Enum

' Contants needed
Const WM_PAINT = &HF
Const WM_USER = &H400
Const WS_CHILD = &H40000000
Const WS_VISIBLE = &H10000000

' Commands to send to the progress
Const PBM_SETRANGE32 = (WM_USER + 6)
Const PBM_GETRANGE = (WM_USER + 7)
Const PBM_SETPOS As Long = (WM_USER + 2)
Const PBM_GETPOS As Long = (WM_USER + 8)
Const PBM_SETBARCOLOR = (WM_USER + 9)
Const PBM_SETMARQUEE = WM_USER + 10

' Used to change the style of the progressbar
Const GWL_STYLE As Long = (-16)

' Default color
Const CLR_DEFAULT = -16777216

' The window handle to the progress bar
Private hWndPB As Long

' Different color-values
Private pBarColor As Long
Private pMarquee As Boolean
Private pTimer As Long
Private bSmooth As Boolean
Private bVertical As Boolean

Private Sub UpdateStyle()

    ' If the control exists
    If hWndPB <> 0 Then
        ' Set the style of this progressbar
        SetWindowLong hWndPB, GWL_STYLE, IIf(bSmooth, PBS_SMOOTH, 0) Or IIf(bVertical, PBS_VERTICAL, 0) Or IIf(pMarquee, PBS_MARQUEE, 0) Or WS_CHILD Or WS_VISIBLE
    End If

    ' Update the style
    UserControl_Paint

End Sub

Private Sub UpdateMarquee()
    
    ' Firstly, update the style
    UpdateStyle
    
    ' If the control exists
    If hWndPB <> 0 Then
        ' Update the marquee
        SendMessage hWndPB, PBM_SETMARQUEE, IIf(pMarquee, 1, 0), ByVal pTimer
    End If

End Sub

Public Property Get Vertical() As Boolean

    ' Get the window style
    Vertical = bVertical
    
End Property

Public Property Let Vertical(ByVal vNewValue As Boolean)
    
    ' Save this setting
    bVertical = vNewValue
    
    ' Update the style
    UpdateStyle
    
End Property

Public Property Get Smooth() As Boolean

    ' Get the window style
    Smooth = bSmooth
    
End Property

Public Property Let Smooth(ByVal vNewValue As Boolean)
    
    ' Save this setting
    bSmooth = vNewValue
    
    ' Update the style
    UpdateStyle
    
End Property

Public Property Get MarqueeEnabled() As Boolean

    ' Return whether or not the marquee is enabled
    MarqueeEnabled = pMarquee

End Property

Public Property Let MarqueeEnabled(ByVal vNewValue As Boolean)

    ' Save the setting
    pMarquee = vNewValue
    
    ' Set the marquee
    UpdateMarquee
    
End Property

Public Property Get MarqueeTimer() As Long

    ' Return the update time of the marquee
    MarqueeTimer = pTimer

End Property

Public Property Let MarqueeTimer(ByVal vNewValue As Long)

    ' Save the setting
    pTimer = vNewValue

    ' Set the marquee
    UpdateMarquee

End Property

Public Property Get BarColor() As Long

    ' Return the barcolor
    BarColor = pBarColor

End Property

Public Property Let BarColor(ByVal vNewValue As Long)

    ' If the control exists
    If hWndPB <> 0 Then
        ' Set the barcolor
        SendMessage hWndPB, PBM_SETBARCOLOR, 0, ByVal vNewValue
    
        ' Save the barcolor for later use
        pBarColor = vNewValue
    End If

End Property

Public Property Get Min() As Long

    ' If the control exists
    If hWndPB <> 0 Then
        ' Get and return the minimun range
        Min = SendMessage(hWndPB, PBM_GETRANGE, 1, ByVal 0&)
    End If

End Property

Public Property Let Min(ByVal vNewValue As Long)

    ' If the control exists
    If hWndPB <> 0 Then
        ' Set the min-property
        SendMessage hWndPB, PBM_SETRANGE32, vNewValue, ByVal Max
    End If

End Property

Public Property Get Max() As Long

    ' If the control exists
    If hWndPB <> 0 Then
        ' Get and return the maximum range
        Max = SendMessage(hWndPB, PBM_GETRANGE, 0, ByVal 0&)
    End If

End Property

Public Property Let Max(ByVal vNewValue As Long)

    ' If the control exists
    If hWndPB <> 0 Then
        ' Set the min-property
        SendMessage hWndPB, PBM_SETRANGE32, Min, ByVal vNewValue
    End If

End Property

Public Property Get Value() As Long

    ' If the control exists
    If hWndPB <> 0 Then
        ' Return the value
        Value = SendMessage(hWndPB, PBM_GETPOS, 0, ByVal 0&)
    End If

End Property

Public Property Let Value(ByVal vNewValue As Long)

    ' If the control exists
    If hWndPB <> 0 Then
        ' Set the value
        SendMessage hWndPB, PBM_SETPOS, vNewValue, ByVal 0&
    End If

End Property

Private Sub UserControl_Initialize()

    ' Firstly, create our window
    hWndPB = CreateWindowEx(0&, "msctls_progress32", UserControl.Name, WS_CHILD Or WS_VISIBLE, _
     0, 0, UserControl.ScaleWidth / Screen.TwipsPerPixelX, UserControl.ScaleHeight / _
      Screen.TwipsPerPixelY, UserControl.hwnd, 0, App.hInstance, ByVal 0&)

End Sub

Private Sub UserControl_Paint()

    ' If the control exists
    If hWndPB <> 0 Then
        ' Paint the control
        SendMessage hWndPB, WM_PAINT, UserControl.hDC, ByVal 0&
    End If

End Sub

Private Sub UserControl_Resize()
    
    ' If the control exists
    If hWndPB <> 0 Then
        ' Resize the window
        MoveWindow hWndPB, 0, 0, UserControl.ScaleWidth / Screen.TwipsPerPixelX, _
         UserControl.ScaleHeight / Screen.TwipsPerPixelY, 1
    End If

End Sub

Private Sub UserControl_Terminate()

    ' Destroy our progressbar
    DestroyWindow hWndPB

    ' Indicate that the window is actually destroyed
    hWndPB = 0

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    ' Read the properties
    Max = PropBag.ReadProperty("Max", 100)
    Min = PropBag.ReadProperty("Min", 0)
    Value = PropBag.ReadProperty("Value", 0)
    BarColor = PropBag.ReadProperty("Barcolor", CLR_DEFAULT)
    MarqueeEnabled = PropBag.ReadProperty("MarqueeEnabled", False)
    MarqueeTimer = PropBag.ReadProperty("MarqueeTimer", 0)
    Smooth = PropBag.ReadProperty("Smooth", False)
    Vertical = PropBag.ReadProperty("Vertical", False)
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    ' These properties can only be saved if the progressbar exists
    If hWndPB <> 0 Then
        PropBag.WriteProperty "Max", Max, 100
        PropBag.WriteProperty "Min", Min, 0
        PropBag.WriteProperty "Value", Value, 0
    End If

    ' Write the rest
    PropBag.WriteProperty "BarColor", BarColor, CLR_DEFAULT
    PropBag.WriteProperty "MarqueeEnabled", MarqueeEnabled, False
    PropBag.WriteProperty "MarqueeTimer", MarqueeTimer, 0
    PropBag.WriteProperty "Smooth", Smooth, False
    PropBag.WriteProperty "Vertical", Vertical, False

End Sub


