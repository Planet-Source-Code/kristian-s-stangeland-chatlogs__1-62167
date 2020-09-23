VERSION 5.00
Begin VB.UserControl usrChart 
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8100
   ScaleHeight     =   5895
   ScaleWidth      =   8100
   Begin VB.PictureBox picChart 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   0
      ScaleHeight     =   377
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   529
      TabIndex        =   0
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "usrChart"
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

' Public events
Public Event DataReallocated()
Public Event Redrawing()
Public Event Redrawed()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

' Public variables
Public MaxValue As Long

' Private variables
Private pColumnCount As Long
Private pDataX() As Variant
Private pDataY() As Variant

' A shortcut-sub
Public Sub SetData(Index, x, y)

    pDataX(Index) = x
    pDataY(Index) = y
    
End Sub

Public Property Get DataX(ByVal Index As Long) As Variant

    ' Return the value
    DataX = pDataX(Index)

End Property

Public Property Let DataX(ByVal Index As Long, ByVal vNewValue As Variant)

    ' Set the value
    pDataX(Index) = vNewValue

End Property

Public Property Get DataY(ByVal Index As Long) As Variant

    ' Return the value
    DataY = pDataY(Index)

End Property

Public Property Let DataY(ByVal Index As Long, ByVal vNewValue As Variant)

    ' Set the value
    pDataY(Index) = vNewValue

End Property

Public Property Get ColumnCount() As Variant

    ' Retrive the row-amout
    ColumnCount = pColumnCount

End Property

Public Property Let ColumnCount(ByVal vNewValue As Variant)

    ' If the new value differ from the already registered, reallocate array
    If vNewValue <> pColumnCount Then
    
        ' Set the new row count
        pColumnCount = vNewValue
    
        ' Reallocate array
        ReDim pDataY(1 To pColumnCount)
        ReDim pDataX(1 To pColumnCount)
        
        ' Inform about the reallocation
        RaiseEvent DataReallocated
        
        ' Redraw everything
        RedrawChart picChart
    
    End If

End Property

Private Sub UserControl_Initialize()

    ' Set the damn standard values
    MaxValue = 100
    ColumnCount = 10
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    ' Read the saved properties
    MaxValue = PropBag.ReadProperty("MaxValue", 100)
    ColumnCount = PropBag.ReadProperty("ColumnCount", 10)

End Sub

Private Sub UserControl_Resize()

    ' Resize the control after the usercontrol
    picChart.Width = UserControl.Width
    picChart.Height = UserControl.Height

    ' Simply redraw the picturebox
    RedrawChart picChart

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    ' Save properties
    PropBag.WriteProperty "MaxValue", MaxValue, 100
    PropBag.WriteProperty "ColumnCount", pColumnCount, 10

End Sub

Public Sub RedrawChart(Optional Control As Object)

    Dim lngRowWidth As Long, lngTextWidth As Long, lngTextHeight As Long, Tell As Long
    
    ' If a control is not reference, ...
    If Control Is Nothing Then
        ' ... use the default one
        Set Control = picChart
    End If
    
    ' Clear the control (if it's not a printer-class though)
    If Not (TypeOf Control Is Printer) Then
        Control.Cls
    End If
    
    ' Calulate the width of each row
    lngRowWidth = Control.ScaleWidth / pColumnCount
    
    ' Retrive the standard text-height
    lngTextHeight = Control.TextHeight("A")

    ' Write the column names
    For Tell = 1 To pColumnCount
    
        ' Set the font-color
        Control.ForeColor = &H8000&
    
        ' Then, get the lenght of the text to draw
        lngTextWidth = Control.TextWidth(pDataX(Tell))
    
        ' Set the position of this draw
        Control.CurrentX = (CDbl(Tell - 0.5) * lngRowWidth) - (lngTextWidth / 2)
        Control.CurrentY = Control.ScaleHeight - lngTextHeight - 8
        
        ' And finally write the text
        Control.Print pDataX(Tell)
    
    Next
    
    ' Then write the line
    For Tell = 2 To pColumnCount
    
        ' Draw the line from two points
        Control.Line ((Tell - 1.5) * lngRowWidth, ValueToPos(CLng(pDataY(Tell - 1)), Control))-((Tell - 0.5) * lngRowWidth, ValueToPos(CLng(pDataY(Tell)), Control)), vbRed
    
    Next

    ' And finally print the data amount, whereby avoiding all values of zero
    For Tell = 1 To pColumnCount
        
        ' Avoid numbers below or equal to zero
        If pDataY(Tell) > 0 Then
    
            ' Set the write-position
            Control.CurrentX = (Tell - 0.5) * lngRowWidth
            Control.CurrentY = ValueToPos(CLng(pDataY(Tell)), Control) - lngTextHeight
            
            ' Move downwards if we cannot see the number
            If Control.CurrentY < 0 Then
                Control.CurrentY = 0
            End If
            
            ' Set the color to print the numbers in
            Control.ForeColor = vbBlue
            
            ' And print the amount
            Control.Print pDataY(Tell)
    
        End If
    
    Next

End Sub

Public Function ValueToPos(lngValue As Long, Control As Object) As Long

    On Error Resume Next

    ' Calulcate the Y-position based on the given value
    ValueToPos = (Control.ScaleHeight - 32) * (1 - (lngValue / MaxValue))

End Function

Private Sub picChart_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Reraise the external event
    RaiseEvent MouseDown(Button, Shift, x, y)

End Sub

Private Sub picChart_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' As the above
    RaiseEvent MouseMove(Button, Shift, x, y)

End Sub

Private Sub picChart_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Likewise again
    RaiseEvent MouseUp(Button, Shift, x, y)

End Sub
