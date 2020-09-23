VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrint"
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

' Default values used in the printing process
Public MargineLeft As Long
Public MargineTop As Long

Public Sub PrintControl(Control As Object)

    Dim ColumnWidth() As Long, TotalWidth As Long, Tell As Long, X As Long, Y As Long
    Dim sColumn As String, Item As cListItem, CellHeight As Long, lngSub As Long

    ' Then, see what type this control is
    If TypeOf Control Is vbalListViewCtl Then
    
        ' Initialize the start-values
        X = MargineLeft
        Y = MargineTop
        
        ' Set the font-name and -size, including draw size
        Printer.fontname = "MS Sans Serif"
        Printer.FontSize = 9
    
        ' Get the height of a cell
        CellHeight = Printer.TextHeight("A") + 6
        
        ' Reallocate array
        ReDim ColumnWidth(1 To Control.Columns.Count)
    
        ' Get the total width of the table, plus each individual width
        For Tell = 1 To Control.Columns.Count
        
            ' Save the width of this column
            ColumnWidth(Tell) = Control.Columns(Tell).Width / 3
         
            ' Add the width of this column
            TotalWidth = TotalWidth + ColumnWidth(Tell)
         
        Next
        
        ' Draw the columns
        For Tell = 1 To Control.Columns.Count
        
            ' Retrive the name of this column
            sColumn = Control.Columns(Tell).Text
            
            ' Print the box around the column
            Printer.FillStyle = 0
            Printer.FillColor = RGB(200, 200, 200)
            Printer.DrawWidth = 4
            Printer.Line (X, Y)-(X + ColumnWidth(Tell), Y + CellHeight), vbBlack, B
        
            ' Set the position of where to draw the text
            Printer.CurrentX = X + (ColumnWidth(Tell) / 2) - (Printer.TextWidth(sColumn) / 2)
            Printer.CurrentY = Y + (6 * Printer.TwipsPerPixelY)
            
            ' Draw the text
            Printer.Print sColumn
            
            ' Increase the X-position
            X = X + ColumnWidth(Tell)
        
        Next
        
        ' Increase the variable and reset X-position
        Y = Y + CellHeight
        X = MargineLeft
    
        ' Then draw each cell, row for row
        For Tell = 1 To Control.ListItems.Count
        
            ' Retrive the item to draw
            Set Item = Control.ListItems(Tell)
        
            ' Draw the left-line
            Printer.DrawWidth = 4
            Printer.Line (X, Y)-(X, Y + CellHeight), vbBlack
        
            ' Then draw the first column
            DrawCell X + (12 * Printer.TwipsPerPixelX), Y + (6 * Printer.TwipsPerPixelY), Item.Text, ColumnWidth(1)
            
            ' Increase X-position
            X = X + ColumnWidth(1)
            
            ' Draw all subelements
            For lngSub = 1 To Item.SubItems.Count
  
                ' Draw the line downwards
                Printer.DrawWidth = 2
                Printer.Line (X, Y)-(X, Y + CellHeight), RGB(200, 200, 200)
                
                ' Draw the subelement-cell
                DrawCell X + (12 * Printer.TwipsPerPixelX), Y + (6 * Printer.TwipsPerPixelY), Item.SubItems(lngSub).Caption, ColumnWidth(lngSub + 1)
            
                ' Increase X-position
                X = X + ColumnWidth(lngSub + 1)

            Next
            
            ' Draw the right-line
            Printer.DrawWidth = 4
            Printer.Line (TotalWidth + MargineLeft, Y)-(TotalWidth + MargineLeft, Y + CellHeight), vbBlack
            
            ' Draw the underline
            If Tell = Control.ListItems.Count Then
                ' Draw a double-line with another color
                Printer.DrawWidth = 4
                Printer.Line (MargineLeft, Y + CellHeight)-(TotalWidth + MargineLeft, Y + CellHeight), vbBlack
            Else
                ' Draw a single line
                Printer.DrawWidth = 2
                Printer.Line (MargineLeft, Y + CellHeight)-(TotalWidth + MargineLeft, Y + CellHeight), RGB(200, 200, 200)
            End If
        
            ' Clear both variables
            Y = Y + CellHeight
            X = MargineLeft
        
        Next
    
    ElseIf TypeOf Control Is usrChart Then
        
        ' Use the landscape-orientation
        Printer.Orientation = vbPRORLandscape
    
        ' Print the chart
        Control.RedrawChart Printer
        
    End If
    
    ' End the drawing-process and start printing
    Printer.EndDoc
    
End Sub

Private Sub DrawCell(X As Long, Y As Long, sText As String, MaxWidth As Long)

    Dim Tell As Long

    ' Move to the correct position ...
    Printer.CurrentX = X
    Printer.CurrentY = Y
    
    ' ... and draw the first part of the text that fits within the maximum width
    For Tell = 1 To Len(sText)
        
        ' If the width of the text exceeds the maximum, print the last amount
        If Printer.TextWidth(Mid(sText, 1, Tell)) > MaxWidth Then
            
            ' Print the text
            Printer.Print Mid(sText, 1, Tell - 1)
            
            ' We're finish printing
            Exit Sub
            
        End If
    
    Next
    
    ' Print the entire text insted
    Printer.Print sText
    
End Sub

Private Sub Class_Initialize()

    ' Initialize the to variables
    MargineLeft = 450
    MargineTop = 550

End Sub
