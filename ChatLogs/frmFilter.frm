VERSION 5.00
Begin VB.Form frmFilter 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Filters"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   389
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1440
      ScaleHeight     =   375
      ScaleWidth      =   4215
      TabIndex        =   1
      Top             =   4200
      Width           =   4215
      Begin VB.CommandButton cmdRemoveFilter 
         Caption         =   "&Remove Filter"
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddFilter 
         Caption         =   "&Add filter"
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   2880
         TabIndex        =   2
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.ListBox lstFilters 
      Height          =   3960
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

' The filter to use
Public currFilters As Collection

' Variable used in the wating-procedure
Private bWait As Boolean

Private Sub Form_Load()

    ' Initialize this form
    ChatApp.InitializeForm Me

End Sub

Public Sub UpdateFilters(ListBox As ListBox)

    Dim Filter As clsFilter
    
    ' Clear the listbox
    ListBox.Clear
    
    ' Go through all filters and add them to the listbox
    For Each Filter In currFilters
    
        ' Add the decription of this filter
        ListBox.AddItem ListBox.ListCount + 1 & ". " & Filter.ToString
    
    Next
    
End Sub

Public Sub WaitForFinish()

    ' Wait until the user is done
    bWait = True
    
    ' The loop used to wait
    Do While bWait
        DoEvents
        Sleep 10
    Loop

End Sub

Private Sub cmdAddFilter_Click()

    Dim Dialog As New clsDialog, Result As PropertyBag, Filter As clsFilter
    
    ' Set the reference
    Set Dialog.ReferenceForm = New frmAddFilter

    ' Show the dialog
    Set Result = Dialog.ShowDialog(New PropertyBag, LanguageConst("FilterCaption"))
    
    ' Firstly, check if the user has pressed OK
    If Result.ReadProperty("Returned") = "OK" Then
    
        ' Create a new filter
        Set Filter = New clsFilter
    
        ' Set the properties
        Filter.FilterType = Val(Result.ReadProperty("cmbFilterType", Filter_None - 1)) + 1
        Filter.ResultValue = CBool(Result.ReadProperty("cmbOperation", 0) = 0)
        Filter.UsePattern = CBool(Result.ReadProperty("chkUsePattern", 0) = 1)
        Filter.ValueOne = Result.ReadProperty("txtValueOne", "")
        Filter.ValueTwo = Result.ReadProperty("txtValueTwo", "")
        
        ' Add the filter to the collection
        currFilters.Add Filter
        
        ' Update the list
        UpdateFilters lstFilters
    
    End If
    
    ' Hide the reference form
    Dialog.ReferenceForm.Hide
    
    ' Remove the reference form
    Set Dialog.ReferenceForm = Nothing
    
    ' Remove the dialog-class
    Set Dialog = Nothing

End Sub

Private Sub cmdOK_Click()

    ' Hide this form
    Unload Me
    
End Sub

Private Sub cmdRemoveFilter_Click()

    ' If a element is in fact selected ...
    If lstFilters.ListIndex >= 0 Then
        
        ' ... remove it from the collection ...
        currFilters.Remove lstFilters.ListIndex + 1
        
        ' ... and update the listbox.
        UpdateFilters lstFilters
    
    End If

End Sub

Private Sub Form_Resize()

    ' Only resize when we can
    If Me.WindowState <> 1 Then

        ' Resize the listbox
        lstFilters.Width = Me.ScaleWidth - 16
        lstFilters.Height = Me.ScaleHeight - lstFilters.Top - picToolbar.Height - 16
        
        ' Set the position of the toolbar
        picToolbar.Left = Me.ScaleWidth - picToolbar.Width - 8
        picToolbar.Top = lstFilters.Top + lstFilters.Height + 8

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' We are finished witht the waiting
    bWait = False

End Sub
