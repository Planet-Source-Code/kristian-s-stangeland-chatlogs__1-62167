VERSION 5.00
Begin VB.Form frmTimeChart 
   Caption         =   "Time Chart"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   458
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   629
   Begin VB.ComboBox cmdPlot 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   6330
      Width           =   2295
   End
   Begin VB.Frame frameTimeChart 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin ChatLogs.usrChart usrChart 
         Height          =   5295
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   9340
      End
      Begin VB.Label lblTimeChart 
         Caption         =   "This is the timeline chart of your Instant Messaging usage"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   150
         TabIndex        =   3
         Top             =   270
         Width           =   4095
      End
   End
   Begin VB.Label lblViewPlotBy 
      Caption         =   "&View plot by:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   6390
      Width           =   2895
   End
End
Attribute VB_Name = "frmTimeChart"
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

Public ScriptEngine As Object
Public Sessions As Collection

Public Function Clone() As frmTimeChart

    ' Create a new form
    Set Clone = New frmTimeChart
    
    ' Initialize it ...
    Set Clone.Sessions = Sessions
    
    ' and finally show it
    Clone.Show

End Function

Private Sub Form_Resize()

    ' Resize the control and the frame
    frameTimeChart.Width = Me.ScaleWidth - 16
    frameTimeChart.Height = Me.ScaleHeight - cmdPlot.Height - 48
    usrChart.Width = (frameTimeChart.Width - 16) * Screen.TwipsPerPixelX
    usrChart.Height = ((frameTimeChart.Height - 16) * Screen.TwipsPerPixelY) - usrChart.Top

    ' Then set the position of the remaining controls
    cmdPlot.Top = frameTimeChart.Top + frameTimeChart.Height + 16
    lblViewPlotBy.Top = cmdPlot.Top + 2

End Sub

Private Sub Form_Load()

    ' Initialize the script-engine
    Set ScriptEngine = CreateObject("MSScriptControl.ScriptControl")

    ' Update our combobox
    UpdateCombobox cmdPlot
    
    ' Initialize this form
    ChatApp.InitializeForm Me

End Sub

Private Sub cmdPlot_Change()

    ' Execute the code below
    cmdPlot_Change

End Sub

Private Sub cmdPlot_Click()

    Dim sData As String, Free As Long

    ' Reset the engine
    ScriptEngine.Language = "VBScript"
    ScriptEngine.Reset
    
    ' Get a free file-handle
    Free = FreeFile
    
    ' Get the content of the code-file
    Open FileOperation.ValidPath(App.Path) & "Data\TimeChart\" & cmdPlot.Text & ".vbs" For Binary As #Free
        
        ' Allocate buffer
        sData = Space(LOF(Free))
        
        ' Retrive data from file
        Get #Free, , sData
    
    Close #Free
    
    ' Then, initialize the code from the code-file
    ScriptEngine.AddCode sData
    
    ' And execute the procedure
    ScriptEngine.Run "UpdateControl", usrChart, Sessions
    
End Sub

Public Sub UpdateCombobox(refCombobox As ComboBox)

    Dim File As Variant

    ' Firstly, clear it
    refCombobox.Clear
    
    ' Then add each vbs-script in the TimeChart-folder to it
    For Each File In FileOperation.RetriveFileList(FileOperation.ValidPath(App.Path) & "Data\TimeChart\", "*.vbs", False, vbNormal)

        ' Add the file-name without the extension nor the path
        refCombobox.AddItem FileOperation.GetNoExtension(FileOperation.GetFileName(CStr(File)))

    Next
    
    ' Select the first element
    refCombobox.ListIndex = 0
    
End Sub

Private Sub usrChart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' If the user right-click on the chart ...
    If Button = 2 Then
    
        ' Save the control the procedure should call/print
        Set frmMain.MenuControl = usrChart
    
        ' ... invoke the associated popup menu
        Me.PopupMenu frmMain.mnuListView
    
    End If

End Sub
