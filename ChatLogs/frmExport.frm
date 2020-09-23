VERSION 5.00
Begin VB.Form frmExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Dialog"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&l"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Frame frameLocation 
      Caption         =   "Lo&cation:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   6495
      Begin VB.PictureBox picLocation 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   240
         ScaleHeight     =   855
         ScaleWidth      =   6135
         TabIndex        =   4
         Top             =   480
         Width           =   6135
         Begin VB.TextBox txtPath 
            Height          =   285
            Left            =   1080
            TabIndex        =   7
            Top             =   0
            Width           =   3735
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "&Browse"
            Height          =   270
            Left            =   4920
            TabIndex        =   6
            Top             =   0
            Width           =   1095
         End
         Begin VB.ComboBox cmbFileType 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   480
            Width           =   4935
         End
         Begin VB.Label lblFolder 
            Caption         =   "F&older: "
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   30
            Width           =   1095
         End
         Begin VB.Label lblFileType 
            Caption         =   "&File type:"
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   510
            Width           =   1215
         End
      End
   End
   Begin VB.Frame frameArchiving 
      Caption         =   "&Archiving:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   6495
      Begin VB.PictureBox picArchiving 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   240
         ScaleHeight     =   975
         ScaleWidth      =   6015
         TabIndex        =   10
         Top             =   480
         Width           =   6015
         Begin VB.CheckBox chkYears 
            Caption         =   "A&rchive by years"
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   5895
         End
         Begin VB.CheckBox chkMonths 
            Caption         =   "Arc&hive by months"
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   360
            Width           =   5895
         End
         Begin VB.CheckBox chkWeeks 
            Caption         =   "&Archi&ve by weeks"
            Height          =   255
            Left            =   0
            TabIndex        =   11
            Top             =   720
            Width           =   5895
         End
      End
   End
End
Attribute VB_Name = "frmExport"
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

Private Sub Form_Load()

    ' Initialize this form
    ChatApp.InitializeForm Me

End Sub

Public Sub Update()

    On Error Resume Next
    Dim Export As Object
    
    ' Clear the combobox
    cmbFileType.Clear
    
    ' Loop through all export modules
    For Each Export In ChatApp.ExportModules
    
        ' Only include the module if it dosen't handle all by itself
        If Not Export.ExportHandleAll Then
    
            ' Add the name of the export module
            cmbFileType.AddItem Export.ExportName
        
        End If
    
    Next
    
    ' Move to the first element
    cmbFileType.ListIndex = 0

End Sub

Private Function WeekNum(dDate As Date) As Long

    ' Calculate the amount of weeks since newyear of a certain date
    WeekNum = DateDiff("ww", DateSerial(Year(dDate), 1, 1), dDate) + 1

End Function

Private Sub cmdBrowse_Click()

    Dim Browse As New clsFolderBrowse
    
    ' Set the description
    Browse.Description = LanguageConst("ChooseExport")
    
    ' Browse for a folder
    If Browse.Browse(False) = Browse_OK Then
    
        ' Set the result path
        txtPath.Text = Browse.ResultPath
    
    End If

End Sub

Private Sub cmdCancel_Click()

    ' Hide form
    Me.Hide

End Sub

Private Sub cmdExport_Click()

    Dim oExport As Object, Session As clsSession, sData As String, sFolder As String
    Dim bYears As Boolean, bMonths As Boolean, bWeeks As Boolean, sRealPath As String
    Dim lngYears As Long, lngMonths As Long, lngWeeks As Long, sFile As String, Tell As Long
    
    ' Retrive the exporter by name
    Set oExport = ChatApp.ExportModules(cmbFileType.Text)

    ' Retrive the settings
    bYears = (chkYears.Value = 1)
    bMonths = (chkMonths.Value = 1)
    bWeeks = (chkWeeks.Value = 1)
    
    ' Get the base path
    sFolder = FileOperation.ValidPath(txtPath)

    ' Initialize progress-form
    frmProgress.Show
    frmProgress.Caption = "Exporting files ..."
    frmProgress.progressBar.Value = 0
    frmProgress.progressBar.Max = frmMain.Logs.Sessions.Count
    DoEvents

    ' See if the export-module handles everything itself
    If oExport.ExportHandleIO Then
    
        ' Just start the process
        oExport.ExportLog frmMain.Logs, sFolder & "Database.mdb", FileOperation

    Else

        ' Loop through all sessions
        For Each Session In frmMain.Logs.Sessions
        
            ' Show the update
            frmProgress.progressBar.Value = frmProgress.progressBar.Value + 1
            frmProgress.lblDescription.Caption = "Exporting session " & frmProgress.progressBar.Value & " of " & frmProgress.progressBar.Max & " ..."
            DoEvents
        
            ' Split the date
            lngYears = Year(Session.SessionDate)
            lngMonths = Month(Session.SessionDate)
            lngWeeks = WeekNum(Session.SessionDate)
        
            ' Get the filename of this session
            sRealPath = sFolder & IIf(bYears, lngYears & "\", "") & IIf(bMonths, lngMonths & _
            IIf(bYears, "", " " & lngYears) & "\", "") & IIf(bWeeks, "Week " & lngWeeks & _
             IIf(bMonths, "", ", " & lngMonths) & IIf(bYears, "", IIf(bMonths, ",", "") & " " _
              & lngYears) & "\", "")
            
            ' Now, get the filename of the first log
            sFile = sRealPath & oExport.RetriveName(Session, 2) & "." & oExport.FileExtension
        
            ' Export the log
            sData = oExport.ExportLog(Session, sFile, FileOperation)
        
            ' Don't continue if there's nothing to save
            If LenB(sData) <> 0 Then
        
                ' Append/save the data
                For Tell = 2 To Session.Participants.Count
                
                    ' Get the new file name
                    sFile = sRealPath & oExport.RetriveName(Session, Tell) & "." & oExport.FileExtension
                
                    ' Save the file
                    FileOperation.SaveFile sFile, sData, oExport.ExportMayAppend
            
                Next
        
            End If
        
        Next
    
    End If
    
    ' Hide the progress-form
    frmProgress.Hide
    
    ' Finally, hide the form
    Me.Hide

End Sub
