VERSION 5.00
Begin VB.Form frmBrowse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Browse for folder"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   349
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmdFileType 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3720
      Width           =   3975
   End
   Begin VB.CheckBox chkAddSubfolders 
      Caption         =   "&Add subfolders"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   3975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   4680
      Width           =   1095
   End
   Begin VB.DirListBox lstFolders 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
   Begin VB.DriveListBox lstDrive 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label lblDescription 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmBrowse"
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

Private Sub cmdCancel_Click()
    
    ' Return the action
    Me.Tag = "CANCEL"

End Sub

Private Sub cmdOK_Click()

    ' Return the action
    Me.Tag = "OK"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Surely, the form now needs to report
    Me.Tag = "CANCEL"

End Sub

Private Sub lstDrive_Change()

    On Error Resume Next

    ' Change the path of the folder-list
    lstFolders.Path = lstDrive.Drive

End Sub
