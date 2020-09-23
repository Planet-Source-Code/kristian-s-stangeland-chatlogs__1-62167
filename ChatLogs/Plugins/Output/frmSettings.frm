VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Output Settings"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame frameMessengerPlus 
      Caption         =   "Messenger Plus:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.PictureBox picMessengerPlus 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   240
         ScaleHeight     =   1695
         ScaleWidth      =   4575
         TabIndex        =   3
         Top             =   480
         Width           =   4575
         Begin VB.CheckBox chkMessageTime 
            Caption         =   "Add the time to each message and event"
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Tag             =   "Value;1"
            Top             =   0
            Width           =   4455
         End
         Begin VB.CheckBox chkWordWrap 
            Caption         =   "&Use word wrap"
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Tag             =   "Value;1"
            Top             =   360
            Width           =   4455
         End
         Begin VB.TextBox txtMaxLenght 
            Height          =   285
            Left            =   3120
            TabIndex        =   5
            Tag             =   "Text;-1"
            Text            =   "-1"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CheckBox chkSaveEvents 
            Caption         =   "&Save events"
            Height          =   255
            Left            =   0
            TabIndex        =   4
            Tag             =   "Value;1"
            Top             =   720
            Width           =   4455
         End
         Begin VB.Label lblMaxLenght 
            Caption         =   "&Maximum lenght of nickname:"
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   1320
            Width           =   3015
         End
      End
   End
End
Attribute VB_Name = "frmSettings"
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

Public Sub SaveAll()

    Dim Control As Object, Values
    
    ' Go through all controls in this form
    For Each Control In Controls
    
        ' If the tag is specified, ...
        If LenB(Control.Tag) Then
    
            ' Get the different values of the tag
            Values = Split(Control.Tag, ";")
    
            ' ... save this control
            SaveSetting "ChatLogs", "Settings", Mid(Control.Name, 4), CallByName(Control, Values(0), VbGet)
    
        End If
    
    Next

End Sub

Public Sub LoadAll()

    Dim Control As Object, Values
    
    ' Go through all controls in this form
    For Each Control In Controls
    
        ' If the tag is specified, ...
        If LenB(Control.Tag) Then
    
            ' Get the different values of the tag
            Values = Split(Control.Tag, ";")
    
            ' ... load the control
            CallByName Control, Values(0), VbLet, GetSetting("ChatLogs", "Settings", Mid(Control.Name, 4), Values(1))
    
        End If
    
    Next

End Sub

Private Sub cmdCancel_Click()

    ' Hide form
    Unload Me

End Sub

Private Sub cmdOK_Click()

    ' Save all
    SaveAll

    ' Hide form
    Unload Me

End Sub
