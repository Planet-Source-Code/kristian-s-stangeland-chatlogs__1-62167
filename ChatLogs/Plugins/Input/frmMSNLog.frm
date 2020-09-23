VERSION 5.00
Begin VB.Form frmMSNLog 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MSN Input Module"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProceed 
      Caption         =   "&Proceed"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtNickName 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Top             =   1800
      Width           =   3855
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Label lblNickName 
      Caption         =   "&Nickname:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label lblEmail 
      Caption         =   "&Email address: "
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmMSNLog.frx":0000
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   6255
   End
End
Attribute VB_Name = "frmMSNLog"
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

Private Sub cmdProceed_Click()

    ' Save settings
    SaveAll
        
    ' Then unload the form
    Unload Me
    
    ' And inform that a action has been taken
    Me.Tag = "PROCEED"

End Sub

Public Sub SaveAll()

    ' Save the settings
    SaveSetting "ChatLogs", "Settings", "PrimaryUserEmail", txtEmail
    SaveSetting "ChatLogs", "Settings", "PrimaryUserNick", txtNickName

End Sub

Public Sub LoadAll()

    ' Save the settings
    txtEmail = GetSetting("ChatLogs", "Settings", "PrimaryUserEmail", "")
    txtNickName = GetSetting("ChatLogs", "Settings", "PrimaryUserNick", "")

End Sub
