VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About ..."
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   473
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   596
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picAnimation 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HasDC           =   0   'False
      Height          =   6015
      Left            =   120
      ScaleHeight     =   397
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   573
      TabIndex        =   1
      Top             =   120
      Width           =   8655
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   7560
      TabIndex        =   0
      Top             =   6540
      Width           =   1260
   End
   Begin VB.Line lineDelimiter 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   8
      X2              =   584
      Y1              =   423
      Y2              =   423
   End
   Begin VB.Line lineDelimiter 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   8
      X2              =   584
      Y1              =   424
      Y2              =   424
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "This program is made by Kristian S. Stangeland and is licenced under the GPL-licence."
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   6600
      Width           =   7095
   End
End
Attribute VB_Name = "frmAbout"
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

' Commands for the animation plugin
Private Const AnimationStart = 1
Private Const AnimationStop = 2
Private Const AnimationResize = 3
Private Const AnimationFullscreen = 4

Private Sub cmdOK_Click()

    ' Kill the form
    Unload Me

End Sub

Public Sub StartAnimation()

    ' Resize the control
    ResizeControls

    ' Start the animation
    ChatApp.Plugins("AboutAnimation").Message AnimationStart, picAnimation.hwnd, 0

End Sub

Public Sub ResizeControls()

    ' Send the size of our control
    ChatApp.Plugins("AboutAnimation").Message AnimationResize, picAnimation.ScaleWidth, picAnimation.ScaleHeight

End Sub

Private Sub StopAnimation()
    
    ' Stop the animation
    ChatApp.Plugins("AboutAnimation").Message AnimationStop, 0, 0

End Sub

Private Sub Form_Load()

    ' Initialize this form
    ChatApp.InitializeForm Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Just stop the animation
    StopAnimation

End Sub
