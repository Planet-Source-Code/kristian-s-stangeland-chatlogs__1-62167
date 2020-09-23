VERSION 5.00
Begin VB.Form frmWindow 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "GDI Window"
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   224
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmWindow"
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

' Used to set the window as fullscreen
Private Const AnimationFullscreen = 4

Public PluginMain As PluginMain

Private Sub Form_DblClick()

    ' Now, set as fullscreen, depending on the current status
    PluginMain.Message AnimationFullscreen, IIf(Me.WindowState = 2, 0, 1), 0

End Sub

