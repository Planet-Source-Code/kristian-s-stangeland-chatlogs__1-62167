VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PluginMain"
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

Public Animation As New clsStarWars

' Commands
Private Const AnimationStart = 1
Private Const AnimationStop = 2
Private Const AnimationResize = 3
Private Const AnimationFullscreen = 4

' The original parent
Private lngParent As Long
Private lngCurrent As Long

Public Property Get Name() As String

    ' The plug-in's name.
    Name = "AboutAnimation"

End Property

Public Property Get Description() As String

    ' Gets the description of the plug-in
    Description = "Creates a class Star Wars-scene with scrollable text"

End Property

Public Property Get Author() As String
        
    ' Returns who made this plugin
    Author = "Kristian S. Stangeland"

End Property

Public Property Get Configurable() As Boolean

    ' Nope, not yet.
    Configurable = False

End Property

Public Sub Initialize(oMainObject As Object)
    
    ' Not needed.

End Sub

' Called when the user has asked to configure the plug-in
Public Sub Configure()

    ' Not yet.
    
End Sub

Private Sub InitializeAnimation()

    With Animation
        ' Remove all current lines
        .ClearLines
        
        ' Add all the lines again
        .AddLine "CHAT LOGS", , 5, 3, 2
        .AddLine ""
        .AddLine "Episode III - The revenge of the nerd", , , , 2
        .AddLine ""
        .AddLine "A long time ago, in a galaxy far far away, hidden deep"
        .AddLine "in a backwater sector of an insignificant galaxy, a"
        .AddLine "man started his war against the grasshoppers. No one"
        .AddLine "knows what happened to this guy, so never mind."
        .AddLine ""
        .AddLine "But, somewhere else, in a different galaxy, a evil"
        .AddLine "program was planned, a vast program that could force"
        .AddLine "order unto otherwise random message logs."
        .AddLine ""
        .AddLine "The creator deduced that such a tool might make it"
        .AddLine "possible to retrieve information that might otherwise"
        .AddLine "have been lost forever, far beneath the seas of chaos"
        .AddLine "that is a chat log. If one was to fetch and organize"
        .AddLine "all the bits and pieces generated throughout the years,"
        .AddLine "it would be an immensive advantage over other methods"
        .AddLine "of getting information, or so he believed."
        .AddLine ""
        .AddLine "Thus the mightly CHAT LOGS was forged, which changed"
        .AddLine "history forever."
        
    End With

End Sub

' Used for the plug-ins to communticate with each other
Public Function Message(uMsgCode As Long, wData As Long, lData As Variant)

    ' Execute the command
    Select Case uMsgCode
    Case AnimationStart
    
        ' Make a reference to this class
        Set frmWindow.PluginMain = Me
        
        ' Show the window to use
        frmWindow.Show
        
        ' Get the original parent
        lngParent = GetParent(frmWindow.hWnd)
        lngCurrent = wData
        
        ' Set the parent of the window
        SetParent frmWindow.hWnd, lngCurrent
        
        ' Now, add all lines
        InitializeAnimation
        
        ' Start the animation with the hDC passed in wData
        Animation.StartAnimation frmWindow.hDC
        
    Case AnimationStop
        
        ' Reset the parent
        SetParent frmWindow.hWnd, lngParent
        
        ' Just stop the animation
        Animation.StopAnimation
    
        ' And hide the window
        frmWindow.Hide
        
    Case AnimationResize
    
        ' Resize the window ...
        frmWindow.Width = wData * Screen.TwipsPerPixelX
        frmWindow.Height = Val(lData) * Screen.TwipsPerPixelY
    
        ' ... and the class
        Animation.WindowWidth = wData
        Animation.WindowHeight = Val(lData)
        
        
    Case AnimationFullscreen
    
        ' See if we should enable fullscreen or not
        If wData = 1 Then
        
            ' Set the original parent of the window
            SetParent frmWindow.hWnd, lngParent
            
            ' And maximize it
            frmWindow.WindowState = 2
            
            ' Now, the width and height is equal to the size of the screen
            Animation.WindowWidth = Screen.Width / Screen.TwipsPerPixelX
            Animation.WindowHeight = Screen.Height / Screen.TwipsPerPixelY
        
        Else
        
            ' Use normal state
            frmWindow.WindowState = 0
        
            ' Move the window back to the new parent
            SetParent frmWindow.hWnd, lngCurrent
        
            ' Set as the original size
            Animation.WindowWidth = frmWindow.ScaleWidth
            Animation.WindowHeight = frmWindow.ScaleHeight
        
        End If
    
    End Select

End Function
