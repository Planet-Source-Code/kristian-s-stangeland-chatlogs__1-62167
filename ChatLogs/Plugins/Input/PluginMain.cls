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

Public Property Get Name() As String

    ' The plug-in's name.
    Name = "InputParsers"

End Property

Public Property Get Author() As String
    
    ' Returns who made this plugin
    Author = "Kristian S. Stangeland"

End Property

Public Property Get Description() As String
    
    ' Gets the description of the plug-in
    Description = "Is capable of parsing both MSN Messenger and Messenger Plus!-logs"

End Property

' If this plug-in is configurable or not
Public Property Get Configurable() As Boolean

    ' It could be in the near future, but right now - NO
    Configurable = False

End Property

' Is invoked at start.
Public Sub Initialize(oMainObject As Object)

    ' Save the main object
    Set MainObject = oMainObject
    
    ' Add the input-parsers
    With MainObject
        .Parsers.Add New clsTextFile, "MessengerPlus"
        .Parsers.Add New clsMSNLog, "MSN Messenger"
        .Parsers.Add New clsContacts, "MSN Contact List"
        .Parsers.Add New clsGaimLog, "Gaim Chat-log"
        .Parsers.Add New clsTrillian, "Trillian"
    End With

End Sub

