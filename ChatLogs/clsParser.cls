VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParser"
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

' Different variables related to the parsing operation
Public ParseType As String
Public LogData As String
Public LogClass As clsLog

Public Sub AboutToBegin()

    ' Inform the parser that we're about to begin
    ChatApp.Parsers(ParseType).ParserInitialize

End Sub

Public Sub BeginParsing(bBytes() As Byte)

    Dim currParser As Object

    ' Firstly, get the parser to use
    Set currParser = ChatApp.Parsers(ParseType)
    
    ' Then execute the parsing depending on what to use
    If LenB(LogData) > 0 Then
        currParser.ParseDocument LogData, LogClass
    Else
        currParser.ParseDocument bBytes, LogClass
    End If
    
End Sub

Public Property Get DeliverString() As Boolean

    ' Return the property
    DeliverString = ChatApp.Parsers(ParseType).ParserDeliverString
    
End Property

Public Property Get HandleIO() As Boolean

    ' Return the property
    HandleIO = ChatApp.Parsers(ParseType).ParserHandleIO

End Property

