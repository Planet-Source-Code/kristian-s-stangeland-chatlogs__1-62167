VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 2  'vbComplexBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This module contains code for developers to attach their own programs (plug-ins) to this application.

' The plug-in's name. This is the string that will identify your class for the application and other plugins.
Public Property Get Name() As String
End Property

' Gets the description of the plug-in
Public Property Get Description() As String
End Property

' Returns who made this plugin
Public Property Get Author() As String
End Property

' If this plug-in is configurable or not
Public Property Get Configurable() As Boolean
End Property

' Is invoked at start. gObjects is a collection of the global classes
Public Sub Initialize(gObjects As Collection)
End Sub

' Called when the user has asked to configure the plug-in
Public Sub Configure()
End Sub

' Used for the plug-ins to communticate with each other
Public Function Message(uMsgCode As Long, wData As Long, lData As Variant)
End Function

' Not always necessary, but can be useful when other plug-ins or the application wants to clone the object.
Public Function Clone() As Object
End Function
