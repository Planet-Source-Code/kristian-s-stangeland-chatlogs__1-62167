VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSettings"
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

Public Property Get Setting(Category As String, Name As String, DefaultValue) As Variant

    ' Retrive the setting
    Setting = GetSetting("ChatLogs", Category, Name, DefaultValue)

End Property

Public Property Let Setting(Category As String, Name As String, DefaultValue, ByVal vNewValue As Variant)

    ' Save the setting
    SaveSetting "ChatLogs", Category, Name, CStr(vNewValue)

End Property

Public Property Get PluginEnabled(PluginName As String) As Boolean

    ' Retrive whether or not this plugin is enabled
    PluginEnabled = CBool(Setting("Plugins", PluginName, True))

End Property

Public Property Let PluginEnabled(PluginName As String, ByVal vNewValue As Boolean)
    
    ' Save the status of this plugin
    Setting("Plugins", PluginName, "") = vNewValue

End Property
