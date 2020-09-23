VERSION 5.00
Object = "{E910F8E1-8996-4EE9-90F1-3E7C64FA9829}#1.1#0"; "vbaListView6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdUse 
      Caption         =   "&Use"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   7080
      Width           =   1215
   End
   Begin TabDlg.SSTab stSettings 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   11880
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   529
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmSettings.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frameLanguage"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Plugins"
      TabPicture(1)   =   "frmSettings.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameAdmin"
      Tab(1).ControlCount=   1
      Begin VB.Frame frameLanguage 
         Caption         =   "Language:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   5535
         Begin VB.ComboBox cmdLanguagePack 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label lblChooseLanguage 
            Caption         =   "&Choose language:"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   540
            Width           =   2295
         End
      End
      Begin VB.Frame frameAdmin 
         Caption         =   "&Manage plugins"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5295
         Left            =   -74760
         TabIndex        =   1
         Top             =   720
         Width           =   5535
         Begin VB.PictureBox picPlugins 
            BorderStyle     =   0  'None
            Height          =   4695
            Left            =   240
            ScaleHeight     =   4695
            ScaleWidth      =   5175
            TabIndex        =   5
            Top             =   480
            Width           =   5175
            Begin VB.CommandButton cmdConfigure 
               Caption         =   "&Configure"
               Height          =   375
               Left            =   0
               TabIndex        =   7
               Top             =   4200
               Width           =   1695
            End
            Begin vbalListViewLib6.vbalListViewCtl lstPlugins 
               Height          =   4095
               Left            =   0
               TabIndex        =   6
               Top             =   0
               Width           =   5055
               _ExtentX        =   8916
               _ExtentY        =   7223
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               View            =   1
               LabelEdit       =   0   'False
               AutoArrange     =   0   'False
               CheckBoxes      =   -1  'True
               HeaderButtons   =   0   'False
               HeaderTrackSelect=   0   'False
               HideSelection   =   0   'False
               InfoTips        =   0   'False
            End
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

    Dim Tell As Long, ListItem As cListItem
    
    ' Firstly, save the plugin-settings
    For Tell = 1 To lstPlugins.ListItems.Count
    
        ' Retrive the listitem
        Set ListItem = lstPlugins.ListItems(Tell)
    
        ' Save each plugin
        ChatApp.Settings.PluginEnabled(ListItem.Text) = ListItem.Checked
    
    Next
    
    ' Then, set the language-pack
    ChatApp.Settings.Setting("Settings", "LanguagePack", "") = cmdLanguagePack.Text
    ChatApp.ReloadLanguage
    
End Sub

Public Sub LoadAll()

    Dim Plugin As Object, Language, sLanguage As String, ListItem As cListItem

    ' Clear the listview
    lstPlugins.ListItems.Clear
    
    ' Add each plugin to it
    For Each Plugin In ChatApp.Plugins
    
        ' Add the listitem
        Set ListItem = lstPlugins.ListItems.Add(, , Plugin.Name)
    
        ' Set different settings
        ListItem.SubItems(1).Caption = Plugin.Author
        ListItem.SubItems(2).Caption = Plugin.Description
        
        ' Whether or not it should be selected
        ListItem.Checked = ChatApp.Settings.PluginEnabled(Plugin.Name)
    
    Next

    ' Update listbox
    UpdateListBox

    ' Nextly, clear the combobox
    cmdLanguagePack.Clear

    ' Then, get all languages and show them
    For Each Language In FileOperation.RetriveFileList(App.Path & "\Data\Languages\", "*.lpk", False)

        ' Now, get the filname only
        sLanguage = FileOperation.GetFileName(CStr(Language))

        ' Add the language
        cmdLanguagePack.AddItem sLanguage
        
        ' If this language correspond to the current language, select it
        If sLanguage = ChatApp.Settings.Setting("Settings", "LanguagePack", "English.lpk") Then
            cmdLanguagePack.ListIndex = cmdLanguagePack.ListCount - 1
        End If

    Next

End Sub

Private Sub cmdCancel_Click()

    ' Unload without saving
    Unload Me

End Sub

Private Sub cmdConfigure_Click()

    ' Execute the configure-prosedure
    ChatApp.Plugins(lstPlugins.SelectedItem.Index).Configure

End Sub

Private Sub cmdOK_Click()

    ' Save the settings
    SaveAll
    
    ' Unload the form
    Unload Me

End Sub

Private Sub cmdUse_Click()

    ' Save all settings
    SaveAll

End Sub

Private Sub Form_Load()

    ' Initialize this form
    ChatApp.InitializeForm Me

End Sub

Private Sub UpdateListBox()

    On Error Resume Next

    ' Disable as default
    cmdConfigure.Enabled = False

    ' Set whether or not the command box should be disabled or not
    cmdConfigure.Enabled = ChatApp.Plugins(lstPlugins.SelectedItem.Index).Configurable

End Sub

Private Sub lstPlugins_ItemClick(Item As vbalListViewLib6.cListItem)
    
    ' Update the listbox
    UpdateListBox

End Sub

Private Sub lstPlugins_KeyDown(KeyCode As Integer, Shift As Integer)

    ' Update the listbox
    UpdateListBox

End Sub

