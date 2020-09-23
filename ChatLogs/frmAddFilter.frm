VERSION 5.00
Begin VB.Form frmAddFilter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add a filter"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   254
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkUsePattern 
      Caption         =   "&Use pattern"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   2520
      Width           =   1815
   End
   Begin VB.ComboBox cmbOperation 
      Height          =   315
      ItemData        =   "frmAddFilter.frx":0000
      Left            =   2280
      List            =   "frmAddFilter.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1920
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtValueTwo 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   1440
      Width           =   3975
   End
   Begin VB.TextBox txtValueOne 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   3975
   End
   Begin VB.ComboBox cmbFilterType 
      Height          =   315
      ItemData        =   "frmAddFilter.frx":0004
      Left            =   2280
      List            =   "frmAddFilter.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label lblOperation 
      Caption         =   "&Include/Exclude:"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   1950
      Width           =   1695
   End
   Begin VB.Label lblValueTwo 
      Caption         =   "V&alue two: "
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblValueOne 
      Caption         =   "&Value one: "
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblFilterTyoe 
      Caption         =   "&Filter type: "
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   510
      Width           =   1695
   End
End
Attribute VB_Name = "frmAddFilter"
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

Private Sub Form_Load()

    ' Initialize this form
    ChatApp.InitializeForm Me

End Sub

Private Sub cmdCancel_Click()
    
    ' Return the action
    Me.Tag = "CANCEL"

End Sub

Private Sub cmdOK_Click()

    ' Return the action
    Me.Tag = "OK"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Surely, the form now needs to report
    Me.Tag = "CANCEL"

End Sub

