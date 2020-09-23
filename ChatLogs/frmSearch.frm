VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkIgnoreCase 
      Caption         =   "&Ignore Case"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   2640
      Width           =   6015
   End
   Begin VB.TextBox txtReceiver 
      Height          =   285
      Left            =   1920
      TabIndex        =   13
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox txtSender 
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox txtToDate 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox txtFromDate 
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CheckBox chkPatternMatch 
      Caption         =   "&Use Pattern Matching"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   2400
      Width           =   6015
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   1500
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   810
      Width           =   1335
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtSearchString 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label lblReciever 
      Caption         =   "&Receiver:"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   1110
      Width           =   1575
   End
   Begin VB.Label lblSender 
      Caption         =   "&Sender:"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   750
      Width           =   1575
   End
   Begin VB.Label lblToDate 
      Caption         =   "&To date/time:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1830
      Width           =   1575
   End
   Begin VB.Label lblFromDate 
      Caption         =   "&From date/time:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1470
      Width           =   1575
   End
   Begin VB.Label lblSearchString 
      Caption         =   "&Find: "
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "frmSearch"
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

Public SearchClass As New clsSearch

Private Sub cmdCancel_Click()

    ' Just hide this form
    Me.Hide

End Sub

Private Sub cmdFind_Click()
    
    Dim frmResult As frmMessageList, Filters As New Collection, bIgnoreCase As Boolean
    
    ' See if we should ignore case
    bIgnoreCase = (chkIgnoreCase.Value = 1)
 
    ' Insert the date filter if we need it
    If LenB(txtFromDate) And LenB(txtToDate) Then
        Filters.Add ChatApp.CreateFilter(Filter_DatePeriod, True, CDate(txtFromDate), CDate(txtToDate))
    End If
    
    ' Insert the receiver-filter if we need it
    If LenB(txtReceiver) Then
        Filters.Add ChatApp.CreateFilter(Filter_ReceiverMatch, True, txtReceiver, , True, bIgnoreCase)
    End If
    
    ' Insert the sender-filter if we need it
    If LenB(txtSender) Then
        Filters.Add ChatApp.CreateFilter(Filter_SenderMatch, True, txtReceiver, , True, bIgnoreCase)
    End If

    ' Then finally insert the query
    Filters.Add ChatApp.CreateFilter(Filter_TextMatch, True, txtSearchString, , CBool(chkPatternMatch.Value = 1), bIgnoreCase)

    ' Initialize class
    With SearchClass
        
        ' Initialize source
        Set .Container = frmMain.Logs
    
        ' Set the filters
        Set .Filters = Filters
        
    End With
    
    ' Execute search
    SearchClass.InitiateSearch
    
    ' Show the result by creating a new form
    Set frmResult = New frmMessageList
    
    ' And update it
    frmResult.Update SearchClass.Messages
    
End Sub

Private Sub Form_Load()

    ' Default values
    txtFromDate.Text = DateSerial(Year(Now) - 5, Month(Now), Day(Now)) + TimeSerial(Hour(Now), Minute(Now), Second(Now))
    txtToDate.Text = Now

    ' Initialize this form
    ChatApp.InitializeForm Me

End Sub
