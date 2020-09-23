VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilter"
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

' The different filter types
Enum eFilterType
    Filter_None
    Filter_DatePeriod
    Filter_TimePeriod
    Filter_SenderMatch
    Filter_ReceiverMatch
    Filter_TextMatch
End Enum

' Variables used to identify this filter
Public FilterType As eFilterType
Public ValueOne As Variant
Public ValueTwo As Variant
Public ResultValue As Boolean
Public UsePattern As Boolean
Public IgnoreCase As Boolean

' The pattern to use
Private Pattern As New RegExp

Public Sub InitializeValues()

    If UsePattern Then

        ' Initialize the pattern
        Pattern.IgnoreCase = IgnoreCase
        Pattern.Pattern = ValueOne

    Else
    
        ' Intialize the values
        Select Case FilterType
        Case Filter_DatePeriod, Filter_TimePeriod
        
            ValueOne = CDate(ValueOne)
            ValueTwo = CDate(ValueTwo)
        
        End Select
    
    End If

End Sub

Public Function Test(sExpression) As Boolean

    Dim dDate As Date

    ' See if we're using a pattern or not
    If UsePattern Then
    
        ' Really, there's not much different when we're testing with a pattern
        Test = Pattern.Test(sExpression)
    
    Else
    
        ' Test the expression differently depending of the filter-type
        Select Case FilterType
        Case Filter_DatePeriod, Filter_TimePeriod
        
            ' Get the date of the expression
            dDate = CDate(sExpression)
        
            ' Simply see if the expression is within the period
            Test = (dDate > ValueOne And dDate < ValueTwo)
        
        Case Filter_SenderMatch, Filter_ReceiverMatch, Filter_TextMatch
        
            ' Return whether or not the filter-dater given is within the expression
            Test = (InStr(1, sExpression, ValueOne, IIf(IgnoreCase, vbTextCompare, vbBinaryCompare)) <> 0)
        
        End Select
    
    End If
    
End Function

Public Function ToString() As String

    ' Make the string ready
    ToString = IIf(ResultValue, LanguageConst("FilterInclude"), LanguageConst("FilterExclude"))

    ' Get the description for the filter
    Select Case FilterType
    Case Filter_DatePeriod, Filter_TimePeriod
    
        ' When using a pattern, date-matching works differently
        If UsePattern Then
            
            ' Filter by date, here using a pattern
            ToString = ToString & LanguageConst("FilterWhen") & IIf(FilterType = Filter_TimePeriod, _
             LanguageConst("FilterTime"), LanguageConst("FilterDate")) & LanguageConst("FilterPattern") & ValueOne
        
        Else
    
            ' Filter by date, using two dates
            ToString = ToString & LanguageConst("FilterWhen") & IIf(FilterType = Filter_TimePeriod, _
             LanguageConst("FilterTime"), LanguageConst("FilterDate")) & LanguageConst("FilterIsBetween") & _
              ValueOne & LanguageConst("FilterAnd") & ValueTwo
        
        End If
        
    Case Filter_SenderMatch, Filter_ReceiverMatch

        ' Pattern-matching is more versatile, but drains power from the CPU
        If UsePattern Then
            
            ' Filter by sender, here using a pattern
            ToString = ToString & LanguageConst("FilterWhen") & IIf(FilterType = Filter_SenderMatch, _
             LanguageConst("FilterSender"), LanguageConst("FilterReceiver")) & _
              LanguageConst("FilterPattern") & ValueOne
        
        Else
    
            ' Filter by sender, using a string that should/shoult not correspond
            ToString = ToString & LanguageConst("FilterWhen") & IIf(FilterType = _
             Filter_SenderMatch, LanguageConst("FilterSender"), LanguageConst("FilterReceiver")) & _
              LanguageConst("FilterCorrespond") & ValueOne
        
        End If

    Case Filter_TextMatch
    
        ' Pattern-matching is more versatile, but drains power from the CPU
        If UsePattern Then
            
            ' Filter by sender, here using a pattern
            ToString = ToString & LanguageConst("FilterTextPattern") & ValueOne
        
        Else
    
            ' Filter by sender, using a string that should/shoult not correspond
            ToString = ToString & LanguageConst("FilterTextCorrespond") & ValueOne
        
        End If
    
    End Select

End Function
