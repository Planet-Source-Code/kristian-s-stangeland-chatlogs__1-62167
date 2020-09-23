VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWordCount"
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

' Public events
Event SessionParsed(ByVal Index As Long)
Event WordCounting(ByVal SessionAmount As Long)
Event WordCounted()
Event SavingResult()

Private Type Word
    Content As String
    Occurrences As Long
End Type

' Filters to use
Public Filters As New Collection

' List-view to put the result in
Public ListView As vbalListViewCtl

' Sessions to search from
Public Sessions As Collection

' The amount of words
Public WordCount As Long

' All located words
Private pWords() As Word, bTable(255) As Byte

Public Property Get WordContent(ByVal Index As Long) As String

    ' Retrive the content of this word
    WordContent = pWords(Index).Content

End Property

Public Property Get WordOccurrences(ByVal Index As Long) As Long

    ' Retrive the occurrences of this word
    WordOccurrences = pWords(Index).Occurrences

End Property

Public Sub Clear()

    ' Firstly, clear the array
    Erase pWords
    
    ' Then, clear the count-variable
    WordCount = -1

End Sub

' The main procedure
Public Sub CountWords()
    
    ' Ignore all errors
    On Error Resume Next
    
    ' All word-lists
    Dim Words(1 To 27) As Variant, Tell As Long, bDate As Date, bAllow As Boolean, lngFilter As Long
    Dim Session As clsSession, oEvent As Object, Filter As clsFilter, FilterCount As Long

    ' Get the amount of filters
    FilterCount = Filters.Count

    ' Go through all filters and initialize the values
    For lngFilter = 1 To FilterCount
        
        ' Initialize the filter
        Filters(lngFilter).InitializeValues
        
    Next
    
    ' Reset the word-count and word-array
    Clear
    
    ' Inform about the amount of sessions to count
    RaiseEvent WordCounting(Sessions.Count)
    
    ' Go through all sessions
    For Tell = 1 To Sessions.Count

        ' Reset variables
        bAllow = True

        ' Get the date of this session
        bDate = Sessions(Tell).SessionDate

        ' Go through all filters
        For lngFilter = 1 To FilterCount
        
            ' See if this is a date-fiter
            If Filters(lngFilter).FilterType = Filter_DatePeriod Then
            
                ' Test the date
                bAllow = bAllow And Filters(lngFilter).Test(bDate)
            
            End If
            
        Next
        
        ' If the filters allow this session, go for it ...
        If bAllow Then

            ' Then go through all messages
            For Each oEvent In Sessions(Tell).Events
                        
                ' If this is a message, ...
                If oEvent.IsMessage Then

                    ' Reset variables
                    bAllow = True
                    
                    ' Go through all filters
                    For lngFilter = 1 To FilterCount
                    
                        ' And test the message
                        Select Case Filters(lngFilter).FilterType
                        Case Filter_ReceiverMatch ' Base the filtering upon the receiver
                            bAllow = bAllow And Filters(lngFilter).Test(ChatApp.JoinParticipants(oEvent.Receiver, ","))
                        
                        Case Filter_SenderMatch ' As above, only here by the sender
                            bAllow = bAllow And Filters(lngFilter).Test(oEvent.Sender.NickName & " " & oEvent.Sender.Email)
                        
                        Case Filter_TextMatch ' And by the by the text
                            bAllow = bAllow And Filters(lngFilter).Test(oEvent.Text)
                        
                        Case Filter_TimePeriod ' And here by the time this message was sent
                            bAllow = bAllow And Filters(lngFilter).Test(GetTimeOnly(oEvent.MessageDate))
                        
                        End Select
                    
                    Next
            
                    ' See if we should search within this message
                    If bAllow Then
        
                        '... search for words within it
                        LocateWords Words, oEvent.Text
        
                    End If
        
                End If
        
            Next
    
        End If
    
        ' Inform about the session beeing parsed
        RaiseEvent SessionParsed(Tell)
    
    Next

    ' Save the results
    SaveResults Words
    
    ' We are finished with all
    RaiseEvent WordCounted
    
End Sub

Private Sub SaveResults(Words)

    Dim Row As Long, Tell As Long, lngIndex As Long, sWord As String
    Dim lngOccurrences As Long, oItem As cListItem
    
    ' Inform about the saving-operation
    RaiseEvent SavingResult

    ' Set the array depending of where to put the result
    If ListView Is Nothing Then
    
        ' Initialize the result-array
        ReDim pWords(WordCount)
    
    Else
        
        ' Hide it to increase the process of adding elements
        ShowWindow ListView.hWndListView, SW_HIDE
    
    End If
    
    ' First, go through all tables
    For Row = 1 To 27
    
        ' Don't go further if this isn't an array
        If IsArray(Words(Row)) Then
        
            ' Then go through all underelements
            For Tell = 0 To UBound(Words(Row))
        
                ' Determine what to do
                If ListView Is Nothing Then
        
                    ' Set the word element
                    pWords(lngIndex).Content = Words(Row)(Tell)(0)
                    pWords(lngIndex).Occurrences = Words(Row)(Tell)(1)
                    
                    ' Increase the index
                    lngIndex = lngIndex + 1
        
                Else
                                    
                    ' Get the data
                    sWord = Words(Row)(Tell)(0)
                    lngOccurrences = Words(Row)(Tell)(1)
                                        
                    ' Just add the element to the listview
                    Set oItem = ListView.ListItems.Add(, , sWord)
                    
                    ' Then set its subelements
                    oItem.SubItems(1).Caption = lngOccurrences
                    
                    ' Speed up sorting
                    oItem.ItemData = lngOccurrences
                    oItem.Tag = sWord
                
                End If
        
            Next
    
        End If
    
    Next
    
    If Not ListView Is Nothing Then
    
        ' Show the listview ...
        ShowWindow ListView.hWndListView, SW_SHOWNORMAL
        
        ' ... and update it
        ListView.Refresh
    
    End If

End Sub

Private Sub SetArray(bByte() As Byte, lngStart As Long, lngEnd As Long, Value As Byte)

    Dim Tell As Long
    
    ' Loop through the array ...
    For Tell = lngStart To lngEnd
    
        ' Then set the value
        bByte(Tell) = Value
    
    Next

End Sub

Private Sub LocateWords(Words() As Variant, sText As String)

    Dim intChars() As Integer, intChar As Integer, SafeArray As SAFEARRAY1D, Temp As Long
    Dim Tell As Long, lngLastPos As Long, lngIndex As Long, sWord As String, Element As Long
    
    ' Set default values
    lngLastPos = 1
    
    ' Initialize the array
    InitializeArray intChars, SafeArray, sText
    
    ' Go through all characters
    For Tell = 1 To Len(sText)
    
        ' Get the current character
        intChar = intChars(Tell)
    
        ' Only proceed if this is a part of the ASCII-table
        If intChar < 256 Then
    
            ' All characters beside those in the alphabet are considered to be "word-delimiters"
            If bTable(intChar) > 0 Then
                
                ' This is a character within a word - do nothing
            
            Else
                
                ' See if we may have a word
                If Tell - lngLastPos > 0 Then
                
                    ' Yup; save this word by firstly finding the right table to put it in
                    Select Case intChars(lngLastPos)
                    Case SmallA To SmallZ: lngIndex = intChars(lngLastPos) - SmallA + 1
                    Case LargeA To LargeZ: lngIndex = intChars(lngLastPos) - LargeA + 1
                    Case Else: lngIndex = 27
                    End Select
                
                    ' Then get the word
                    sWord = LCase(Mid(sText, lngLastPos, Tell - lngLastPos))
                    
                    ' First we need to see if this table actually is an array
                    If IsArray(Words(lngIndex)) Then
                    
                        ' And search for it in the table
                        For Element = LBound(Words(lngIndex)) To UBound(Words(lngIndex))
                    
                            ' See if this is the same word
                            If Words(lngIndex)(Element)(0) = sWord Then
                            
                                ' Increase the count
                                Words(lngIndex)(Element)(1) = Words(lngIndex)(Element)(1) + 1
                            
                                ' We are finished
                                Exit For
                            
                            End If
                    
                        Next
                        
                        ' No words was found
                        If Element > UBound(Words(lngIndex)) Then
                        
                            ' Increase the array
                            Temp = IncreaseArray(Words(lngIndex))
                            
                            ' Set the element
                            Words(lngIndex)(Temp) = Array(sWord, 1)
                            
                            ' Increase the word count
                            WordCount = WordCount + 1
                        
                        End If
                    
                    Else
                    
                        ' Initialize the table
                        IncreaseArray Words(lngIndex)
                        
                        ' Set this element
                        Words(lngIndex)(0) = Array(sWord, 1)
                    
                        ' Increase the word count
                        WordCount = WordCount + 1
                    
                    End If
                
                End If
                
                ' Save this position for later
                lngLastPos = Tell + 1
                
            End If
        
        End If
    
    Next
    
    ' Clean up
    ZeroMemory VarPtrArray(intChars), 4

End Sub

Private Function IncreaseArray(aArray As Variant) As Long

    ' First, see if this actually is an array
    If IsArray(aArray) Then
    
        ' Return the index
        IncreaseArray = UBound(aArray) + 1
        
        ' Initialize the array
        ReDim Preserve aArray(IncreaseArray)
    
    Else
        
        ' If not, initialize it
        ReDim aArray(0)
        
    End If

End Function

Private Sub Class_Initialize()

    ' Executed as standard on startup
    Clear

    ' Set the letter-table
    SetArray bTable, SmallA, SmallZ, 1
    SetArray bTable, LargeA, LargeZ, 1
    SetArray bTable, 192, 246, 1
    SetArray bTable, 247, 255, 1

End Sub
