VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Log(s) to search in
Public Container As clsContainer

' All messages found to correspond to the query
Public Messages As Collection

' Filters used in combination with the above settings
Public Filters As Collection

Public Sub InitiateSearch()

    On Error Resume Next

    ' Variables used for, among other things, literation
    Dim Session As clsSession, Message As Object, lngFilter As Long
    Dim FilterCount As Long, bAllow As Boolean, bDate As Date, lngTextFilter As Long
    
    ' Create a new collection
    Set Messages = New Collection
    
    ' Retrive the amount of filters
    FilterCount = Filters.Count
    
    ' Go through all filters and initialize the values
    For lngFilter = 1 To FilterCount
        
        ' Initialize the filter
        Filters(lngFilter).InitializeValues
        
        ' See if this is a text-filter (used in optimization)
        If Filters(lngFilter).FilterType = Filter_TextMatch Then
        
            ' Save its index
            lngTextFilter = lngFilter
            
        End If
        
    Next
    
    ' Go through all sessions in ALL logs
    For Each Session In Container.Sessions
    
        ' Reset variables
        bAllow = True

        ' Get the date of this session
        bDate = Session.SessionDate
    
        ' Go through all filters, excluding this session if the date isn't matching
        For lngFilter = 1 To FilterCount
        
            ' See if this is a date-fiter
            If Filters(lngFilter).FilterType = Filter_DatePeriod Then
            
                ' Test the date
                bAllow = bAllow And Filters(lngFilter).Test(bDate)
            
            End If
            
        Next
        
        ' Only continue if the date-filters let the session pass
        If bAllow Then
    
            ' Then run through each message in each session
            For Each Message In Session.Events
        
                ' Be sure that this actually is a message
                If TypeOf Message Is clsMessage Then
                    
                    ' Reset variables
                    bAllow = True
                    
                    ' Go through all filters
                    For lngFilter = 1 To FilterCount
                    
                        ' And test the message
                        Select Case Filters(lngFilter).FilterType
                        Case Filter_ReceiverMatch ' Base the filtering upon the receiver
                            bAllow = bAllow And Filters(lngFilter).Test(ChatApp.JoinParticipants(Message.Receiver, ","))
                        
                        Case Filter_SenderMatch ' As above, only here by the sender
                            bAllow = bAllow And Filters(lngFilter).Test(Message.Sender.NickName & " " & Message.Sender.Email)
                        
                        Case Filter_TextMatch ' And by the by the text
                            bAllow = bAllow And Filters(lngFilter).Test(Message.Text)
                        
                        Case Filter_TimePeriod ' And here by the time this message was sent
                            bAllow = bAllow And Filters(lngFilter).Test(GetTimeOnly(Message.MessageDate))
                        
                        End Select
                    
                    Next
                    
                    ' If it passed through the filters, ...
                    If bAllow Then
                    
                        ' See if we actually have a text-filter (we can only use ONE at a time)
                        If lngTextFilter >= 1 Then
                            
                            ' Test the query
                            If Filters(lngTextFilter).Test(Message.Text) Then
                                Messages.Add Message
                            End If
                            
                        Else
                            
                            ' Add the message nevertheless
                            Messages.Add Message
                            
                        End If
                    
                    End If
                    
                End If
        
            Next

        End If

    Next

End Sub

Private Function ReceiverToString(colReceiver As Collection) As String

    Dim Receiver As clsUser
    
    ' Go through each user
    For Each Receiver In colReceiver
    
        ' Add the email and nickname
        ReceiverToString = ReceiverToString & Receiver.NickName & vbCrLf & Receiver.Email & vbCrLf

    Next

End Function
