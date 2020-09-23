Sub UpdateControl(usrControl, Sessions)

    Dim Weekdays(), lngWeekday, dLastDate, lngHighest, Tell, Message, Session
    
    ' Allocate the data-array containing the information to gather
    ReDim Weekdays(6)
    
    ' Go through all sessions
    For Each Session In Sessions

        ' Then go through all events
        For Each Message In Session.Events
            
            ' See if this is a message, and not an event
            If Message.IsMessage Then
            
		If GetDateOnly(Message.MessageDate) <> dLastDate Then
		    ' Calculate the weekday
                    dLastDate = GetDateOnly(Message.MessageDate)
                    lngWeekday = Weekday(dLastDate)
                End If 

                ' If so, add the message to the data-array
                Weekdays(lngWeekday - 1) = Weekdays(lngWeekday - 1) + 1
            
            End If

        Next

    Next
    
    ' Get the highest value
    For Tell = 0 To 6
        If Weekdays(Tell) > lngHighest Then
            lngHighest = Weekdays(Tell)
        End If
    Next
    
    ' Set the amount of columns and the maximum value
    usrControl.ColumnCount = 7
    usrControl.MaxValue = lngHighest
    
    ' Now, present the gathered information
    For Tell = 0 To 6
    
        ' Set the name and value of this column
        usrControl.SetData Tell + 1, WeekdayName(Tell + 1), Weekdays(Tell)
    
    Next

    ' Redraw the control
    usrControl.RedrawChart

End Sub

Function GetDateOnly(dDate)

    GetDateOnly = DateSerial(Year(dDate), Month(dDate), Day(dDate))

End Function