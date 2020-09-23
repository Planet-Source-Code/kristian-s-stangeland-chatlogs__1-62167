Sub UpdateControl(usrControl, Sessions)

    Dim Months(), lngMonth, lngHighest, dLastDate, Tell, Message, Session

    ' Allocate the data-array containing the information to gather
    ReDim Months(11)
    
    ' Go through all sessions
    For Each Session In Sessions

        ' Then go through all events
        For Each Message In Session.Events
            
            ' See if this is a message, and not an event
            If Message.IsMessage Then
            
		' Get the month of this date
                lngMonth = Month(Message.MessageDate)

                ' If so, add the message to the data-array
                Months(lngMonth - 1) = Months(lngMonth - 1) + 1
            
            End If

        Next

    Next
    
    ' Get the highest value
    For Tell = 0 To 11
        If Months(Tell) > lngHighest Then
            lngHighest = Months(Tell)
        End If
    Next
    
    ' Set the amount of columns and the maximum value
    usrControl.ColumnCount = 12
    usrControl.MaxValue = lngHighest
    
    ' Now, present the gathered information
    For Tell = 0 To 11
    
        ' Set the name and value of this column
        usrControl.SetData Tell + 1, MonthName(Tell + 1, True), Months(Tell)
    
    Next

    ' Redraw the control
    usrControl.RedrawChart

End Sub