Sub UpdateControl(usrControl, Sessions)

    Dim Weeks(), lngWeek, lngHighest, dLastDate, Tell, Message, Session
    
    ' Get the amout of weeks this year
    WeekAmout = WeekNum(DateSerial(Year(Date), 12, 31)) - 1

    ' Allocate the data-array containing the information to gather
    ReDim Weeks(WeekAmout)
    
    ' Go through all sessions
    For Each Session In Sessions

        ' Then go through all events
        For Each Message In Session.Events
            
            ' See if this is a message, and not an event
            If Message.IsMessage Then
            
		If GetDateOnly(Message.MessageDate) <> dLastDate Then
		    ' Calculate the weekday
                    dLastDate = GetDateOnly(Message.MessageDate)
                    lngWeek = WeekNum(dLastDate)
                End If 
            
                ' If so, add the message to the data-array
                Weeks(lngWeek - 1) = Weeks(lngWeek - 1) + 1
            
            End If

        Next

    Next
    
    ' Get the highest value
    For Tell = 0 To WeekAmout
        If Weeks(Tell) > lngHighest Then
            lngHighest = Weeks(Tell)
        End If
    Next
    
    ' Set the amount of columns and the maximum value
    usrControl.ColumnCount = WeekAmout + 1
    usrControl.MaxValue = lngHighest
    
    ' Now, present the gathered information
    For Tell = 0 To WeekAmout
    
        ' Set the name and value of this column
        usrControl.SetData Tell + 1, Tell + 1, Weeks(Tell)
    
    Next

    ' Redraw the control
    usrControl.RedrawChart

End Sub

Function GetDateOnly(dDate)

    GetDateOnly = DateSerial(Year(dDate), Month(dDate), Day(dDate))

End Function

Function WeekNum(dDate)

    WeekNum = DateDiff("ww", DateSerial(Year(dDate), 1, 1), dDate) + 1

End Function