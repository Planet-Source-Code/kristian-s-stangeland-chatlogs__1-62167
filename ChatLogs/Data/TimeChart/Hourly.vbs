Sub UpdateControl(usrControl, Sessions)

    Dim Hours(), lngHour, lngHighest, Tell, Message, Session
    
    ' Allocate the data-array containing the information to gather
    ReDim Hours(23)
    
    ' Go through all sessions
    For Each Session In Sessions

        ' Then go through all events
        For Each Message In Session.Events
            
            ' See if this is a message, and not an event
            If Message.IsMessage Then
            
                ' Get the hour this message was posted
                lngHour = Hour(Message.MessageDate)
            
                ' If so, add the message to the data-array
                Hours(lngHour) = Hours(lngHour) + 1
            
            End If

        Next

    Next
    
    ' Get the highest value
    For Tell = 0 To 23
        If Hours(Tell) > lngHighest Then
            lngHighest = Hours(Tell)
        End If
    Next
    
    ' Set the amount of columns and the maximum value
    usrControl.ColumnCount = 24
    usrControl.MaxValue = lngHighest
    
    ' Now, present the gathered information
    For Tell = 0 To 23
    
        ' Set the name and value of this column
        usrControl.SetData Tell + 1, Tell, Hours(Tell)
    
    Next

    ' Redraw the control
    usrControl.RedrawChart

End Sub