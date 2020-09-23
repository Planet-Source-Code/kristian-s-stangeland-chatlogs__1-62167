Attribute VB_Name = "modGlobal"
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

Public Function CountRecords(Database As Database, sTable As String) As Long

    Dim RecordCount As Recordset

   ' Open the recordset.
    Set RecordCount = Database.OpenRecordset("SELECT COUNT (*) FROM " & sTable)
        
    ' Return the result.
    CountRecords = Val(RecordCount.Fields(0))

End Function

Public Function IsUser(User As Object, aNickNames As Variant, strEmail As String) As Boolean
 
    Dim Tell As Long, NickName
    
    ' First, test the email
    If strEmail = User.Email And User.Email <> "" Then
    
        ' This user correspond, all right
        IsUser = True
    
    Else
    
        ' If not, we need to go through all nicknames
        For Each NickName In User.NickNames
    
            ' Now, go through each and every nicknames we're supposed to correspond
            For Tell = LBound(aNickNames) To UBound(aNickNames)
                
                ' See if the nickname correspond
                If InStr(1, NickName, aNickNames(Tell), vbTextCompare) <> 0 Then
                
                    ' Yup, this is the user
                    IsUser = True
                
                End If
            
            Next
    
        Next
    
    End If

End Function

Public Function JoinID(Collection As Collection, Delimiter As String)

    Dim Tell As Long
    
    ' Go through all elements
    For Tell = 1 To Collection.Count
    
        ' Add the ID of each of them
        JoinID = JoinID & Collection(Tell).Tag & IIf(Tell < Collection.Count, Delimiter, "")
    
    Next

End Function

Public Function JoinNames(Collection As Collection, Delimiter As String)

    Dim Tell As Long
    
    ' Go through all elements
    For Tell = 1 To Collection.Count
    
        ' Add the name of each of them
        JoinNames = JoinNames & Collection(Tell) & IIf(Tell < Collection.Count, Delimiter, "")
    
    Next

End Function

Public Function GetSessionDate(Session As Object) As Date

    ' Extract the date
    If Session.Events(1).IsMessage Then
        GetSessionDate = Session.Events(1).MessageDate
    Else
        GetSessionDate = Session.Events(1).EventDate
    End If

End Function
