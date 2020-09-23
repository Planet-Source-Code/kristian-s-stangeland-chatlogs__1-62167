VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Copyright (C) 2004 Kristian. S.Stangeland

'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

' How many enum operations that is currently under operation on this collection
Private m_nCurrentEnumOperations As Long

' A pointer to the array
Private aElements() As Element
Private ElementCount As Long

' Our parent object
Private pSession As clsSession

Public Function Count() As Long
    
    ' Calculate the amount of elements
    Count = ElementCount

End Function

Public Function Item(ByVal Index As Long) As Object
Attribute Item.VB_UserMemId = 0

    Dim oMessage As clsMessage, oEvent As clsEvent

    ' See if the index is valid
    If Index < 1 Or Index > ElementCount Then
        ' Ignore the call if not
        Exit Function
    End If

    ' Firstly, check the type of this element
    If aElements(Index).IsMessage Then
    
        ' Create a new message-class
        Set oMessage = New clsMessage
        
        ' Set the index
        oMessage.MessageID = Index
        
        ' Initialize the class
        oMessage.InitializeClass aElements, pSession
        
        ' Return the object
        Set Item = oMessage

    Else
    
        ' Create a new event-class
        Set oEvent = New clsEvent
        
        ' Set the index
        oEvent.EventID = Index
    
        ' Initialize the class
        oEvent.InitializeClass aElements, pSession
        
        ' Return the object
        Set Item = oEvent
    
    End If

End Function

Public Property Get NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Dim oEnumerator As clsEnumator
    
    ' Create a new enumator to use in the enumation-process
    Set oEnumerator = New clsEnumator
    
    ' Initialize enumator
    oEnumerator.InitializeEnumeration ObjPtr(Me), 1, 1
    
    ' Count this operation
    m_nCurrentEnumOperations = m_nCurrentEnumOperations + 1
    
    ' Return the enumerator object's IEnumVARIANTReDef interface
    Set NewEnum = oEnumerator
    
    ' Clean up
    Set oEnumerator = Nothing
  
End Property

Friend Sub InitializeClass(refElements() As Element, oSession As clsSession, lngCount As Long)

    ' Make a reference to the array filled with the elements
    CopyMemory ByVal VarPtrArray(aElements), ByVal VarPtrArray(refElements), 4

    ' Make a reference to the parent session
    Set pSession = oSession
    
    ' Set the count-variable
    ElementCount = lngCount

End Sub

Public Sub DecrementEnumCounter()

    ' We're finish with one operation
    m_nCurrentEnumOperations = m_nCurrentEnumOperations - 1

End Sub

Private Sub Class_Initialize()
  
    ' Add this collection to the list
    LookupList.AddPointerToLookupList ObjPtr(Me)
  
End Sub

Private Sub Class_Terminate()

    ' Remove the pointer to the array
    ZeroMemory ByVal VarPtrArray(aElements), 4

    ' Remove this collection from the list
    LookupList.RemovePointerFromLookupList ObjPtr(Me)
  
End Sub

