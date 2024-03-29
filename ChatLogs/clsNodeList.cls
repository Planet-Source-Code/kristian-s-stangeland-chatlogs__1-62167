VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNodeList"
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

' All nodes
Private Nodes() As Long

' The main object
Private pMain As clsDocument

Friend Sub Initialize(lpNodes() As Long, HTMLMain As clsDocument)
    
    ' Do not care about errors
    On Error Resume Next
    
    ' Only try to initialize array if the source does exist
    If Not (Not lpNodes) Then
    
        ' Reallocate node array
        ReDim Nodes(LBound(lpNodes) To UBound(lpNodes))
        
        ' Copy data
        CopyMemory Nodes(LBound(Nodes)), lpNodes(LBound(lpNodes)), (UBound(lpNodes) - LBound(lpNodes) + 1) * LenB(Nodes(LBound(Nodes)))

    Else
    
        ' Just erase array
        Erase Nodes
    
    End If

    ' Then make a reference to the main object
    Set pMain = HTMLMain

End Sub

Public Function IsEqual(refNode As clsNodeList) As Boolean

    Dim Tell As Long

    ' First, check the lenght
    If Lenght = refNode.Lenght Then
    
        ' Then check indexes
        For Tell = 0 To Lenght
        
            If Nodes(Tell) <> refNode.Item(Tell).NodeIndex Then
            
                ' Nope, the node is not equal
                Exit Function
            
            End If
        
        Next
    
        ' Yes, it is equal
        IsEqual = True

    End If

End Function

Public Function Item(ByVal Index As Long) As clsElement

    If Index >= 0 And Index <= Lenght Then
        ' Return the node
        Set Item = pMain.GetElementByIndex(Nodes(Index))
    End If

End Function

Public Property Get Lenght() As Long

    ' Return the amout of nodes in this list
    Lenght = UBound(Nodes)

End Property

Public Property Get NewEnum() As IEnumVARIANT

    Dim oEnumerator As clsEnumator
    
    ' Create a new enumator to use in the enumation-process
    Set oEnumerator = New clsEnumator
    
    ' Initialize enumator
    oEnumerator.InitializeEnumeration ObjPtr(Me), 0, 0
    
    ' Count this operation
    m_nCurrentEnumOperations = m_nCurrentEnumOperations + 1
    
    ' Return the enumerator object's IEnumVARIANTReDef interface
    Set NewEnum = oEnumerator
    
    ' Clean up
    Set oEnumerator = Nothing
  
End Property

Public Sub DecrementEnumCounter()

    ' We're finish with one operation
    m_nCurrentEnumOperations = m_nCurrentEnumOperations - 1

End Sub

Private Sub Class_Initialize()
  
    ' Add this collection to the list
    LookupList.AddPointerToLookupList ObjPtr(Me)
  
End Sub

Private Sub Class_Terminate()

    ' Remove this collection from the list
    LookupList.RemovePointerFromLookupList ObjPtr(Me)
  
End Sub

