VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinkedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   Copyright 2015 Joshua Waxman
'
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.

'   LinkedList is build out of LinkedListNodes

Dim Head As LinkedListNode
Dim Header As LinkedListNode
Dim Tail As LinkedListNode
Dim Trailer As LinkedListNode
Dim item_count As Long

Private Sub Class_Initialize()
    Set Header = New LinkedListNode
    Set Trailer = New LinkedListNode
    Set Header.NextNode = Trailer
    Set Head = Header
    Set Tail = Trailer
End Sub

Sub Assign(ByRef destination As Variant, ByRef source As Variant)
    If VarType(source) = vbObject Then
        Set destination = source
    Else
        destination = source
    End If
End Sub

Public Sub InsertFront(value As Variant, Optional key As Variant)
    Dim node As New LinkedListNode
    node.key = key
    If VarType(value) = vbObject Then
        Set node.value = value
    Else
        node.value = value
    End If
    
    InsertAfterNode node, Head
    
    item_count = item_count + 1
End Sub

Public Function RemoveFront() As Variant
    If item_count = 0 Then
        Err.Raise -407, , "Cannot remove an element of an empty linked list"
    End If
    
    Assign RemoveFront, Head.NextNode.value
    Set Head.NextNode = Head.NextNode.NextNode
    item_count = item_count - 1
End Function

' This is O(N)
Public Sub InsertRear(value As Variant, Optional key As Variant)
    ' traverse to one before the end
    Dim n As LinkedListNode
    Set n = Head
    While Not n.NextNode Is Trailer
        Set n = n.NextNode
    Wend
    
    Dim node As New LinkedListNode
    node.key = key
    If VarType(value) = vbObject Then
        Set node.value = value
    Else
        node.value = value
    End If

    InsertAfterNode node, n
    item_count = item_count + 1
End Sub

Private Sub InsertAfterNode(newNode As LinkedListNode, afterNode As LinkedListNode)
    Set newNode.NextNode = afterNode.NextNode
    Set afterNode.NextNode = newNode
End Sub

Public Function Count() As Long
    Count = item_count
End Function

Public Sub PrintList()
    Dim node As LinkedListNode
    Set node = Header.NextNode
    
    While Not node Is Trailer
        Debug.Print node.value
        Set node = node.NextNode
    Wend
End Sub

Public Function ToArray() As Variant()
    If item_count = 0 Then
        Err.Raise -409, , "Cannot convert empty linked list to array"
    End If
    
    Dim returnValue() As Variant
    ReDim returnValue(item_count - 1)
    Dim node As LinkedListNode
    Set node = Header.NextNode
    
    Dim i As Long
    While Not node Is Trailer
        If VarType(node.value) = vbObject Then
            Set returnValue(i) = node.value
        Else
            returnValue(i) = node.value
        End If
        
        Set node = node.NextNode
        i = i + 1
    Wend
    
    ToArray = returnValue
End Function
