VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RandomQueue"
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

'   Dequeueing from a RandomQueue will return a **randomly** chosen element
'   of the Queue. Even though it is not removing from the front or back
'   of the internal array, the asymptotic running time is O(1)

'   The implementation is identical to that of a regular Queue, except for
'   the Dequeue function. There, we swap the element at the front of the
'   queue with a random element, and then return it.

'   Peek does not make sense, so it is removed. For now, so is ToArray
'   and ToCollection.

Option Explicit

Private arr() As Variant
Private item_count As Long
Private start_element As Long

Private Sub Class_Initialize()
    ReDim arr(9) ' start with 10 elements
    item_count = 0
    start_element = 0
End Sub

Public Sub Enqueue(v As Variant)
    ' if item_count is beyond the bounds of the array, then double it
    If item_count > UBound(arr) Then
        ReDim Preserve arr(item_count * 2)
    End If
    
    ' wrap around the circular array
    Dim position As Long
    position = (start_element + item_count) Mod (UBound(arr) + 1)
    
    
    If VarType(v) = vbObject Then
        Set arr(position) = v
    Else
        arr(position) = v
    End If
    
    item_count = item_count + 1
End Sub

Public Function Dequeue() As Variant
    If item_count = 0 Then
        Err.Raise -405, , "Cannot dequeue from an empty queue"
    End If
        
    ' wrap around the circular array
    Dim position As Long
    position = start_element + item_count - 1
    If position = -1 Then
        position = UBound(arr)
    End If
    
    ' choose a random element
    Dim random_item As Long
    random_item = Int(Rnd * item_count) Mod item_count
    random_item = start_element + random_item Mod (UBound(arr) + 1)
    
    Swap random_item, position
    AssignFromPosition Dequeue, position
    
    arr(position) = Empty
    item_count = item_count - 1
End Function

Private Sub AssignFromPosition(ByRef v As Variant, ByVal position As Long)
    If VarType(arr(position)) = vbObject Then
        Set v = arr(position)
    Else
        v = arr(position)
    End If
End Sub

Private Sub AssignToPosition(ByVal position As Long, ByVal v As Variant)
    If VarType(v) = vbObject Then
        Set arr(position) = v
    Else
        arr(position) = v
    End If
End Sub

Private Sub Swap(ByVal p As Long, ByVal q As Long)
    ' assign p to temporary
    Dim v As Variant
    AssignFromPosition v, p
    
    ' set array(p) = array(q)
    AssignToPosition p, arr(q)
    
    ' set array(q) = temporary
    AssignToPosition q, v
End Sub


Public Function Count() As Long
    Count = item_count
End Function

'Public Function Peek() As Variant
'    If item_count = 0 Then
'        Err.Raise -405, , "Cannot peek into an empty queue"
'    End If
'
'    ' wrap around the circular array
'    Dim position As Long
'    position = start_element + item_count - 1
'    If position = -1 Then
'        position = UBound(arr)
'    End If
'
'    If VarType(arr(position)) = vbObject Then
'        Set Peek = arr(position)
'    Else
'        Peek = arr(position)
'    End If
'End Function
'
'Public Function ToCollection() As Collection
'    Dim c As New Collection
'    Dim i As Long, Count As Long
'
'    i = start_element
'    While Count < item_count
'        c.Add arr(i)
'        i = (i + 1) Mod (UBound(arr) + 1)
'        Count = Count + 1
'    Wend
'
'    Set ToCollection = c
'End Function
'
'
'Public Function ToArray() As Variant()
'    Dim c As Collection
'    Set c = ToCollection
'    If c.Count = 0 Then
'        Exit Function
'    End If
'
'    Dim retVal() As Variant
'    ReDim retVal(c.Count - 1)
'    Dim v As Variant, i As Long
'
'    For Each v In c
'        If VarType(c.Item(i)) = vbObject Then
'            Set retVal(i) = c.Item(i)
'        Else
'            retVal(i) = c.Item(i)
'        End If
'    Next v
'
'    Set ToArray = retVal
'End Function

