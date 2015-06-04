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

'   The implementation is based on a Stack. Except rather than pop, there
'   is the Dequeue function. We swap the element at the top of the
'   stack with a random element, and then return it.


Option Explicit

Private arr() As Variant
Private item_count As Long

Private Sub Class_Initialize()
    ReDim arr(9) ' start with 10 elements
    item_count = 0
End Sub

Public Sub Enqueue(v As Variant)
    ' if item_count is beyond the bounds of the array, then double it
    If item_count > UBound(arr) Then
        ReDim Preserve arr(item_count * 2)
    End If
    
    If VarType(v) = vbObject Then
        Set arr(item_count) = v
    Else
        arr(item_count) = v
    End If
    
    item_count = item_count + 1
End Sub

Public Function Dequeue() As Variant
    If item_count = 0 Then
        Err.Raise -405, , "Cannot Dequeue from an empty RandomQueue"
    End If
    
    item_count = item_count - 1
    
    Dim random_item As Long
    random_item = Int(Rnd * item_count)
    Swap random_item, item_count
    AssignFromPosition Dequeue, item_count
    
    arr(item_count) = Empty
End Function

Public Function Count() As Long
    Count = item_count
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



