VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MaxPriorityQueue"
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

' The MaxPriorityQueue will be based on a heap, based on an array
Option Explicit
Dim keys() As Variant
Dim values() As Variant
Dim item_count As Long

Public Function Count() As Long
    Count = item_count
End Function

Private Function IsValid(ByVal position As Long) As Boolean
    If position > UBound(keys) Then Exit Function
    If IsEmpty(keys(position)) Then Exit Function
    IsValid = True
End Function

Private Function IsLeaf(ByVal position As Long) As Boolean
    ' it is a leaf if both left and right children aren't valid
    ' since heap grows to the right, if left child not valid
    ' then right is not valid either
    IsLeaf = Not IsValid(position * 2)
End Function

Private Sub Class_Initialize()
    ' we will waste array(0) to make calculations easier
    ' and the upper bound will always be an even number
    ReDim keys(10)
    ReDim values(10)
    item_count = 0
End Sub

Private Sub AssignToPosition(ByVal position As Long, ByVal v As Variant)
    If VarType(v) = vbObject Then
        Set values(position) = v
    Else
        values(position) = v
    End If
End Sub

Private Sub AssignFromPosition(ByRef v As Variant, ByVal position As Long)
    If VarType(values(position)) = vbObject Then
        Set v = values(position)
    Else
        v = values(position)
    End If
End Sub


Public Sub Insert(ByVal k As Variant, ByVal v As Variant)
    ' check if the array is full, in which case need to double it
    If item_count = UBound(keys) Then
        ReDim Preserve keys(item_count * 2)
        ReDim Preserve values(item_count * 2)
    End If
    
    item_count = item_count + 1
    
    ' keys must be of type variant
    keys(item_count) = k
    AssignToPosition item_count, v

    ' the insert may have caused a violation in the heap property
    Upheap
End Sub

Public Function RemoveMax() As Variant
    ' check that the PriorityQueue is not empty
    If item_count = 0 Then
        Err.Raise -406, , "Cannot remove max from empty Priority Queue"
    End If
    
' swap last element of heap with first (top) element
    ' save and clear value now at last position in heap
    ' then downheap
    Swap 1, item_count
    AssignFromPosition RemoveMax, item_count
    
    values(item_count) = Empty
    keys(item_count) = Empty
    
    Downheap
End Function

Private Sub Swap(ByVal p As Long, ByVal q As Long)
    ' assign key and value of p to temporaries
    Dim k As Variant, v As Variant
    k = keys(p)
    AssignFromPosition v, p
    
    ' set array(p) = array(q)
    keys(p) = keys(q)
    AssignToPosition p, values(q)
    
    ' set array(q) = temporary
    keys(q) = k
    AssignToPosition q, v
End Sub

Private Function GetMaxChild(ByVal position As Long) As Long
    Dim right_child As Long, left_child As Long
    
    ' find max of siblings
    left_child = position * 2
    right_child = left_child + 1
    
    ' heap grows to right so only right child might be null or off the end of the array
    If Not IsValid(right_child) Then
        GetMaxChild = left_child
    Else
        GetMaxChild = IIf(keys(left_child) >= keys(right_child), left_child, right_child)
    End If
End Function

Private Sub Upheap()
    If item_count = 0 Then ' cannot upheap an empty heap
        Exit Sub
    End If

    ' this is a maxheap, so at every point, the parent should be larger than its two children
    Dim position As Long, parent As Long
    Dim max_child As Long
    position = item_count
    
    While position > 1
        parent = position \ 2 ' integer division
        If keys(position) > keys(parent) Then ' heap condition violated, must fix
            ' calculate the max of the two singlings
            max_child = GetMaxChild(parent)
            
            ' swap parent with max_child
            Swap parent, max_child
            
            ' and move up the heap
            position = parent
        Else ' heap condition fulfilled, so we can end right here
            Exit Sub
        End If
    Wend
End Sub

Private Sub Downheap()
    If item_count = 0 Then ' nothing to do since empty
        Exit Sub
    End If
    
    ' this is a maxheap, so at every point, the parent should be larger than its two children
    Dim position As Long, parent As Long
    Dim max_child As Long, left_child As Long
    position = 1
    
    While Not IsLeaf(position) ' if we are at a leaf, then no more downheap
        ' otherwise, see if either of the two children is larger than parent
        max_child = GetMaxChild(position)
        If keys(max_child) > keys(position) Then
            Swap max_child, position
            position = max_child
        Else ' heap condition satisfied
            Exit Sub
        End If
    Wend
End Sub
