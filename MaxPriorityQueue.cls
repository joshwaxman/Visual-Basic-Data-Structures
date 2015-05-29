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

Private Sub Class_Initialize()
    ReDim keys(9)
    ReDim values(9)
    item_count = 0
End Sub

Public Sub Insert(ByVal k As Variant, ByVal v As Variant)
    ' check if the array is full, in which case need to double it
    If item_count > UBound(keys) Then
        ReDim Preserve keys(item_count * 2)
        ReDim Preserve values(item_count * 2)
    End If
    
    If VarType(k) = vbObject Then
        Set keys(item_count) = k
    Else
        keys(item_count) = k
    End If
    
    If VarType(v) = vbObject Then
        Set values(item_count) = v
    Else
        values(item_count) = v
    End If
    
    item_count = item_count + 1
    ' the insert may have caused a violation in the heap property
    Upheap
End Sub

Private Sub Swap(ByVal p As Long, ByVal q As Long)
    ' assign key and value of p to temporaries
    Dim k As Variant, v As Variant
    k = keys(p)
    If VarType(values(p)) = vbObject Then
        Set v = values(p)
    Else
        v = values(p)
    End If
    
    ' set array(p) = array(q)
    keys(p) = keys(q)
    If VarType(values(q)) = vbObject Then
        Set values(p) = values(q)
    Else
        values(p) = values(q)
    End If
    
    ' set array(q) = temporary
    keys(q) = k
    If VarType(v) = vbObject Then
        Set values(q) = v
    Else
        values(q) = v
    End If
End Sub

Private Sub Upheap()
    If item_count = 0 Then ' cannot upheap an empty heap
        Exit Sub
    End If

    ' this is a maxheap, so at every point, the parent should be larger than its two children
    Dim position As Long, parent As Long
    Dim max_child As Long, min_child As Long
    position = item_count - 1
    
    While position <> 0
        parent = position / 2
        If keys(position) > keys(parent) Then
            ' find max of siblings
            max_child = IIf(keys(parent * 2) >= keys(parent * 2 + 1), parent * 2, parent * 2 + 1)
        
            ' swap parent with max_child
            Swap parent, max_child
            
            ' and move up the heap
            position = parent
        End If
        
    Wend
    
End Sub
