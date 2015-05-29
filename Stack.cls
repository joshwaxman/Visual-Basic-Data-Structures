VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
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

Private arr() As Variant
Private item_count As Long

Private Sub Class_Initialize()
    ReDim arr(9) ' start with 10 elements
    item_count = 0
End Sub

Public Sub Push(v As Variant)
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

Public Function Pop() As Variant
    If item_count = 0 Then
        Err.Raise -405, , "Cannot pop an empty stack"
    End If
    
    item_count = item_count - 1
    
    If VarType(arr(item_count)) = vbObject Then
        Set Pop = arr(item_count)
    Else
        Pop = arr(item_count)
    End If
    
    Set arr(item_count) = Nothing
End Function

Public Function count() As Long
    count = item_count
End Function
