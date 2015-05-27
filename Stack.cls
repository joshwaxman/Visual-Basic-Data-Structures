VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Public Function Count() As Long
    Count = item_count
End Function
