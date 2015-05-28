VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
    
    If VarType(arr(position)) = vbObject Then
        Set Dequeue = arr(position)
    Else
        Dequeue = arr(position)
    End If
    
    Set arr(position) = Nothing
    item_count = item_count - 1
End Function

Public Function count() As Long
    count = item_count
End Function

Public Function Peek() As Variant
    If item_count = 0 Then
        Err.Raise -405, , "Cannot peek into an empty queue"
    End If
        
    ' wrap around the circular array
    Dim position As Long
    position = start_element + item_count - 1
    If position = -1 Then
        position = UBound(arr)
    End If
    
    If VarType(arr(position)) = vbObject Then
        Set Peek = arr(position)
    Else
        Peek = arr(position)
    End If
End Function

Public Function ToCollection() As Collection
    Dim c As New Collection
    Dim i As Long, count As Long
    
    i = start_element
    While count < item_count
        c.Add arr(i)
        i = (i + 1) Mod (UBound(arr) + 1)
        count = count + 1
    Wend
    
    Set ToCollection = c
End Function


Public Function ToArray() As Variant()
    Dim c As Collection
    Set c = ToCollection
    If c.count = 0 Then
        Exit Function
    End If
    
    Dim retVal() As Variant
    ReDim retVal(c.count - 1)
    Dim v As Variant, i As Long
    
    For Each v In c
        If VarType(c.Item(i)) = vbObject Then
            Set retVal(i) = c.Item(i)
        Else
            retVal(i) = c.Item(i)
        End If
    Next v
    
    Set ToArray = retVal
End Function
