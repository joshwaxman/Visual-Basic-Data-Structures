VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnionFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private arr() As Long
Private weight() As Long

Public Sub Init(size As Long)
    Debug.Assert size > 0
    If size < 1 Then
        Err.Raise -401, , "Size of UnionFind must be at least 1"
    End If
    
    ReDim arr(size - 1)
    ReDim weight(size - 1)
    Dim i As Long
    ' assign each item as its own parent
    For i = LBound(arr) To UBound(arr)
        arr(i) = i
        weight(i) = 1 ' one, since each is its own tree
    Next i
End Sub

Private Sub ShortenTree(ByVal p As Long, p_root As Long)
    ' assert that p_root is the root of p
    Debug.Assert p_root = Root(p)
    
    Dim n As Long
    While p <> p_root
        ' save arr(p) in a variable because about to overwrite
        n = arr(p)
        ' point p's parent directly to root
        arr(p) = p_root
        ' if p's old parent was not the root, then we should adjust
        ' the weight of p's parent since it no longer contains p
        If n <> p_root Then
            weight(n) = weight(n) - weight(p)
        End If
        ' p should point to its parent
        p = n
    Wend
End Sub

Public Sub Union(p As Long, q As Long)
    ' make sure that both p and q are within the bounds of the array
    If p < LBound(arr) Or q < LBound(arr) Or p > UBound(arr) Or q > UBound(arr) Then
        Err.Raise -4000, , "Parameter for union is outside the bounds"
    End If
    
    If p = q Then ' they are the same, so there is nothing to do
        Exit Sub
    End If
    
    ' check if p's root is the same as q's root
    Dim p_root As Long, q_root As Long
    p_root = Root(p)
    q_root = Root(q)
    
    ShortenTree p, p_root
    ShortenTree q, q_root

    If p_root = q_root Then ' they are already connected, so there is nothing to do
        Exit Sub
    End If
        
    If weight(p_root) < weight(q_root) Then
        arr(p_root) = q_root
        weight(q_root) = weight(q_root) + weight(p_root)
    Else
        arr(q_root) = p_root
        weight(p_root) = weight(p_root) + weight(q_root)
    End If
End Sub

Public Function Find(ByVal p As Long, ByVal q As Long) As Boolean
    ' make sure that both p and q are within the bounds of the array
    If p < LBound(arr) Or q < LBound(arr) Or p > UBound(arr) Or q > UBound(arr) Then
        Err.Raise -4000, , "Parameter for Find is outside the bounds"
    End If
    
    Dim p_root As Long, q_root As Long
    p_root = Root(p)
    q_root = Root(q)

    ' this is overkill, since already shortening on Union
    ShortenTree p, p_root
    ShortenTree q, q_root

    Find = p_root = q_root
End Function

Private Function Root(ByVal p As Long) As Long
    Debug.Assert p >= LBound(arr) And p <= UBound(arr)
    If p < LBound(arr) Or p > UBound(arr) Then
        Err.Raise -402, , "Parameter for Root is outside of bounds"
    End If

    While p <> arr(p)
        p = arr(p)
    Wend
    
    Root = p
End Function


Public Function ListConnected(ByVal p As Long) As Collection
    Dim c As New Collection
    Dim i As Long
    Dim p_root As Long, q_root As Long
    p_root = Root(p)
    ShortenTree p, p_root
    
    For i = LBound(arr) To UBound(arr)
        q_root = Root(i)
        ShortenTree i, q_root
        If q_root = p_root Then
            c.Add i
        End If
    Next i
    Set ListConnected = c
End Function

Private Sub Class_Initialize()
    Init 10
End Sub
