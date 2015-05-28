Attribute VB_Name = "Module1"
Option Explicit

Sub Main()
    Dim UF As New UnionFind
    UF.Init 10
    
    UF.Union 1, 2
    UF.Union 1, 9
    
    Debug.Print UF.Find(2, 9)
    Debug.Print UF.Find(2, 7)
    
    Dim c As Collection
    Set c = UF.ListConnected(2)
    Debug.Print c.count '
    
    Set c = UF.ListConnected(3)
    Debug.Print c.count
    
    Debug.Print UF.ComponentSize(2)
End Sub


Sub TestStack()
    Dim s As New Stack
    s.Push ("hello")
    s.Push "goodbye"
    s.Push s
    Dim v As Object
    Set v = s.Pop
    Debug.Print v



End Sub
