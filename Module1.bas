Attribute VB_Name = "Module1"
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
    Debug.Print c.Count '
    
    Set c = UF.ListConnected(3)
    Debug.Print c.Count
    
    Debug.Print UF.ComponentSize(2)
End Sub


Sub TestMaxPriorityQueue()
    Dim q As New MaxPriorityQueue

    q.Insert 6, 6
    q.Insert 7, 7
    q.Insert 2, 2
    q.Insert 4, 4
    q.Insert 14, "boo"
    
    Dim d As Variant
    d = q.RemoveMax
    Debug.Print d
    
    d = q.RemoveMax
    Debug.Print d
    
    d = q.RemoveMax
    Debug.Print d
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

Sub TestRandomQueue()
    Dim i As Long, v As Long
    Dim rq As New RandomQueue
    For i = 1 To 100
        rq.Enqueue i
    Next i
    
    While rq.Count <> 0
        v = rq.Dequeue
        Debug.Print v
    Wend
    
End Sub
