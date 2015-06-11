Attribute VB_Name = "TestLinkedList"
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

'   LinkedList is build out of LinkedListNodes

Option Explicit

Function TestInsertFront() As String
    Dim list As New LinkedList
    list.InsertFront 5
    list.InsertFront 6
    
    Dim arr() As Variant
    arr = list.ToArray
    
    If arr(0) = 6 And arr(1) = 5 And list.Count = 2 Then
        TestInsertFront = "Passed"
    Else
        TestInsertFront = "Failed"
    End If
End Function

Function TestRemoveFront()
    Dim list As New LinkedList
    list.InsertFront 5
    list.InsertFront 6

    Dim v As Long
    v = list.RemoveFront
    If v = 6 And list.Count = 1 Then
        TestRemoveFront = "Passed"
    Else
        TestRemoveFront = "Failed"
    End If
End Function

Function TestCount()
    Dim list As New LinkedList
    
    list.InsertFront 5
    list.InsertFront 6
    
    If list.Count = 2 Then
        TestCount = "Passed"
    Else
        TestCount = "Failed"
    End If
End Function

Function TestInsertRear()
    Dim list As New LinkedList
    Dim i As Long
    For i = 0 To 9
        list.InsertRear i
    Next i
    
    Dim arr() As Variant
    arr = list.ToArray()
    
    Dim bPassed As Boolean
    bPassed = True
    For i = 0 To 9
        If arr(i) <> i Then
            bPassed = False
        End If
    Next i
    
    If bPassed Then
        TestInsertRear = "Passed"
    Else
        TestInsertRear = "Failed"
    End If
End Function


Sub TestAll()
    Debug.Print "Unit Test", , "Status"
    Debug.Print "----------------------------------"
    Debug.Print "TestInsertFront", TestInsertFront
    Debug.Print "TestRemoveFront", TestRemoveFront
    Debug.Print "TestCount", , TestCount
    Debug.Print "TestInsertRear", TestInsertRear
End Sub
