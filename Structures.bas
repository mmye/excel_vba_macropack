Attribute VB_Name = "Structures"
Function CreateCollection() As Collection
'Create native VBA Collection
    Set CreateCollection = New Collection
End Function
Function CreateDictionary() As Object
'Supported methods:
'.Exists(Key)
'.Count
'.Add(Key, Val)
'.Remove(Key)
'.RemoveAll
'.Keys
'.Items
    Set CreateDictionary = CreateObject("Scripting.Dictionary")
End Function
Function CreateArrayList() As Object
'Supported methods:
'.Count
'.Add(Val)
'.Remove(Val)
    Set CreateArrayList = CreateObject("System.Collections.ArrayList")
End Function
Function CreateQueue() As Object
'Supported methods:
'.Count
'.Contains(Val)
'.Enqueue(Val)
'.Peek
'.Dequeue
'.Clear
    Set CreateQueue = CreateObject("System.Collections.Queue")
End Function
Function CreateStack() As Object
'Supported methods:
'.Count
'.Contains(Val)
'.Push(Val)
'.Peek
'.Pop
'.Clear
    Set CreateStack = CreateObject("System.Collections.Stack")
End Function
Function CreateSortedList() As Object
'Supported methods:
'.Add(Key, Val)
'.ContainsKey(Key)
'.ContainsValue(Val)
'.GetKey(Index)
'.GetValue(Index)
'.Clear
    Set CreateSortedList = CreateObject("System.Collections.SortedList")
End Function
Function CreateHashTable() As Object
'Supported methods:
'.Add(Key, Val)
'.ContainsKey(Key)
'.ContainsValue(Val)
'.GetKey(Index)
'.GetValue(Index)
'.Clear
    Set CreateHashTable = CreateObject("System.Collections.HashTable")
End Function

