Attribute VB_Name = "modCollection"
Option Explicit
Option Compare Text
Public Sub dhc_Remove(ByRef cCol As Collection, ByRef item As Variant)
Dim k As Long, colItem As Variant

    For k = cCol.Count To 1 Step -1
        colItem = cCol(k)
        If VarType(colItem) = VarType(item) Then
            If colItem = item Then
                Call cCol.Remove(k)
                Exit For
            End If
        End If
    Next k

End Sub

Public Sub dhc_RemoveAll(ByRef cCol As Collection)
Dim k As Long

    For k = cCol.Count To 1 Step -1
        Call cCol.Remove(k)
    Next k

End Sub


Public Sub dhc_Add(ByRef cCol As Collection, ByRef item As Variant)
Dim k As Long

    For k = cCol.Count To 1 Step -1
        If cCol(k) = item Then Exit Sub
    Next k
    cCol.Add item

End Sub

Public Function dhc_Contains(ByRef cCol As Collection, ByRef item As Variant) As Boolean
' A helper function to detect the precense of items in a collection
Dim k As Long
    
    For k = cCol.Count To 1 Step -1
        If item = cCol(k) Then
            dhc_Contains = True
            Exit Function
        End If
    Next k

End Function
