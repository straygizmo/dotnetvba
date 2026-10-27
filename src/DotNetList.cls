VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DotNetList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' System.Collections.Generic.List<T> over Variant items. Obtain one with Net.NewList().
' Indexes are 0-based. Item is the default member and For Each is supported.

Private Const INITIAL_CAPACITY As Long = 4

Private mItems() As Variant
Private mCount As Long
' Snapshot that For Each enumerates. A Collection enumerator does not keep its
' Collection alive, so the snapshot is held here and rebuilt only after a change,
' which keeps nested For Each loops over an unchanged list working.
Private mEnumItems As Collection
Private mEnumStale As Boolean

Private Sub Class_Initialize()
    ReDim mItems(0 To INITIAL_CAPACITY - 1)
End Sub

Public Property Get Count() As Long
    Count = mCount
End Property

Public Property Get Item(ByVal index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    RequireIndex index, "Item"
    If IsObject(mItems(index)) Then
        Set Item = mItems(index)
    Else
        Item = mItems(index)
    End If
End Property

Public Property Let Item(ByVal index As Long, ByVal value As Variant)
    RequireIndex index, "Item"
    mItems(index) = value
    mEnumStale = True
End Property

Public Property Set Item(ByVal index As Long, ByVal value As Variant)
    RequireIndex index, "Item"
    Set mItems(index) = value
    mEnumStale = True
End Property

Public Sub Add(ByVal value As Variant)
    EnsureCapacity mCount + 1
    AssignVariant mItems(mCount), value
    mCount = mCount + 1
    mEnumStale = True
End Sub

' Accepts a DotNetList or a 1-D array. Adding a list to itself is supported.
Public Sub AddRange(ByVal items As Variant)
    Dim snapshot As Variant
    Dim i As Long
    If TypeName(items) = "DotNetList" Then
        snapshot = items.ToArray()
    ElseIf IsArray(items) Then
        snapshot = items
        If ArrayRank(snapshot) > 1 Then
            ThrowDotNet DotNetArgument, "DotNetList.AddRange", "The array must be one-dimensional. (Parameter 'items')"
        End If
    Else
        ThrowDotNet DotNetArgument, "DotNetList.AddRange", "The value must be a DotNetList or an array. (Parameter 'items')"
    End If
    If ArrayRank(snapshot) = 0 Then Exit Sub
    For i = LBound(snapshot) To UBound(snapshot)
        Add snapshot(i)
    Next i
End Sub

' index may equal Count (appends).
Public Sub Insert(ByVal index As Long, ByVal value As Variant)
    Dim i As Long
    If index < 0 Or index > mCount Then
        ThrowDotNet DotNetArgumentOutOfRange, "DotNetList.Insert", "Index must be within the bounds of the List. (Parameter 'index')"
    End If
    EnsureCapacity mCount + 1
    For i = mCount To index + 1 Step -1
        AssignVariant mItems(i), mItems(i - 1)
    Next i
    AssignVariant mItems(index), value
    mCount = mCount + 1
    mEnumStale = True
End Sub

Public Function Remove(ByVal value As Variant) As Boolean
    Dim index As Long
    index = IndexOf(value)
    If index < 0 Then Exit Function
    RemoveSpan index, 1
    Remove = True
End Function

Public Sub RemoveAt(ByVal index As Long)
    RequireIndex index, "RemoveAt"
    RemoveSpan index, 1
End Sub

Public Sub RemoveRange(ByVal index As Long, ByVal itemCount As Long)
    RequireRange index, itemCount, "RemoveRange"
    RemoveSpan index, itemCount
End Sub

' Removes all items but keeps the capacity.
Public Sub Clear()
    ReDim mItems(0 To UBound(mItems))
    mCount = 0
    Set mEnumItems = Nothing
    mEnumStale = True
End Sub

Public Function Contains(ByVal value As Variant) As Boolean
    Contains = (IndexOf(value) >= 0)
End Function

Public Function IndexOf(ByVal value As Variant) As Long
    Dim i As Long
    For i = 0 To mCount - 1
        If ItemsEqual(mItems(i), value) Then
            IndexOf = i
            Exit Function
        End If
    Next i
    IndexOf = -1
End Function

Public Function LastIndexOf(ByVal value As Variant) As Long
    Dim i As Long
    For i = mCount - 1 To 0 Step -1
        If ItemsEqual(mItems(i), value) Then
            LastIndexOf = i
            Exit Function
        End If
    Next i
    LastIndexOf = -1
End Function

' Returns a shallow copy of a range of items as a new list.
Public Function GetRange(ByVal index As Long, ByVal itemCount As Long) As DotNetList
    Dim result As DotNetList
    Dim i As Long
    RequireRange index, itemCount, "GetRange"
    Set result = New DotNetList
    For i = index To index + itemCount - 1
        result.Add mItems(i)
    Next i
    Set GetRange = result
End Function

' Stable merge sort. Strings compare ordinally (vbBinaryCompare), other values numerically;
' in a mix of strings and non-strings, the non-strings (numbers) sort before the strings.
' Lists containing objects cannot be sorted. The list is unchanged if Sort raises.
Public Sub Sort()
    Dim work() As Variant
    Dim buffer() As Variant
    Dim i As Long
    If mCount < 2 Then Exit Sub
    For i = 0 To mCount - 1
        If IsObject(mItems(i)) Then
            ThrowDotNet DotNetInvalidOperation, "DotNetList.Sort", "Failed to compare two elements in the array: objects cannot be compared."
        End If
    Next i
    ReDim work(0 To mCount - 1)
    ReDim buffer(0 To mCount - 1)
    For i = 0 To mCount - 1
        work(i) = mItems(i)
    Next i
    MergeSort work, buffer, 0, mCount - 1
    For i = 0 To mCount - 1
        mItems(i) = work(i)
    Next i
    mEnumStale = True
End Sub

Public Sub Reverse()
    Dim lo As Long
    Dim hi As Long
    Dim temp As Variant
    lo = 0
    hi = mCount - 1
    Do While lo < hi
        AssignVariant temp, mItems(lo)
        AssignVariant mItems(lo), mItems(hi)
        AssignVariant mItems(hi), temp
        lo = lo + 1
        hi = hi - 1
    Loop
    mEnumStale = True
End Sub

' Returns a 0-based Variant array; an empty list returns an array with UBound -1.
Public Function ToArray() As Variant()
    Dim result() As Variant
    Dim i As Long
    If mCount = 0 Then
        ToArray = Array()
        Exit Function
    End If
    ReDim result(0 To mCount - 1)
    For i = 0 To mCount - 1
        AssignVariant result(i), mItems(i)
    Next i
    ToArray = result
End Function

' Like String.Join(separator, list). Null joins as "", objects and arrays raise InvalidOperation.
Public Function Join(Optional ByVal separator As String = ",") As String
    Dim parts() As String
    Dim i As Long
    If mCount = 0 Then Exit Function
    ReDim parts(0 To mCount - 1)
    For i = 0 To mCount - 1
        parts(i) = ItemText(i)
    Next i
    Join = VBA.Join(parts, separator)
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Dim i As Long
    If mEnumItems Is Nothing Or mEnumStale Then
        Set mEnumItems = New Collection
        For i = 0 To mCount - 1
            mEnumItems.Add mItems(i)
        Next i
        mEnumStale = False
    End If
    Set NewEnum = mEnumItems.[_NewEnum]
End Property

' ===== Private helpers =====

Private Sub AssignVariant(ByRef target As Variant, ByVal value As Variant)
    If IsObject(value) Then
        Set target = value
    Else
        target = value
    End If
End Sub

Private Sub EnsureCapacity(ByVal required As Long)
    Dim capacity As Long
    capacity = UBound(mItems) + 1
    If required <= capacity Then Exit Sub
    Do While capacity < required
        capacity = capacity * 2
    Loop
    ReDim Preserve mItems(0 To capacity - 1)
End Sub

Private Sub RequireIndex(ByVal index As Long, ByVal memberName As String)
    If index < 0 Or index >= mCount Then
        ThrowDotNet DotNetArgumentOutOfRange, "DotNetList." & memberName, "Index was out of range. Must be non-negative and less than the size of the collection. (Parameter 'index')"
    End If
End Sub

Private Sub RequireRange(ByVal index As Long, ByVal itemCount As Long, ByVal memberName As String)
    If index < 0 Then
        ThrowDotNet DotNetArgumentOutOfRange, "DotNetList." & memberName, "Non-negative number required. (Parameter 'index')"
    End If
    If itemCount < 0 Then
        ThrowDotNet DotNetArgumentOutOfRange, "DotNetList." & memberName, "Non-negative number required. (Parameter 'itemCount')"
    End If
    If index > mCount - itemCount Then
        ThrowDotNet DotNetArgumentOutOfRange, "DotNetList." & memberName, "Offset and length were out of bounds for the list. (Parameter 'itemCount')"
    End If
End Sub

Private Sub RemoveSpan(ByVal index As Long, ByVal itemCount As Long)
    Dim i As Long
    If itemCount = 0 Then Exit Sub
    For i = index To mCount - itemCount - 1
        AssignVariant mItems(i), mItems(i + itemCount)
    Next i
    For i = mCount - itemCount To mCount - 1
        mItems(i) = Empty
    Next i
    mCount = mCount - itemCount
    mEnumStale = True
End Sub

' 0 = not an allocated array, otherwise the number of dimensions.
Private Function ArrayRank(ByRef arr As Variant) As Long
    Dim rank As Long
    Dim probe As Long
    On Error GoTo Done
    Do
        probe = UBound(arr, rank + 1)
        rank = rank + 1
    Loop
Done:
    ArrayRank = rank
End Function

' Objects compare by identity; a string never equals a non-string; other values use "=".
Private Function ItemsEqual(ByRef a As Variant, ByRef b As Variant) As Boolean
    Dim aIsObject As Boolean
    Dim bIsObject As Boolean
    aIsObject = IsObject(a)
    bIsObject = IsObject(b)
    If aIsObject And bIsObject Then
        ItemsEqual = (a Is b)
        Exit Function
    End If
    If aIsObject Or bIsObject Then Exit Function
    If (VarType(a) = vbString) <> (VarType(b) = vbString) Then Exit Function
    On Error Resume Next
    ItemsEqual = (a = b)
    On Error GoTo 0
End Function

Private Sub MergeSort(ByRef work() As Variant, ByRef buffer() As Variant, ByVal lo As Long, ByVal hi As Long)
    Dim midIndex As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    If lo >= hi Then Exit Sub
    midIndex = (lo + hi) \ 2
    MergeSort work, buffer, lo, midIndex
    MergeSort work, buffer, midIndex + 1, hi
    i = lo
    j = midIndex + 1
    k = lo
    Do While i <= midIndex And j <= hi
        If CompareItems(work(i), work(j)) <= 0 Then
            buffer(k) = work(i)
            i = i + 1
        Else
            buffer(k) = work(j)
            j = j + 1
        End If
        k = k + 1
    Loop
    Do While i <= midIndex
        buffer(k) = work(i)
        i = i + 1
        k = k + 1
    Loop
    Do While j <= hi
        buffer(k) = work(j)
        j = j + 1
        k = k + 1
    Loop
    For k = lo To hi
        work(k) = buffer(k)
    Next k
End Sub

' Strings sort after non-strings; strings compare ordinally, other values with < and >.
Private Function CompareItems(ByRef a As Variant, ByRef b As Variant) As Long
    Dim aIsString As Boolean
    Dim bIsString As Boolean
    aIsString = (VarType(a) = vbString)
    bIsString = (VarType(b) = vbString)
    If aIsString And bIsString Then
        CompareItems = StrComp(a, b, vbBinaryCompare)
    ElseIf aIsString Then
        CompareItems = 1
    ElseIf bIsString Then
        CompareItems = -1
    Else
        On Error GoTo Fail
        If a < b Then
            CompareItems = -1
        ElseIf a > b Then
            CompareItems = 1
        End If
    End If
    Exit Function
Fail:
    ThrowDotNet DotNetInvalidOperation, "DotNetList.Sort", "Failed to compare two elements in the array."
End Function

Private Function ItemText(ByVal index As Long) As String
    If IsObject(mItems(index)) Then
        ThrowDotNet DotNetInvalidOperation, "DotNetList.Join", "Objects cannot be converted to text."
    End If
    If IsNull(mItems(index)) Then Exit Function
    On Error GoTo Fail
    ItemText = CStr(mItems(index))
    Exit Function
Fail:
    ThrowDotNet DotNetInvalidOperation, "DotNetList.Join", "The element at index " & index & " cannot be converted to text."
End Function
