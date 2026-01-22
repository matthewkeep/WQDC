Option Explicit
' DictionaryShim: lightweight stand-in for Scripting.Dictionary.
' Provides Mac-compatible dictionary functionality.

Private mKeys() As String
Private mValues() As Variant
Private mCount As Long

Private Sub Class_Initialize()
    mCount = 0
    ReDim mKeys(0)
    ReDim mValues(0)
End Sub

Public Function Exists(ByVal key As Variant) As Boolean
    Exists = (KeyIndex(key) > 0)
End Function

Public Sub Add(ByVal key As Variant, ByVal value As Variant)
    If Exists(key) Then
        Err.Raise 457, "DictionaryShim", "Key already exists"
    End If

    mCount = mCount + 1
    If mCount > UBound(mKeys) Then
        ReDim Preserve mKeys(mCount * 2)
        ReDim Preserve mValues(mCount * 2)
    End If

    mKeys(mCount) = CStr(key)
    If IsObject(value) Then
        Set mValues(mCount) = value
    Else
        mValues(mCount) = value
    End If
End Sub

Public Property Get Item(ByVal key As Variant) As Variant
    Dim idx As Long
    idx = KeyIndex(key)
    If idx = 0 Then Err.Raise 5, "DictionaryShim", "Key not found"
    If IsObject(mValues(idx)) Then
        Set Item = mValues(idx)
    Else
        Item = mValues(idx)
    End If
End Property

Public Property Let Item(ByVal key As Variant, ByVal value As Variant)
    Dim idx As Long
    idx = KeyIndex(key)
    If idx = 0 Then
        Add CStr(key), value
    Else
        mValues(idx) = value
    End If
End Property

Public Property Set Item(ByVal key As Variant, ByVal value As Object)
    Dim idx As Long
    idx = KeyIndex(key)
    If idx = 0 Then
        Add CStr(key), value
    Else
        Set mValues(idx) = value
    End If
End Property

Public Property Get Count() As Long
    Count = mCount
End Property

Public Function Keys() As Variant
    Dim arr() As String, idx As Long
    If mCount = 0 Then Exit Function
    ReDim arr(0 To mCount - 1)
    For idx = 1 To mCount
        arr(idx - 1) = mKeys(idx)
    Next idx
    Keys = arr
End Function

Public Function Items() As Variant
    Dim arr() As Variant, idx As Long
    If mCount = 0 Then Exit Function
    ReDim arr(0 To mCount - 1)
    For idx = 1 To mCount
        If IsObject(mValues(idx)) Then
            Set arr(idx - 1) = mValues(idx)
        Else
            arr(idx - 1) = mValues(idx)
        End If
    Next idx
    Items = arr
End Function

Public Sub RemoveAll()
    mCount = 0
    ReDim mKeys(0)
    ReDim mValues(0)
End Sub

Private Function KeyIndex(ByVal key As Variant) As Long
    Dim txt As String, idx As Long
    txt = CStr(key)
    For idx = 1 To mCount
        If mKeys(idx) = txt Then
            KeyIndex = idx
            Exit Function
        End If
    Next idx
End Function
