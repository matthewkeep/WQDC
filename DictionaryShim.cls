Option Explicit
' DictionaryShim: lightweight stand-in for Scripting.Dictionary.

Private mKeys() As String
Private mValues() As Variant
Private mCount As Long
Private mCompareMode As VbCompareMethod

Private Sub Class_Initialize()
    mCount = 0
    mCompareMode = vbBinaryCompare
    ReDim mKeys(0)
    ReDim mValues(0)
End Sub

Public Property Get CompareMode() As VbCompareMethod
    CompareMode = mCompareMode
End Property

Public Property Let CompareMode(ByVal mode As VbCompareMethod)
    mCompareMode = mode
End Property

Public Function Exists(ByVal key As Variant) As Boolean
    Exists = (KeyIndex(key) > 0)
End Function

Public Sub SetObjectItem(ByVal key As Variant, ByVal value As Object)
    Dim idx As Long
    Dim keyStr As String
    keyStr = CStr(key)
    idx = KeyIndex(key)
    If idx = 0 Then
        ' Add new
        mCount = mCount + 1
        If mCount > UBound(mKeys) Then
            ReDim Preserve mKeys(mCount * 2)
            ReDim Preserve mValues(mCount * 2)
        End If
        mKeys(mCount) = keyStr
        Set mValues(mCount) = value
    Else
        Set mValues(idx) = value
    End If
End Sub

Public Function GetObjectItem(ByVal key As Variant) As Object
    Dim idx As Long
    idx = KeyIndex(key)
    If idx = 0 Then Exit Function
    If IsObject(mValues(idx)) Then
        Set GetObjectItem = mValues(idx)
    End If
End Function

' Return the stored item as a Variant (preserves objects inside Variant)
Public Function RawItem(ByVal key As Variant) As Variant
    Dim idx As Long
    idx = KeyIndex(key)
    If idx = 0 Then Exit Function
    If IsObject(mValues(idx)) Then
        Set RawItem = mValues(idx)
    Else
        RawItem = mValues(idx)
    End If
End Function

Public Sub Add(ByVal key As Variant, ByVal value As Variant)
    If Exists(key) Then
        Err.Raise 457, "DictionaryShim", "This key is already associated with an element of this collection"
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
    If idx = 0 Then Err.Raise 5, "DictionaryShim", "Invalid procedure call or argument"
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
        ' New key: use Add which will check IsObject
        Add CStr(key), value
    Else
        ' Existing key: assign value (not object)
        mValues(idx) = value
    End If
End Property

Public Property Set Item(ByVal key As Variant, ByVal value As Variant)
    Dim idx As Long
    Dim keyStr As String
    
    keyStr = CStr(key)
    idx = KeyIndex(key)
    
    If idx = 0 Then
        ' New key: add directly to arrays to preserve object reference
        If Exists(key) Then
            Err.Raise 457, "DictionaryShim", "This key is already associated with an element of this collection"
        End If
        
        mCount = mCount + 1
        If mCount > UBound(mKeys) Then
            ReDim Preserve mKeys(mCount * 2)
            ReDim Preserve mValues(mCount * 2)
        End If
        
        mKeys(mCount) = keyStr
        Set mValues(mCount) = value
    Else
        ' Existing key: replace with Set
        Set mValues(idx) = value
    End If
End Property

Public Property Get Count() As Long
    Count = mCount
End Property

Public Function Items() As Variant
    Dim arr() As Variant
    Dim idx As Long

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

Public Function Keys() As Variant
    Dim arr() As String
    Dim idx As Long

    If mCount = 0 Then Exit Function

    ReDim arr(0 To mCount - 1)
    For idx = 1 To mCount
        arr(idx - 1) = mKeys(idx)
    Next idx

    Keys = arr
End Function

Public Sub RemoveAll()
    mCount = 0
    ReDim mKeys(0)
    ReDim mValues(0)
End Sub

' Indexed accessors for JSON serialization (Json.bas)
Public Function KeyAt(ByVal idx As Long) As String
    If idx >= 1 And idx <= mCount Then
        KeyAt = mKeys(idx)
    End If
End Function

Public Function IsObjectAt(ByVal idx As Long) As Boolean
    If idx >= 1 And idx <= mCount Then
        IsObjectAt = IsObject(mValues(idx))
    End If
End Function

Public Function ValueTypeAt(ByVal idx As Long) As VbVarType
    If idx >= 1 And idx <= mCount Then
        ValueTypeAt = VarType(mValues(idx))
    End If
End Function

Public Function StringValueAt(ByVal idx As Long) As String
    If idx >= 1 And idx <= mCount Then
        If Not IsObject(mValues(idx)) Then
            StringValueAt = CStr(mValues(idx))
        End If
    End If
End Function

Public Function NumericValueAt(ByVal idx As Long) As Double
    If idx >= 1 And idx <= mCount Then
        If IsNumeric(mValues(idx)) Then
            NumericValueAt = CDbl(mValues(idx))
        End If
    End If
End Function

Public Function BoolValueAt(ByVal idx As Long) As Boolean
    If idx >= 1 And idx <= mCount Then
        If VarType(mValues(idx)) = vbBoolean Then
            BoolValueAt = mValues(idx)
        End If
    End If
End Function

Public Function DateValueAt(ByVal idx As Long) As Date
    If idx >= 1 And idx <= mCount Then
        If VarType(mValues(idx)) = vbDate Then
            DateValueAt = mValues(idx)
        End If
    End If
End Function

Public Function ObjectValueAt(ByVal idx As Long) As Object
    If idx >= 1 And idx <= mCount Then
        If IsObject(mValues(idx)) Then
            Set ObjectValueAt = mValues(idx)
        End If
    End If
End Function

Public Function IndexOf(ByVal key As Variant) As Long
    IndexOf = KeyIndex(key)
End Function

Private Function KeyIndex(ByVal key As Variant) As Long
    Dim normalized As String
    Dim idx As Long

    normalized = NormalizeKeyText(key)

    For idx = 1 To mCount
        If NormalizeKeyText(mKeys(idx)) = normalized Then
            KeyIndex = idx
            Exit Function
        End If
    Next idx
End Function

Private Function NormalizeKeyText(ByVal key As Variant) As String
    Dim txt As String

    txt = CStr(key)
    If mCompareMode = vbTextCompare Then
        NormalizeKeyText = UCase$(txt)
    Else
        NormalizeKeyText = txt
    End If
End Function
