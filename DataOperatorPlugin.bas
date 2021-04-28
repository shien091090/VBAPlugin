Attribute VB_Name = "DataOperatorPlugin"
'----------------------------------------------------------------------------------
'DataOperatorPlugin
'各種資料相關操作的Function
'----------------------------------------------------------------------------------

'合併陣列並挑掉重複元素
'arr1 = 陣列1
'arr2 = 陣列2
'返回合併後並挑掉重複元素的元素陣列(Variant)
Public Function MergeArray(arr1 As Variant, arr2 As Variant) As Variant

    Dim result() As Variant
    
    '若輸入陣列不是同型別時, 則報錯
    If (VarType(arr1) + VarType(arr2)) / 2 <> VarType(arr1) Then
        GoTo OuputStep
    
    End If
           
    Dim coll As New Collection '建立儲存陣列元素的集合
    
    ' 忽略重複的錯誤
    On Error Resume Next

    ' 將陣列元素放入Collection中
    For i = LBound(arr1) To UBound(arr1)
        coll.Add CStr(arr1(i)), CStr(arr1(i))
    Next i
    
    For i = LBound(arr2) To UBound(arr2)
        coll.Add CStr(arr2(i)), CStr(arr2(i))
    Next i
    
    ' 設定結果陣列大小
    ReDim result(coll.count - 1)
    
    i = LBound(result)
    For Each Item In coll
        result(i) = Item
        i = i + 1
    Next

OuputStep:
    MergeArray = result

End Function

'返回指定數量的Tab空白
'count = Tab空白數量
Function GetTabSpace(count As Integer) As String
    
    If count <= 0 Then
        GetTabSpace = ""
        Exit Function
    End If
    
    Dim result As String
    result = ""
    
    Dim i As Integer
    For i = 1 To count
        result = result & vbTab
    Next i
    
    GetTabSpace = result

End Function

'測試陣列是否無效
'arr = 陣列物件
Public Function IsVarArrayValid(arr As Variant) As Boolean

    On Error GoTo catchError
    
    Dim i As Integer
    i = UBound(arr)
    
    IsVarArrayValid = True
    Exit Function
    
catchError:
    IsVarArrayValid = False

End Function

'將字串表示的陣列轉換成陣列物件
'arrStr = 陣列字串
Public Function ConvertStringToArray(arrStr As String) As Variant

    Dim content As String
    content = Trim(arrStr)
    content = Replace(content, "[", Empty)
    content = Replace(content, "]", Empty)

    Dim resultArr As Variant
    resultArr = Split(content, ",")

    ConvertStringToArray = resultArr

End Function

'取代字串
Public Function Remove(origin As String, keyWord As String) As String

    Dim newContent As String
    newContent = Replace(origin, keyWord, Empty)

    Remove = newContent

End Function

