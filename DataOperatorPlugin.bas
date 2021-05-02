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
    For Each item In coll
        result(i) = item
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

'測試輸入值是否可以轉成INT
'var = 輸入值
'返回是否可以轉型的布林值
Public Function TryParseInt(var As Variant) As Boolean

    Dim result As Boolean
    
    On Error GoTo catchError

    Dim i As Integer
    i = CInt(var)
    
    result = True
    GoTo returnResult
    
catchError:
    result = False
    GoTo returnResult
    
returnResult:
    TryParseInt = result

End Function

'測試陣列是否無效
'arr = 陣列物件
Public Function IsVarArrayValid(arr As Variant) As Boolean

    Dim result As Boolean
    
    On Error GoTo catchError
    
    Dim i As Integer
    i = UBound(arr)
    
    result = True
    GoTo returnResult
    
catchError:
    result = False
    GoTo returnResult
    
returnResult:
    IsVarArrayValid = result

End Function

'測試陣列是否為指定型別的有效陣列
'arr = 陣列物件
'fieldType = 指定型別
Public Function IsAppointTypeArrayValid(arr As Variant, fieldType As String) As Boolean

    Dim result As Boolean
    
    On Error GoTo catchError
    
    For i = 0 To UBound(arr)
    
        Dim value As Variant
        value = arr(i)
        
        If TypeName(arr(i)) <> fieldType Then
        
            Dim tempVar As Variant
            
            Select Case fieldType
                Case "Integer"
                    tempVar = CInt(arr(i))
                Case "String"
                    tempVar = CStr(arr(i))
                Case "Single"
                    tempVar = CSng(arr(i))
                Case "Long"
                    tempVar = CLng(arr(i))
                Case "Double"
                    tempVar = CDbl(arr(i))
                Case "Date"
                    tempVar = CDate(arr(i))
                Case "Boolean"
                    tempVar = CBool(arr(i))
                Case "Byte"
                    tempVar = CByte(arr(i))
            
            End Select
        
        End If
    
    Next i
    
    result = True
    GoTo returnResult
    
catchError:
    result = False
    GoTo returnResult
    
returnResult:
    IsAppointTypeArrayValid = result

End Function

'將字串表示的陣列轉換成陣列物件
'arrStr = 陣列字串
'輸出String陣列
Public Function ConvertStringToArray(arrStr As String) As Variant

    Dim content As String
    content = Trim(arrStr)
    content = Replace(content, "[", Empty)
    content = Replace(content, "]", Empty)
    
    Dim dotPos As Integer
    dotPos = InStr(1, content, ",")
    
    Dim resultArr As Variant
    
    If dotPos = 0 Then
        If content = Empty Then
            ConvertStringToArray = resultArr
            Exit Function
            
        End If
        
        ReDim resultArr(0)
        resultArr(0) = content
    
    Else
        resultArr = Split(content, ",")
    
    End If

    ConvertStringToArray = resultArr

End Function

'過濾陣列, 將陣列中不屬於指定型別的元素篩掉返回新的陣列
'(目前發現當輸入陣列為String, 而元素為Empty時, On Error會失效, 目前無解...)
'arr = 原陣列
'fieldType = 指定型別
'輸出將所有元素Convert指定型別後的新陣列
Public Function FilterArray(arr As Variant, fieldType As String) As Variant

    Dim result() As Variant
    
    Dim i As Integer
    i = 0
    
    On Error GoTo forEachEnd
    
    Dim v As Variant
    For Each v In arr
    
        Dim tempVar As Variant
    
        Select Case fieldType
            Case "Integer"
                tempVar = CInt(v)
            Case "String"
                tempVar = CStr(v)
            Case "Single"
                tempVar = CSng(v)
            Case "Long"
                tempVar = CLng(v)
            Case "Double"
                tempVar = CDbl(v)
            Case "Date"
                tempVar = CDate(v)
            Case "Boolean"
                tempVar = CBool(v)
            Case "Byte"
                tempVar = CByte(v)
                
        End Select
        
        ReDim Preserve result(i)
        result(i) = tempVar
        
        i = i + 1

forEachEnd:
    
    Next
    
    If i = 0 Then
        GoTo returnNullArr
    End If
    
    FilterArray = result
    Exit Function
    
returnNullArr:
    Dim nullArr() As Variant
    FilterArray = nullArr

End Function

'輸入String字串, 篩掉無法Convert的部分後輸出Int Array
'arr = String字串
'輸出Int陣列
Public Function StringArrayToIntArray(arr As Variant) As Integer()

    On Error GoTo returnNull
    
    If IsVarArrayValid(arr) = False Then
        GoTo returnNull
    End If
    
    If IsAppointTypeArrayValid(arr, "String") = False Then
        GoTo returnNull
    End If
    
    Dim result() As Integer
    Dim index As Integer
    index = 0
    
    On Error GoTo forEnd
    
    Dim i As Integer
    For i = 0 To UBound(arr)
    
        If arr(i) <> Empty And TryParseInt(arr(i)) = True Then
            ReDim Preserve result(index)
            result(index) = CInt(arr(i))
            
            index = index + 1
        
        End If
        
forEnd:
        
    Next i
    
    StringArrayToIntArray = result
    Exit Function
    
returnNull:
    Dim nullArr() As Integer
    StringArrayToIntArray = nullArr

End Function

'刪除字串中的指定部分
'origin = 原始字串
'KeyWord = 欲刪除的關鍵字
'返回新的字串(string)
Public Function ContentPartRemove(origin As String, KeyWord As String) As String

    Dim newContent As String
    newContent = Replace(origin, KeyWord, Empty)

    ContentPartRemove = newContent

End Function

'陣列中是否包含指定元素
'arr = 目標陣列
'item = 指定元素
'返回是否存在的布林值
Public Function IsArrayContain(arr As Variant, item As Variant) As Boolean

    If IsVarArrayValid(arr) = False Then
        IsArrayContain = False
        Exit Function
    End If
    
    IsArrayContain = UBound(Filter(arr, item)) > -1
    
End Function
