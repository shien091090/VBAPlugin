Attribute VB_Name = "DebugMessagePlugin"
'----------------------------------------------------------------------------------
'DebugMessagePlugin
'顯示測試用訊息的Function
'----------------------------------------------------------------------------------

'顯示陣列元素
'elements = 陣列
Public Sub PrintArrayElements(elements As Variant)

    On Error GoTo catchError
    
    Dim message As String
    message = "Count = " & UBound(elements) + 1 & vbCrLf
    
    Dim i As Integer
    For i = LBound(elements) To UBound(elements)
        message = message & "[" & i & "] = " & elements(i) & vbCrLf
    Next
    
    MsgBox message
    
catchError:
    MsgBox "[ERROR]錯誤的輸入參數"

End Sub


