Attribute VB_Name = "DebugMessagePlugin"
'----------------------------------------------------------------------------------
'DebugMessagePlugin
'��ܴ��եΰT����Function
'----------------------------------------------------------------------------------

'��ܰ}�C����
'elements = �}�C
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
    MsgBox "[ERROR]���~����J�Ѽ�"

End Sub


