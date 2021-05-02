Attribute VB_Name = "DataOperatorPlugin"
'----------------------------------------------------------------------------------
'DataOperatorPlugin
'�U�ظ�Ƭ����ާ@��Function
'----------------------------------------------------------------------------------

'�X�ְ}�C�ìD�����Ƥ���
'arr1 = �}�C1
'arr2 = �}�C2
'��^�X�֫�ìD�����Ƥ����������}�C(Variant)
Public Function MergeArray(arr1 As Variant, arr2 As Variant) As Variant

    Dim result() As Variant
    
    '�Y��J�}�C���O�P���O��, �h����
    If (VarType(arr1) + VarType(arr2)) / 2 <> VarType(arr1) Then
        GoTo OuputStep
    
    End If
           
    Dim coll As New Collection '�إ��x�s�}�C���������X
    
    ' �������ƪ����~
    On Error Resume Next

    ' �N�}�C������JCollection��
    For i = LBound(arr1) To UBound(arr1)
        coll.Add CStr(arr1(i)), CStr(arr1(i))
    Next i
    
    For i = LBound(arr2) To UBound(arr2)
        coll.Add CStr(arr2(i)), CStr(arr2(i))
    Next i
    
    ' �]�w���G�}�C�j�p
    ReDim result(coll.count - 1)
    
    i = LBound(result)
    For Each item In coll
        result(i) = item
        i = i + 1
    Next

OuputStep:
    MergeArray = result

End Function

'��^���w�ƶq��Tab�ť�
'count = Tab�ťռƶq
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

'���տ�J�ȬO�_�i�H�নINT
'var = ��J��
'��^�O�_�i�H�૬�����L��
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

'���հ}�C�O�_�L��
'arr = �}�C����
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

'���հ}�C�O�_�����w���O�����İ}�C
'arr = �}�C����
'fieldType = ���w���O
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

'�N�r���ܪ��}�C�ഫ���}�C����
'arrStr = �}�C�r��
'��XString�}�C
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

'�L�o�}�C, �N�}�C�����ݩ���w���O�������z����^�s���}�C
'(�ثe�o�{���J�}�C��String, �Ӥ�����Empty��, On Error�|����, �ثe�L��...)
'arr = ��}�C
'fieldType = ���w���O
'��X�N�Ҧ�����Convert���w���O�᪺�s�}�C
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

'��JString�r��, �z���L�kConvert���������XInt Array
'arr = String�r��
'��XInt�}�C
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

'�R���r�ꤤ�����w����
'origin = ��l�r��
'KeyWord = ���R��������r
'��^�s���r��(string)
Public Function ContentPartRemove(origin As String, KeyWord As String) As String

    Dim newContent As String
    newContent = Replace(origin, KeyWord, Empty)

    ContentPartRemove = newContent

End Function

'�}�C���O�_�]�t���w����
'arr = �ؼа}�C
'item = ���w����
'��^�O�_�s�b�����L��
Public Function IsArrayContain(arr As Variant, item As Variant) As Boolean

    If IsVarArrayValid(arr) = False Then
        IsArrayContain = False
        Exit Function
    End If
    
    IsArrayContain = UBound(Filter(arr, item)) > -1
    
End Function
