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
    If (VarType(arr1) + VarType(arr2)/ 2 <> VarType(arr1) Then
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
    ReDim result(coll.Count - 1)
    
    i = LBound(result)
    For Each Item In coll
        result(i) = Item
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