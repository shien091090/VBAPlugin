Attribute VB_Name = "CellOperatorPlugin"
'----------------------------------------------------------------------------------
'CellOperatorPlugin
'�U���x�s������ާ@��Function
'----------------------------------------------------------------------------------

'�q�Y�@�x�s��X�o, ���o�Ԧܫ��w��V�̫�@��(Ctrl+shift+down)���x�s��d��
'startCell = ��l�x�s��
'dir = ��V
'��^�Ԧܫ��w��V�̫�@��(Ctrl+shift+down)���x�s��d��
Public Function GetRangeLineByStartCell(startCell As Range, dir As XlDirection) As Range
    
    Dim col As Integer
    Dim row As Integer
    col = 0
    row = 0
    
    Select Case dir
        Case xlDown
            row = 1
            
        Case xlUp
            row = -1

        Case xlToRight
            col = 1

        Case xlToLeft
            col = -1

    End Select
    
    If startCell.Offset(columnOffset:=col, rowOffset:=row).value = "" Then
        Set GetRangeLineByStartCell = startCell
        Exit Function
    
    End If
    
    Set GetRangeLineByStartCell = Range(startCell, startCell.End(dir))

End Function

'�]�w�@���l�x�s��, �V�k�j�M�аO(Flag), �j�M�쪺�[�\�d��i�A�]�w�C�����q, ��^������d�򪺦a�}
'startCell = ���l�x�s��
'filterFlag = ���j�M���аO
'offsetValue = �C(�W�U)�����q
Public Function GetAddressByFlag(startCell As Range, filterFlag As String, offsetValue As Integer) As String

    Dim targetRange As Range

    If startCell.count > 0 Then
        GetAddressByFlag = Empty
        Exit Function
    End If

    If startCell.Offset(columnOffset:=1).value = Empty Then
        GetAddressByFlag = Empty
        Exit Function
    End If
    
    Set targetRange = Range(startCell, startCell.End(xlToRight).Offset(columnOffset:=1))

    Dim firstCell As Range
    Dim endCell As Range
    Dim rng As Range
    Dim foundFirst As Boolean

    Set firstCell = startCell
    Set endCell = startCell
    foundFirst = False

    For Each rng In targetRange.Cells
       
        If rng.value <> filterFlag And foundFirst = True Then
            Set endCell = rng
            Exit For
        ElseIf rng.value = filterFlag And foundFirst = False Then
            Set firstCell = rng
            foundFirst = True
        End If

    Next rng

    If foundFirst = False Then
        GetAddressByFlag = Empty
        Exit Function
    End If

    Dim result As String
    result = Range( _
    firstCell.Offset(rowOffset:=offsetValue), _
    endCell.Offset(rowOffset:=offsetValue, columnOffset:=-1)).address(External:=True)

    GetAddressByFlag = result

End Function

'�q���w�C����j�M����r, �j�M���A��^����ΦC�����쩳�������x�s�椸���}�C
'�|�Ҩӻ�, �p�G���w�q�Y��j�M, �h�j�M�쪺��}���|��^�k������x�s��d�򪺤���
'searchKey = ���j�M������r
'searchRange = ���j�M������γ�C(�Y�P�ɬ��h��&�h�C�ɪ�^�ŭ�)
'��^�x�s�椸���}�C
Public Function GetRangeLineBySearch(searchKey As String, searchRange As Range) As Variant
    
    Dim rowOrColumn As Boolean

    If searchRange.Columns.count > 1 And searchRange.Rows.count > 1 Then
        GoTo ReturnNull
    ElseIf searchRange.Columns.count > 1 And searchRange.Rows.count = 1 Then
        rowOrColumn = True
    ElseIf searchRange.Columns.count = 1 And searchRange.Rows.count > 1 Then
        rowOrColumn = False
    Else
        GoTo ReturnNull
    End If
    
    Dim dir As XlDirection
    If rowOrColumn = True Then
        dir = xlDown
    Else
        dir = xlToRight
    End If

    Dim resultRange As Range
    Set resultRange = searchRange.Find(what:=searchKey)

    If resultRange Is Nothing Then
        GoTo ReturnNull
    End If

    Dim col As Integer
    Dim row As Integer
    col = 0
    row = 0

    Select Case dir
        Case xlDown
            row = 1

        Case xlToRight
            col = 1

    End Select

    Dim arrayCount As Integer
    If resultRange.value = Empty Then
        GoTo ReturnNull
        
    ElseIf resultRange.Offset(columnOffset:=col, rowOffset:=row).value = Empty Then
        arrayCount = 1
        
    Else
        Dim rangeLine As Range
        Set rangeLine = Range(resultRange, resultRange.End(dir))
        arrayCount = rangeLine.Cells.count - 1
        
    End If
    
    Dim contentArray() As String
    ReDim contentArray(arrayCount - 1)
    
    Dim i As Integer
    For i = 1 To arrayCount
        If rowOrColumn = True Then
            contentArray(i - 1) = resultRange.Cells(1 + i, 1).value
        Else
            contentArray(i - 1) = resultRange.Cells(1, 1 + i).value
        End If
        
    Next i

    Dim result As Variant
    result = contentArray

    GetRangeLineBySearch = result
    Exit Function
    
ReturnNull:
    Dim nullArr(0) As String
    GetRangeLineBySearch = nullArr
    Exit Function

End Function