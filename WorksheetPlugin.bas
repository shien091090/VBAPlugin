Attribute VB_Name = WorkSheetPlugin
'----------------------------------------------------------------------------------
'WorkSheetPlugin
'��K�����b�x�s�椽���ϥΪ�Function
'----------------------------------------------------------------------------------

'��J����r�d�߫��w�d�򪺦r��, �Y���j�M��h��^�V�U�Ԧ̤ܳU�趵��(Ctrl+shift+down)���d���}
'HINT : �i�Ω�ֳt�إߤU�Ԧ����
'searchRange = ���j�M���d��
'searchKey = ���j�M������r
'rowOffset = �j�M�쪺�r��b�@"Ctrl+shift+down"�e���W�U�����q
'��^�x�s��d���}(string)
Public Function GetAddressBySearch(searchRange As Range ,searchKey As String, rowOffset As Integer) As String
   
    Dim c As Range
    Set c = searchRange.Find(searchKey, LookIn:=xlValues)
    
    If c Is Nothing Then
        GetAddressBySearch = Empty
        Exit Function
    End If
    
    If c.Offset(rowOffset:=rowOffset + 1).value = Empty Then
        GetAddressBySearch = c.Offset(rowOffset:=rowOffset).address(External:=True)
        Exit Function
    End If

    Dim startCell As Range
    Set startCell = c.Offset(rowOffset:=rowOffset)
    
    Dim targetRange As Range
    Set targetRange = Range(startCell, startCell.End(xlDown))
    
    GetAddressBySearch = targetRange.address(External:=True)

End Function

'���w�x�s��d��, ���o�d�򤤩Ҧ������ƪ��Ȩóz�L��J���ި��o��������
'address = ���w�x�s��d���}
'index = ����
'��^���޹�������(string)
Public Function GetValueByAddress(address As String, index As Integer) As String

    Dim targetRange As Range
    Set targetRange = Range(address)

    If targetRange Is Nothing Then
        GetValueByAddress = Empty
        Exit Function
    End If

    Dim i As Integer
    i = 0

    Dim value As String
    Dim rng As Range

    For Each rng In targetRange.Cells

        If value <> rng.value Then
            If i = index Then
                GetValueByAddress = rng.value
                Exit Function
            End If

            value = rng.value
            i = i + 1

        End If

    Next rng

    GetValueByAddress = Empty

End Function

'�]�w�U�Ԧ����
'targetRange = ���]�w�U�Կ�檺�x�s��d��
'dropDownDataAddress = ��椺�e�ѦҪ��x�s���}
Public Function SetDropDown(targetRange As Range, dropDownDataAddress As String) As String
   
    With targetRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=" & dropDownDataAddress
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

End Function