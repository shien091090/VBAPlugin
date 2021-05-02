Attribute VB_Name = WorkSheetPlugin
'----------------------------------------------------------------------------------
'WorkSheetPlugin
'方便直接在儲存格公式使用的Function
'----------------------------------------------------------------------------------

'輸入關鍵字查詢指定範圍的字串, 若有搜尋到則返回向下拉至最下方項目(Ctrl+shift+down)的範圍位址
'HINT : 可用於快速建立下拉式選單
'searchRange = 欲搜尋的範圍
'searchKey = 欲搜尋的關鍵字
'rowOffset = 搜尋到的字串在作"Ctrl+shift+down"前的上下偏移量
'返回儲存格範圍位址(string)
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

'指定儲存格範圍, 取得範圍中所有不重複的值並透過輸入索引取得對應的值
'address = 指定儲存格範圍位址
'index = 索引
'返回索引對應的值(string)
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

'設定下拉式選單
'targetRange = 欲設定下拉選單的儲存格範圍
'dropDownDataAddress = 選單內容參考的儲存格位址
Public Function SetDropDown(targetRange As Range, dropDownDataAddress As String) As String
   
    With targetRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=" & dropDownDataAddress
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

End Function