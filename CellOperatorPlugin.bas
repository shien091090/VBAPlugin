Attribute VB_Name = "CellOperatorPlugin"
'----------------------------------------------------------------------------------
'CellOperatorPlugin
'各種儲存格相關操作的Function
'----------------------------------------------------------------------------------

'從某一儲存格出發, 取得拉至指定方向最後一格(Ctrl+shift+down)的儲存格範圍
'startCell = 初始儲存格
'dir = 方向
'返回拉至指定方向最後一格(Ctrl+shift+down)的儲存格範圍
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

'設定一��始儲存格, 向右搜尋標記(Flag), 搜尋到的涵蓋範圍可再設定列偏移量, 返回偏移後範圍的地址
'startCell = ��始儲存格
'filterFlag = 欲搜尋的標記
'offsetValue = 列(上下)偏移量
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

'從指定列或欄搜尋關鍵字, 搜尋到後再返回該欄或列延伸到底之間的儲存格元素陣列
'舉例來說, 如果指定從某欄搜尋, 則搜尋到的位址��會返回右側整條儲存格範圍的元素
'searchKey = 欲搜尋的關鍵字
'searchRange = 欲搜尋的單欄或單列(若同時為多欄&多列時返回空值)
'返回儲存格元素陣列
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