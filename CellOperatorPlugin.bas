Attribute VB_Name = "CellOperatorPlugin"
'----------------------------------------------------------------------------------
'CellOperatorPlugin
'¦UºØÀx¦s®æ¬ÛÃö¾Þ§@ªºFunction
'----------------------------------------------------------------------------------

'±q¬Y¤@Àx¦s®æ¥Xµo, ¨ú±o©Ô¦Ü«ü©w¤è¦V³Ì«á¤@®æ(Ctrl+shift+down)ªºÀx¦s®æ½d³ò
'startCell = ªì©lÀx¦s®æ
'dir = ¤è¦V
'ªð¦^©Ô¦Ü«ü©w¤è¦V³Ì«á¤@®æ(Ctrl+shift+down)ªºÀx¦s®æ½d³ò
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

'³]©w¤@þ©lÀx¦s®æ, ¦V¥k·j´M¼Ð°O(Flag), ·j´M¨ìªº²[»\½d³ò¥i¦A³]©w¦C°¾²¾¶q, ªð¦^°¾²¾«á½d³òªº¦a§}
'startCell = þ©lÀx¦s®æ
'filterFlag = ±ý·j´Mªº¼Ð°O
'offsetValue = ¦C(¤W¤U)°¾²¾¶q
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

'±q«ü©w¦C©ÎÄæ·j´MÃöÁä¦r, ·j´M¨ì«á¦Aªð¦^¸ÓÄæ©Î¦C©µ¦ù¨ì©³¤§¶¡ªºÀx¦s®æ¤¸¯À°}¦C
'Á|¨Ò¨Ó»¡, ¦pªG«ü©w±q¬YÄæ·j´M, «h·j´M¨ìªº¦ì§}þ·|ªð¦^¥k°¼¾ã±øÀx¦s®æ½d³òªº¤¸¯À
'searchKey = ±ý·j´MªºÃöÁä¦r
'searchRange = ±ý·j´Mªº³æÄæ©Î³æ¦C(­Y¦P®É¬°¦hÄæ&¦h¦C®Éªð¦^ªÅ­È)
'ªð¦^Àx¦s®æ¤¸¯À°}¦C
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