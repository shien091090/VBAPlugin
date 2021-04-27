Attribute VB_Name = "FileOutputManager"

'----------------------------------------------------------------------------------
'FileOutputManager
'檔案輸出相關Function
'----------------------------------------------------------------------------------

'檢查路徑並建立資料夾
'fileName = 檔案名稱
'subFolderName = 直屬資料夾名稱
'返回文字檔案資料(TextStream)
Public Function BuildPath(fileName As String, subFolderName As String) As TextStream

    Dim filePath As String
    Dim subPath As String
    
    filePath = Application.ActiveWorkbook.path
    subPath = "\" & subFolderName
    
    If Not FolderExists(filePath & subPath) Then
        FolderCreate (filePath & subPath)
    End If
    
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    Dim fileStream As TextStream
    Set fileStream = fso.CreateTextFile(filePath & subPath & fileName)
    
    Set BuildPath = fileStream

End Function

'(文字資料寫入方法)逐一從儲存格範圍的內容寫入
'stream = 欲寫入的資料
'sourceRange = 指定儲存格範圍
'prefix = 前綴詞
'suffix = 後綴詞
Public Sub WriteInByRangeContent(ByVal stream As TextStream, sourceRange As Range, prefix As String, suffix As String)

    Dim rngCell As Range
    
    For Each rngCell In sourceRange.Cells

        stream.WriteLine prefix & rngCell.value & suffix

    Next rngCell

End Sub

'檢查是否存在資料夾
Private Function FolderExists(ByVal path As String) As Boolean

    FolderExists = False
    Dim fso As New FileSystemObject

    If fso.FolderExists(path) Then FolderExists = True

End Function

'創建資料夾
Private Sub FolderCreate(ByVal path As String)

    Dim fso As New FileSystemObject

    If FolderExists(path) Then
        Exit Sub
    Else
        fso.CreateFolder path
        Exit Sub
    End If

End Sub