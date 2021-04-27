Attribute VB_Name = "FileOutputManager"

'----------------------------------------------------------------------------------
'FileOutputManager
'�ɮ׿�X����Function
'----------------------------------------------------------------------------------

'�ˬd���|�ëإ߸�Ƨ�
'fileName = �ɮצW��
'subFolderName = ���ݸ�Ƨ��W��
'��^��r�ɮ׸��(TextStream)
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

'(��r��Ƽg�J��k)�v�@�q�x�s��d�򪺤��e�g�J
'stream = ���g�J�����
'sourceRange = ���w�x�s��d��
'prefix = �e���
'suffix = ����
Public Sub WriteInByRangeContent(ByVal stream As TextStream, sourceRange As Range, prefix As String, suffix As String)

    Dim rngCell As Range
    
    For Each rngCell In sourceRange.Cells

        stream.WriteLine prefix & rngCell.value & suffix

    Next rngCell

End Sub

'�ˬd�O�_�s�b��Ƨ�
Private Function FolderExists(ByVal path As String) As Boolean

    FolderExists = False
    Dim fso As New FileSystemObject

    If fso.FolderExists(path) Then FolderExists = True

End Function

'�Ыظ�Ƨ�
Private Sub FolderCreate(ByVal path As String)

    Dim fso As New FileSystemObject

    If FolderExists(path) Then
        Exit Sub
    Else
        fso.CreateFolder path
        Exit Sub
    End If

End Sub