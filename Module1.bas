Attribute VB_Name = "Module1"
Sub ExportToCSV()
    MsgBox "�G�N�X�|�[�g�J�n���܂��B", vbInformation, "CSV�G�N�X�|�[�g"

    Dim filePath As String
    filePath = Environ("USERPROFILE") & "\Desktop\words.csv"
    
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    On Error GoTo ErrorHandler
    
    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    
    Dim i As Long
    For i = 3 To lastRow
        Print #fileNumber, Join(Application.Index(Range("A:D").Value, i), ",")
    Next i
    
    Close #fileNumber
    
    MsgBox "�����I�f�X�N�g�b�v�� 'words.csv' ���쐬���܂����B", vbInformation, "����"
    Exit Sub

ErrorHandler:
    MsgBox "�G���[�����F" & Err.Description, vbCritical, "���s"
    On Error GoTo 0
End Sub

