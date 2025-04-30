Attribute VB_Name = "Module1"
Sub ExportToCSV()
    MsgBox "エクスポート開始します。", vbInformation, "CSVエクスポート"

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
    
    MsgBox "完了！デスクトップに 'words.csv' を作成しました。", vbInformation, "完了"
    Exit Sub

ErrorHandler:
    MsgBox "エラー発生：" & Err.Description, vbCritical, "失敗"
    On Error GoTo 0
End Sub

