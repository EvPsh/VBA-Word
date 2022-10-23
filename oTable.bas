Sub Обработка_таблиц()
  ' Обработка всех таблиц документа: выравнивание по ширине листа.
  '''

    Dim objTable As Word.Table
    For Each objTable In ActiveDocument.Tables
        objTable.Select
        objTable.AutoFitBehavior (wdAutoFitWindow)
'        MsgBox _
'        "Кол-во строк =" & objTable.Rows.Count & vbCrLf & _
'        "Кол-во столбцов =" & objTable.Columns.Count, , ""
    Next
End Sub
