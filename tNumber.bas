Sub tNumber()
  ' Возвращает номер таблицы, в которой стоит курсор.
  '''
    Dim i As Integer, tableNumber As Integer
    
    tableNumber = 0
    If Selection.Information(wdWithInTable) = True Then
        For i = 1 To ActiveDocument.Tables.Count
            If Selection.InRange(ActiveDocument.Tables(i).Range) Then
                tableNumber = i
                Exit For
            End If
        Next
    End If
    
    MsgBox IIf((tableNumber = 0), "курсор вне таблицы", "курсор внутри таблицы № " & tableNumber), vbInformation
    
End Sub
