Private Sub Summa1()
' Исходные данные: есть одна таблица с одним столбцом, в столбце числа,
' последняя строка таблицы пустая
'
'''
  Dim x As Single, i As Integer
  Dim myTable As Table
  Dim myStr As String
  Dim m As Integer
  x = 0
  
  Set myTable = ActiveDocument.Tables(1) 'работаем со второй таблицей документа, обзываем её myTable
  m = myTable.Rows.Count 'кол-во строк в таблице
  
  For i = 1 To m - 1 'со второй строки до предпоследней (последняя строка - для вывода суммы)
    x = x + myTable.Cell(i, 1).Range.Calculate 'вычисляем содержимое 1-го столбца i-й строки и добавляем к x
  Next i
  
  x = FormatNumber(x, 3) 'форматируем отображение х
  myTable.Cell(m, 1).Range = x ' записываем х в последнюю строку
  MsgBox ("Закончил"), , ""
End Sub
