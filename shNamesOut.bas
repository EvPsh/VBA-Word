Private Sub SheetsNameOut()
' Вывод имён листов в новую книгу
'''

Dim lName() As String ' массив с именами листов
Dim x as Integer

    ReDim lName(ActiveWorkbook.Worksheets.Count) ' переинициализация массива по числу листов в книге
	For x = 1 To ActiveWorkbook.Worksheets.Count
		lName(x) = Sheets(x).Name ' заполняем массив названиями листов Excel
	Next x
        
    ' Вывод результатов в новую книгу excel
    Workbooks.Add
    For x = 1 To UBound(lName())
        Cells(x, 1) = lName(x) ' вывод имён листов из массива в 1ый столбец новой книги Excel
    Next x

End Sub