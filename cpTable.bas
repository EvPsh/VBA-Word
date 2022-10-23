Sub cpTable()
  ' Ф-ция копирования таблицы из одного файла
  ' в другой
  '''
   
    Dim docSrc As Document, docRes As Document, rngTable As Range
    Dim strFN As String

   
  '1. Отключение обновления экрана. Ускоряет выполнение макроса. 
    ' Если необходимо видеть, что происходит - можно не отключать обновление экрана
    Application.ScreenUpdating = False

    '2. Юзер выбирает файл, в котором таблица.
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Документы Word", "*.doc"
        If .Show = 0 Then
            Exit Sub
        End If
        strFN = .SelectedItems(1)
    End With
   
    '3. Присваивание имени "docRes" активного файлу (в который надо вставить таблицу).
        ' После открытия другого файла, он станет неактивным.
    Set docRes = ActiveDocument
   
    '4. Открытие файла, в котором таблица. При этом присваиваем файлу имя "docSrc".
    Set docSrc = Documents.Open(FileName:=strFN)
   
    '5. Копирование таблицы из одного файла в другой.
    With docRes.Range.find
        ' Текст-метка, куда надо вставить таблицу.
          .Text = "~tbl~"
        ' Поиск текста-метки.
        .Execute
        ' Присваиваем имя "rngTable" фрагменту, в котором находится текст-метка.
            ' Parent - это найденный текст.
        Set rngTable = .Parent
    End With
   
    '6. Убираем цветовую заливку.
    rngTable.HighlightColorIndex = wdNoHighlight
   
    '7. Вставка таблицы. Копируется первая таблица из файла-источника.
    docSrc.Tables(1).Range.Copy
    rngTable.Paste
   
    '8. Очистка буфера обмена. Если таблица большая, то при закрытии ворда
        ' будет сообщение, что в буфере много данных.
        ' Просто копируем первый символ.
    docSrc.Range.Characters(1).Copy
   
    '9. Закрытие файла-источника.
    docSrc.Close SaveChanges:=False
   
    '10. Включение обновления экрана.
    Application.ScreenUpdating = True
   
End Sub
