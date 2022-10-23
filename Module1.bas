Sub vstavka_kvadratov()
Dim oFD As FileDialog
Dim x, lf As Long
Dim SH As Shapes
Dim j1s, dName

    'назначаем переменной ссылку на экземпляр диалога
Set oFD = Application.FileDialog(msoFileDialogFolderPicker)
With oFD 'используем короткое обращение к объекту
    .Title = "Выбрать папку" '"заголовок окна диалога
    .ButtonName = "Выбрать папку"
    .Filters.Clear 'очищаем установленные ранее типы файлов
    .InitialFileName = "C:\" '"назначаем первую папку отображения
    .InitialView = msoFileDialogViewLargeIcons 'вид диалогового окна(доступно 9 вариантов)
    If oFD.Show = 0 Then Exit Sub 'показывает диалог
    x = .SelectedItems(1) 'считываем путь к папке
End With
Set oFD = Nothing

dName = Dir(x & "\*.doc*")

Do While dName <> ""
    Documents.Open FileName:=x & "\" & dName
    j1s = Word.ActiveDocument.Shapes.Count
    Set SH = ActiveDocument.Shapes
    Selection.EndKey Unit:=wdStory ' отматываем на конец документа
    
    If j1s = 0 Then
        SH.AddShape(msoShapeRectangle, 0, 780, 600, 100).Select 'добавление прямоугольника по низу листа
        SH.Item(1).Select
    Else
        SH.Item(j1s).Select
        Selection.Delete ' удаление последней фигуры, надо ли?
        SH.AddShape(msoShapeRectangle, 0, 780, 600, 100).Select
        SH.Item(j1s).Select
    End If

    Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = wdThemeColorBackground1
    Selection.ShapeRange.Fill.ForeColor.TintAndShade = 0#
    Selection.ShapeRange.Fill.Visible = msoTrue
    Selection.ShapeRange.Fill.Solid
    
    Selection.ShapeRange.Line.ForeColor.RGB = RGB(0, 0, 255)
    Selection.ShapeRange.Line.ForeColor.ObjectThemeColor = wdThemeColorBackground1
    Selection.ShapeRange.Line.ForeColor.TintAndShade = 0#
    Selection.ShapeRange.Line.Visible = msoTrue
    
    
    ActiveDocument.Save
    ActiveWindow.Close
        'Documents.Close
    dName = Dir
Loop
Set SH = Nothing
MsgBox ("Закончил"), , ""

End Sub





