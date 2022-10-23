Attribute VB_Name = "Module1"
Sub vstavka_kvadratov()
Dim oFD As FileDialog
Dim x, lf As Long
Dim SH As Shapes
Dim j1s, dName

    '��������� ���������� ������ �� ��������� �������
Set oFD = Application.FileDialog(msoFileDialogFolderPicker)
With oFD '���������� �������� ��������� � �������
    .Title = "������� �����" '"��������� ���� �������
    .ButtonName = "������� �����"
    .Filters.Clear '������� ������������� ����� ���� ������
    .InitialFileName = "C:\" '"��������� ������ ����� �����������
    .InitialView = msoFileDialogViewLargeIcons '��� ����������� ����(�������� 9 ���������)
    If oFD.Show = 0 Then Exit Sub '���������� ������
    x = .SelectedItems(1) '��������� ���� � �����
End With
Set oFD = Nothing

dName = Dir(x & "\*.doc*")

Do While dName <> ""
    Documents.Open FileName:=x & "\" & dName
    j1s = Word.ActiveDocument.Shapes.Count
    Set SH = ActiveDocument.Shapes
    Selection.EndKey Unit:=wdStory ' ���������� �� ����� ���������
    
    If j1s = 0 Then
        SH.AddShape(msoShapeRectangle, 0, 780, 600, 100).Select '���������� �������������� �� ���� �����
        SH.Item(1).Select
    Else
        SH.Item(j1s).Select
        Selection.Delete ' �������� ��������� ������, ���� ��?
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
MsgBox ("��������"), , ""

End Sub
