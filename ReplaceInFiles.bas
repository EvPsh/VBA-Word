Sub data_izm() 
' изменение ненужных дат
' перебираем все файлы в папке
' ищем текст, удаляем
' меняем его на нужный
'''
Dim dName
dName = Dir("d:\tmp\")

Do While dName <> ""
    Documents.Open FileName:="d:\tmp\" & dName
        With Selection
            .Find.Text = "дата проведения работ"
            .Find.Execute
            '.MoveLeft unit:=wdCharacter, Count:=1
            .EndKey Unit:=wdLine, Extend:=wdExtend
            .Delete Unit:=wdCharacter, Count:=1
            .ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Text = "Дата проведения работ" & Chr(9) & "00.00.21 г." & Chr(13) ' здесь выставляется необходимая дата. chr(9) - это tab, chr(13) - enter
        End With
        
    ActiveDocument.Save
    Documents.Close
dName = Dir
Loop

End Sub
