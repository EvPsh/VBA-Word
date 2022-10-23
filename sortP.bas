Sub SortP()
  ' сортировка массива "пузырьком"
  '''
Dim i As Integer
Dim j As Integer
Dim u As Integer
Dim m() As Integer

' заполнение массива случайными числами
ReDim m(5) ' переинициализируем массив
For i = 1 To 5
  m(i) = Int(Rnd * 10) + 1
  Debug.Print m(i) ' вывод несортированного массива в immediate window
Next i

' сортировка "пузырьком" по возрастанию. элемент m(0) не использую.
For i = 1 To UBound(m())
    For j = 1 To UBound(m()) - 1
        If m(j) > m(j + 1) Then
            u = m(j)
            m(j) = m(j + 1)
            m(j + 1) = u
        End If
    Next j
Next i
    
' вывод отсортированного массива в immediate window (ctrl+G, или View -> Immediate Window
For i = 1 To UBound(m())
    Debug.Print "sorted: " & m(i)
Next i

end sub
