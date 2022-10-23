Function Get_DirS(path As String)
  ' Ф-ция получения списка файлов в массив
  '''
Dim a() As String, D As String, U As Long
D = Dir(path, vbDirectory)
While D <> ""
  If GetAttr(path & "\" & D) And vbDirectory Then
    ReDim Preserve a(U)
    a(U) = path & D
    U = U + 1
  End If
  D = Dir
Wend
Get_DirS = a
End Function
  
Sub test()
    ' ф-ция показывает, как использовать
    ' ф-цию Get_DirS
    '''
Dim a() As String, i as integer

  a() = Get_DirS("d:\tmp\", "*.doc") ' путь, маска

  For i = 0 To UBound(a())
      Debug.Print a(i) ' возврат файлов *.doc из папки
  Next
end sub
