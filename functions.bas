Option Explicit

'Data matching - highSpeed_matching()

Sub highSpeed_matching()
  Dim i As Long
  Dim values As Variant
  Dim dic As Object
  Set dic = CreateObject("Scripting.Dictionary")
  
  ' Store the range in a dynamic array
  With Workbooks(WBNAME).Worksheets(LOADSHEETNAME)
    values = .Range("A2", .Cells(.Rows.Count, 1).End(xlUp)).Resize(, 2).Value
  End With
  ' Store keys and values in Dictionary
  For i = 1 To UBound(values)
    dic(values(i, 1)) = values(i, 2)
  Next

  ' Search and store in output array
  With Workbooks(WBNAME).Worksheets(OUTPUTSHEETNAME)
    With .Range("A2", .Cells(.Rows.Count, 1).End(xlUp))
      values = .Value
      For i = 1 To UBound(values)
        If dic.Exists(values(i, 1)) Then
            values(i, 1) = dic(values(i, 1))
        Else
            values(i, 1) = Empty
        End If
      Next
      .Offset(, 1).Value = values  'Write the result in column B
    End With
  End With
End Sub
