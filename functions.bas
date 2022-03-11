Option Explicit

' /**
'  * Data matching - highSpeed_matching()
' */
Private Sub highSpeed_matching()
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


' /**
'  * Searches within the 2nd dimension array and returns the index number.
'  *  - Mainly used to find a header name in dynamic arrays.
'  * @param {Variant} var   dynamic arrays
'  * @param {String}  stxt  target header name
' */
Private Function lookup_HeaderIndex(ByRef var As Variant, ByRef stxt As String)

  Dim i As Long
  Dim lngCol As Long

  lngCol = -1
  For i = 1 To UBound(var, 2)
      If var(1, i) = stxt Then
          lngCol = i
          Exit For
      End If
  Next

  gGetVarColNum = lngCol

End Function
