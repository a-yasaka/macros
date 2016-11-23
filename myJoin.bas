Function myJoin(”ÍˆÍ As Range, Optional ‹æØ‚è•¶š As String, Optional ‹æØ‚è•¶š” As Integer) As Variant
Dim c As Range, buf As String
   If ”ÍˆÍ.Rows.Count = 1 Or ”ÍˆÍ.Columns.Count = 1 Then
      For Each c In ”ÍˆÍ
         buf = buf & ‹æØ‚è•¶š & c.Value
      Next c
      If ‹æØ‚è•¶š <> "" Then
         myJoin = Mid$(buf, ‹æØ‚è•¶š” + 1)
         Else
         myJoin = buf
      End If
      Else
      myJoin = CVErr(xlErrRef)  'ƒGƒ‰[’l
   End If
End Function
