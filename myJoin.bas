Function myJoin(�͈� As Range, Optional ��؂蕶�� As String, Optional ��؂蕶���� As Integer) As Variant
Dim c As Range, buf As String
   If �͈�.Rows.Count = 1 Or �͈�.Columns.Count = 1 Then
      For Each c In �͈�
         buf = buf & ��؂蕶�� & c.Value
      Next c
      If ��؂蕶�� <> "" Then
         myJoin = Mid$(buf, ��؂蕶���� + 1)
         Else
         myJoin = buf
      End If
      Else
      myJoin = CVErr(xlErrRef)  '�G���[�l
   End If
End Function
