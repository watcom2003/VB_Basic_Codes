Sub BMI()
Documents("codeVBA1-2559.docm").Activate
ActiveDocument.Bookmarks("bmi").Select
Selection.MoveDown
w = Val(Selection.Tables(1).Cell(1, 2).Range.Text)
h = Val(Selection.Tables(1).Cell(2, 2).Range.Text) / 100
bmiVal = w / (h ^ 2)

If (bmiVal < 18.5) Then
mess = "UnderWeight"
ElseIf (bmiVal < 25) Then
mess = "Normal"
ElseIf (bmiVal < 30) Then
mess = "OverWeight"
Else
mess = "Obese"
End If
Selection.MoveEnd wdTable
Selection.MoveDown
Selection.MoveEnd wdLine, 3
Selection.Delete

Selection.TypeText "BMI = " & Format(bmiVal, "#,##0.00") & vbCr & mess
End Sub
