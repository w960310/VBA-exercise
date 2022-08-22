Option Explicit

Sub 標顏色()
    Dim i As Integer, j As Integer
       
    i = 2
   
    While Cells(i, 3) <> ""
        j = 3
        While Cells(i, j) <> ""
            If Cells(i, j) >= 90 Then
            
                Cells(i, j).Select
                With Selection.Font
                .Color = 255
                End With
            ElseIf Cells(i, j) < 90 Then
            
                Cells(i, j).Select
                With Selection.Font
                     .ThemeColor = xlThemeColorAccent1
                End With
            
            End If
            j = j + 1
        Wend
        i = i + 1
    Wend

End Sub

