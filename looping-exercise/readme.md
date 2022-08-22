Option Explicit

Sub 自動依學生班級建立各班級成績單()

    Dim i As Integer, sht As Worksheet
    Set sht = Worksheets("成績表")
    i = 2
    While sht.Cells(i, 1) <> ""
        
        On Error Resume Next
        
        If Worksheets(sht.Cells(i, 1).Value) Is Nothing Then
        
            Worksheets.Add after:=Worksheets(Worksheets.Count)
            ActiveSheet.Name = sht.Cells(i, 1)
            
        End If
    
        i = i + 1
    Wend
    
         
    Worksheets("成績表").Select
    Range("A1:E1").Select
    Selection.Copy
    
    For i = 2 To Worksheets.Count
        Sheets(i).Select
        Cells(1, "A").Select
        ActiveSheet.Paste
    Next
        

    Dim s As String
    
    i = 2
    While Worksheets("成績表").Cells(i, 1) <> ""
        Worksheets("成績表").Select
        Range(Cells(i, "A"), Cells(i, "E")).Select
        Selection.Copy
        s = Cells(i, 1).Value
        Worksheets(s).Select
        
        Selection.End(xlToRight).Select
        Selection.End(xlDown).Select
        Selection.End(xlToLeft).Select
        Selection.End(xlUp).Select
        ActiveCell.Offset(1, 0).Range("A1").Select
        ActiveSheet.Paste
        i = i + 1
    Wend
    
End Sub



