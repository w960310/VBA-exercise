# VBA-exercise
## I. looping-exercise
### EX1.自動依學生成績標色
將學生成績按照分數高低改變顏色，90分以上改為紅色，90分以下改為藍色。   
  
執行巨集後:  
![image](looping-pictures/自動依學生成績標色(標色後).jpg)  
  
code:  
  
Option Explicit

Sub 標顏色() Dim i As Integer, j As Integer

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
  
  
### EX2.自動依學生班級建立各班級成績單
將混雜的全校學生成績，按照A欄的班級，自動建立各個班級的成績單。  
並再使用完畢後，可以點擊還原按鈕回到初始狀態，以便下次套用其他資料使用。   
  
執行巨集前:  
![image](looping-pictures/自動依學生班級建立各班級成績單(巨集前).jpg)    
  
執行巨集後:  
![image](looping-pictures/自動依學生班級建立各班級成績單(巨集後).jpg)
