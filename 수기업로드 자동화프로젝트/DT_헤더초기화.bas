Sub DT_헤더초기화()

    On Error Resume Next
    
    Sheets("사업장").UsedRange.Find(What:="사업장", LookAt:=xlWhole).Resize(, 17).Cut
    Sheets("사업장").Range("B4:R4").Insert Shift:=xlDown
   
End Sub