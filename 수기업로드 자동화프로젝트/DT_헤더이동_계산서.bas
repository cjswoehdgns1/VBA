Sub DT_헤더이동_계산서()
'
' 매크로2 매크로
'
' 바로 가기 키: Ctrl+e
'
    Dim cellvalue As String
    On Error Resume Next
    
    Sheets("국세청 금액").Select
    Cells(1, 4).End(xlDown).Select
    
    cellvalue = ActiveCell.Value
    
    
    Sheets("사업장").UsedRange.Find(What:="사업장", LookAt:=xlWhole).Resize(, 17).Cut
    Sheets("사업장").UsedRange.Find(What:=cellvalue, LookAt:=xlWhole).Offset(1, 0).Insert Shift:=xlDown

    Sheets("사업장").Select

' xlwhole= 전체일치, xlpart=부분일치
' entirerow 메서드 활용(행전체 선택)

    
    
    
    
End Sub