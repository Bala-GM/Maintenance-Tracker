Attribute VB_Name = "Module2"
Sub SORT()
Attribute SORT.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SORT Macro
'

'
    ActiveWindow.SmallScroll Down:=-9
    Range("K1:M2").Select
    ActiveSheet.Range("$A$3:$H$382").AutoFilter Field:=4, Criteria1:=RGB(255, _
        199, 206), Operator:=xlFilterCellColor
End Sub
Sub Clear()
Attribute Clear.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Clear Macro
'

'
    ActiveSheet.ShowAllData
End Sub
