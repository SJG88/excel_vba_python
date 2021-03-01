'This script demos how to interact with tick boxes in excel used to hide and unhide
'tabs in an excel workbook. 

'Written by Shane Gore

Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then
    Worksheets("Report_sheet1").Visible = True
Else
    Worksheets("Report_sheet1").Visible = xlSheetVeryHidden
End If
    
End Sub

Private Sub CheckBox2_Click()
If CheckBox2.Value = True Then
    Worksheets("Report_sheet2").Visible = True
Else
    Worksheets("Report_sheet2").Visible = xlSheetVeryHidden
End If
End Sub


Private Sub CheckBox3_Click()
If CheckBox3.Value = True Then
    Worksheets("Report_sheet3").Visible = True
Else
    Worksheets("Report_sheet3").Visible = xlSheetVeryHidden
End If
End Sub
