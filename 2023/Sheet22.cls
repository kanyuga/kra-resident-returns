VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Go to Next Sheet Start
Sub next_sheet()
    Sheet38.Activate
    Sheet38.Cells(3, 7).Select
End Sub
'Go to Next Sheet End

'Add Row Functionality For Self Start
Sub addRow_WAT_ListS()
Worksheets("E2_CA_WTA_SLM").Unprotect (Pwd)
Call InsertRowsAndFillFormulas("WAT.ListBS")
Worksheets("E2_CA_WTA_SLM").Protect (Pwd)
End Sub
'Add Row Functionality For Self End

'Add Row Functionality For Wife Start
Sub addRow_WAT_ListW()
Worksheets("E2_CA_WTA_SLM").Unprotect (Pwd)
Call InsertRowsAndFillFormulas("WAT.ListBW")
Worksheets("E2_CA_WTA_SLM").Protect (Pwd)
End Sub
'Add Row Functionality For Wife End

Private Sub Worksheet_Change(ByVal Target As Range)
    
    Worksheets("E2_CA_WTA_SLM").Unprotect (Pwd)
    Worksheets("E2_CA_WTA_SLM").Activate
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    If (Target.Columns.Count = 1) Then
        
        rangeName = Worksheets("E2_CA_WTA_SLM").Range("WAT.ListBS").Address
        startRow = Worksheets("E2_CA_WTA_SLM").Range(rangeName).row
        endRow = Worksheets("E2_CA_WTA_SLM").Range(rangeName).Rows.Count + startRow - 1
        startColumn = Worksheets("E2_CA_WTA_SLM").Range(rangeName).column
        lastColumn = startColumn + Worksheets("E2_CA_WTA_SLM").Range(rangeName).Columns.Count - 1
                         
        For i = startRow To endRow
            Select Case Target.Address
                'Validation for Date of Purchase,Date of Purchase should be before or within return period Start
                Case "$D$" & i
                    If (Target.value <> "") Then
                        result = TestDate(Target.value)
                        If (result = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            If (Date <> "") Then
                                If TestDateBeforeRtnPeriodEnd(Target.value) = False Then
                                    'MsgBox "Date of Purchase should be before or within Return Period."
                                    MsgBox "Date of Purchase in Row " & i & " should be before or within Return period " & Sheet14.Range("RetInf.RetStartDate").value & " to " & Sheet14.Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                End If
                            End If
                        End If
                    End If
                'Validation for Date of Purchase,Date of Purchase should be before or within return period End
                    
                'Validation for Disposal at Cose During the year,Disposal at Cost during the Year should be less than or equal to Actual Cost Start
                Case "$G$" & i
                    cellRow = Target.row
                    If (Worksheets("E2_CA_WTA_SLM").Range("G" & cellRow & ":G" & cellRow).value <> "" And Worksheets("E2_CA_WTA_SLM").Range("F" & cellRow & ":F" & cellRow).value <> "") Then
                        If (Worksheets("E2_CA_WTA_SLM").Range("G" & cellRow & ":G" & cellRow).value > Worksheets("E2_CA_WTA_SLM").Range("F" & cellRow & ":F" & cellRow).value) Then
                            MsgBox "Disposal at Cost during the Year should be less than or equal to Actual Cost."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'Validation for Disposal at Cose During the year,Disposal at Cost during the Year should be less than or equal to Actual Cost End
                    
                'Valdation for Actual Cost,Actual Cost should be greater than or equal to Disposal at Cost during the Year Start
                Case "$F$" & i
                    cellRow = Target.row
                    If (Worksheets("E2_CA_WTA_SLM").Range("F" & cellRow & ":F" & cellRow).value <> "" And Worksheets("E2_CA_WTA_SLM").Range("G" & cellRow & ":G" & cellRow).value <> "") Then
                        If (Worksheets("E2_CA_WTA_SLM").Range("F" & cellRow & ":F" & cellRow).value < Worksheets("E2_CA_WTA_SLM").Range("G" & cellRow & ":G" & cellRow).value) Then
                            MsgBox " Actual Cost should be greater than or equal to Disposal at Cost during the Year."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'Valdation for Actual Cost,Actual Cost should be greater than or equal to Disposal at Cost during the Year End
            End Select
        Next
    
        rangeName = Worksheets("E2_CA_WTA_SLM").Range("WAT.ListBW").Address
        startRow = Worksheets("E2_CA_WTA_SLM").Range(rangeName).row
        endRow = Worksheets("E2_CA_WTA_SLM").Range(rangeName).Rows.Count + startRow - 1
        startColumn = Worksheets("E2_CA_WTA_SLM").Range(rangeName).column
        lastColumn = startColumn + Worksheets("E2_CA_WTA_SLM").Range(rangeName).Columns.Count - 1
                         
        For i = startRow To endRow
            Select Case Target.Address
                'Validation for Date of Purchase,Date of Purchase should be before or within return period Start
                Case "$D$" & i
                    If (Target.value <> "") Then
                        result = TestDate(Target.value)
                        If (result = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            If (Date <> "") Then
                                If TestDateBeforeRtnPeriodEnd(Target.value) = False Then
                                    'MsgBox "Date of Purchase should be before or within Return Period."
                                    MsgBox "Date of Purchase in Row " & i & " should be before or within Return period " & Sheet14.Range("RetInf.RetStartDate").value & " to " & Sheet14.Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                End If
                            End If
                        End If
                    End If
                'Validation for Date of Purchase,Date of Purchase should be before or within return period End
                     
                'Validation for Disposal at Cose During the year,Disposal at Cost during the Year should be less than or equal to Actual Cost Start
                Case "$G$" & i
                    cellRow = Target.row
                    If (Worksheets("E2_CA_WTA_SLM").Range("G" & cellRow & ":G" & cellRow).value <> "" And Worksheets("E2_CA_WTA_SLM").Range("F" & cellRow & ":F" & cellRow).value <> "") Then
                        If (Worksheets("E2_CA_WTA_SLM").Range("G" & cellRow & ":G" & cellRow).value > Worksheets("E2_CA_WTA_SLM").Range("F" & cellRow & ":F" & cellRow).value) Then
                            MsgBox "Disposal at Cost during the Year should be less than or equal to Actual Cost."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'Validation for Disposal at Cose During the year,Disposal at Cost during the Year should be less than or equal to Actual Cost End
                    
                'Valdation for Actual Cost,Actual Cost should be greater than or equal to Disposal at Cost during the Year Start
                Case "$F$" & i
                    cellRow = Target.row
                    If (Worksheets("E2_CA_WTA_SLM").Range("F" & cellRow & ":F" & cellRow).value <> "" And Worksheets("E2_CA_WTA_SLM").Range("G" & cellRow & ":G" & cellRow).value <> "") Then
                        If (Worksheets("E2_CA_WTA_SLM").Range("F" & cellRow & ":F" & cellRow).value < Worksheets("E2_CA_WTA_SLM").Range("G" & cellRow & ":G" & cellRow).value) Then
                            MsgBox " Actual Cost should be greater than or equal to Disposal at Cost during the Year."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'Valdation for Actual Cost,Actual Cost should be greater than or equal to Disposal at Cost during the Year End
            End Select
        Next
    
    End If
    
    Worksheets("E2_CA_WTA_SLM").Protect (Pwd)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

