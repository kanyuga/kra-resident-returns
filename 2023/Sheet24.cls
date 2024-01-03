VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub addRow_DTAADtls_ListS()
    Worksheets("R_DTAA_Credits").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("DTAACredits.DetailsS")
    Worksheets("R_DTAA_Credits").Protect (Pwd)
End Sub

Sub addRow_DTAADtls_ListW()
    Worksheets("R_DTAA_Credits").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("DTAACredits.DetailsW")
    Worksheets("R_DTAA_Credits").Protect (Pwd)
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    
    Worksheets("R_DTAA_Credits").Unprotect (Pwd)
    Worksheets("R_DTAA_Credits").Activate
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    If (Target.Columns.Count = 1) Then

        Dim cellRow As Integer
        Dim startRow As Integer
        Dim endRow As Integer
        Dim rangeAddress As String
    
        Dim startDate As String
        Dim endDate As String
        Dim StartDepositDate As String
        Dim DsysDate As String
    
        If (Sheet14.Range("RetInf.RetEndDate").value <> "") Then
            If (TestDate(Sheet14.Range("RetInf.RetEndDate").value) = True) Then
                endDate = CDate(Format(Sheet14.Range("RetInf.RetEndDate").value, "dd/mm/yyyy"))
            End If
        End If
    
        
        Dim resultMsg As String
        rangeName = Worksheets("R_DTAA_Credits").Range("DTAACredits.DetailsS").Address
        startRow = Worksheets("R_DTAA_Credits").Range(rangeName).row
        endRow = Worksheets("R_DTAA_Credits").Range(rangeName).Rows.Count + startRow - 1
    
        startColumn = Worksheets("R_DTAA_Credits").Range(rangeName).column
        lastColumn = startColumn + Worksheets("R_DTAA_Credits").Range(rangeName).Columns.Count - 1
        
        'Self Start

        For i = startRow To endRow
            Select Case Target.Address
                Case "$C$" & i
                    If (Target.value <> "") Then
                        result = TestDate(Target.value)
                        If (result = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            If (endDate <> "") Then
                                If TestDateBeforeRtnPeriodEnd(Target.value) = False Then
                                    MsgBox "Date of Deduction  in Row " & i & " should be before or within Return Period."
                                    Target.value = ""
                                    Target.Select
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Sheet14.Activate
                                Sheet14.Range("RetInf.RetStartDate").Select
                                GoTo endSub
                            End If
                        End If
                    End If
            End Select
        Next
        'Self End
        
        'Wife Start
        
        rangeName = Worksheets("R_DTAA_Credits").Range("DTAACredits.DetailsW").Address
        startRow = Worksheets("R_DTAA_Credits").Range(rangeName).row
        endRow = Worksheets("R_DTAA_Credits").Range(rangeName).Rows.Count + startRow - 1
    
        startColumn = Worksheets("R_DTAA_Credits").Range(rangeName).column
        lastColumn = startColumn + Worksheets("R_DTAA_Credits").Range(rangeName).Columns.Count - 1

        For i = startRow To endRow
            Select Case Target.Address
                Case "$C$" & i
                    If (Target.value <> "") Then
                        result = TestDate(Target.value)
                        If (result = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            If (endDate <> "") Then
                                If TestDateBeforeRtnPeriodEnd(Target.value) = False Then
                                    MsgBox "Date of Deduction  in Row " & i & " should be before or within Return Period."
                                    Target.value = ""
                                    Target.Select
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Sheet14.Activate
                                Sheet14.Range("RetInf.RetStartDate").Select
                                GoTo endSub
                            End If
                        End If
                    End If
            End Select
        Next
        'Wife End
     
        'new code added for Bank Section Enable/Disable Start
'        If Sheet20.Range("FinalTax.TaxRefundDueS").value <> "" Then
'            If Sheet20.Range("FinalTax.TaxRefundDueS").value < 0 Then
'                If Sheet14.Range("BankS").Locked = True Then
'                    Call lockUnlock_cell_rng("A_Basic_Info", "BankS", False)
'                    Call lockUnlock_cell_rng("A_Basic_Info", "BranchS", False)
'                    Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.CityS", False)
'                    Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNameS", False)
'                    Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNumberS", False)
'                End If
'            Else
'                If Sheet14.Range("BankS").Locked = False Then
'                    Call lockUnlock_cell_rng("A_Basic_Info", "BankS", True)
'                    Call lockUnlock_cell_rng("A_Basic_Info", "BranchS", True)
'                    Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.CityS", True)
'                    Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNameS", True)
'                    Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNumberS", True)
'                End If
'            End If
'        End If
    'new code added for Bank Section Enable/Disable End
    End If
endSub:
    
    Worksheets("R_DTAA_Credits").Protect (Pwd)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


