VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Add Row Functionality For Self Start

Sub addRow_EmpIncome_ListS()
    Worksheets("F_Employment_Income").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("EmpIncome.ListS")
    rangeName = Worksheets("F_Employment_Income").Range("EmpIncome.ListS").Address
    startRowIndex = Range(rangeName).row
    endRowIndex = Range(rangeName).Rows.Count + startRowIndex - 1
    Dim Str As String
    Str = "A" & startRowIndex & ":A" & endRowIndex
    ActiveWorkbook.Names("PINofEmployerS").Delete
    Range(Str).name = "PINofEmployerS"
    Worksheets("F_Employment_Income").Protect (Pwd)
End Sub
'Add Row Functionality For Self End

'Add Row Functionality For Wife Start
Sub addRow_EmpIncome_ListW()
    Worksheets("F_Employment_Income").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("EmpIncome.ListW")
    rangeName = Worksheets("F_Employment_Income").Range("EmpIncome.ListW").Address
    startRowIndex = Range(rangeName).row
    endRowIndex = Range(rangeName).Rows.Count + startRowIndex - 1
    Dim Str As String
    Str = "A" & startRowIndex & ":A" & endRowIndex
    ActiveWorkbook.Names("PINofEmployerW").Delete
    Range(Str).name = "PINofEmployerW"
    Worksheets("F_Employment_Income").Protect (Pwd)
    End Sub
'Add Row Functionality For Wife End

Private Sub Worksheet_Change(ByVal Target As Range)

    
    Worksheets("F_Employment_Income").Unprotect (Pwd)
    Worksheets("F_Employment_Income").Activate
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'Added by Ruth and Lawrence on 29/12/2020
    'Validates if Input in the added 2020 cells is Numeric
    If Target.Address = Range("EmpIncome.ListSTOJANMAR").Address Then
        Call validateIsNumeric(Target)
    End If
    If Target.Address = Range("EmpIncome.ListSPENSJANMAR").Address Then
        Call validateIsNumeric(Target)
    End If
    If Target.Address = Range("EmpIncome.ListSTOAPRDEC").Address Then
        Call validateIsNumeric(Target)
    End If
    If Target.Address = Range("EmpIncome.ListSPENSAPRDEC").Address Then
        Call validateIsNumeric(Target)
    End If
    
    If Target.Address = Range("EmpIncome.ListWTOJANMAR").Address Then
        Call validateIsNumeric(Target)
    End If
    If Target.Address = Range("EmpIncome.ListWPENJANMAR").Address Then
        Call validateIsNumeric(Target)
    End If
    If Target.Address = Range("EmpIncome.ListWTOAPRDEC").Address Then
        Call validateIsNumeric(Target)
    End If
    If Target.Address = Range("EmpIncome.ListWPENAPRDEC").Address Then
        Call validateIsNumeric(Target)
    End If

    If (Target.Columns.Count = 1) Then
        PINNoS = Sheet14.Range("RetInf.PIN").value
        PINNoW = Sheet14.Range("RetInf.SpousePIN").value
        
        Dim resultMsg As String
        rangeName = Range("EmpIncome.ListS").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1
    
        For i = startRow To endRow
            Select Case Target.Address
                'PIN convert in uppar Case Start
                Case "$A$" & i
                     Target.value = StrConv(Target.Text, vbUpperCase)
                     If Target.value <> "" Then
                        For Each r In Range("EmpIncome.ListS").Rows
                            For Each c In r.Cells
                                If c.row > endRow Then
                                    Exit For
                                End If
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                Else
                                    If UCase(Target.value <> "") And UCase(c.value) = UCase(Target.value) Then
                                        MsgBox "You have already entered details of PIN of Employer " & Target.value & " in Row " & c.row
                                        Target.value = ""
                                        Target.Select
                                        Exit For
                                        Exit For
                                    End If
                                End If
                            Next
                        Next
                    End If
                    
                    If (Target.value <> "") Then
                        If (Target.value = PINNoS) Then
                            MsgBox "PIN entered should not be same as Taxpayer's PIN."
                            Target.value = ""
                            Target.Select
                        End If
                     End If
                     If (Target.value <> "") Then
                        If (Target.value = PINNoW) Then
                            MsgBox "PIN entered should not be same as Wife's PIN."
                            Target.value = ""
                            Target.Select
                        End If
                     End If
                     If (Target.value <> "") Then
                        result = TestPIN(Target.value)
                        If (result = False) Then
                            MsgBox "Please enter PIN of Employer in proper format."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                 'PIN convert in uppar Case End
                Case "$G$" & i
                    If (Target.value <> "") Then
                        If (Target.value = 0) Then
                            'Do nothing
                        Else
                            If (Target.value < 300000) Then
                                MsgBox "'Pension if in excess of 300,000 (Ksh)' value should be 300,000 Or greater than 300,000."
                                Target.value = 0#
                                Target.Select
                            End If
                        End If
                    End If
            End Select
        Next


    rangeName = Range("EmpIncome.ListW").Address
    startRow = Range(rangeName).row
    endRow = Range(rangeName).Rows.Count + startRow - 1

    For i = startRow To endRow
        Select Case Target.Address
            'PIN convert in uppar Case Start
            Case "$A$" & i
                 Target.value = StrConv(Target.Text, vbUpperCase)
                 
                 If Target.value <> "" Then
                    For Each r In Range("EmpIncome.ListW").Rows
                        For Each c In r.Cells
                            If c.row > endRow Then
                                Exit For
                            End If
                            If c.Address = Target.Address Then 'Skip comparison with self value
                            Else
                                If UCase(Target.value <> "") And UCase(c.value) = UCase(Target.value) Then
                                    MsgBox "You have already entered details of PIN of Employer " & Target.value & " in Row " & c.row
                                    Target.value = ""
                                    Target.Select
                                    Exit For
                                    Exit For
                                End If
                            End If
                        Next
                    Next
                End If
                
                If (Target.value <> "") Then
                    If (Target.value = PINNoS) Then
                        MsgBox "PIN entered should not be same as Taxpayer's PIN."
                        Target.value = ""
                        Target.Select
                    End If
                 End If
                 If (Target.value <> "") Then
                    If (Target.value = PINNoW) Then
                        MsgBox "PIN entered should not be same as Wife's PIN."
                        Target.value = ""
                        Target.Select
                    End If
                 End If
                 If (Target.value <> "") Then
                    result = TestPIN(Target.value)
                    If (result = False) Then
                        MsgBox "Please enter PIN of Employer in proper format."
                        Target.value = ""
                        Target.Select
                    End If
                End If
            'PIN convert in uppar Case End
            Case "$G$" & i
                    If (Target.value <> "") Then
                        If (Target.value = 0) Then
                                'Do nothing
                            Else
                                If (Target.value < 300000) Then
                                    MsgBox "'Pension if in excess of 300,000 (Ksh)' value should be 300,000 Or greater than 300,000."
                                    Target.value = 0#
                                    Target.Select
                                End If
                            End If
                    End If
        End Select
    Next
    'new code added for Bank Section Enable/Disable Start
'            If Sheet20.Range("FinalTax.TaxRefundDueS").value <> "" Then
'               If Sheet20.Range("FinalTax.TaxRefundDueS").value < 0 Then
'                    If Sheet14.Range("BankS").Locked = True Then
'                        Call lockUnlock_cell_rng("A_Basic_Info", "BankS", False)
'                        Call lockUnlock_cell_rng("A_Basic_Info", "BranchS", False)
'                        Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.CityS", False)
'                        Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNameS", False)
'                        Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNumberS", False)
'                    End If
'                Else
'                    If Sheet14.Range("BankS").Locked = False Then
'                        Call lockUnlock_cell_rng("A_Basic_Info", "BankS", True)
'                        Call lockUnlock_cell_rng("A_Basic_Info", "BranchS", True)
'                        Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.CityS", True)
'                        Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNameS", True)
'                        Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNumberS", True)
'                    End If
'                End If
'            End If
    'new code added for Bank Section Enable/Disable End
    End If

endSub:
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Worksheets("F_Employment_Income").Protect (Pwd)
End Sub

