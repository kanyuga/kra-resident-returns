VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Add Row Functionality For Self Start
Sub addRow_PayeDed_ListS()
    Worksheets("M_Details_of_PAYE_Deducted").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("PayeDed.ListS")
    Worksheets("M_Details_of_PAYE_Deducted").Protect (Pwd)
End Sub
'Add Row Functionality For Self End

'Add Row Functionality For Wife Start
Sub addRow_PayeDed_ListW()
    Worksheets("M_Details_of_PAYE_Deducted").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("PayeDed.ListW")
    Worksheets("M_Details_of_PAYE_Deducted").Protect (Pwd)
End Sub
'Add Row Functionality For Wife End

'Add Row Functionality For Self Start
Sub addRow_Schedule7_9_EmployerIncomeDtls()
    Worksheets("M_Details_of_PAYE_Deducted").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("Sch79.EmployerIncomeDtlsList")
    Worksheets("M_Details_of_PAYE_Deducted").Protect (Pwd)
End Sub
'Add Row Functionality For Self End

'Add Row Functionality For Wife Start
Sub addRow_Schedule8_WithHoldingDtls()
    Worksheets("M_Details_of_PAYE_Deducted").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("Sch79.WithHoldingAmtDtlsList")
    Worksheets("M_Details_of_PAYE_Deducted").Protect (Pwd)
End Sub
'Add Row Functionality For Wife End

Private Sub Worksheet_Change(ByVal Target As Range)
    
    Worksheets("M_Details_of_PAYE_Deducted").Unprotect (Pwd)
    Worksheets("M_Details_of_PAYE_Deducted").Activate
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    If (Target.Columns.Count = 1) Then
        PINNoS = Sheet14.Range("RetInf.PIN").value
        PINNoW = Sheet14.Range("RetInf.SpousePIN").value
        
        'Code for Self Start
        rangeName = Worksheets("M_Details_of_PAYE_Deducted").Range("PayeDed.ListS").Address
        startRow = Worksheets("M_Details_of_PAYE_Deducted").Range(rangeName).row
        endRow = Worksheets("M_Details_of_PAYE_Deducted").Range(rangeName).Rows.Count + startRow - 1
        startColumn = Worksheets("M_Details_of_PAYE_Deducted").Range(rangeName).column
        lastColumn = startColumn + Worksheets("M_Details_of_PAYE_Deducted").Range(rangeName).Columns.Count - 1
        For i = startRow To endRow
            Select Case Target.Address
                'PIN convert in uppar Case Start
                Case "$A$" & i
                     Target.value = StrConv(Target.Text, vbUpperCase)
                     If Target.value <> "" Then
                        For Each r In Range("PayeDed.ListS").Rows
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
            End Select
        Next
    'Code for Self End

    'Code for Wife Start
    rangeName = Worksheets("M_Details_of_PAYE_Deducted").Range("PayeDed.ListW").Address
    startRow = Worksheets("M_Details_of_PAYE_Deducted").Range(rangeName).row
    endRow = Worksheets("M_Details_of_PAYE_Deducted").Range(rangeName).Rows.Count + startRow - 1
    startColumn = Worksheets("M_Details_of_PAYE_Deducted").Range(rangeName).column
    lastColumn = startColumn + Worksheets("M_Details_of_PAYE_Deducted").Range(rangeName).Columns.Count - 1
                     
    For i = startRow To endRow
        Select Case Target.Address
            'PIN convert in uppar Case Start
            Case "$A$" & i
                Target.value = StrConv(Target.Text, vbUpperCase)
                 If Target.value <> "" Then
                    For Each r In Range("PayeDed.ListW").Rows
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
        End Select
    Next
    'Code for Wife End
    'new code added for Bank Section Enable/Disable Start
'        If Sheet20.Range("FinalTax.TaxRefundDueS").value <> "" Then
'           If Sheet20.Range("FinalTax.TaxRefundDueS").value < 0 Then
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
    
    Worksheets("M_Details_of_PAYE_Deducted").Protect (Pwd)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
