VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Add Row Functionality For Self Start
Sub addRow_InstallmentTax_ListS()
    Worksheets("N_Installment_Tax_Credits").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("InstallmentTax.ListS")
    Worksheets("N_Installment_Tax_Credits").Protect (Pwd)
End Sub
'Add Row Functionality For Self End

'Add Row Functionality For Self Start
Sub addRow_InstallmentTax_ListW()
    Worksheets("N_Installment_Tax_Credits").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("InstallmentTax.ListW")
    Worksheets("N_Installment_Tax_Credits").Protect (Pwd)
End Sub
'Add Row Functionality For Self End


Private Sub Worksheet_Change(ByVal Target As Range)

    
    Worksheets("N_Installment_Tax_Credits").Unprotect (Pwd)
    Worksheets("N_Installment_Tax_Credits").Activate
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
    
        If (Sheet14.Range("RetInf.RetStartDate").value <> "") Then
            If (TestDate(Sheet14.Range("RetInf.RetStartDate").value) = True) Then
                startDate = CDate(Format(Sheet14.Range("RetInf.RetStartDate").value, "dd/mm/yyyy"))
                DsysDate = Date
            End If
        End If

        If (Sheet14.Range("RetInf.RetEndDate").value <> "") Then
            If (TestDate(Sheet14.Range("RetInf.RetEndDate").value) = True) Then
                endDate = CDate(Format(Sheet14.Range("RetInf.RetEndDate").value, "dd/mm/yyyy"))
            End If
        End If
    
        If (Sheet14.Range("RetInf.DepositStartDate").value <> "") Then
            If (TestDate(Sheet14.Range("RetInf.DepositStartDate").value) = True) Then
                StartDepositDate = CDate(Format(Sheet14.Range("RetInf.DepositStartDate").value, "dd/mm/yyyy"))
            End If
        End If
        
        'Code for Self Start
        rangeName = Worksheets("N_Installment_Tax_Credits").Range("InstallmentTax.ListS").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1

        For i = startRow To endRow
            Select Case Target.Address
            'Validation for PRN Number,PRN number should be unique and convert in uppar Case for Self Start
                Case "$A$" & i
                    If Target.value <> "" Then
                        Target.value = UCase(Target.value)
                        For Each r In Range("InstallmentTax.ListS").Rows
                            For Each c In r.Cells
                                If c.row > endRow Then
                                    Exit For
                                End If
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                Else
                                    If Target.value <> "" And c.value = Target.value Then
                                        MsgBox "You have already entered details of PRN " & Target.value & " in Row " & c.row
                                        Target.value = ""
                                        Target.Select
                                        Exit For
                                    Exit For
                                    End If
                                End If
                            Next
                        Next
                        'Validation for PRN Number,PRN number should be unique and convert in uppar Case for Self End
                        
                        'Validation for PRN Number,PRN number should be unique and convert in uppar Case for Wife Start
                        For Each r In Range("InstallmentTax.ListW").Rows
                            For Each c In r.Cells
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                Else
                                    If Target.value <> "" And c.value = Target.value Then
                                        MsgBox "You have already entered details of PRN " & Target.value & " in Row " & c.row
                                        Target.value = ""
                                        Target.Select
                                        Exit For
                                    Exit For
                                    End If
                                End If
                            Next
                        Next
                        'Validation for PRN Number,PRN number should be unique and convert in uppar Case for Wife End
                        
                        'Validation for PRN Number,Same PRN should not be allowed in other PRN related Section for Self Start
                        
                        rangeName = Sheet41.Range("DtlIncomePaid.IncomePaidAdvanceListS").Address
                        startRowRange = Sheet41.Range(rangeName).row
                        endRowRange = Sheet41.Range(rangeName).Rows.Count + startRow - 1
                        For j = startRowRange To endRowRange
                            If Target.value <> "" And Sheet41.Range("A" & j & ":A" & j) = Target.value Then
                                MsgBox "You have already entered details of PRN " & Target.value & " in " & j & "th Row of Q_IT_Payment_Credits"
                                Target.value = ""
                                Target.Select
                                Exit For
                            End If
                        Next
                        'Validation for PRN Number,Same PRN should not be allowed in other PRN related Section for Self End
                        
                        'Validation for PRN Number,Same PRN should not be allowed in other PRN related Section for Self Start
                        
                        rangeName = Sheet41.Range("DtlIncomePaid.IncomePaidSelfAssmntListS").Address
                        startRowRange = Sheet41.Range(rangeName).row
                        endRowRange = Sheet41.Range(rangeName).Rows.Count + startRowRange - 1
                        For j = startRowRange To endRowRange
                            If Target.value <> "" And Sheet41.Range("A" & j & ":A" & j) = Target.value Then
                                MsgBox "You have already entered details of PRN " & Target.value & " in " & j & "th Row of Q_IT_Payment_Credits"
                                Target.value = ""
                                Target.Select
                                Exit For
                            End If
                        Next
                        'Validation for PRN Number,Same PRN should not be allowed in other PRN related Section for Self End
                        
                        'Validation for PRN Number,Same PRN should not be allowed in other PRN related Section for Wife Start
                        rangeName = Sheet41.Range("DtlIncomePaid.IncomePaidAdvanceListW").Address
                        startRowRange = Sheet41.Range(rangeName).row
                        endRowRange = Sheet41.Range(rangeName).Rows.Count + startRow - 1
                        For j = startRowRange To endRowRange
                            If Target.value <> "" And Sheet41.Range("A" & j & ":A" & j) = Target.value Then
                                MsgBox "You have already entered details of PRN " & Target.value & " in " & j & "th Row of Q_IT_Payment_Credits"
                                Target.value = ""
                                Target.Select
                                Exit For
                            End If
                        Next
                        'Validation for PRN Number,Same PRN should not be allowed in other PRN related Section for Wife End
                        
                        'Validation for PRN Number,Same PRN should not be allowed in other PRN related Section for Wife Start
                        
                        rangeName = Sheet41.Range("DtlIncomePaid.IncomePaidSelfAssmntListW").Address
                        startRowRange = Sheet41.Range(rangeName).row
                        endRowRange = Sheet41.Range(rangeName).Rows.Count + startRowRange - 1
                        For j = startRowRange To endRowRange
                            If Target.value <> "" And Sheet41.Range("A" & j & ":A" & j) = Target.value Then
                                MsgBox "You have already entered details of PRN " & Target.value & " in " & j & "th Row of Q_IT_Payment_Credits"
                                Target.value = ""
                                Target.Select
                                Exit For
                            End If
                        Next
                        'Validation for PRN Number,Same PRN should not be allowed in other PRN related Section for Wife End
                End If
                'Validation for Date of Deposit,Date of Deposit should be before 1year back from Return period and current Date Start
                Case "$B$" & i
                    Dim tSValue As String
                    If (Target.value <> "") Then
                         result = TestDate(Target.value)
                         If (result = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            tSValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                            If (startDate <> "" And endDate <> "") Then
                                If TestDateBtwnRtnPeriod(Target.value) = False Then
                                    MsgBox "Date of Deposit in Row " & i & " should be within Return Period " & Sheet14.Range("RetInf.RetStartDate").value & " to " & Sheet14.Range("RetInf.RetEndDate").value
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
            'Validation for Date of Deposit,Date of Deposit should be before 1year back from Return period and current Date End
        End Select
    Next
    'Code for Self End

    'Code for Wife Start
    rangeName = Worksheets("N_Installment_Tax_Credits").Range("InstallmentTax.ListW").Address
    startRow = Worksheets("N_Installment_Tax_Credits").Range(rangeName).row
    endRow = Worksheets("N_Installment_Tax_Credits").Range(rangeName).Rows.Count + startRow - 1
                    
    For i = startRow To endRow
        Select Case Target.Address
            'Validation for PRN Number,PRN number should be unique and convert in uppar Case for Self Start
            Case "$A$" & i
                If Target.value <> "" Then
                    Target.value = UCase(Target.value)
                    For Each r In Range("InstallmentTax.ListS").Rows
                        For Each c In r.Cells
                            If c.Address = Target.Address Then 'Skip comparison with self value
                            Else
                                If Target.value <> "" And c.value = Target.value Then
                                    MsgBox "You have already entered details of PRN " & Target.value & " in Row " & c.row
                                    Target.value = ""
                                    Target.Select
                                    Exit For
                                Exit For
                                End If
                            End If
                        Next
                    Next
            'Validation for PRN Number,PRN number should be unique and convert in uppar Case for Self End
            
            'Validation for PRN Number,PRN number should be unique and convert in uppar Case for Wife Start
                    For Each r In Range("InstallmentTax.ListW").Rows
                        For Each c In r.Cells
                                If c.row > endRow Then
                                    Exit For
                                End If
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                Else
                                    If Target.value <> "" And c.value = Target.value Then
                                        MsgBox "You have already entered details of PRN " & Target.value & " in Row " & c.row
                                        Target.value = ""
                                        Target.Select
                                        Exit For
                                    Exit For
                                    End If
                                End If
                        Next
                    Next
            'Validation for PRN Number,PRN number should be unique and convert in uppar Case for Wife End
                    
            'Validation for PRN Number,Same PRN should not be allowed in other PRN related Section for Self Start
                    
                    rangeName = Sheet41.Range("DtlIncomePaid.IncomePaidAdvanceListS").Address
                    startRowRange = Sheet41.Range(rangeName).row
                    endRowRange = Sheet41.Range(rangeName).Rows.Count + startRow - 1
                    For j = startRowRange To endRowRange
                        If Target.value <> "" And Sheet41.Range("A" & j & ":A" & j) = Target.value Then
                            MsgBox "You have already entered details of PRN " & Target.value & " in " & j & "th Row of Q_IT_Payment_Credits"
                            Target.value = ""
                            Target.Select
                            Exit For
                        End If
                    Next
            'Validation for PRN Number,Same PRN should not be allowed in other PRN related Section for Self End
            
            'Validation for PRN Number,Same PRN should not be allowed in other PRN related Section for Self Start
                    rangeName = Sheet41.Range("DtlIncomePaid.IncomePaidSelfAssmntListS").Address
                    startRowRange = Sheet41.Range(rangeName).row
                    endRowRange = Sheet41.Range(rangeName).Rows.Count + startRowRange - 1
                    For j = startRowRange To endRowRange
                        If Target.value <> "" And Sheet41.Range("A" & j & ":A" & j) = Target.value Then
                            MsgBox "You have already entered details of PRN " & Target.value & " in " & j & "th Row of Q_IT_Payment_Credits"
                            Target.value = ""
                            Target.Select
                            Exit For
                        End If
                    Next
            'Validation for PRN Number,Same PRN should not be allowed in other PRN related Section for Self End
           
            'Validation for PRN Number,Same PRN should not be allowed in other PRN related Section for Wife Start
                    rangeName = Sheet41.Range("DtlIncomePaid.IncomePaidAdvanceListW").Address
                    startRowRange = Sheet41.Range(rangeName).row
                    endRowRange = Sheet41.Range(rangeName).Rows.Count + startRow - 1
                    For j = startRowRange To endRowRange
                        If Target.value <> "" And Sheet41.Range("A" & j & ":A" & j) = Target.value Then
                            MsgBox "You have already entered details of PRN " & Target.value & " in " & j & "th Row of Q_IT_Payment_Credits"
                            Target.value = ""
                            Target.Select
                            Exit For
                        End If
                    Next
            'Validation for PRN Number,Same PRN should not be allowed in other PRN related Section for Wife End
            
             'Validation for PRN Number,Same PRN should not be allowed in other PRN related Section for Wife Start
                    rangeName = Sheet41.Range("DtlIncomePaid.IncomePaidSelfAssmntListW").Address
                    startRowRange = Sheet41.Range(rangeName).row
                    endRowRange = Sheet41.Range(rangeName).Rows.Count + startRowRange - 1
                    For j = startRowRange To endRowRange
                        If Target.value <> "" And Sheet41.Range("A" & j & ":A" & j) = Target.value Then
                            MsgBox "You have already entered details of PRN " & Target.value & " in " & j & "th Row of Q_IT_Payment_Credits"
                            Target.value = ""
                            Target.Select
                            Exit For
                        End If
                    Next
                    'Validation for PRN Number,Same PRN should not be allowed in other PRN related Section for Wife End
            End If
            'Validation for Date of Deposit,Date of Deposit should be before 1year back from Return period and current Date Start
            Case "$B$" & i
                Dim tWValue As String
                If (Target.value <> "") Then
                    result = TestDate(Target.value)
                    If (result = False) Then
                        MsgBox "Please enter date value in dd/mm/yyyy format."
                        Target.value = ""
                        Target.Select
                    Else
                        tWValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                        If (startDate <> "" And endDate <> "") Then
                            If TestDateBtwnRtnPeriod(Target.value) = False Then
                                MsgBox "Date of Deposit in Row " & i & " should be within Return Period " & Sheet14.Range("RetInf.RetStartDate").value & " to " & Sheet14.Range("RetInf.RetEndDate").value
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
            'Validation for Date of Deposit,Date of Deposit should be before 1year back from Return period and current Date End
        End Select
    Next
    'Code for Wife End
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
    
    Worksheets("N_Installment_Tax_Credits").Protect (Pwd)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
