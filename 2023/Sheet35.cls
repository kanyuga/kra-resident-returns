VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet35"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Add Row Functionality For Self Start
Sub addRow_VehicleAdvTaxPaid_ListS()
    Worksheets("P_Advance_Tax_Credits").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("VehicleAdvTaxPaid.ListS")
    Worksheets("P_Advance_Tax_Credits").Protect (Pwd)
End Sub
'Add Row Functionality For Self End

'Add Row Functionality For Wife Start
Sub addRow_VehicleAdvTaxPaid_ListW()
    Worksheets("P_Advance_Tax_Credits").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("VehicleAdvTaxPaid.ListW")
    Worksheets("P_Advance_Tax_Credits").Protect (Pwd)
End Sub
'Add Row Functionality For Wife End

Private Sub Worksheet_Change(ByVal Target As Range)

    
    Worksheets("P_Advance_Tax_Credits").Unprotect (Pwd)
    Worksheets("P_Advance_Tax_Credits").Activate
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    If (Target.Columns.Count = 1) Then
        rangeName = Worksheets("P_Advance_Tax_Credits").Range("VehicleAdvTaxPaid.ListS").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1
        startColumn = Range(rangeName).column
        lastColumn = startColumn + Range(rangeName).Columns.Count - 1
        'Code for Self Start
        For i = startRow To endRow
            Select Case Target.Address
                'Validation for Vehicle Registration Number,Vehicle Registration number should be unique for Self Start
                Case "$A$" & i
                    If Target.value <> "" Then
                        Target.value = UCase(Target.value)
                        For Each r In Range("VehicleAdvTaxPaid.ListS").Rows
                            For Each c In r.Cells
                                If c.row > endRow Then
                                    Exit For
                                End If
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                Else
                                    If Target.value <> "" And c.value = Target.value Then
                                        MsgBox "You have already entered details of Vehicle Registration Number  " & Target.value & " in Row " & c.row
                                        Target.value = ""
                                        Target.Select
                                        Exit For
                                        Exit For
                                    End If
                                End If
                            Next
                        Next
                    'Validation for Vehicle Registration Number,Vehicle Registration number should be unique for Self End
                    End If
                    
                'usage code will be set in hidden field in sheet Start
                Case "$C$" & i
'                    If (Target.value <> "") Then
'                        rowvalue = find_UsageCode(Target.value)
'                        usageCode = Sheet18.Cells(rowvalue, usageCodeCol).value
'                        If (rowvalue <> 0) Then
'                            Range("H" & i & ":H" & i).value = usageCode
'                        End If
'                    Else
'                        Worksheets("P_Advance_Tax_Credits").Range("H" & i & ":H" & i).Locked = True
'                        Worksheets("P_Advance_Tax_Credits").Range("H" & i & ":H" & i).value = ""
'                    End If
                    usageCode = Worksheets("P_Advance_Tax_Credits").Range("H" & i & ":H" & i).value
                    If (usageCode = "VAN") Then
                        Call lockUnlock_cell_rng(Worksheets("P_Advance_Tax_Credits").name, "D" & i & ":D" & i, False)
                        Call lockUnlock_cell_rng(Worksheets("P_Advance_Tax_Credits").name, "E" & i & ":E" & i, True)
                        Target.Next.Select
                    ElseIf (usageCode = "SALOON") Then
                        Call lockUnlock_cell_rng(Worksheets("P_Advance_Tax_Credits").name, "D" & i & ":D" & i, True)
                        Call lockUnlock_cell_rng(Worksheets("P_Advance_Tax_Credits").name, "E" & i & ":E" & i, False)
                        Target.Next.Select
                    Else
                        Call lockUnlock_cell_rng(Worksheets("P_Advance_Tax_Credits").name, "D" & i & ":D" & i, False)
                        Call lockUnlock_cell_rng(Worksheets("P_Advance_Tax_Credits").name, "E" & i & ":E" & i, True)
                        Target.Next.Select
                    End If
                'usage code will be set in hidden field in sheet End
            End Select
        Next
        'Code for Self End
    
        'Code for Wife Start
        rangeName = Worksheets("P_Advance_Tax_Credits").Range("VehicleAdvTaxPaid.ListW").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1
        startColumn = Range(rangeName).column
        lastColumn = startColumn + Range(rangeName).Columns.Count - 1
                         
        For i = startRow To endRow
            Select Case Target.Address
                'Validation for Vehicle Registration Number,Vehicle Registration number should be unique for Self Start
                Case "$A$" & i
                    If Target.value <> "" Then
                        Target.value = UCase(Target.value)
                        'Validation for Vehicle Registration Number,Vehicle Registration number should be unique for Self End
                        
                        'Validation for Vehicle Registration Number,Vehicle Registration number should be unique for Wife Start
                        For Each r In Range("VehicleAdvTaxPaid.ListW").Rows
                            For Each c In r.Cells
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                Else
                                    If Target.value <> "" And c.value = Target.value Then
                                        MsgBox "You have already entered details of Vehicle Registration Number " & Target.value & " in Row " & c.row
                                        Target.value = ""
                                        Target.Select
                                        Exit For
                                    Exit For
                                    End If
                                End If
                            Next
                        Next
                'Validation for Vehicle Registration Number,Vehicle Registration number should be unique for Wife End
                    End If
            'usage code will be set in hidden field in sheet Start
                Case "$C$" & i
'                    If (Target.value <> "") Then
'                        rowvalue = find_UsageCode(Target.value)
'                        usageCode = Sheet18.Cells(rowvalue, usageCodeCol).value
'                        If (rowvalue <> 0) Then
'                            Range("H" & i & ":H" & i).value = usageCode
'                        End If
'                    Else
'                        Worksheets("P_Advance_Tax_Credits").Range("H" & i & ":H" & i).Locked = True
'                        Worksheets("P_Advance_Tax_Credits").Range("H" & i & ":H" & i).value = ""
'                    End If
                    usageCode = Worksheets("P_Advance_Tax_Credits").Range("H" & i & ":H" & i).value
                    If (usageCode = "VAN") Then
                        Call lockUnlock_cell_rng(Worksheets("P_Advance_Tax_Credits").name, "D" & i & ":D" & i, False)
                        Call lockUnlock_cell_rng(Worksheets("P_Advance_Tax_Credits").name, "E" & i & ":E" & i, True)
                        Target.Next.Select
                    ElseIf (usageCode = "SALOON") Then
                        Call lockUnlock_cell_rng(Worksheets("P_Advance_Tax_Credits").name, "D" & i & ":D" & i, True)
                        Call lockUnlock_cell_rng(Worksheets("P_Advance_Tax_Credits").name, "E" & i & ":E" & i, False)
                        Target.Next.Select
                    End If
                    
            'usage code will be set in hidden field in sheet End
            End Select
        Next
        'Code for Wife End
        'new code added for Bank Section Enable/Disable Start
'                If Sheet20.Range("FinalTax.TaxRefundDueS").value <> "" Then
'                   If Sheet20.Range("FinalTax.TaxRefundDueS").value < 0 Then
'                        If Sheet14.Range("BankS").Locked = True Then
'                            Call lockUnlock_cell_rng("A_Basic_Info", "BankS", False)
'                            Call lockUnlock_cell_rng("A_Basic_Info", "BranchS", False)
'                            Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.CityS", False)
'                            Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNameS", False)
'                            Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNumberS", False)
'                        End If
'                    Else
'                        If Sheet14.Range("BankS").Locked = False Then
'                            Call lockUnlock_cell_rng("A_Basic_Info", "BankS", True)
'                            Call lockUnlock_cell_rng("A_Basic_Info", "BranchS", True)
'                            Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.CityS", True)
'                            Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNameS", True)
'                            Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNumberS", True)
'                        End If
'                    End If
'                End If
                Worksheets("P_Advance_Tax_Credits").Activate
        'new code added for Bank Section Enable/Disable End
    End If
endSub:
    
    Worksheets("P_Advance_Tax_Credits").Protect (Pwd)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

