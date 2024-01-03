VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Add Row Functionality For Self Start
Sub addRow_EstateTrust_ListS()
Worksheets("H_Estate_Trust_Income").Unprotect (Pwd)
Call InsertRowsAndFillFormulas("EstateTrust.ListS")
Worksheets("H_Estate_Trust_Income").Protect (Pwd)
End Sub
'Add Row Functionality For Self End

'Add Row Functionality For Wife Start
Sub addRow_EstateTrust_ListW()
Worksheets("H_Estate_Trust_Income").Unprotect (Pwd)
Call InsertRowsAndFillFormulas("EstateTrust.ListW")
Worksheets("H_Estate_Trust_Income").Protect (Pwd)
End Sub
'Add Row Functionality For Wife End

Private Sub Worksheet_Change(ByVal Target As Range)
    
    
    Worksheets("H_Estate_Trust_Income").Unprotect (Pwd)
    Worksheets("H_Estate_Trust_Income").Activate
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    If (Target.Columns.Count = 1) Then
        PINNoS = Sheet14.Range("RetInf.PIN").value
        PINNoW = Sheet14.Range("RetInf.SpousePIN").value
        
        'Code for Self Start
        rangeName = Worksheets("H_Estate_Trust_Income").Range("EstateTrust.ListS").Address
        startRow = Worksheets("H_Estate_Trust_Income").Range(rangeName).row
        endRow = Worksheets("H_Estate_Trust_Income").Range(rangeName).Rows.Count + startRow - 1
        startColumn = Worksheets("H_Estate_Trust_Income").Range(rangeName).column
        lastColumn = startColumn + Worksheets("H_Estate_Trust_Income").Range(rangeName).Columns.Count - 1
                         
        For i = startRow To endRow
            Select Case Target.Address
                'PIN convert in uppar Case Start
                Case "$A$" & i
                     Target.value = StrConv(Target.Text, vbUpperCase)
                     If (Target.value <> "") Then
                     result = TestOtherPINNonId(Target.value)
                        If (result = False) Then
                            MsgBox "Please enter Non Individual (Estate/Trust/Settlement)'s PIN in proper format."
                            Target.value = ""
                            Target.Select
                        End If
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
                'PIN convert in uppar Case End
            End Select
        Next
        'Code for Self End
    
        'Code for Wife Start
        rangeName = Worksheets("H_Estate_Trust_Income").Range("EstateTrust.ListW").Address
        startRow = Worksheets("H_Estate_Trust_Income").Range(rangeName).row
        endRow = Worksheets("H_Estate_Trust_Income").Range(rangeName).Rows.Count + startRow - 1
        startColumn = Worksheets("H_Estate_Trust_Income").Range(rangeName).column
        lastColumn = startColumn + Worksheets("H_Estate_Trust_Income").Range(rangeName).Columns.Count - 1
                         
        For i = startRow To endRow
            Select Case Target.Address
                'PIN convert in uppar Case Start
                Case "$A$" & i
                    Target.value = StrConv(Target.Text, vbUpperCase)
                    If (Target.value <> "") Then
                     result = TestOtherPINNonId(Target.value)
                        If (result = False) Then
                            MsgBox "Please enter Non Individual (Estate/Trust/Settlement)'s PIN in proper format."
                            Target.value = ""
                            Target.Select
                        End If
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
                'PIN convert in uppar Case End
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
        'new code added for Bank Section Enable/Disable End
    End If
endSub:
    
    Worksheets("H_Estate_Trust_Income").Protect (Pwd)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub



