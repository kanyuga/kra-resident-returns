VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Add Row Functionality For Self Start
Sub addRow_Name_Expenses_ListS()
    Worksheets("T_Income_Computation_Self").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("TaxComp.OthExpListS")
    Worksheets("T_Income_Computation_Self").Protect (Pwd)
End Sub
'Add Row Functionality For Self End

'Add Row Functionality For Self Start
Sub addRow_Name_Deductions_ListS()
    Worksheets("T_Income_Computation_Self").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("TaxComp.OthDedListS")
    Worksheets("T_Income_Computation_Self").Protect (Pwd)
End Sub
'Add Row Functionality For Self End

Private Sub Worksheet_Change(ByVal Target As Range)

    
    Worksheets("T_Income_Computation_Self").Unprotect (Pwd)
    Worksheets("T_Income_Computation_Self").Activate
    Application.EnableEvents = False
    Application.ScreenUpdating = False
 
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
 
endSub:
    
    Worksheets("T_Income_Computation_Self").Protect (Pwd)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
