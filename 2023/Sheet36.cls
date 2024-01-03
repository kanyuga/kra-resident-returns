VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet36"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub addRow_WithHolding_ListS()
    Worksheets("O_WHT_Credits").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("WithHolding.ListS")
    Worksheets("O_WHT_Credits").Protect (Pwd)
End Sub
Sub addRow_WithHolding_ListW()
Worksheets("O_WHT_Credits").Unprotect (Pwd)
Call InsertRowsAndFillFormulas("WithHolding.ListW")
Worksheets("O_WHT_Credits").Protect (Pwd)
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    
    Worksheets("O_WHT_Credits").Unprotect (Pwd)
    Worksheets("O_WHT_Credits").Activate
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    If (Target.Columns.Count = 1) Then
    
        PINNoS = Sheet14.Range("RetInf.PIN").value
        PINNoW = Sheet14.Range("RetInf.SpousePIN").value
        
        rangeName = Worksheets("O_WHT_Credits").Range("WithHolding.ListS").Address
        startRow = Worksheets("O_WHT_Credits").Range(rangeName).row
        endRow = Worksheets("O_WHT_Credits").Range(rangeName).Rows.Count + startRow - 1
        
        startColumn = Worksheets("O_WHT_Credits").Range(rangeName).column
        lastColumn = startColumn + Worksheets("O_WHT_Credits").Range(rangeName).Columns.Count - 1
                     
        For i = startRow To endRow
            Select Case Target.Address
                '=============== convert pin to upper case ===========================
                Case "$A$" & i
                    Target.value = StrConv(Target.Text, vbUpperCase)
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
                            MsgBox "Please enter PIN of Withholder in proper format."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    If (Target.value <> "") Then
                        cert = Application.Range("$C$" & i).value
                        For Each r In Range("WithHolding.ListS").Rows
                            For Each c In r.Cells
                                rowNum = c.row
                                certToCompare = Application.Range("$C$" & rowNum).value
                                If c.row > endRow Then
                                    Exit For
                                End If
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                Else
                                    If Target.value <> "" And c.value = Target.value And cert = certToCompare Then
                                        MsgBox "The PIN of Withholder " & Target.value & " has already been entered for Withholding Certificate Number " & cert & " in Row " & c.row
                                        Target.value = ""
                                        Target.Select
                                        Exit For
                                    End If
                                End If
                            Next
                        Next
                      End If
                Case "$D$" & i
                    If (Target.value <> "") Then
                       If TestDate(Target.value) = False Then
                        MsgBox "Please enter date value in dd/mm/yyyy format."
                        Target.value = ""
                        Target.Select
                       Else
                        If Target.value <> "" Then
                         If checkFutureDate(Target.value) = False Then
                                MsgBox "Withholding Transaction Date should not be future date."
                                Target.value = ""
                                Target.Select
                            End If
                        End If
                       End If
                    End If
                Case "$C$" & i
                    If Target.value <> UCase(Target.value) Then
                        Target.value = UCase(Target.value)
                    End If
                    If (Target.value <> "") Then
                        PIN = Application.Range("$A$" & i).value
                        For Each r In Range("WithHolding.ListS").Rows
                            For Each c In r.Cells
                                rowNum = c.row
                                pinToCompare = Application.Range("$A$" & rowNum).value
                                If c.row > endRow Then
                                    Exit For
                                End If
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                Else
                                    If Target.value <> "" And c.value = Target.value Then 'And pin = pinToCompare Then
                                        MsgBox "The Withholding Certificate Number " & Target.value & " has already been entered in Row " & c.row
                                        Target.value = ""
                                        Target.Select
                                        Exit For
                                    End If
                                End If
                            Next
                        Next
                    End If
            End Select
        Next

    rangeName = Worksheets("O_WHT_Credits").Range("WithHolding.ListW").Address
    startRow = Range(rangeName).row
    endRow = Range(rangeName).Rows.Count + startRow - 1
    
    startColumn = Range(rangeName).column
    lastColumn = startColumn + Range(rangeName).Columns.Count - 1
                         
    For i = startRow To endRow
        Select Case Target.Address
            '=============== convert pin to upper case ===========================
            Case "$A$" & i
                 Target.value = StrConv(Target.Text, vbUpperCase)
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
                            MsgBox "Please enter PIN of Withholder in proper format."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    If (Target.value <> "") Then
                        cert = Application.Range("$C$" & i).value
                        For Each r In Range("WithHolding.ListW").Rows
                            For Each c In r.Cells
                                rowNum = c.row
                                certToCompare = Application.Range("$C$" & rowNum).value
                                If c.row > endRow Then
                                    Exit For
                                End If
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                Else
                                    If Target.value <> "" And c.value = Target.value And cert = certToCompare Then
                                        MsgBox "The PIN of Withholder " & Target.value & " has already been entered for Withholding Certificate Number " & cert & " in Row " & c.row
                                        Target.value = ""
                                        Target.Select
                                        Exit For
                                    End If
                                End If
                            Next
                        Next
                    End If
            Case "$D$" & i
                If (Target.value <> "") Then
                   If TestDate(Target.value) = False Then
                    MsgBox "Please enter date value in dd/mm/yyyy format."
                    Target.value = ""
                    Target.Select
                   Else
                    If Target.value <> "" Then
                        If checkFutureDate(Target.value) = False Then
                            MsgBox "Withholding Transaction Date should not be future date."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                  End If
                End If
                
            Case "$C$" & i
                If Target.value <> UCase(Target.value) Then
                    Target.value = UCase(Target.value)
                End If
                If (Target.value <> "") Then
                    PIN = Application.Range("$A$" & i).value
                    For Each r In Range("WithHolding.ListW").Rows
                        For Each c In r.Cells
                            rowNum = c.row
                            pinToCompare = Application.Range("$A$" & rowNum).value
                            If c.row > endRow Then
                                Exit For
                            End If
                            If c.Address = Target.Address Then 'Skip comparison with self value
                            Else
                                If Target.value <> "" And c.value = Target.value And PIN = pinToCompare Then
                                    MsgBox "The Withholding Certificate Number " & Target.value & " has already been entered for PIN of Withholder " & PIN & " in Row " & c.row
                                    Target.value = ""
                                    Target.Select
                                    Exit For
                                End If
                            End If
                        Next
                    Next
                End If
        End Select
    Next
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
                
                Worksheets("O_WHT_Credits").Activate
        'new code added for Bank Section Enable/Disable End
    End If
ensub:
    
    Worksheets("O_WHT_Credits").Protect (Pwd)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

