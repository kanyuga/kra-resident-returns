VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Add Row Functionality For Self Start
Sub addRow_MortgageIntDtls_ListS()
Worksheets("J_Computation_of_Mortgage").Unprotect (Pwd)
Call InsertRowsAndFillFormulas("MortgageIntDtls.ListS")
Worksheets("J_Computation_of_Mortgage").Protect (Pwd)
End Sub
'Add Row Functionality For Self End

'Add Row Functionality For Self Start
Sub addRow_MortgageIntDtls_ListW()
Worksheets("J_Computation_of_Mortgage").Unprotect (Pwd)
Call InsertRowsAndFillFormulas("MortgageIntDtls.ListW")
Worksheets("J_Computation_of_Mortgage").Protect (Pwd)
End Sub
'Add Row Functionality For Self End

Private Sub Worksheet_Change(ByVal Target As Range)
    
    Worksheets("J_Computation_of_Mortgage").Unprotect (Pwd)
    Worksheets("J_Computation_of_Mortgage").Activate
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    'Added by Ruth and Lawrence on 29/12/2020
    'Validates if Input in the added 2020 cells is Numeric
    If Target.Address = Range("MortgageIntDtls.ListSTOJANMAR").Address Then
        Call validateIsNumeric(Target)
    End If
    If Target.Address = Range("MortgageIntDtls.ListSTOAPRDEC").Address Then
        Call validateIsNumeric(Target)
    End If
    If Target.Address = Range("MortgageIntDtls.ListWTOJANMAR").Address Then
        Call validateIsNumeric(Target)
    End If
    If Target.Address = Range("MortgageIntDtls.ListWTOAPRDEC").Address Then
        Call validateIsNumeric(Target)
    End If



    If (Target.Columns.Count = 1) Then
        PINNoS = Sheet14.Range("RetInf.PIN").value
        PINNoW = Sheet14.Range("RetInf.SpousePIN").value
    
        'Code for Self Start
        rangeName = Worksheets("J_Computation_of_Mortgage").Range("MortgageIntDtls.ListS").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1

        For i = startRow To endRow
            Select Case Target.Address
            'PIN convert in uppar Case Start
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
                        result = TestOtherPINNonId(Target.value)
                        If (result = False) Then
                            MsgBox "Please enter valid PIN of Lender."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                'PIN convert in uppar Case End
            
                'Validation for Mortgage Account,Mortgage Account number should be unique and convert in uppar Case for Self Start
                Case "$C$" & i
                    If Target.value <> "" Then
                        Target.value = UCase(Target.value)
                        For Each r In Range("MortgageIntDtls.ListS").Rows
                            For Each c In r.Cells
                                If c.row > endRow Then
                                    Exit For
                                End If
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                Else
                                    If Target.value <> "" And c.value = Target.value Then
                                        MsgBox "You have already entered details of Mortgage Account Number  " & Target.value & " in Row " & c.row
                                        Target.value = ""
                                        Target.Select
                                        Exit For
                                    Exit For
                                    End If
                                End If
                            Next
                        Next
                        'Validation for Mortgage Account,Mortgage Account number should be unique and convert in uppar Case for Self End
                        
                        'Validation for Mortgage Account,Mortgage Account number should be unique and convert in uppar Case for Wife Start
                        For Each r In Range("MortgageIntDtls.ListW").Rows
                            For Each c In r.Cells
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                Else
                                    If Target.value <> "" And c.value = Target.value Then
                                        MsgBox "You have already entered details of Mortgage Account Number " & Target.value & " in Row " & c.row
                                        Target.value = ""
                                        Target.Select
                                        Exit For
                                    Exit For
                                    End If
                                End If
                            Next
                        Next
                      'Validation for Mortgage Account,Mortgage Account number should be unique and convert in uppar Case for Wife End
                    End If
            End Select
        Next
        'Code for Self End
    
        'Code for Wife Start
        rangeName = Worksheets("J_Computation_of_Mortgage").Range("MortgageIntDtls.ListW").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1
    
        For i = startRow To endRow
            Select Case Target.Address
                'PIN convert in uppar Case Start
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
                        result = TestOtherPINNonId(Target.value)
                        If (result = False) Then
                            MsgBox "Please enter valid PIN of Lender."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                'PIN convert in uppar Case End
                
                'Validation for Mortgage Account,Mortgage Account number should be unique and convert in uppar Case for Self Start
                Case "$C$" & i
                    If Target.value <> "" Then
                        Target.value = UCase(Target.value)
                        For Each r In Range("MortgageIntDtls.ListS").Rows
                            For Each c In r.Cells
                                    If c.row > endRow Then
                                        Exit For
                                    End If
                                    If c.Address = Target.Address Then 'Skip comparison with self value
                                    Else
                                        If Target.value <> "" And c.value = Target.value Then
                                            MsgBox "You have already entered details of Mortgage Account Number  " & Target.value & " in Row " & c.row
                                            Target.value = ""
                                            Target.Select
                                            Exit For
                                        Exit For
                                        End If
                                    End If
                            Next
                        Next
                    'Validation for Mortgage Account,Mortgage Account number should be unique and convert in uppar Case for Self End
                    
                    'Validation for Mortgage Account,Mortgage Account number should be unique and convert in uppar Case for Wife Start
                        For Each r In Range("MortgageIntDtls.ListW").Rows
                            For Each c In r.Cells
                                    If c.Address = Target.Address Then 'Skip comparison with self value
                                    Else
                                        If Target.value <> "" And c.value = Target.value Then
                                            MsgBox "You have already entered details of Mortgage Account Number " & Target.value & " in Row " & c.row
                                            Target.value = ""
                                            Target.Select
                                            Exit For
                                        Exit For
                                        End If
                                    End If
                            Next
                        Next
                    'Validation for Mortgage Account,Mortgage Account number should be unique and convert in uppar Case for Wife End
                    End If
            End Select
        Next
        'Code for Wife End
    End If
    
    Worksheets("J_Computation_of_Mortgage").Protect (Pwd)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
