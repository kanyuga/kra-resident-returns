VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Add Row Functionality For Self Start
Sub addRow_HomeOwnershipSavingPlan_ListS()
    Worksheets("K_Home_Ownership_Saving_Plan").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("HomeOwnershipSavingPlan.ListS")
    Worksheets("K_Home_Ownership_Saving_Plan").Protect (Pwd)
End Sub
'Add Row Functionality For Self End

'Add Row Functionality For Wife Start
Sub addRow_HomeOwnershipSavingPlan_ListW()
    Worksheets("K_Home_Ownership_Saving_Plan").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("HomeOwnershipSavingPlan.ListW")
    Worksheets("K_Home_Ownership_Saving_Plan").Protect (Pwd)
End Sub
'Add Row Functionality For Wife End

Private Sub Worksheet_Change(ByVal Target As Range)
    
    Worksheets("K_Home_Ownership_Saving_Plan").Unprotect (Pwd)
    Worksheets("K_Home_Ownership_Saving_Plan").Activate
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Dim sumDeposit As Variant
    Dim sumDepositWife As Variant
    Dim period As String
    Dim yearInt As Integer
    Dim year As String
    Dim dateArray() As String
    
    'Added by Ruth and Lawrence on 29/12/2020
    'Validates if Input in the added 2020 cells is Numeric
    If Target.Address = Range("HomeOwnershipSavingPlan.ListSTOJANMAR").Address Then
        Call validateIsNumeric(Target)
    End If
    If Target.Address = Range("HomeOwnershipSavingPlan.ListSTOAPRDEC").Address Then
        Call validateIsNumeric(Target)
    End If
    If Target.Address = Range("HomeOwnershipSavingPlan.ListWTOJANMAR").Address Then
        Call validateIsNumeric(Target)
    End If
    If Target.Address = Range("HomeOwnershipSavingPlan.ListWTOAPRDEC").Address Then
        Call validateIsNumeric(Target)
    End If
    
    If (Target.Columns.Count = 1) Then
        PINNoS = Sheet14.Range("RetInf.PIN").value
        PINNoW = Sheet14.Range("RetInf.SpousePIN").value
        'Code for Self Start
        rangeName = Worksheets("K_Home_Ownership_Saving_Plan").Range("HomeOwnershipSavingPlan.ListS").Address
        startRow = Worksheets("K_Home_Ownership_Saving_Plan").Range(rangeName).row
        endRow = Worksheets("K_Home_Ownership_Saving_Plan").Range(rangeName).Rows.Count + startRow - 1
        
        sumDeposit = Worksheets("K_Home_Ownership_Saving_Plan").Range("HomeOwnershipSavingPlan.ListSTO").value
        sumDepositWife = Worksheets("K_Home_Ownership_Saving_Plan").Range("HomeOwnershipSavingPlan.ListWTO").value
        period = Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
        dateArray = Split(period, "/")
        year = dateArray(2)
        yearInt = CInt(year)

        For i = startRow To endRow
            Select Case Target.Address
                'PIN convert in uppar Case Start
                Case "$A$" & i
                    Target.value = StrConv(Target.Text, vbUpperCase)
                    If (Target.value <> "") Then
                    result = TestOtherPINNonId(Target.value)
                        If (result = False) Then
                            MsgBox "Please enter Non Individual Financial Institution's PIN in proper format."
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
                     If (Target.value <> "") Then
                        result = TestPIN(Target.value)
                        If (result = False) Then
                            MsgBox "Please enter PIN in proper format."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                    
                'PIN convert in uppar Case End
                
                'Validation for Account Number,Account Number should be Unique and convert in Uppar Case for Self Start
                Case "$C$" & i
                    If Target.value <> "" Then
                        Target.value = UCase(Target.value)
                        For Each r In Range("HomeOwnershipSavingPlan.ListS").Rows
                            For Each c In r.Cells
                                If c.row > endRow Then
                                    Exit For
                                End If
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                Else
                                    If Target.value <> "" And c.value = Target.value Then
                                        MsgBox "You have already entered details of Account Number  " & Target.value & " in Row " & c.row
                                        Target.value = ""
                                        Target.Select
                                        Exit For
                                    Exit For
                                    End If
                                End If
                            Next
                        Next
                    'Validation for Account Number,Account Number should be Unique and convert in Uppar Case for Self End
                    
                    'Validation for Account Number,Account Number should be Unique and convert in Uppar Case for Wife Start
                        For Each r In Range("HomeOwnershipSavingPlan.ListW").Rows
                            For Each c In r.Cells
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                Else
                                    If Target.value <> "" And c.value = Target.value Then
                                        MsgBox "You have already entered details of Account Number " & Target.value & " in Row " & c.row
                                        Target.value = ""
                                        Target.Select
                                        Exit For
                                    Exit For
                                    End If
                                End If
                            Next
                        Next
                    'Validation for Account Number,Account Number should be Unique and convert in Uppar Case for Wife End
                    End If
            End Select
        Next
        'Code for Self End

    'Code for Wife Start
    rangeName = Worksheets("K_Home_Ownership_Saving_Plan").Range("HomeOwnershipSavingPlan.ListW").Address
    startRow = Worksheets("K_Home_Ownership_Saving_Plan").Range(rangeName).row
    endRow = Worksheets("K_Home_Ownership_Saving_Plan").Range(rangeName).Rows.Count + startRow - 1

    For i = startRow To endRow
        Select Case Target.Address
            'PIN convert in uppar Case Start
            Case "$A$" & i
                Target.value = StrConv(Target.Text, vbUpperCase)
                If (Target.value <> "") Then
                    result = TestOtherPINNonId(Target.value)
                    If (result = False) Then
                        MsgBox "Please enter Non Individual Financial Institution's PIN in proper format."
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
                 If (Target.value <> "") Then
                    result = TestPIN(Target.value)
                    If (result = False) Then
                        MsgBox "Please enter PIN in proper format."
                        Target.value = ""
                        Target.Select
                    End If
                End If
                
                
             'PIN convert in uppar Case End
            
            'Validation for Account Number,Account Number should be Unique and convert in Uppar Case for Self Start
            Case "$C$" & i
                If Target.value <> "" Then
                    Target.value = UCase(Target.value)
                    For Each r In Range("HomeOwnershipSavingPlan.ListS").Rows
                        For Each c In r.Cells
                                If c.row > endRow Then
                                    Exit For
                                End If
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                
                                Else
                                    If Target.value <> "" And c.value = Target.value Then
                                        MsgBox "You have already entered details of Account Number  " & Target.value & " in Row " & c.row
                                        Target.value = ""
                                        Target.Select
                                        Exit For
                                    Exit For
                                    End If
                                End If
                        Next
                    Next
                'Validation for Account Number,Account Number should be Unique and convert in Uppar Case for Self End

                'Validation for Account Number,Account Number should be Unique and convert in Uppar Case for Wife Start
                    For Each r In Range("HomeOwnershipSavingPlan.ListW").Rows
                        For Each c In r.Cells
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                
                                Else
                                    If Target.value <> "" And c.value = Target.value Then
                                        MsgBox "You have already entered details of Account Number " & Target.value & " in Row " & c.row
                                        Target.value = ""
                                        Target.Select
                                        Exit For
                                    Exit For
                                    End If
                                End If
                        Next
                    Next
                'Validation for Account Number,Account Number should be Unique and convert in Uppar Case for Wife End
                End If
        End Select
    Next
    'Code for Wife End
    
     Worksheets("T_Tax_Computation").Unprotect (Pwd)
    If (yearInt < 2019) Then
        If (sumDeposit > 48000) Then
            Worksheets("T_Tax_Computation").Range("TaxComp.HomeOwnershipSavingPlanListSTO").value = 48000
        Else
            Worksheets("T_Tax_Computation").Range("TaxComp.HomeOwnershipSavingPlanListSTO").value = sumDeposit
        End If
        If (sumDepositWife > 48000) Then
            Worksheets("T_Tax_Computation").Range("TaxComp.HomeOwnershipSavingPlanListWTO").value = 48000
        Else
            Worksheets("T_Tax_Computation").Range("TaxComp.HomeOwnershipSavingPlanListWTO").value = sumDepositWife
        End If
    End If
    If (yearInt = 2019) Then
        If (sumDeposit > 72000) Then
            Worksheets("T_Tax_Computation").Range("TaxComp.HomeOwnershipSavingPlanListSTO").value = 72000
        Else
            Worksheets("T_Tax_Computation").Range("TaxComp.HomeOwnershipSavingPlanListSTO").value = sumDeposit
        End If
        If (sumDepositWife > 72000) Then
            Worksheets("T_Tax_Computation").Range("TaxComp.HomeOwnershipSavingPlanListWTO").value = 72000
        Else
            Worksheets("T_Tax_Computation").Range("TaxComp.HomeOwnershipSavingPlanListWTO").value = sumDepositWife
        End If
    End If
    If (yearInt >= 2020) Then
        If (sumDeposit > 96000) Then
            Worksheets("T_Tax_Computation").Range("TaxComp.HomeOwnershipSavingPlanListSTO").value = 96000
        Else
            Worksheets("T_Tax_Computation").Range("TaxComp.HomeOwnershipSavingPlanListSTO").value = sumDeposit
        End If
        If (sumDepositWife > 96000) Then
            Worksheets("T_Tax_Computation").Range("TaxComp.HomeOwnershipSavingPlanListWTO").value = 96000
        Else
            Worksheets("T_Tax_Computation").Range("TaxComp.HomeOwnershipSavingPlanListWTO").value = sumDepositWife
        End If
    End If
    Worksheets("T_Tax_Computation").Protect (Pwd)
    
End If

endSub:
    Worksheets("K_Home_Ownership_Saving_Plan").Protect (Pwd)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
