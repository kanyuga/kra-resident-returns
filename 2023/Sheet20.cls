VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    
    Worksheets("T_Tax_Computation").Unprotect (Pwd)
    Worksheets("T_Tax_Computation").Activate
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
   'Added by Ruth and Lawrence on 29/12/2020
   'If 2020 validate that the pension in sheet T_Tax_Computation equals the sum of Pension at F_Employment_Income
    If (Application.Range("SecA.RtnYear").value = "2020") Then
        If Target.Address = Range("DedDtl.PensionS").Address Then
            Call validateIsEqual(Target, (Application.Range("EmpIncome.ListSPENSJANMAR").value + Application.Range("EmpIncome.ListSPENSAPRDEC").value))
        End If
        If Target.Address = Range("DedDtl.PensionW").Address Then
            Call validateIsEqual(Target, (Application.Range("EmpIncome.ListWPENJANMAR").value + Application.Range("EmpIncome.ListWPENAPRDEC").value))
        End If
    End If

'    If ActiveSheet.Range("FinalTax.TaxRefundDueS").value <> "" Then
'        If ActiveSheet.Range("FinalTax.TaxRefundDueS").value < 0 Then
'            If Sheet14.Range("BankS").Locked = True Then
'                Call lockUnlock_cell_rng("A_Basic_Info", "BankS", False)
'                Call lockUnlock_cell_rng("A_Basic_Info", "BranchS", False)
'                Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.CityS", False)
'                Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNameS", False)
'                Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNumberS", False)
'            End If
'        Else
'            If Sheet14.Range("BankS").Locked = False Then
'                Call lockUnlock_cell_rng("A_Basic_Info", "BankS", True)
'                Call lockUnlock_cell_rng("A_Basic_Info", "BranchS", True)
'                Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.CityS", True)
'                Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNameS", True)
'                Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNumberS", True)
'            End If
'        End If
'    End If
    'Added by vaishali gohil
    If ActiveSheet.Range("TaxComp.PersonalReliefS").value <> "" Then
        Dim PersonalRelief As Double
        Dim RtnYear As Integer
        PersonalRelief = ActiveSheet.Range("TaxComp.PersonalReliefS").value
        Worksheets("A_Basic_Info").Unprotect Password:=Pwd
        Worksheets("A_Basic_Info").Activate
        RtnYear = ActiveSheet.Range("SecA.RtnYear").value
        Worksheets("A_Basic_Info").Protect (Pwd)
        Worksheets("T_Tax_Computation").Activate
        Call validatePRS(PersonalRelief, RtnYear)
        
    End If
    If ActiveSheet.Range("TaxComp.PersonalReliefW").value <> "" Then
        Dim PersonalReliefW As Double
        Dim RtnYearW As Integer
        PersonalReliefW = ActiveSheet.Range("TaxComp.PersonalReliefW").value
        Worksheets("A_Basic_Info").Unprotect Password:=Pwd
        Worksheets("A_Basic_Info").Activate
        RtnYearW = ActiveSheet.Range("SecA.RtnYear").value
        Worksheets("A_Basic_Info").Protect (Pwd)
        Worksheets("T_Tax_Computation").Activate
        Call validatePRW(PersonalReliefW, RtnYearW)
    End If
    If (ActiveSheet.Range("TaxComp.WithHoldingListSTO").value <> "") Then
        If ActiveSheet.Range("TaxComp.WithHoldingListSTO").value > 0 Then
            Worksheets("A_Basic_Info").Unprotect Password:=Pwd
            If (Worksheets("A_Basic_Info").Range("RetInf.DeclareSelfBusIncome").value = "No" Or Worksheets("A_Basic_Info").Range("RetInf.DeclareSelfBusIncome").value = "") Then
                Worksheets("A_Basic_Info").Protect (Pwd)
                Worksheets("T_Tax_Computation").Activate
                ActiveSheet.Range("TaxComp.WithHoldingListSTO").value = 0#
                MsgBox "Please select 'Do you have any income other than employment income?' as 'Yes' as you have Withholding Credits. Once you select 'Yes', you will be allowed to declare Withholding Credit."
                Worksheets("A_Basic_Info").Activate
                Worksheets("A_Basic_Info").Range("RetInf.DeclareSelfBusIncome").Select
            Else
                Worksheets("T_Tax_Computation").Activate
            End If
        End If
        
    End If
    'Worksheets("A_Basic_Info").Unprotect Password:=Pwd
    'If (Worksheets("A_Basic_Info").Range("RetInf.DeclareSelfBusIncome").value = "Yes" Or Worksheets("A_Basic_Info").Range("RetInf.DeclareSelfBusIncome").value = "") Then
           'Worksheets("A_Basic_Info").Protect (Pwd)
            If (Worksheets("T_Tax_Computation").Range("TaxComp.MRI").value <> "") Then
                If (Worksheets("T_Tax_Computation").Range("TaxComp.MRI").value > 0) Then
                    Worksheets("B_Profit_Loss_Account_Self").Unprotect Password:=Pwd
                    If (Worksheets("B_Profit_Loss_Account_Self").Range("PLA.GrossTurnoverBPRentalToS").value = 0 Or Worksheets("B_Profit_Loss_Account_Self").Range("PLA.GrossTurnoverBPRentalToS").value = "") Then
                        Worksheets("B_Profit_Loss_Account_Self").Protect (Pwd)
                        Worksheets("T_Tax_Computation").Activate
                        ActiveSheet.Range("TaxComp.MRI").value = 0#
                        If (Worksheets("A_Basic_Info").Range("RetInf.DeclareSelfBusIncome").value = "" Or Worksheets("A_Basic_Info").Range("RetInf.DeclareSelfBusIncome").value = "No") Then
                        MsgBox "Please select 'Do you have any income other than employment income?' as 'Yes' as you have Withholding Credits. Once you select 'Yes', you will be allowed to declare 'Gross Turnover of Business or Profession'  for Rental Income  in  B_Profit_Loss_Account_Self as you have Rental Income Withholding Credits. Once you declare the Rental Income , you will be allowed to declare the Rental Income Withholding Credit."
                        Worksheets("A_Basic_Info").Activate
                        Worksheets("A_Basic_Info").Range("RetInf.DeclareSelfBusIncome").Select
                        Else
                        MsgBox "Please declare 'Gross Turnover of Business or Profession'  for Rental Income  in  B_Profit_Loss_Account_Self as you have Rental Income Withholding Credits. Once you declare the Rental Income , you will be allowed to declare the Rental Income Withholding Credit."
                        Worksheets("B_Profit_Loss_Account_Self").Activate
                        Worksheets("B_Profit_Loss_Account_Self").Range("PLA.GrossTurnoverBPRentalToS").Select
                        End If
                    Else
                        Worksheets("T_Tax_Computation").Activate
                    End If
                End If
            End If
        'End If
        'Worksheets("A_Basic_Info").Unprotect Password:=Pwd
    'If (Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeBusIncome").value = "Yes" Or Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeBusIncome").value = "") Then
           'Worksheets("A_Basic_Info").Protect (Pwd)
            If (Worksheets("T_Tax_Computation").Range("TaxComp.MRIWTO").value <> "") Then
                If (Worksheets("T_Tax_Computation").Range("TaxComp.MRIWTO").value > 0) Then
                    Worksheets("B_Profit_Loss_Account_Wife").Unprotect Password:=Pwd
                    If (Worksheets("B_Profit_Loss_Account_Wife").Range("PLA.GrossTurnoverBPRentalToW").value = 0) Then
                        Worksheets("B_Profit_Loss_Account_Wife").Protect (Pwd)
                        Worksheets("T_Tax_Computation").Activate
                        ActiveSheet.Range("TaxComp.MRIWTO").value = 0#
                        If (Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeBusIncome").value = "" Or Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeBusIncome").value = "No") Then
                        MsgBox "Please select 'Does your wife have any income other than employment income? ' as 'Yes' as you have Withholding Credits. Once you select 'Yes', you will be allowed to declare 'Gross Turnover of Business or Profession'  for Rental Income  in  B_Profit_Loss_Account_Wife as you have Rental Income Withholding Credits. Once you declare the Rental Income , you will be allowed to declare the Rental Income Withholding Credit."
                        Worksheets("A_Basic_Info").Activate
                        Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeBusIncome").Select
                        Else
                        MsgBox "Please declare 'Gross Turnover of Business or Profession'  for Rental Income  in  B_Profit_Loss_Account_Wife as you have Rental Income Wittholding Credits. Once you declare the Rental Income , you will be allowed to declare the Rental Income Withholding Credit"
                        Worksheets("B_Profit_Loss_Account_Wife").Activate
                        Worksheets("B_Profit_Loss_Account_Wife").Range("PLA.GrossTurnoverBPRentalToW").Select
                        End If
                    Else
                        Worksheets("T_Tax_Computation").Activate
                    End If
                End If
            End If
        'End If
    'Added by vaishali gohil end
    
endSub:
    
    Worksheets("T_Tax_Computation").Protect (Pwd)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


