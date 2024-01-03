VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Add Row Functionality Start
Sub addRow_Oth_Exp_List()
    Worksheets("B_Profit_Loss_Account_Self").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("PLA.OtherExpensesListS")
    Worksheets("B_Profit_Loss_Account_Self").Protect (Pwd)
End Sub
'Add Row Functionality End

'Add Row Functionality Start
Sub addRow_Oth_Income_List()
    Worksheets("B_Profit_Loss_Account_Self").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("PLA.OtherIncomeListS")
    Worksheets("B_Profit_Loss_Account_Self").Protect (Pwd)
End Sub
'Add Row Functionality End
'Added by vaishali gohil
Private Sub Worksheet_Change(ByVal Target As Range)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    'Worksheets("A_Basic_Info").Unprotect Password:=Pwd
    'Worksheets("A_Basic_Info").Activate
    'If (ActiveSheet.Range("RetInf.DeclareSelfBusIncome").value = "Yes") Then
           'Worksheets("A_Basic_Info").Protect (Pwd)
           Worksheets("B_Profit_Loss_Account_Self").Activate
           Worksheets("B_Profit_Loss_Account_Self").Unprotect (Pwd)
           If (ActiveSheet.Range("PLA.GrossTurnoverBPRentalToS").value = "" Or ActiveSheet.Range("PLA.GrossTurnoverBPRentalToS").value = 0) Then
                Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
                Worksheets("T_Tax_Computation").Activate
                If ActiveSheet.Range("TaxComp.MRI").value > 0 Then
                    ActiveSheet.Range("TaxComp.MRI").value = 0#
                End If
                Worksheets("T_Tax_Computation").Protect (Pwd)
           End If
           Worksheets("B_Profit_Loss_Account_Self").Protect (Pwd)
    'End If
endSub:
    Worksheets("B_Profit_Loss_Account_Self").Activate
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
'Added by vaishali gohil end

