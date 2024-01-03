VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'workbook Open Start
Private Sub Workbook_Open()
    Pwd = "��������i@"
    Call Initialize
    checkOfficeVersion
    
    Call ToggleCutCopyAndPaste(True)
    Dim ActiveWorkbook As Workbook
    Application.EnableEvents = False
    Application.Calculation = xlManual
    Set ActiveWorkbook = ThisWorkbook
    ActiveWorkbook.Unprotect (Pwd)
    
    MakeSheetsVisible
    
    Sheet18.Visible = xlSheetHidden
    Sheet5.Visible = xlSheetHidden
    Sheet39.Visible = xlSheetHidden
    Sheet1.Visible = xlSheetHidden
'***********************************************************
'                       AMENDMENT CODE
'***********************************************************

    Application.ScreenUpdating = True
    Dim WS As Worksheet
    'Amendment Function Start
    On Error Resume Next
    Set WS = Worksheets("Amendment")
    On Error GoTo 0
    If Not WS Is Nothing Then
       If WS.Cells(1, 1).value <> "" Or WS.Cells(2, 1).value <> "" Then
            'unlock Spouse Fields initially
            Sheet14.toggleSpouseFields (False)
            Call fillDataInFields("Amendment")
            If (Worksheets("A_Basic_Info").Range("SecA.RtnYear").value < 2015) Then
                Call lockUnlock_cell_rng_without_clearing_contents("S_Previous_Years_Losses", "C4:H8", True)
            Else
                  Call lockUnlock_cell_rng_without_clearing_contents("S_Previous_Years_Losses", "C4:H8", False)
            End If
            If UCase(Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value) = "YES" Then
                    Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value = "Yes"
                     If (Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value <> "") Then
                         If (Worksheets("A_Basic_Info").Range("SecA.RtnYear").value < "2015") Then
                            Call lockUnlock_cell_rng_without_clearing_contents("S_Previous_Years_Losses", "C23:H27", True)
                         Else
                             Call lockUnlock_cell_rng_without_clearing_contents("S_Previous_Years_Losses", "C23:H27", False)
                         End If
                    End If
                    Call lockUnlock_cell_rng_without_clearing_contents("T_Tax_Computation", "TaxComp.WithHoldingListWTO", False)
                    Call lockUnlock_cell_rng_without_clearing_contents("T_Tax_Computation", "TaxComp.VehicleAdvTaxPaidListWTO", False)
            Call lockUnlock_cell_rng_without_clearing_contents("T_Tax_Computation", "TaxComp.MRIWTO", False)
                  
                ElseIf UCase(Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value) = "NO" _
                Or Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value = "" Then
    
                    Call lockUnlock_cell_rng_without_clearing_contents("T_Tax_Computation", "TaxComp.WithHoldingListWTO", True)
                    Call lockUnlock_cell_rng_without_clearing_contents("T_Tax_Computation", "TaxComp.VehicleAdvTaxPaidListWTO", True)
                    Call lockUnlock_cell_rng_without_clearing_contents("T_Tax_Computation", "TaxComp.MRIWTO", True)
                   
            End If
            Call lockUnlock_cell_rng_without_clearing_contents("O_WHT_Credits", "A3:A4", True)
            Call lockUnlock_cell_rng_without_clearing_contents("O_WHT_Credits", "B3:B4", True)
            Call lockUnlock_cell_rng_without_clearing_contents("O_WHT_Credits", "C3:C4", True)
            Call lockUnlock_cell_rng_without_clearing_contents("O_WHT_Credits", "D3:D4", True)
            Call lockUnlock_cell_rng_without_clearing_contents("O_WHT_Credits", "E3:E4", True)
            Call lockUnlock_cell_rng_without_clearing_contents("O_WHT_Credits", "A11:A12", True)
            Call lockUnlock_cell_rng_without_clearing_contents("O_WHT_Credits", "B11:B12", True)
            Call lockUnlock_cell_rng_without_clearing_contents("O_WHT_Credits", "C11:C12", True)
            Call lockUnlock_cell_rng_without_clearing_contents("O_WHT_Credits", "D11:D12", True)
            Call lockUnlock_cell_rng_without_clearing_contents("O_WHT_Credits", "E11:E12", True)
            Call lockUnlock_cell_rng_without_clearing_contents("P_Advance_Tax_Credits", "A3:A4", True)
            Call lockUnlock_cell_rng_without_clearing_contents("P_Advance_Tax_Credits", "B3:B4", True)
            Call lockUnlock_cell_rng_without_clearing_contents("P_Advance_Tax_Credits", "C3:C4", True)
            Call lockUnlock_cell_rng_without_clearing_contents("P_Advance_Tax_Credits", "D3:D4", True)
            Call lockUnlock_cell_rng_without_clearing_contents("P_Advance_Tax_Credits", "E3:E4", True)
            Call lockUnlock_cell_rng_without_clearing_contents("P_Advance_Tax_Credits", "F3:F4", True)
            Call lockUnlock_cell_rng_without_clearing_contents("P_Advance_Tax_Credits", "A11:A12", True)
            Call lockUnlock_cell_rng_without_clearing_contents("P_Advance_Tax_Credits", "B11:B12", True)
            Call lockUnlock_cell_rng_without_clearing_contents("P_Advance_Tax_Credits", "C11:C12", True)
            Call lockUnlock_cell_rng_without_clearing_contents("P_Advance_Tax_Credits", "D11:D12", True)
            Call lockUnlock_cell_rng_without_clearing_contents("P_Advance_Tax_Credits", "E11:E12", True)
            Call lockUnlock_cell_rng_without_clearing_contents("P_Advance_Tax_Credits", "F11:F12", True)
            Call lockUnlock_cell_rng_without_clearing_contents("T_Tax_Computation", "TaxComp.WithHoldingListSTO", False)
            Call lockUnlock_cell_rng_without_clearing_contents("T_Tax_Computation", "TaxComp.VehicleAdvTaxPaidListSTO", False)
            Call lockUnlock_cell_rng_without_clearing_contents("T_Tax_Computation", "TaxComp.MRI", False)
            Worksheets("Amendment").Visible = xlSheetHidden
            '29/12/2023
'            If (Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value = "No") Then
'                Worksheets("B_Profit_Loss_Account_Self").Visible = xlSheetHidden
'                Worksheets("C_Balance_Sheet").Visible = xlSheetHidden
'                Worksheets("D_Stock_Analysis").Visible = xlSheetHidden
'                Worksheets("E1_IDA_CA").Visible = xlSheetHidden
'                Worksheets("E2_CA_WTA_WDV").Visible = xlSheetHidden
'                Worksheets("E_Summary_of_Capital_Allowance").Visible = xlSheetHidden
'                Worksheets("N_Installment_Tax_Credits").Visible = xlSheetHidden
'                Worksheets("O_WHT_Credits").Visible = xlSheetHidden
'                Worksheets("S_Previous_Years_Losses").Visible = xlSheetHidden
'                Worksheets("T_Income_Computation_Self").Visible = xlSheetHidden
'            Else
'                Worksheets("B_Profit_Loss_Account_Self").Visible = xlSheetVisible
'                Worksheets("C_Balance_Sheet").Visible = xlSheetVisible
'                Worksheets("D_Stock_Analysis").Visible = xlSheetVisible
'                Worksheets("E1_IDA_CA").Visible = xlSheetVisible
'                Worksheets("E2_CA_WTA_WDV").Visible = xlSheetVisible
'                Worksheets("E_Summary_of_Capital_Allowance").Visible = xlSheetVisible
'                Worksheets("N_Installment_Tax_Credits").Visible = xlSheetVisible
'                Worksheets("O_WHT_Credits").Visible = xlSheetVisible
'                Worksheets("S_Previous_Years_Losses").Visible = xlSheetVisible
'                Worksheets("T_Income_Computation_Self").Visible = xlSheetVisible
'            End If
            '29/12/2023
            
            'new code added for Bank Section Enable/Disable Start
'            If Sheet14.Range("BankS").value <> "" Then
'                Call lockUnlock_cell_rng_without_clearing_contents("A_Basic_Info", "BankS", False)
'                Call lockUnlock_cell_rng_without_clearing_contents("A_Basic_Info", "BranchS", False)
'                Call lockUnlock_cell_rng_without_clearing_contents("A_Basic_Info", "BankDtl.CityS", False)
'                Call lockUnlock_cell_rng_without_clearing_contents("A_Basic_Info", "BankDtl.AccNameS", False)
'                Call lockUnlock_cell_rng_without_clearing_contents("A_Basic_Info", "BankDtl.AccNumberS", False)
'            Else
'                Call lockUnlock_cell_rng_without_clearing_contents("A_Basic_Info", "BankS", True)
'                Call lockUnlock_cell_rng_without_clearing_contents("A_Basic_Info", "BranchS", True)
'                Call lockUnlock_cell_rng_without_clearing_contents("A_Basic_Info", "BankDtl.CityS", True)
'                Call lockUnlock_cell_rng_without_clearing_contents("A_Basic_Info", "BankDtl.AccNameS", True)
'                Call lockUnlock_cell_rng_without_clearing_contents("A_Basic_Info", "BankDtl.AccNumberS", True)
'            End If
            
            'New code added by Atul Jha for moduler design of IT-R Template
            Call showModulerSectionsAmendment
            Dim endDate As String
            Dim mm As String
            Dim year As String
            endDate = CDate(Format(Sheet14.Range("RetInf.RetEndDate").value, "dd/mm/yyyy"))
            If (endDate <> "") Then
                mm = Format(CDate(Trim(endDate)), "MM")
                year = Format(CDate(Trim(endDate)), "yyyy")
            End If
            If UCase(Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value) = "YES" Then
                Worksheets("E2_CA_WTA_WDV").Unprotect (Pwd)
                If year >= 2020 Then
                        If (year = 2020 And mm >= 4) Or year > 2020 Then
                            Worksheets("E2_CA_WTA_WDV").Range("B16").value = "25"
                            Worksheets("E2_CA_WTA_WDV").Range("C16").value = "10"
                        Else
                            Worksheets("E2_CA_WTA_WDV").Range("B16").value = "37.5"
                            Worksheets("E2_CA_WTA_WDV").Range("C16").value = "30"
                        End If
                Else
                            Worksheets("E2_CA_WTA_WDV").Range("B16").value = "37.5"
                            Worksheets("E2_CA_WTA_WDV").Range("C16").value = "30"
                End If
                        Worksheets("E2_CA_WTA_WDV").Range("D16").value = "25"
                        Worksheets("E2_CA_WTA_WDV").Range("E16").value = "12.5"
                        Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_WDV", "B16:E16", True)
            End If
            Worksheets("E2_CA_WTA_WDV").Protect (Pwd)
        End If
    End If
        
    'new code added for Bank Section Enable/Disable End
    'Amendment Function End

    MsgBox "Please DO NOT Cut and Paste any values in the workbook."
    Application.EnableEvents = False
    Worksheets("Sheet1").Unprotect (Pwd)
    Worksheets("Sheet1").Range("A1:Z3").ClearContents
    Worksheets("Data").Visible = xlSheetHidden
    Worksheets("Errors").Unprotect (Pwd)
    Worksheets("Errors").Range("A2:E65535").ClearContents
    Worksheets("Errors").Protect (Pwd)
    Worksheets("Validations").Calculate
    Worksheets("ValidationList").Calculate

    Dim i As Integer
    For i = 1 To Worksheets.Count
        Worksheets(i).Protect (Pwd)
    Next
    
    Call resetRtnYrComboBox("Data", "A")
    
    'Sheet14.Activate
    'Sheet14.Cells(3, 2).Select
    ActiveWorkbook.Protect (Pwd)
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
'workbook Open End

Private Sub Workbook_Activate()
    Call ToggleCutCopyAndPaste(True)
    Worksheets("A_Basic_Info").Activate
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call ToggleCutCopyAndPaste(True)
    If Not isHideUnhidePerformed Then
        Call MacrosDisClose
    End If
End Sub

'Right Click Disable Start
Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
'    Cancel = True
'    MsgBox "Right clicking any where is disabled for  the workbook"
End Sub
'Right Click Disable End

'selection Disable Start
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)

    Call ToggleCutCopyAndPaste(False)
End Sub
'selection Disable End

'Dobuel Click Disable Start
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
   Cancel = True
   MsgBox "Double clicking any where is not allowed in the workbook"
End Sub
'Dobuel Click Disable End

'Added by Atul : Start
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    If isCalledOnClose And Success Then
        ThisWorkbook.Saved = False
        Call HideAllSheets
        isCalledOnClose = False
        ThisWorkbook.Saved = True
        ThisWorkbook.Save
'        Application.EnableEvents = False
'        ThisWorkbook.Close SaveChanges:=True
    End If
End Sub
'Added by Atul : End





