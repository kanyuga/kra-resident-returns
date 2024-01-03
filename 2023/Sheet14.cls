VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Add Row Functionality For Part 4 Self Start

Sub addRow_RentPaid_RentListS()
    Worksheets("A_Basic_Info").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("RentPaid.RentListS")
    Worksheets("A_Basic_Info").Protect (Pwd)
End Sub
'Add Row Functionality For Part 4 Self End

'Add Row Functionality For Part 4 Wife Start
Sub addRow_RentPaid_RentListW()
    Worksheets("A_Basic_Info").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("RentPaid.RentListW")
    Worksheets("A_Basic_Info").Protect (Pwd)
End Sub
'Add Row Functionality For Part 4 Wife End

'Add Row Functionality For Part 5 Self Start
Sub addRow_RentalIncome_ListS()
    Worksheets("A_Basic_Info").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("RentalIncome.ListS")
    Worksheets("A_Basic_Info").Protect (Pwd)
End Sub
'Add Row Functionality For Part 5 Self End

'Add Row Functionality For Part 5 Wife Start
Sub addRow_RentalIncome_ListW()
    Worksheets("A_Basic_Info").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("RentalIncome.ListW")
    Worksheets("A_Basic_Info").Protect (Pwd)
End Sub
'Add Row Functionality For Part 5 Wife End

'Add Row Functionality For Part 6 Self Start
Sub addRow_ExemptCerti_ListS()
    Worksheets("A_Basic_Info").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("ExemptCerti.ListS")
    Worksheets("A_Basic_Info").Protect (Pwd)
End Sub
'Add Row Functionality For Part 6 Self End

'Add Row Functionality For Part 6 Wife Start
Sub addRow_ExemptCerti_ListW()
    Worksheets("A_Basic_Info").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("ExemptCerti.ListW")
    Worksheets("A_Basic_Info").Protect (Pwd)
End Sub
'Add Row Functionality For Part 6 Wife End

Private Sub Worksheet_Change(ByVal Target As Range)

    Worksheets("A_Basic_Info").Unprotect Password:=Pwd
    Worksheets("A_Basic_Info").Activate
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim lockCell As Boolean
    Dim tTDWValue As String
    Dim tTDWNXValue As String
    Dim dateOfDateTo As String
    
    'Added by lawrence 22122020 for the covid tax rates affecting IT1
    'Declare an array of sheets that have the rows to be hidden or unhidden
    Dim workSheetArray() As Variant
    'Initilize the array with the data of the sheets.
    'This variable at this sub only has the sheets with wife income
    workSheetArray = Array("F_Employment_Income", "J_Computation_of_Mortgage", "K_Home_Ownership_Saving_Plan")
    
    If (Target.Columns.Count = 1) Then
        PINNoS = Worksheets("A_Basic_Info").Range("RetInf.PIN").value
        PINNoW = Worksheets("A_Basic_Info").Range("RetInf.SpousePIN").value
        
        If (Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value <> "") Then
            If (TestDate(Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value) = True) Then
                startDate = CDate(Format(Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value, "dd/mm/yyyy"))
            End If
        End If
    
        If (Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value <> "") Then
            If (TestDate(Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value) = True) Then
                endDate = CDate(Format(Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value, "dd/mm/yyyy"))
            End If
        End If
    
        
          If (Worksheets("A_Basic_Info").Range("RetInf.RtnPrdToActStart").value <> "") Then
            If (TestDate(Worksheets("A_Basic_Info").Range("RetInf.RtnPrdToActStart").value) = True) Then
                RPToActStart = CDate(Format(Worksheets("A_Basic_Info").Range("RetInf.RtnPrdToActStart").value, "dd/mm/yyyy"))
            End If
        End If
        
        If (Worksheets("A_Basic_Info").Range("RetInf.RtnPrdToAct").value <> "") Then
            If (TestDate(Worksheets("A_Basic_Info").Range("RetInf.RtnPrdToAct").value) = True) Then
                RPToAct = CDate(Format(Worksheets("A_Basic_Info").Range("RetInf.RtnPrdToAct").value, "dd/mm/yyyy"))
            End If
        End If
        If (Worksheets("A_Basic_Info").Range("RetInf.auditStartDate").value <> "") Then
            If (TestDate(Worksheets("A_Basic_Info").Range("RetInf.auditStartDate").value) = True) Then
                auditStartDate = CDate(Format(Worksheets("A_Basic_Info").Range("RetInf.auditStartDate").value, "dd/mm/yyyy"))
            End If
        End If
        If (Worksheets("A_Basic_Info").Range("RetInf.dateforAuditCertificate").value <> "") Then
            If (TestDate(Worksheets("A_Basic_Info").Range("RetInf.dateforAuditCertificate").value) = True) Then
                auditEndDate = CDate(Format(Worksheets("A_Basic_Info").Range("RetInf.dateforAuditCertificate").value, "dd/mm/yyyy"))
            End If
        End If
        
        If Target.Address = "$B$5" Then
            If (Target.value <> "") Then
            DsysDate = Date
                If TestDate(Target.value) = False Then
                    MsgBox "Please enter date value in dd/mm/yyyy format."
                    Target.value = ""
                    Target.Select
                Else
                    If (TestDateLimit(Target.value) = False) Then
                        MsgBox "Please enter date in Return Period From with year on or after 2001."
                        Target.value = ""
                        Target.Select
                    End If
                    If (DateValue(startDate) > DateValue(DsysDate)) Then
                        MsgBox "You cannot file the return for the period whose date of Return Period From and / or Return Period To is a future date."
                        Target.value = ""
                        Target.Select
                    End If
                    
                    Dim rtnPrdFrom As String
                    Dim rtnPrdTo As String
                    Dim rtnDepositDate As String
                    Dim mon As String
                    Dim mm As String
                    Dim mmU As String
                    Dim year As String
                    Dim Depstyear As String
                    Dim DepstDate As String
            
                    
                    rtnPrdFrom = Target.value
                    dd = Format(CDate(Trim(startDate)), "DD")
                    mm = Format(CDate(Trim(startDate)), "MM")
                    mmU = Format(CDate(Trim(startDate)), "MM")
                    year = DatePart("yyyy", startDate)
                    Dim rtnTo  As String
                    rtnTo = mm & "/" & year
                    Depstyear = mm & "/" & year + 1
                    DepstDate = dd & "/" & mm & "/" & year - 1
                    
                    Worksheets("A_Basic_Info").Range("RetInf.RtnPrdToAct").value = "31/12/" & year
                    Worksheets("A_Basic_Info").Range("SecA.RtnYear").value = year
                    Worksheets("A_Basic_Info").Range("RetInf.DepositStartDate").value = DepstDate
                    
                End If
    
                If (endDate <> "") Then
                    Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value = ""
                    Worksheets("A_Basic_Info").Range("RetInf.rtnPrdToActStart").value = ""
                    
                End If
                
                'Added by Palak
              Call getyear(year)
            'End If
               
                'Added by vaishali gohil
               If (year < 2015) Then
                Call lockUnlock_cell_rng("S_Previous_Years_Losses", "C4:H8", True)
                Else
                Call lockUnlock_cell_rng("S_Previous_Years_Losses", "C4:H8", False)
            End If
                Call getyear2(year)
                'Added by vaishali gohil end
            Else
                    Worksheets("A_Basic_Info").Range("RetInf.RtnPrdToAct").value = ""
                    Worksheets("A_Basic_Info").Range("SecA.RtnYear").value = ""
                    Worksheets("A_Basic_Info").Range("RetInf.DepositStartDate").value = ""
                    Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value = ""
            End If
            'Added by on 03/02/2023 ticket 64901
            If (year <> "") Then
                Worksheets("L_Computation_of_Insu_Relief").Unprotect (Pwd)
                If (year > 2021) Then
                    Worksheets("L_Computation_of_Insu_Relief").Range("C2:C2").value = "Type of policy " & vbNewLine & "(Life /Education /Health /Nhif)"
                    Worksheets("L_Computation_of_Insu_Relief").Range("D2:D2").value = "Insurance Policy Number/NHIF Membership Number"
                Else
                    Worksheets("L_Computation_of_Insu_Relief").Range("C2:C2").value = "Type of policy (Life/Education/Health)"
                    Worksheets("L_Computation_of_Insu_Relief").Range("D2:D2").value = "Insurance Policy Number"
                End If
                Worksheets("L_Computation_of_Insu_Relief").Protect (Pwd)
            End If
            'End 03/02/2023
            
            'Added on 17/07/2023 ticket 69358
            If (Target.value <> "") Then
                If (year >= 2023) Then
                    Worksheets("A_Basic_Info").Unprotect (Pwd)
                    If (Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value = "Yes") Then
                        MsgBox "You cannot declare wife's income after 01/07/2023."
                        Worksheets("A_Basic_Info").Activate
                        Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").Select
                        Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value = "No"
                        Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").Locked = True
                        Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").Interior.ColorIndex = 48
                        Call b17("$B$17", workSheetArray, year, mm)
                    Else
                        Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").Locked = True
                        Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").Interior.ColorIndex = 48
                    End If
                    Worksheets("A_Basic_Info").Protect (Pwd)
                Else
                    Worksheets("A_Basic_Info").Unprotect (Pwd)
                    Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").Locked = False
                    Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").Interior.ColorIndex = 2
                    Worksheets("A_Basic_Info").Protect (Pwd)
                End If
            End If
            
        '        Dim sysdt As Date
        '        sysdt = Format(Now(), "dd/mm/yyyy")
        '        If (sysdt >= CDate(Format("01/07/2023", "dd/mm/yyyy"))) Then
        '            Worksheets("A_Basic_Info").Unprotect (Pwd)
        '            Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").Locked = True
        '            Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").Interior.ColorIndex = 48
        '            Worksheets("A_Basic_Info").Protect (Pwd)
        '
        '        Else
        '            Worksheets("A_Basic_Info").Unprotect (Pwd)
        '            Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").Locked = False
        '            Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").Interior.ColorIndex = 2
        '            Worksheets("A_Basic_Info").Protect (Pwd)
        '        End If
        'MsgBox "You cannot declare wife's income after 01/07/2023."

        'End on 17/07/2023 ticket 69358
        
            Target.Select
        End If
        
        If Target.Address = "$B$6" Then
            If (Target.value <> "") Then
            DsysDate = Date
                
                
                If TestDate(Target.value) = False Then
                    MsgBox "Please enter date value in dd/mm/yyyy format."
                    Target.value = ""
                    Target.Select
                Else
                    If (startDate = "") Then
                        MsgBox "Please enter Return Period From first."
                        Target.value = ""
                        Target.Select
                    End If
                    rtnPrdFrom = Target.value
                    mm = Format(CDate(Trim(endDate)), "MM")
                    dd = Format(CDate(Trim(endDate)), "dd")
                    year = Format(CDate(Trim(endDate)), "yyyy")
                    yearU = Format(CDate(Trim(endDate)), "yyyy")
                    mmU = Format(CDate(Trim(endDate)), "MM")
                     Call enabledisable(mm, year)
                    'by Palak for Enh-6 SR2 starts
                    If UCase(Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value) = "YES" Then
                        toggleSpouseFields (True)
                    End If
                    Call toggleCellsInE1Sheet(year, mm)
                    rtnTo = mm & "/" & year
                    If (year < 2015) Then
                        Call lockUnlock_cell_rng("S_Previous_Years_Losses", "C4:H8", True)
                    Else
                          Call lockUnlock_cell_rng("S_Previous_Years_Losses", "C4:H8", False)
                    End If
                    If (Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value = "Yes") Then
                 If (year < 2015) Then
                    Call lockUnlock_cell_rng("S_Previous_Years_Losses", "C23:H27", True)
                 Else
                     Call lockUnlock_cell_rng("S_Previous_Years_Losses", "C23:h27", False)
                 End If
            End If
             Worksheets("A_Basic_Info").Activate
             Worksheets("A_Basic_Info").Unprotect Password:=Pwd
                            Worksheets("A_Basic_Info").Range("RetInf.RtnPrdToActStart").value = "01/01/" & year
                    
                     'temp code added Start
                    mon = 0
                    
                    If (mm + 1) >= 8 Then
                        For i = 8 To (mm + 1)
                            mon = mon + 1
                        Next i
                        year = year + 1
                    Else
                        mon = (mm + 6)
                    End If
                    If (Len(mon) = 1) Then
                       mon = "0" & mon
                    End If
                    rtnTo = mon & "/" & year
                    
                    If IsDate(dd & "/" & rtnTo) Then
                            Worksheets("A_Basic_Info").Range("RetInf.dateforAuditCertificate").value = dd & "/" & rtnTo
                    Else
                        If (mon = "02") Then
                            Worksheets("A_Basic_Info").Range("RetInf.dateforAuditCertificate").value = 28 & "/" & rtnTo
                        ElseIf (mon = "04") Then
                            Worksheets("A_Basic_Info").Range("RetInf.dateforAuditCertificate").value = 30 & "/" & rtnTo
                        ElseIf (mon = "06") Then
                            Worksheets("A_Basic_Info").Range("RetInf.dateforAuditCertificate").value = 30 & "/" & rtnTo
                        ElseIf (mon = "09") Then
                            Worksheets("A_Basic_Info").Range("RetInf.dateforAuditCertificate").value = 30 & "/" & rtnTo
                        ElseIf (mon = "11") Then
                            Worksheets("A_Basic_Info").Range("RetInf.dateforAuditCertificate").value = 30 & "/" & rtnTo
                        End If
                        
                    End If
                    
                    'new code added for Audit Date Start
                    Dim firstDate As String, secondDate As String
                    Dim auditStDt As String
                    If (Target.value <> "") Then
                        firstDate = DateValue(Target.value)
                        secondDate = DateAdd("d", 1, firstDate)
                        auditStDt = Format(secondDate, "dd/mm/yyyy")
                    End If
                    
                    'temp code added start
                    If (auditStDt <> "") Then
                        If IsDate(auditStDt) = True Then
                            Length = Len(auditStDt)
                            For i = 1 To Length
                                        If (Mid(auditStDt, i, 1) = "-") Then
                                            auditStDt = Replace(auditStDt, "-", "/")
                                        End If
                            Next i
                        End If
                    End If
                    'temp code added end
                    
                    Worksheets("A_Basic_Info").Range("RetInf.auditStartDate").value = auditStDt
                    
                    
                    'new code added for set day in Sec B and Sec C Start
                    Worksheets("A_Basic_Info").Range("RetInf.ReturnToDay").value = dd
                    Worksheets("A_Basic_Info").Range("RtnInf.MonthCode").value = mm
                    Worksheets("A_Basic_Info").Range("RetInf.YearIncome").value = yearU
                     
                     If (dd <> "") Then
                        If (dd = "02" Or dd = "22") Then
                            Worksheets("A_Basic_Info").Range("RetInf.ReturnToDayStr").value = "nd"
                        ElseIf (dd = "01" Or dd = "21" Or dd = "31") Then
                            Worksheets("A_Basic_Info").Range("RetInf.ReturnToDayStr").value = "st"
                        Else
                            Worksheets("A_Basic_Info").Range("RetInf.ReturnToDayStr").value = "th"
                        End If
                    End If
                    'new code added for set day in Sec B and Sec C End
                    
                    'new code for set month in word start
                    Call setmonth(mm)
                    'new code for set month in word End
                    
                    If (DateValue(endDate) > DateValue(DsysDate)) Then
                        MsgBox "You cannot file the return for the period whose date of Return Period From and / or Return Period To is a future date."
                        Target.value = ""
                        Target.Select
                    End If
                    If (endDate <> "" And RPToAct <> "") Then
                        If (DateValue(endDate) > DateValue(RPToAct)) Then
                            MsgBox "The year in Return Period From and Return Period To should be same."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    If (startDate <> "" And endDate <> "") Then
                        If (DateValue(endDate) < DateValue(startDate)) Then
                                MsgBox "The date in Return Period To should be after the date in Return Period From."
                                Target.value = ""
                                Target.Select
                        End If
                    End If
                End If
            Else
                Worksheets("A_Basic_Info").Range("RetInf.RtnPrdToActStart").value = ""
                Worksheets("A_Basic_Info").Range("RetInf.dateforAuditCertificate").value = ""
                Worksheets("A_Basic_Info").Range("RetInf.auditStartDate").value = ""
                Worksheets("A_Basic_Info").Range("RetInf.ReturnToDay").value = ""
                Worksheets("A_Basic_Info").Range("RetInf.ReturnToDayStr").value = ""
                Worksheets("A_Basic_Info").Range("RtnInf.MonthCode").value = ""
                Worksheets("A_Basic_Info").Range("RetInf.MonthDesc").value = ""
                Worksheets("A_Basic_Info").Range("RetInf.YearIncome").value = ""
                Worksheets("A_Basic_Info").Range("RetInf.MonthDesc").value = ""
            End If
                Target.Select
        End If
        
        Dim otherVal As String
        
        'Added on 13/07/2023 ticket 64843
            Dim rtnPrdTo1 As String
            Dim year1 As String
        If (Worksheets("A_Basic_Info").Range("RetInf.RtnPrdToAct").value <> "") Then
            rtnPrdTo1 = CDate(Format(Worksheets("A_Basic_Info").Range("RetInf.RtnPrdToAct").value, "dd/mm/yyyy"))
            'rtnPrdTo1 = Worksheets("A_Basic_Info").Range("RetInf.RtnPrdToAct").value
            If (rtnPrdTo1 <> "") Then
                year1 = Format(CDate(Trim(rtnPrdTo1)), "yyyy")
                'mm = Format(CDate(Trim(rtnPrdTo)), "MM")
                If (year1 <> "" And mm <> "") Then
                    Worksheets("E1_IDA_CA").Unprotect (Pwd)
                        If ((year1 = 2022 And mm > 6) Or year1 > 2022) Then
                            With Worksheets("E1_IDA_CA").Range("locationCombo").Validation
                                    .Delete
                                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                                    xlBetween, Formula1:="=LocationJuly2022"
                                    .IgnoreBlank = True
                                    .InCellDropdown = True
                                    .InputTitle = ""
                                    .ErrorTitle = ""
                                    .InputMessage = ""
                                    .ErrorMessage = ""
                                    .ShowInput = True
                                    .ShowError = True
                            End With
                        ElseIf ((year1 = 2022 And mm >= 1) Or (year1 = 2022 And mm1 < 7)) Then
                            With Worksheets("E1_IDA_CA").Range("locationCombo").Validation
                                    .Delete
                                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                                    xlBetween, Formula1:="=LocationJan2022"
                                    .IgnoreBlank = True
                                    .InCellDropdown = True
                                    .InputTitle = ""
                                    .ErrorTitle = ""
                                    .InputMessage = ""
                                    .ErrorMessage = ""
                                    .ShowInput = True
                                    .ShowError = True
                            End With
                        ElseIf ((year1 = 2020 And mm > 3) Or (year1 < 2022 And year1 > 2020)) Then
                            With Worksheets("E1_IDA_CA").Range("locationCombo").Validation
                                    .Delete
                                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                                    xlBetween, Formula1:="=Location2020"
                                    .IgnoreBlank = True
                                    .InCellDropdown = True
                                    .InputTitle = ""
                                    .ErrorTitle = ""
                                    .InputMessage = ""
                                    .ErrorMessage = ""
                                    .ShowInput = True
                                    .ShowError = True
                            End With
                            
                        ElseIf ((year1 = 2020 And mm < 5) Or year1 < 2020) Then
                            With Worksheets("E1_IDA_CA").Range("locationCombo").Validation
                                    .Delete
                                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                                    xlBetween, Formula1:="=Location"
                                    .IgnoreBlank = True
                                    .InCellDropdown = True
                                    .InputTitle = ""
                                    .ErrorTitle = ""
                                    .InputMessage = ""
                                    .ErrorMessage = ""
                                    .ShowInput = True
                                    .ShowError = True
                            End With
                        End If
                        Worksheets("E1_IDA_CA").Protect (Pwd)
                    End If
                End If
           End If
        'End on 13/07/2023 ticket 64843

        
        If Target.Address = "$B$7" Then
        
            Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
            Worksheets("T_Tax_Computation").Activate
            ActiveSheet.Range("TaxComp.WithHoldingListSTO").value = 0#
            Worksheets("T_Tax_Computation").Protect (Pwd)
            otherVal = Worksheets("A_Basic_Info").Range("B19").value
            Call showHideWorksheet(Target.value, "B_Profit_Loss_Account_Self")
            Call showHideWorksheetBI(Target.value, otherVal, "C_Balance_Sheet")
            Call showHideWorksheetBI(Target.value, otherVal, "D_Stock_Analysis")
            Call showHideWorksheetBI(Target.value, otherVal, "E1_IDA_CA")
            Call showHideWorksheetBI(Target.value, otherVal, "E2_CA_WTA_WDV")
            If checkIfReturnPeriodIsFilled = False Then
            Else
            mm = Format(CDate(Trim(endDate)), "MM")
                    dd = Format(CDate(Trim(endDate)), "dd")
                    year = Format(CDate(Trim(endDate)), "yyyy")
            If mm > 3 And year = 2020 Or year > 2020 Then
                        ActiveWorkbook.Unprotect (Pwd)
                        Worksheets("E2_CA_WTA_SLM").Unprotect (Pwd)
                        Worksheets("E2_CA_WTA_SLM").Visible = xlSheetHidden
                        Call lockUnlock_cell_rng("E2_CA_WTA_SLM", "WAT.ListBS", True, "10")
                        Call lockUnlock_cell_rng("E2_CA_WTA_SLM", "WAT.ListBW", True, "10")
                        Worksheets("E2_CA_WTA_SLM").Protect (Pwd)
                    Else
                     ActiveWorkbook.Unprotect (Pwd)
                        Worksheets("E2_CA_WTA_SLM").Unprotect (Pwd)
                        Worksheets("E2_CA_WTA_SLM").Visible = xlSheetVisible
                        Call lockUnlock_cell_rng("E2_CA_WTA_SLM", "WAT.ListBS", False, "10")
                        If Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value = "Yes" Then
                         Call lockUnlock_cell_rng("E2_CA_WTA_SLM", "WAT.ListBW", False, "10")
                        End If
                        Worksheets("E2_CA_WTA_SLM").Protect (Pwd)
                    End If
            Call showHideWorksheetBI(Target.value, otherVal, "E_Summary_of_Capital_Allowance")
            Call showHideWorksheetBI(Target.value, otherVal, "N_Installment_Tax_Credits")
            Call showHideWorksheetBI(Target.value, otherVal, "O_WHT_Credits")
            Call showHideWorksheetBI(Target.value, otherVal, "S_Previous_Years_Losses")
            Call showHideWorksheet(Target.value, "T_Income_Computation_Self")
            'If (Target.value = "Yes") Then
            If (Worksheets("T_Tax_Computation").Range("TaxComp.MRI").value > 0) Then
                    Worksheets("B_Profit_Loss_Account_Self").Unprotect Password:=Pwd
                    If (Worksheets("B_Profit_Loss_Account_Self").Range("PLA.GrossTurnoverBPRentalToS").value = 0) Then
                        Worksheets("B_Profit_Loss_Account_Self").Protect (Pwd)
                        Worksheets("T_Tax_Computation").Activate
                        ActiveSheet.Range("TaxComp.MRI").value = 0#
                    End If
            End If
        End If
            If (Target.value = "No") Then
            If (Worksheets("T_Tax_Computation").Range("TaxComp.WithHoldingListSTO").value > 0) Then
                    Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
                    Worksheets("T_Tax_Computation").Range("TaxComp.WithHoldingListSTO").value = 0#
                    Worksheets("T_Tax_Computation").Protect (Pwd)
            End If
            End If
            Worksheets("A_Basic_Info").Activate
            Worksheets("A_Basic_Info").Range("RetInf.DeclareSelfBusIncome").Select
                
        ElseIf Target.Address = "$B$8" Then
            Call showHideWorksheet(Target.value, "G_Partnership_Income")
            Worksheets("A_Basic_Info").Activate
            Worksheets("A_Basic_Info").Range("dispSec.PartnershipIncome").Select
            
        ElseIf Target.Address = "$B$9" Then
            Call showHideWorksheet(Target.value, "H_Estate_Trust_Income")
            Worksheets("A_Basic_Info").Activate
            Worksheets("A_Basic_Info").Range("dispSec.EstateTrustIncome").Select
            
        ElseIf Target.Address = "$B$10" Then
            Call showHideWorksheet(Target.value, "I_Computation_of_Car_Benefit")
            Worksheets("A_Basic_Info").Activate
            Worksheets("A_Basic_Info").Range("dispSec.CarBenefit").Select
             
        ElseIf Target.Address = "$B$11" Then
            Call showHideWorksheet(Target.value, "J_Computation_of_Mortgage")
            Worksheets("A_Basic_Info").Activate
            Worksheets("A_Basic_Info").Range("dispSec.Mortgage").Select
              
        ElseIf Target.Address = "$B$12" Then
            Call showHideWorksheet(Target.value, "K_Home_Ownership_Saving_Plan")
            If (Target.value = "No") Then
                 Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
                Worksheets("T_Tax_Computation").Range("TaxComp.HomeOwnershipSavingPlanListSTO").value = 0#
                Worksheets("T_Tax_Computation").Range("TaxComp.HomeOwnershipSavingPlanListWTO").value = 0#
                Worksheets("T_Tax_Computation").Protect (Pwd)
            End If
            Worksheets("A_Basic_Info").Activate
            Worksheets("A_Basic_Info").Range("dispSec.OwnershipSaving").Select
              
        ElseIf Target.Address = "$B$13" Then
            Call showHideWorksheet(Target.value, "L_Computation_of_Insu_Relief")
            Worksheets("A_Basic_Info").Activate
            Worksheets("A_Basic_Info").Range("dispSec.InsuRelief").Select
             
        ElseIf Target.Address = "$B$14" Then
            Call showHideWorksheet(Target.value, "P_Advance_Tax_Credits")
            Worksheets("A_Basic_Info").Activate
            Worksheets("A_Basic_Info").Range("dispSec.AdvanceTax").Select
        
        ElseIf Target.Address = "$B$15" Then
            Call showHideWorksheet(Target.value, "R_DTAA_Credits")
            Worksheets("A_Basic_Info").Activate
            Worksheets("A_Basic_Info").Range("dispSec.DTAA").Select
            
        ElseIf Target.Address = "$B$19" Then
            otherVal = Worksheets("A_Basic_Info").Range("B7").value
            Call showHideWorksheet(Target.value, "B_Profit_Loss_Account_Wife")
            Call showHideWorksheetBI(Target.value, otherVal, "C_Balance_Sheet")
            Call showHideWorksheetBI(Target.value, otherVal, "D_Stock_Analysis")
            Call showHideWorksheetBI(Target.value, otherVal, "E1_IDA_CA")
            Call showHideWorksheetBI(Target.value, otherVal, "E2_CA_WTA_WDV")
             mm = Format(CDate(Trim(endDate)), "MM")
                    dd = Format(CDate(Trim(endDate)), "dd")
                    year = Format(CDate(Trim(endDate)), "yyyy")
             If mm > 3 And year = 2020 Or year > 2020 Then
                        ActiveWorkbook.Unprotect (Pwd)
                        Worksheets("E2_CA_WTA_SLM").Unprotect (Pwd)
                        Worksheets("E2_CA_WTA_SLM").Visible = xlSheetHidden
                        Call lockUnlock_cell_rng("E2_CA_WTA_SLM", "WAT.ListBS", True, "10")
                        Call lockUnlock_cell_rng("E2_CA_WTA_SLM", "WAT.ListBW", True, "10")
                        Worksheets("E2_CA_WTA_SLM").Protect (Pwd)
                    Else
                     ActiveWorkbook.Unprotect (Pwd)
                        Worksheets("E2_CA_WTA_SLM").Unprotect (Pwd)
                        Worksheets("E2_CA_WTA_SLM").Visible = xlSheetVisible
                        Call lockUnlock_cell_rng("E2_CA_WTA_SLM", "WAT.ListBS", False, "10")
                        If Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value = "Yes" Then
                         Call lockUnlock_cell_rng("E2_CA_WTA_SLM", "WAT.ListBW", False, "10")
                        End If
                        Worksheets("E2_CA_WTA_SLM").Protect (Pwd)

                    End If
            Call showHideWorksheetBI(Target.value, otherVal, "E_Summary_of_Capital_Allowance")
            Call showHideWorksheetBI(Target.value, otherVal, "N_Installment_Tax_Credits")
            Call showHideWorksheetBI(Target.value, otherVal, "O_WHT_Credits")
            Call showHideWorksheetBI(Target.value, otherVal, "S_Previous_Years_Losses")
            Call showHideWorksheet(Target.value, "T_Income_Computation_Wife")
            'If (Target.value = "Yes") Then
            If (Worksheets("T_Tax_Computation").Range("TaxComp.MRIWTO").value > 0) Then
                    Worksheets("B_Profit_Loss_Account_Wife").Unprotect Password:=Pwd
                    If (Worksheets("B_Profit_Loss_Account_Wife").Range("PLA.GrossTurnoverBPRentalToW").value = 0) Then
                        Worksheets("B_Profit_Loss_Account_Wife").Protect (Pwd)
                        Worksheets("T_Tax_Computation").Activate
                        ActiveSheet.Range("TaxComp.MRIWTO").value = 0#
                    End If
                End If
            'End If
            Worksheets("A_Basic_Info").Activate
            Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeBusIncome").Select
        End If
        
        
        Dim resultMsg As String
        Select Case Target.Address
            'Check PIN Format and convert in Uppar Case Start
            Case "$B$3"
                If Target.value <> "" Then
                    If Mid(Target.value, 1, 1) <> "A" And Mid(Target.value, 1, 1) <> "a" Then
                        MsgBox "Please enter valid Individual Personal Identification Number."
                        Target.value = ""
                        Target.Select
                    End If
                End If
                 
                Target.value = StrConv(Target.Text, vbUpperCase)
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
                        MsgBox "Please enter Personal Identification Number in proper format."
                        Target.value = ""
                        Target.Select
                    End If
                End If
            'Check PIN Format and convert in Uppar Case End
            
            'Check PIN Format and convert in Uppar Case Start
            Case "$B$18"
                 If Target.value <> "" Then
                     If Target.value <> "" Then
                        If Mid(Target.value, 1, 1) <> "A" And Mid(Target.value, 1, 1) <> "a" Then
                            MsgBox "Please enter valid Individual Wife's PIN."
                            Target.value = ""
                            Target.Select
                        End If
                     End If
                    Target.value = UCase(Target.value)
                    If (Target.value <> "") Then
                        If (Target.value = PINNoS) Then
                            MsgBox "PIN entered should not be same as Taxpayer's PIN."
                            Target.value = ""
                            Target.Select
                        End If
                     End If
                     If (Target.value <> "") Then
                        result = TestPIN(Target.value)
                        If (result = False) Then
                            MsgBox "Please enter Personal Identification Number  in proper format."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                 End If
            'Check PIN Format and convert in Uppar Case End
            
            'Return Type code set in hidden field in sheet Start
            Case "$B$4"
                If (Target.value <> "") Then
                    rowvalue = find_RetTypeCode(Target.value)
                    RetTypeCode = Sheet18.Cells(rowvalue, RetTypeCodeCol).value 'RetTypeCodeCol = 5th column of data sheet
                     If (rowvalue <> 0) Then
                             Worksheets("A_Basic_Info").Range("C4:C4").value = RetTypeCode
                             Worksheets("A_Basic_Info").Range("C4:C4").Locked = True
                     End If
                Else
                    '** make ret type code field blank
                    Worksheets("A_Basic_Info").Range("C4:C4").Locked = True
                    Worksheets("A_Basic_Info").Range("C4:C4").value = ""
                End If
            'Return Type code set in hidden field in sheet End
            
            'Base on select value from Have you been issued exemption certificate or not, other exemption certificate related section will be Enable/Disable Start
            Case "$B$16"
                Application.ThisWorkbook.Unprotect (Pwd)
                lockCell = True
                If Target.value <> "" Then
                    If UCase(Worksheets("A_Basic_Info").Range("RetInf.DeclareExemptionCerti").value) = "YES" Then
                        Worksheets("A_Basic_Info").Range("RetInf.DeclareExemptionCerti").value = "Yes"
                        lockCell = False
                    End If
                End If
                Call lockUnlock_cell_rng("A_Basic_Info", "ExemptCerti.ListS", lockCell, 4)
                Application.ThisWorkbook.Protect (Pwd)
                Worksheets("A_Basic_Info").Activate
                Worksheets("A_Basic_Info").Range("RetInf.DeclareExemptionCerti").Select
            'Base on select value from Have you been issued exemption certificate or not, other exemption certificate related section will be Enable/Disable End
            
            'Base on select value from Do you want to Declare wife's income yes/no then Wife related other section will be Enable/Disable Start
            Case "$B$17"
                year = Format(CDate(Trim(endDate)), "yyyy")
                mm = Format(CDate(Trim(endDate)), "MM")
                Call b17("$B$17", workSheetArray, year, mm)
            'Base on select value from Do you want to Declare wife's income yes/no then Wife related other section will be Enable/Disable End
            
            'Base on select value from Does your wife have been issued exemption certificate or not, other exemption certificate related section will be Enable/Disable Start
            Case "$B$20"
                Application.ThisWorkbook.Unprotect (Pwd)
                lockCell = True
                If Target.value <> "" Then
                    If UCase(Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeExemptionCerti").value) = "YES" Then
                        Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeExemptionCerti").value = "Yes"
                        lockCell = False
                    End If
                End If
                Call lockUnlock_cell_rng("A_Basic_Info", "ExemptCerti.ListW", lockCell, 4)
                Application.ThisWorkbook.Protect (Pwd)
                Worksheets("A_Basic_Info").Activate
                Worksheets("A_Basic_Info").Range("RetInf.DeclareExemptionCerti").Select
            'Base on select value from Does your wife have been issued exemption certificate or not, other exemption certificate related section section will be Enable/Disable End
            
            'Check Date of Audit Certificate validate,Date can be enter from Return Period End Date to 6 Month for Self Start
            Case "$B$32"
                    Dim tValue As String
                    If Target.value <> "" Then
                        If TestDate(Target.value) = False Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            If (endDate <> "") Then
                                tValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                                If (DateValue(tValue) > DateValue(endDate) And DateValue(tValue) <= DateValue(auditEndDate)) Then
                                    'do nothing
                                Else
                                    MsgBox "Date of Audit Certificate should be between " & Worksheets("A_Basic_Info").Range("RetInf.auditStartDate").value & " and " & Worksheets("A_Basic_Info").Range("RetInf.dateforAuditCertificate").value
                                    Target.value = ""
                                    Target.Select
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                            End If
                        End If
                    End If
             'Check Date of Audit Certificate validate,Date can be enter from Return Period End Date to 6 Month for Self End
             
            'Check Date of Audit Certificate validate,Date can be enter from Return Period End Date to 6 Month for Wife Start
            Case "$C$32"
                    Dim tWValue As String
                    If Target.value <> "" Then
                        If TestDate(Target.value) = False Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            If (endDate <> "") Then
                                tWValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                                If (DateValue(tWValue) > DateValue(endDate) And DateValue(tWValue) <= DateValue(auditEndDate)) Then
                                    'do nothing
                                Else
                                    MsgBox "Date of Audit Certificate should be between " & Worksheets("A_Basic_Info").Range("RetInf.auditStartDate").value & " and " & Worksheets("A_Basic_Info").Range("RetInf.dateforAuditCertificate").value
                                    Target.value = ""
                                    Target.Select
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                            End If
                        End If
                    End If
            'Check Date of Audit Certificate validate,Date can be enter from Return Period End Date to 6 Month for Wife End
        End Select
        
    
        'Code for Part 4 Self  Start
        rangeName = Range("RentPaid.RentListS").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1
    
        For i = startRow To endRow
            Select Case Target.Address
                'PIN Convert in Uppar Case Start
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
                            MsgBox "Please enter PIN of Landlord in proper format."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                 'PIN Convert in Uppar Case End
                 
                'Find County Id and set in hidden field in sheet Start
                Case "$G$" & i
                    If (Target.value <> "") Then
                        rowvalue2 = find_CountyID(Target.value)
                        If rowvalue2 <> -1 Then
                            selectedCountyId = Sheet18.Cells(rowvalue2, 15).value  '15th column of data sheet
                          If (rowvalue2 <> 0) Then
                            Worksheets("A_Basic_Info").Range("R" & (i) & ":R" & i).value = selectedCountyId
                            Target.Next.value = ""
                          End If
                            Worksheets("A_Basic_Info").Cells(i, 8).Select
                         Else
                            Target.value = ""
                         End If
                    Else
                        Worksheets("A_Basic_Info").Range("R" & (i) & ":R" & i).value = ""
                        Worksheets("A_Basic_Info").Range("T" & (i) & ":T" & i).value = ""
                        Worksheets("A_Basic_Info").Range("U" & (i) & ":U" & i).value = ""
                        Target.Next.value = ""
                        Worksheets("A_Basic_Info").Cells(i, 7).Select
                    End If
                 'Find County Id and set in hidden field in sheet End
                
                'find_DistrictID and set in hidden field in sheet Start
                Case "$H$" & i             'Set district id
                    If (Target.value <> "") Then
                        rowvalue2 = find_DistrictID(Target.value)
                       If rowvalue2 <> -1 Then
                        selectedDistrictId = Sheet18.Cells(rowvalue2, 27).value  '27th column of data sheet
                         If (rowvalue2 <> 0) Then
                            Worksheets("A_Basic_Info").Range("T" & (i) & ":T" & i).value = selectedDistrictId
                            Target.Next.value = ""
                         End If
                         Worksheets("A_Basic_Info").Cells(i, 9).Select
                       Else
                         Target.value = ""
                        End If
                    Else
                        Worksheets("A_Basic_Info").Range("T" & (i) & ":T" & i).value = ""
                        Worksheets("A_Basic_Info").Range("U" & (i) & ":U" & i).value = ""
                        Target.Next.value = ""
                        Worksheets("A_Basic_Info").Cells(i, 8).Select
                    End If
                    'find_DistrictID and set in hidden field in sheet End
                    
                'find Location ID and Set in Hidden Field in sheet Start
                Case "$I$" & i             'Set district id
                    If (Target.value <> "") Then
                       rowvalue2 = find_LocalID(Target.value)
                       If rowvalue2 <> -1 Then
                        selectedlocalId = Sheet18.Cells(rowvalue2, 31).value  '32nd column of data sheet
                         If (rowvalue2 <> 0) Then
                            Worksheets("A_Basic_Info").Range("U" & (i) & ":U" & i).value = selectedlocalId
                         End If
                         Worksheets("A_Basic_Info").Cells(i, 10).Select
                        Else
                            Target.value = ""
                        End If
                    Else
                        Worksheets("A_Basic_Info").Range("U" & (i) & ":U" & i).value = ""
                        Worksheets("A_Basic_Info").Cells(i, 9).Select
                    End If
                'find Location ID and Set in Hidden Field in sheet End
                
                'find Towan Id and Set in Hidden field in sheet Start
                Case "$L$" & i                 'Set postal code
                    If (Target.value <> "") Then
                        rowvalue111 = find_Town(Target.value)
                        If rowvalue111 <> -1 Then
                            selectedTown = Sheet18.Cells(rowvalue111, 13).value
                            If (rowvalue111 <> 0) Then
                                Target.Next.value = selectedTown
                            End If
                            Target.Select
                        Else
                            Target.value = ""
                        End If
                    Else
                        Target.Next = ""
                    End If
                'find Town Id and Set in Hidden field in sheet End
                
                Case "$M$" & i                 'Set Post Office Name
                If (Target.value <> "") Then
                    rowvalue111 = find_PostalCode(Target.value)
                    If rowvalue111 <> -1 Then
                        postOffice = Worksheets("Data").Cells(rowvalue111, 14).value
                        If (rowvalue111 <> 0) Then
                            Target.Previous.value = postOffice
                        End If
                        Target.Select
                    Else
                        Target.value = ""
                    End If
                Else
                    Target.Previous = ""
                End If
                
                'check Date valid Date format or not and check other validation date related date from should be between Return period Start
                Case "$N$" & i
                    Dim tLDSValue As String
                    Dim tLDSNXValue As String
                    If (Target.value <> "") Then
                         result = TestDate(Target.value)
                         If (result = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            tLDSValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                            If (startDate <> "" And endDate <> "") Then
                                If (DateValue(tLDSValue) < DateValue(startDate) Or DateValue(tLDSValue) > DateValue(endDate)) Then
                                    MsgBox "Date From in Row " & i & " should be within Return period " & Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value & " to " & Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                Else
                                    If Target.Next <> "" And TestDate(Target.Next) = True Then
                                        tLDSNXValue = CDate(Format(Target.Next, "dd/mm/yyyy"))
                                        If DateValue(tLDSValue) > DateValue(tLDSNXValue) Then
                                            MsgBox "Date From in Row " & i & " should be same as or before Date To "
                                            Target.value = ""
                                            Target.Select
                                        End If
                                    End If
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                            End If
                        End If
                    End If
                'check Date valid Date format or not and check other validation date related date from should be between Return period End
                    
                'check Date valid Date format or not and check other validation date related date to should be between Return period Start
                Case "$O$" & i                 'Date To validation
                    Dim tLDTOSValue As String
                    Dim tLDTOPRSValue As String
                    If (Target.value <> "") Then
                         result = TestDate(Target.value)
                         If (result = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            tLDTOSValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                            If (startDate <> "" And endDate <> "") Then
                                If (DateValue(tLDTOSValue) < DateValue(startDate) Or DateValue(tLDTOSValue) > DateValue(endDate)) Then
                                    MsgBox "Date To in Row " & i & " should be within Return period " & Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value & " to " & Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                Else
                                    If Target.Previous <> "" And TestDate(Target.Previous) = True Then
                                        tLDTOPRSValue = CDate(Format(Target.Previous, "dd/mm/yyyy"))
                                        If DateValue(tLDTOSValue) < DateValue(tLDTOPRSValue) Then
                                            MsgBox "Date To in Row " & i & " should be same as or after Date From "
                                            Target.value = ""
                                            Target.Select
                                        End If
                                    End If
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Cells("RetInf.RetStartDate").Select
                            End If
                        End If
                    End If
                    'check Date valid Date format or not and check other validation date related date to should be between Return period End
            End Select
        Next
        'Code for Part 4 Self  End
    
        'Code for Part 4 Wife  Start
        rangeName = Range("RentPaid.RentListW").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1
        
        For i = startRow To endRow
            Select Case Target.Address
                'PIN Convert in Uppar Case Start
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
                            MsgBox "Please enter PIN of Landlord in proper format."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                 'PIN Convert in Uppar Case End
                
                'Find County Id and set in hidden field in sheet Start
                Case "$G$" & i             'Set county id
                    If (Target.value <> "") Then
                       rowvalue2 = find_CountyID(Target.value)
                       If rowvalue2 <> -1 Then
                        selectedCountyId = Sheet18.Cells(rowvalue2, 15).value  '15th column of data sheet
                         If (rowvalue2 <> 0) Then
                            Worksheets("A_Basic_Info").Range("R" & (i) & ":R" & i).value = selectedCountyId
                            Target.Next.value = ""
                         End If
                         Worksheets("A_Basic_Info").Cells(i, 8).Select
                        Else
                            Target.value = ""
                        End If
                    Else
                        Worksheets("A_Basic_Info").Range("R" & (i) & ":R" & i).value = ""
                        Worksheets("A_Basic_Info").Range("T" & (i) & ":T" & i).value = ""
                        Worksheets("A_Basic_Info").Range("U" & (i) & ":U" & i).value = ""
                        Target.Next.value = ""
                        Worksheets("A_Basic_Info").Cells(i, 7).Select
                    End If
                'Find County Id and set in hidden field in sheet End
                
                'find_DistrictID and set in hidden field in sheet Start
                Case "$H$" & i             'Set district id
                    If (Target.value <> "") Then
                       rowvalue2 = find_DistrictID(Target.value)
                       If rowvalue2 <> -1 Then
                        selectedDistrictId = Sheet18.Cells(rowvalue2, 27).value  '27th column of data sheet
                         If (rowvalue2 <> 0) Then
                            Worksheets("A_Basic_Info").Range("T" & (i) & ":T" & i).value = selectedDistrictId
                            Target.Next.value = ""
                         End If
                         Worksheets("A_Basic_Info").Cells(i, 9).Select
                       Else
                            Target.value = ""
                        End If
                    Else
                        Worksheets("A_Basic_Info").Range("T" & (i) & ":T" & i).value = ""
                        Worksheets("A_Basic_Info").Range("U" & (i) & ":U" & i).value = ""
                        Target.Next.value = ""
                        Worksheets("A_Basic_Info").Cells(i, 8).Select
                    End If
                'find_DistrictID and set in hidden field in sheet End
                
                'find Location ID and Set in Hidden Field in sheet Start
                Case "$I$" & i             'Set district id
                    If (Target.value <> "") Then
                       rowvalue2 = find_LocalID(Target.value)
                       If rowvalue2 <> -1 Then
                        selectedlocalId = Sheet18.Cells(rowvalue2, 31).value  '32nd column of data sheet
                         If (rowvalue2 <> 0) Then
                            Worksheets("A_Basic_Info").Range("U" & (i) & ":U" & i).value = selectedlocalId
                         End If
                         Worksheets("A_Basic_Info").Cells(i, 10).Select
                        Else
                            Target.value = ""
                        End If
                    Else
                        Worksheets("A_Basic_Info").Range("U" & (i) & ":U" & i).value = ""
                        Worksheets("A_Basic_Info").Cells(i, 9).Select
                    End If
                'find Location ID and Set in Hidden Field in sheet End
                
                'find Towan Id and Set in Hidden field in sheet Start
                Case "$L$" & i                 'Set postal code
                    If (Target.value <> "") Then
                        rowvalue111 = find_Town(Target.value)
                        If rowvalue111 <> -1 Then
                            selectedTown = Sheet18.Cells(rowvalue111, 13).value
                            If (rowvalue111 <> 0) Then
                                Target.Next.value = selectedTown
                            End If
                            Target.Select
                        Else
                            Target.value = ""
                        End If
                    Else
                        Target.Next = ""
                    End If
                'find Towan Id and Set in Hidden field in sheet End
                
                Case "$M$" & i                 'Set Post Office Name
                If (Target.value <> "") Then
                    rowvalue111 = find_PostalCode(Target.value)
                    If rowvalue111 <> -1 Then
                        postOffice = Worksheets("Data").Cells(rowvalue111, 14).value
                        If (rowvalue111 <> 0) Then
                            Target.Previous.value = postOffice
                        End If
                        Target.Select
                    Else
                        Target.value = ""
                    End If
                Else
                    Target.Previous = ""
                End If
                
                'check Date valid Date format or not and check other validation date related date from should be between Return period Start
                Case "$N$" & i              'Date From validation
                    Dim tLDWValue As String
                    Dim tLDNXWValue As String
                    If (Target.value <> "") Then
                         result = TestDate(Target.value)
                         If (result = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            tLDWValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                            If (startDate <> "" And endDate <> "") Then
                                If (DateValue(tLDWValue) < DateValue(startDate) Or DateValue(tLDWValue) > DateValue(endDate)) Then
                                    MsgBox "Date From in Row " & i & " should be within Return period " & Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value & " to " & Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                Else
                                    If Target.Next <> "" And TestDate(Target.Next) = True Then
                                        tLDNXWValue = CDate(Format(Target.Next, "dd/mm/yyyy"))
                                        If DateValue(tLDWValue) > DateValue(tLDNXWValue) Then
                                            MsgBox "Date From in Row " & i & " should be same as or before Date To "
                                            Target.value = ""
                                            Target.Select
                                        End If
                                    End If
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                            End If
                        End If
                    End If
                'check Date valid Date format or not and check other validation date related date from should be between Return period End
                
                'check Date valid Date format or not and check other validation date related date to should be between Return period Start
                Case "$O$" & i                 'Date To validation
                    Dim tLDTOWValue As String
                    Dim tLDTOPRWValue As String
                    If (Target.value <> "") Then
                         result = TestDate(Target.value)
                         If (result = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            tLDTOWValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                            If (startDate <> "" And endDate <> "") Then
                                If (DateValue(tLDTOWValue) < DateValue(startDate) Or DateValue(tLDTOWValue) > DateValue(endDate)) Then
                                    MsgBox "Date To in Row " & i & " should be within Return period " & Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value & " to " & Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                Else
                                    If Target.Previous <> "" And TestDate(Target.Previous) = True Then
                                        tLDTOPRWValue = CDate(Format(Target.Previous, "dd/mm/yyyy"))
                                        If DateValue(tLDTOWValue) < DateValue(tLDTOPRWValue) Then
                                            MsgBox "Date To in Row " & i & " should be same as or after Date From "
                                            Target.value = ""
                                            Target.Select
                                        End If
                                    End If
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                            End If
                        End If
                    End If
                'check Date valid Date format or not and check other validation date related date to should be between Return period End
            End Select
        Next
        'Code for Part 4 Wife  End
    
        'Code for Part 5 Self  Start
        rangeName = Range("RentalIncome.ListS").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1
    
        For i = startRow To endRow
            Select Case Target.Address
                'PIN Convert in Uppar Case Start
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
                            MsgBox "Please enter PIN of Tenant in proper format."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                'PIN Convert in Uppar Case End
                
                'Find County Id and set in hidden field in sheet Start
                Case "$G$" & i             'Set county id
                    If (Target.value <> "") Then
                       rowvalue2 = find_CountyID(Target.value)
                       If rowvalue2 <> -1 Then
                        selectedCountyId = Sheet18.Cells(rowvalue2, 15).value  '15th column of data sheet
                         If (rowvalue2 <> 0) Then
                            Worksheets("A_Basic_Info").Range("R" & (i) & ":R" & i).value = selectedCountyId
                            Target.Next.value = ""
                         End If
                         Worksheets("A_Basic_Info").Cells(i, 8).Select
                        Else
                            Target.value = ""
                        End If
                    Else
                        Worksheets("A_Basic_Info").Range("R" & (i) & ":R" & i).value = ""
                        Worksheets("A_Basic_Info").Range("T" & (i) & ":T" & i).value = ""
                        Worksheets("A_Basic_Info").Range("U" & (i) & ":U" & i).value = ""
                        Target.Next.value = ""
                        Worksheets("A_Basic_Info").Cells(i, 7).Select
                    End If
                'Find County Id and set in hidden field in sheet End
                
                'find_DistrictID and set in hidden field in sheet Start
                Case "$H$" & i
                    If (Target.value <> "") Then
                       rowvalue2 = find_DistrictID(Target.value)
                       If rowvalue2 <> -1 Then
                        selectedDistrictId = Sheet18.Cells(rowvalue2, 27).value  '27th column of data sheet
                         If (rowvalue2 <> 0) Then
                            Worksheets("A_Basic_Info").Range("T" & (i) & ":T" & i).value = selectedDistrictId
                            Target.Next.value = ""
                         End If
                         Worksheets("A_Basic_Info").Cells(i, 9).Select
                       Else
                            Target.value = ""
                        End If
                    Else
                        Worksheets("A_Basic_Info").Range("T" & (i) & ":T" & i).value = ""
                        Worksheets("A_Basic_Info").Range("U" & (i) & ":U" & i).value = ""
                        Target.Next.value = ""
                        Worksheets("A_Basic_Info").Cells(i, 8).Select
                    End If
                'find_DistrictID and set in hidden field in sheet End
                
                'find Location ID and Set in Hidden Field in sheet Start
                Case "$I$" & i             'Set district id
                    If (Target.value <> "") Then
                       rowvalue2 = find_LocalID(Target.value)
                       If rowvalue2 <> -1 Then
                        selectedlocalId = Sheet18.Cells(rowvalue2, 31).value  '32nd column of data sheet
                         If (rowvalue2 <> 0) Then
                            Worksheets("A_Basic_Info").Range("U" & (i) & ":U" & i).value = selectedlocalId
                         End If
                         Worksheets("A_Basic_Info").Cells(i, 10).Select
                        Else
                            Target.value = ""
                        End If
                    Else
                        Worksheets("A_Basic_Info").Range("U" & (i) & ":U" & i).value = ""
                        Worksheets("A_Basic_Info").Cells(i, 9).Select
                    End If
                'find Location ID and Set in Hidden Field in sheet End
                
                'find Towan Id and Set in Hidden field in sheet Start
                Case "$L$" & i                 'Set postal code
                    If (Target.value <> "") Then
                        rowvalue111 = find_Town(Target.value)
                        If rowvalue111 <> -1 Then
                            selectedTown = Sheet18.Cells(rowvalue111, 13).value
                            If (rowvalue111 <> 0) Then
                                Target.Next.value = selectedTown
                            End If
                            Target.Select
                        Else
                            Target.value = ""
                        End If
                    Else
                        Target.Next = ""
                    End If
                'find Towan Id and Set in Hidden field in sheet End
                
                Case "$M$" & i                 'Set Post Office Name
                If (Target.value <> "") Then
                    rowvalue111 = find_PostalCode(Target.value)
                    If rowvalue111 <> -1 Then
                        postOffice = Worksheets("Data").Cells(rowvalue111, 14).value
                        If (rowvalue111 <> 0) Then
                            Target.Previous.value = postOffice
                        End If
                        Target.Select
                    Else
                        Target.value = ""
                    End If
                Else
                    Target.Previous = ""
                End If
                
                'check Date valid Date format or not and check other validation date related date from should be between Return period Start
                Case "$N$" & i              'Date From validation
                    Dim tTDSValue As String
                    Dim tTDSNXValue As String
                    If (Target.value <> "") Then
                         result = TestDate(Target.value)
                         If (result = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            tTDSValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                            If (startDate <> "" And endDate <> "") Then
                                If (DateValue(tTDSValue) < DateValue(startDate) Or DateValue(tTDSValue) > DateValue(endDate)) Then
                                    MsgBox "Date From in Row " & i & " should be within Return period " & Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value & " to " & Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                Else
                                    If Target.Next <> "" And TestDate(Target.Next) = True Then
                                        tTDSNXValue = CDate(Format(Target.Next, "dd/mm/yyyy"))
                                        If DateValue(tTDSValue) > DateValue(tTDSNXValue) Then
                                            MsgBox "Date From in Row " & i & " should be same as or before Date To "
                                            Target.value = ""
                                            Target.Select
                                        End If
                                    End If
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                            End If
                        End If
                    End If
                'check Date valid Date format or not and check other validation date related date from should be between Return period End
                
                'check Date valid Date format or not and check other validation date related date to should be between Return period Start
                Case "$O$" & i                 'Date To validation
                    Dim tTDTOSValue As String
                    Dim tTDTOPRSValue As String
                    If (Target.value <> "") Then
                         result = TestDate(Target.value)
                         If (result = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            tTDTOSValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                            If (startDate <> "" And endDate <> "") Then
                                If (DateValue(tTDTOSValue) < DateValue(startDate) Or DateValue(tTDTOSValue) > DateValue(endDate)) Then
                                    MsgBox "Date To in Row " & i & " should be within Return period " & Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value & " to " & Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                Else
                                    If Target.Previous <> "" And TestDate(Target.Previous) = True Then
                                        tTDTOPRSValue = CDate(Format(Target.Previous, "dd/mm/yyyy"))
                                        If DateValue(tTDTOSValue) < DateValue(tTDTOPRSValue) Then
                                            MsgBox "Date To in Row " & i & " should be same as or after Date From "
                                            Target.value = ""
                                            Target.Select
                                        End If
                                    End If
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                            End If
                        End If
                    End If
                'check Date valid Date format or not and check other validation date related date to should be between Return period End
            End Select
        Next
        'Code for Part 5 Self  End
    
        'Code for Part 5 Wife  Start
        rangeName = Range("RentalIncome.ListW").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1
    
        For i = startRow To endRow
            Select Case Target.Address
                'PIN Convert in Uppar Case Start
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
                            MsgBox "Please enter PIN of Tenant in proper format."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                 'PIN Convert in Uppar Case End
                
                'Find County Id and set in hidden field in sheet Start
                Case "$G$" & i             'Set county id
                    If (Target.value <> "") Then
                        rowvalue2 = find_CountyID(Target.value)
                        If rowvalue2 <> -1 Then
                        selectedCountyId = Sheet18.Cells(rowvalue2, 15).value  '15th column of data sheet
                         If (rowvalue2 <> 0) Then
                            Worksheets("A_Basic_Info").Range("R" & (i) & ":R" & i).value = selectedCountyId
                            Target.Next.value = ""
                         End If
                         Worksheets("A_Basic_Info").Cells(i, 8).Select
                        Else
                            Target.value = ""
                        End If
                    Else
                        Worksheets("A_Basic_Info").Range("R" & (i) & ":R" & i).value = ""
                        Worksheets("A_Basic_Info").Range("T" & (i) & ":T" & i).value = ""
                        Worksheets("A_Basic_Info").Range("U" & (i) & ":U" & i).value = ""
                        Target.Next.value = ""
                        Worksheets("A_Basic_Info").Cells(i, 7).Select
                    End If
                'Find County Id and set in hidden field in sheet End
                
                'find_DistrictID and set in hidden field in sheet Start
                Case "$H$" & i             'Set district id
                    If (Target.value <> "") Then
                       rowvalue2 = find_DistrictID(Target.value)
                       If rowvalue2 <> -1 Then
                        selectedDistrictId = Sheet18.Cells(rowvalue2, 27).value  '27th column of data sheet
                         If (rowvalue2 <> 0) Then
                            Worksheets("A_Basic_Info").Range("T" & (i) & ":T" & i).value = selectedDistrictId
                            Target.Next.value = ""
                         End If
                         Worksheets("A_Basic_Info").Cells(i, 9).Select
                       Else
                            Target.value = ""
                        End If
                    Else
                        Worksheets("A_Basic_Info").Range("T" & (i) & ":T" & i).value = ""
                        Worksheets("A_Basic_Info").Range("U" & (i) & ":U" & i).value = ""
                        Target.Next.value = ""
                        Worksheets("A_Basic_Info").Cells(i, 8).Select
                    End If
                'find_DistrictID and set in hidden field in sheet End
                
                'find Location ID and Set in Hidden Field in sheet Start
                Case "$I$" & i             'Set district id
                    If (Target.value <> "") Then
                       rowvalue2 = find_LocalID(Target.value)
                       If rowvalue2 <> -1 Then
                        selectedlocalId = Sheet18.Cells(rowvalue2, 31).value  '32nd column of data sheet
                         If (rowvalue2 <> 0) Then
                            Worksheets("A_Basic_Info").Range("U" & (i) & ":U" & i).value = selectedlocalId
                         End If
                         Worksheets("A_Basic_Info").Cells(i, 10).Select
                        Else
                            Target.value = ""
                        End If
                    Else
                        Worksheets("A_Basic_Info").Range("U" & (i) & ":U" & i).value = ""
                        Worksheets("A_Basic_Info").Cells(i, 9).Select
                    End If
                'find Location ID and Set in Hidden Field in sheet End
                
                'find Towan Id and Set in Hidden field in sheet Start
                Case "$L$" & i                 'Set postal code
                    If (Target.value <> "") Then
                        rowvalue111 = find_Town(Target.value)
                        If rowvalue111 <> -1 Then
                            selectedTown = Sheet18.Cells(rowvalue111, 13).value
                            If (rowvalue111 <> 0) Then
                                Target.Next.value = selectedTown
                            End If
                            Target.Select
                        Else
                            Target.value = ""
                        End If
                    Else
                        Target.Next = ""
                    End If
                'find Towan Id and Set in Hidden field in sheet End
                
                Case "$M$" & i                 'Set Post Office Name
                If (Target.value <> "") Then
                    rowvalue111 = find_PostalCode(Target.value)
                    If rowvalue111 <> -1 Then
                        postOffice = Worksheets("Data").Cells(rowvalue111, 14).value
                        If (rowvalue111 <> 0) Then
                            Target.Previous.value = postOffice
                        End If
                        Target.Select
                    Else
                        Target.value = ""
                    End If
                Else
                    Target.Previous = ""
                End If
                
                'check Date valid Date format or not and check other validation date related date from should be between Return period Start
                Case "$N$" & i
                    If (Target.value <> "") Then
                         result = TestDate(Target.value)
                         If (result = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            tTDWValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                            If (startDate <> "" And endDate <> "") Then
                                If (DateValue(tTDWValue) < DateValue(startDate) Or DateValue(tTDWValue) > DateValue(endDate)) Then
                                    MsgBox "Date From in Row " & i & " should be within Return period " & Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value & " to " & Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                Else
                                    If Target.Next <> "" And TestDate(Target.Next) = True Then
                                        tTDWNXValue = CDate(Format(Target.Next, "dd/mm/yyyy"))
                                        If DateValue(tTDWValue) > DateValue(tTDWNXValue) Then
                                            MsgBox "Date From in Row " & i & " should be same as or before Date To "
                                            Target.value = ""
                                            Target.Select
                                        End If
                                    End If
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                            End If
                        End If
                    End If
                'check Date valid Date format or not and check other validation date related date from should be between Return period End
                
                'check Date valid Date format or not and check other validation date related date to should be between Return period Start
                Case "$O$" & i                 'Date To validation
                    Dim tTDTOWValue As String
                    Dim tTDTOPRWValue As String
                    If (Target.value <> "") Then
                         result = TestDate(Target.value)
                         If (result = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            tTDTOWValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                            If (startDate <> "" And endDate <> "") Then
                                If (DateValue(tTDTOWValue) < DateValue(startDate) Or DateValue(tTDTOWValue) > DateValue(endDate)) Then
                                    MsgBox "Date To in Row " & i & " should be within Return period " & Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value & " to " & Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                Else
                                    If Target.Previous <> "" And TestDate(Target.Previous) = True Then
                                        tTDTOPRWValue = CDate(Format(Target.Previous, "dd/mm/yyyy"))
                                        If DateValue(tTDTOWValue) < DateValue(tTDTOPRWValue) Then
                                            MsgBox "Date To in Row " & i & " should be same as or after Date From "
                                            Target.value = ""
                                            Target.Select
                                        End If
                                    End If
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                            End If
                        End If
                    End If
                'check Date valid Date format or not and check other validation date related date to should be between Return period End
            End Select
        Next
        'Code for Part 5 Wife  End
    
        
        'Code for Part 6 Self  Start
        rangeName = Range("ExemptCerti.ListS").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1
        
        rangeName1 = Range("ExemptCerti.ListW").Address
        startRow1 = Range(rangeName1).row
        endRow1 = Range(rangeName1).Rows.Count + startRow1 - 1
        
        Dim tDFromDateCompare As String
        Dim tDToDateCompare As String
        Dim tDToDateCompare1 As String
        For i = startRow To endRow
            Select Case Target.Address
                'check Exemption Certificate Number Start
                Case "$A$" & i
                    If (Target.value <> "") Then
                        Target.value = StrConv(Target.Text, vbUpperCase)
                        For Each r In Range("ExemptCerti.ListS").Rows
                            For Each c In r.Cells
                                rowNum = c.row
                                certToCompare = ""
                                certToCompare = Application.Range("$A$" & rowNum).value
                                If c.row > endRow Then
                                    Exit For
                                End If
                                If c.row = Target.row Then
                                    'Skip comparison with self value
                                'ElseIf Target.value <> "" And certToCompare <> "" And Target.value <> certToCompare Then
                                '    MsgBox "Exemption Certificate Number should be unique for Taxpayer."
                                '    Target.value = ""
                                '    Target.Select
                                End If
                            Next
                        Next
                        For Each r In Range("ExemptCerti.ListW").Rows
                            For Each c In r.Cells
                                rowNum = c.row
                                certToCompare = ""
                                certToCompare = Application.Range("$A$" & rowNum).value
                                If c.row > endRow1 Then
                                    Exit For
                                End If
                                If Target.value <> "" And certToCompare <> "" And Target.value = certToCompare Then
                                    MsgBox "Taxpayer's Exemption Certificate Number should not be same as Wife's Exemption Certificate Number."
                                    Target.value = ""
                                    Target.Select
                                End If
                            Next
                        Next
                    End If
                'check Exemption Certificate Number End
            
                'check Date valid Date format or not and check other validation date related date from should be between Return period Start
                Case "$B$" & i
                    If (Target.value <> "") Then
                         result = TestDate(Target.value)
                         If (result = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            tTDWValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                            If (startDate <> "" And endDate <> "") Then
                                If (DateValue(tTDWValue) < DateValue(startDate) Or DateValue(tTDWValue) > DateValue(endDate)) Then
                                    MsgBox "Date From in Row " & i & " should be within Return period " & Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value & " to " & Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                ElseIf Target.Next <> "" And TestDate(Target.Next) = True Then
                                        tTDWNXValue = CDate(Format(Target.Next, "dd/mm/yyyy"))
                                        If DateValue(tTDWValue) > DateValue(tTDWNXValue) Then
                                            MsgBox "Date From in Row " & i & " should be same as or before Date To "
                                            Target.value = ""
                                            Target.Select
                                        End If
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                            End If
                        End If
                    End If
                    If Target.value <> "" Then
                        dateString = Mid(Target.value, 3, 8)
                        firstDateToCompare = "01" & dateString
                        'If (firstDateToCompare <> "" And Target.value <> "") Then
                            'If (DateValue(Target.value) <> DateValue(firstDateToCompare)) Then
                             '   MsgBox "Date From should be first date of the month."
                            '    Target.value = ""
                           '     Target.Select
                          '      GoTo endSub
                         '   End If
                        'End If
                        For Each r In Range("ExemptCerti.ListS").Rows
                            For Each c In r.Cells
                                rowNum = c.row
                                tTDWValue = ""
                                tDFromDateCompare = ""
                                tDToDateCompare = ""
                                tDToDateCompare1 = ""
                                tTDWValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                                If (Worksheets("A_Basic_Info").Range("$B$" & rowNum).value <> "") Then
                                    tDFromDateCompare = CDate(Format(Worksheets("A_Basic_Info").Range("$B$" & rowNum).value, "dd/mm/yyyy"))
                                End If
                                If (Worksheets("A_Basic_Info").Range("$C$" & rowNum).value <> "") Then
                                    tDToDateCompare = CDate(Format(Worksheets("A_Basic_Info").Range("$C$" & rowNum).value, "dd/mm/yyyy"))
                                End If
                                If (Worksheets("A_Basic_Info").Range("$C$" & i).value <> "") Then
                                    tDToDateCompare1 = CDate(Format(Worksheets("A_Basic_Info").Range("$C$" & i).value, "dd/mm/yyyy"))
                                End If
                                If c.row > endRow Then
                                    Exit For
                                End If
                                If c.row = Target.row Then
                                    'Skip comparison with self value
                                Else
                                    If Target.value <> "" And tTDWValue <> "" And tDFromDateCompare <> "" And tDToDateCompare <> "" Then
                                        If (DateValue(tTDWValue) >= DateValue(tDFromDateCompare) And DateValue(tTDWValue) <= DateValue(tDToDateCompare)) Then
                                            MsgBox "Exemption Certificate Number for this period has been declared in row " & c.row
                                            Target.value = ""
                                            Target.Select
                                            GoTo endSub
                                        End If
                                    End If
                                    If Target.value <> "" And tTDWValue <> "" And tDFromDateCompare <> "" And tDToDateCompare <> "" And tDToDateCompare1 <> "" Then
                                        If (DateValue(tDToDateCompare1) > DateValue(tDToDateCompare) And DateValue(tTDWValue) < DateValue(tDFromDateCompare)) Then
                                            MsgBox "Exemption Certificate Number for this period has been declared in row " & c.row
                                            Target.value = ""
                                            Target.Select
                                            GoTo endSub
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                'check Date valid Date format or not and check other validation date related date from should be between Return period End
                
                'check Date valid Date format or not and check other validation date related date to should be between Return period Start
                Case "$C$" & i       'Date To validation
                    If (Target.value <> "") Then
                         result = TestDate(Target.value)
                         If (result = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            tTDTOWValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                            If (startDate <> "" And endDate <> "") Then
                                If (DateValue(tTDTOWValue) < DateValue(startDate) Or DateValue(tTDTOWValue) > DateValue(endDate)) Then
                                    MsgBox "Date To in Row " & i & " should be within Return period " & Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value & " to " & Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                ElseIf Target.Previous <> "" And TestDate(Target.Previous) = True Then
                                    tTDTOPRWValue = CDate(Format(Target.Previous, "dd/mm/yyyy"))
                                    If DateValue(tTDTOWValue) < DateValue(tTDTOPRWValue) Then
                                        MsgBox "Date To in Row " & i & " should be same as or after Date From "
                                        Target.value = ""
                                        Target.Select
                                    End If
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                            End If
                        End If
                    End If
                    If (Target.value <> "") Then
'                        dateString = Mid(Target.value, 3, 8)
'                        mm = Mid(Target.value, 4, 2)
'                        If (mm = 1 Or mm = 3 Or mm = 5 Or mm = 7 Or mm = 8 Or mm = 10 Or mm = 12) Then
'                            lastDateToCompare = "31" & dateString
'                        ElseIf (mm = 2) Then
'                            yy = Mid(Target.value, 7, 4)
'                            If (yy Mod 4 = 0) Then
'                                lastDateToCompare = "29" & dateString
'                            Else
'                                lastDateToCompare = "28" & dateString
'                            End If
'                        Else
'                            lastDateToCompare = "30" & dateString
'                        End If
'                        If (lastDateToCompare <> "" And Target.value <> "") Then
'                            If (DateValue(Target.value) <> DateValue(lastDateToCompare)) Then
'                                MsgBox "Date To should be end Date of the month."
'                                Target.value = ""
'                                Target.Select
'                                GoTo endSub
'                            End If
'                        End If
                        
                        For Each r In Range("ExemptCerti.ListS").Rows
                            For Each c In r.Cells
                                rowNum = c.row
                                
                                tDFromDateCompare = ""
                                tDToDateCompare = ""
                                tDToDateCompare1 = ""
                                If (Worksheets("A_Basic_Info").Range("$B$" & rowNum).value <> "") Then
                                    tDFromDateCompare = CDate(Format(Worksheets("A_Basic_Info").Range("$B$" & rowNum).value, "dd/mm/yyyy"))
                                End If
                                If (Worksheets("A_Basic_Info").Range("$C$" & rowNum).value <> "") Then
                                    tDToDateCompare = CDate(Format(Worksheets("A_Basic_Info").Range("$C$" & rowNum).value, "dd/mm/yyyy"))
                                End If
                                If (Worksheets("A_Basic_Info").Range("$B$" & i).value <> "") Then
                                    tDFromDateCompare1 = CDate(Format(Worksheets("A_Basic_Info").Range("$B$" & i).value, "dd/mm/yyyy"))
                                End If
                                If c.row > endRow Then
                                    Exit For
                                End If
                                If c.row = Target.row Then
                                    'Skip comparison with self value
                                Else
                                    If Target.value <> "" And tDFromDateCompare <> "" And tDToDateCompare <> "" Then
                                        If (DateValue(Target.value) >= DateValue(tDFromDateCompare) And DateValue(Target.value) <= DateValue(tDToDateCompare)) Then
                                            MsgBox "Exemption Certificate Number for this period has been declared in row " & c.row
                                            Target.value = ""
                                            Target.Select
                                            GoTo endSub
                                        End If
                                    End If
                                    If Target.value <> "" And tDFromDateCompare <> "" And tDToDateCompare <> "" And tDFromDateCompare1 <> "" Then
                                        If (DateValue(tDFromDateCompare1) < DateValue(tDFromDateCompare) And DateValue(Target.value) > DateValue(tDToDateCompare)) Then
                                            MsgBox "Exemption Certificate Number for this period has been declared in row " & c.row
                                            Target.value = ""
                                            Target.Select
                                            GoTo endSub
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                'check Date valid Date format or not and check other validation date related date to should be between Return period End
endSub:
            End Select
        Next
        'Code for Part 6 Self  End
        
        'Code for Part 6 Wife  Start
        rangeName = Range("ExemptCerti.ListW").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1
        
        rangeName1 = Range("ExemptCerti.ListS").Address
        startRow1 = Range(rangeName1).row
        endRow1 = Range(rangeName1).Rows.Count + startRow1 - 1
        
        For i = startRow To endRow
            Select Case Target.Address
            'check Exemption Certificate Number Start
                Case "$A$" & i
                        If (Target.value <> "") Then
                        Target.value = StrConv(Target.Text, vbUpperCase)
                        For Each r In Range("ExemptCerti.ListW").Rows
                            For Each c In r.Cells
                                rowNum = c.row
                                certToCompare = ""
                                certToCompare = Application.Range("$A$" & rowNum).value
                                If c.row > endRow Then
                                    Exit For
                                End If
                                If c.row = Target.row Then
                                    'Skip comparison with self value
                                ElseIf Target.value <> "" And certToCompare <> "" And Target.value <> certToCompare Then
                                    MsgBox "Exemption Certificate Number should be unique for Taxpayer."
                                    Target.value = ""
                                    Target.Select
                                End If
                            Next
                        Next
                        For Each r In Range("ExemptCerti.ListS").Rows
                            For Each c In r.Cells
                                rowNum = c.row
                                certToCompare = ""
                                certToCompare = Application.Range("$A$" & rowNum).value
                                If c.row > endRow1 Then
                                    Exit For
                                End If
                                If Target.value <> "" And certToCompare <> "" And Target.value = certToCompare Then
                                    MsgBox "Wife's Exemption Certificate Number should not be same as Taxpayer's Exemption Certificate Number."
                                    Target.value = ""
                                    Target.Select
                                End If
                            Next
                        Next
                    End If
                'check Exemption Certificate Number End
            
                'check Date valid Date format or not and check other validation date related date from should be between Return period Start
                Case "$B$" & i
                    If (Target.value <> "") Then
                         result = TestDate(Target.value)
                         If (result = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            tTDWValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                            If (startDate <> "" And endDate <> "") Then
                                If (DateValue(tTDWValue) < DateValue(startDate) Or DateValue(tTDWValue) > DateValue(endDate)) Then
                                    MsgBox "Date From in Row " & i & " should be within Return period " & Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value & " to " & Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                ElseIf Target.Next <> "" And TestDate(Target.Next) = True Then
                                        tTDWNXValue = CDate(Format(Target.Next, "dd/mm/yyyy"))
                                        If DateValue(tTDWValue) > DateValue(tTDWNXValue) Then
                                            MsgBox "Date From in Row " & i & " should be same as or before Date To "
                                            Target.value = ""
                                            Target.Select
                                        End If
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                            End If
                        End If
                    End If
                    If Target.value <> "" Then
                        dateString = Mid(Target.value, 3, 8)
                        firstDateToCompare = "01" & dateString
                        
                        For Each r In Range("ExemptCerti.ListW").Rows
                            For Each c In r.Cells
                                rowNum = c.row
                                tDFromDateCompare = ""
                                tDToDateCompare = ""
                                tDToDateCompare1 = ""
                                tDateValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                                If (Worksheets("A_Basic_Info").Range("$B$" & rowNum).value <> "") Then
                                    tDFromDateCompare = CDate(Format(Worksheets("A_Basic_Info").Range("$B$" & rowNum).value, "dd/mm/yyyy"))
                                End If
                                If (Worksheets("A_Basic_Info").Range("$C$" & rowNum).value <> "") Then
                                    tDToDateCompare = CDate(Format(Worksheets("A_Basic_Info").Range("$C$" & rowNum).value, "dd/mm/yyyy"))
                                End If
                                If (Worksheets("A_Basic_Info").Range("$C$" & i).value <> "") Then
                                    tDToDateCompare1 = CDate(Format(Worksheets("A_Basic_Info").Range("$C$" & i).value, "dd/mm/yyyy"))
                                End If
                                If c.row > endRow Then
                                    Exit For
                                End If
                                If c.row = Target.row Then
                                    'Skip comparison with self value
                                Else
                                    If Target.value <> "" And tDateValue <> "" And tDFromDateCompare <> "" And tDToDateCompare <> "" Then
                                        If (DateValue(tDateValue) >= DateValue(tDFromDateCompare) And DateValue(tDateValue) <= DateValue(tDToDateCompare)) Then
                                            MsgBox "Exemption Certificate Number for this period has been declared in row " & c.row
                                            Target.value = ""
                                            Target.Select
                                            GoTo endSub
                                        End If
                                    End If
                                    If Target.value <> "" And tDateValue <> "" And tDFromDateCompare <> "" And tDToDateCompare <> "" And tDToDateCompare1 <> "" Then
                                        If (DateValue(tDToDateCompare1) > DateValue(tDToDateCompare) And DateValue(tDateValue) < DateValue(tDFromDateCompare)) Then
                                            MsgBox "Exemption Certificate Number for this period has been declared in row " & c.row
                                            Target.value = ""
                                            Target.Select
                                            GoTo endSub
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                'check Date valid Date format or not and check other validation date related date from should be between Return period End
                
                'check Date valid Date format or not and check other validation date related date to should be between Return period Start
                Case "$C$" & i       'Date To validation
                    If (Target.value <> "") Then
                         result = TestDate(Target.value)
                         If (result = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            tTDTOWValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                            If (startDate <> "" And endDate <> "") Then
                                If (DateValue(tTDTOWValue) < DateValue(startDate) Or DateValue(tTDTOWValue) > DateValue(endDate)) Then
                                    MsgBox "Date To in Row " & i & " should be within Return period " & Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value & " to " & Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                ElseIf Target.Previous <> "" And TestDate(Target.Previous) = True Then
                                    tTDTOPRWValue = CDate(Format(Target.Previous, "dd/mm/yyyy"))
                                    If DateValue(tTDTOWValue) < DateValue(tTDTOPRWValue) Then
                                        MsgBox "Date To in Row " & i & " should be same as or after Date From "
                                        Target.value = ""
                                        Target.Select
                                    End If
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                            End If
                        End If
                    End If
                    If (Target.value <> "") Then
                        dateString = Mid(Target.value, 3, 8)
                        mm = Mid(Target.value, 4, 2)
                        If (mm = 1 Or mm = 3 Or mm = 5 Or mm = 7 Or mm = 8 Or mm = 10 Or mm = 12) Then
                            lastDateToCompare = "31" & dateString
                        ElseIf (mm = 2) Then
                            yy = Mid(Target.value, 7, 4)
                            If (yy Mod 4 = 0) Then
                                lastDateToCompare = "29" & dateString
                            Else
                                lastDateToCompare = "28" & dateString
                            End If
                        Else
                            lastDateToCompare = "30" & dateString
                        End If
                        
                        For Each r In Range("ExemptCerti.ListW").Rows
                            For Each c In r.Cells
                                rowNum = c.row
                                tDFromDateCompare = ""
                                tDToDateCompare = ""
                                tDToDateCompare1 = ""
                                If (Worksheets("A_Basic_Info").Range("$B$" & rowNum).value <> "") Then
                                    tDFromDateCompare = CDate(Format(Worksheets("A_Basic_Info").Range("$B$" & rowNum).value, "dd/mm/yyyy"))
                                End If
                                If (Worksheets("A_Basic_Info").Range("$C$" & rowNum).value <> "") Then
                                    tDToDateCompare = CDate(Format(Worksheets("A_Basic_Info").Range("$C$" & rowNum).value, "dd/mm/yyyy"))
                                End If
                                If (Worksheets("A_Basic_Info").Range("$B$" & i).value <> "") Then
                                    tDFromDateCompare1 = CDate(Format(Worksheets("A_Basic_Info").Range("$B$" & i).value, "dd/mm/yyyy"))
                                End If
                                If c.row > endRow Then
                                    Exit For
                                End If
                                If c.row = Target.row Then
                                    'Skip comparison with self value
                                Else
                                    If Target.value <> "" And tDFromDateCompare <> "" And tDToDateCompare <> "" Then
                                        If (DateValue(Target.value) >= DateValue(tDFromDateCompare) And DateValue(Target.value) <= DateValue(tDToDateCompare)) Then
                                            MsgBox "Exemption Certificate Number for this period has been declared in row " & c.row
                                            Target.value = ""
                                            Target.Select
                                            GoTo endSubW
                                        End If
                                    End If
                                    If Target.value <> "" And tDFromDateCompare <> "" And tDToDateCompare <> "" And tDFromDateCompare1 <> "" Then
                                        If (DateValue(tDFromDateCompare1) < DateValue(tDFromDateCompare) And DateValue(Target.value) > DateValue(tDToDateCompare)) Then
                                            MsgBox "Exemption Certificate Number for this period has been declared in row " & c.row
                                            Target.value = ""
                                            Target.Select
                                            GoTo endSubW
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                'check Date valid Date format or not and check other validation date related date to should be between Return period End
endSubW:
            End Select
        Next
        'Code for Part 6 Wife  End
        
        Select Case Target.Address
            'base on selected bank name from combo bank id will be set in hidden field in sheet and bank branch field will be enable Start
            Case "$B$23"
                Dim bankName As String
                Dim selectedBankNameId As String
                bankName = Worksheets("A_Basic_Info").Range("BankS").value
                If (bankName = "") Then
                    Worksheets("A_Basic_Info").Range("BranchS").value = ""
                End If
                Target.Select
            'base on selected bank name from combo bank id will be set in hidden field in sheet and bank branch field will be enable End
        
            'base on selected branch name from combo branch id will be set in hidden field in sheet Start
            Case "$B$24"
                Target.Select
                'base on selected branch name from combo branch id will be set in hidden field in sheet End
        
            'AuditPIN(self and wife) will be convert in Uppar Case Start
            Case "$B$30", "$C$30"
                If Target.value <> "" Then
                    Target.value = UCase(Target.value)
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
                End If
            'AuditPIN(self and wife) will be convert in Uppar Case End
             
            'check Date format for Date of Audit certificate and date can not be future date start
            Case "$B$32", "$C$32"
                If (Target.value <> "") Then
                     result = TestDate(Target.value)
                     If (result = False) Then
                        MsgBox "Please enter date value in dd/mm/yyyy format."
                        Target.value = ""
                        Target.Select
                    Else
                        If TestBfrDate(Target.value) = False Then
                            MsgBox "Date cannot be a future date"
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                End If
            'check Date format for Date of Audit certificate and date can not be future date End
        End Select
    End If
    
    Worksheets("A_Basic_Info").Protect Password:=Pwd
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

'toggleSpouseFields function using for base on selected value from do you want to declare wife income then wife related section and field will be Enable/Disable Start
Public Sub toggleSpouseFields(lockCell As Boolean)
    
    Call lockUnlock_cell_rng("A_Basic_Info", "B18:B18", lockCell)
    Call lockUnlock_cell_rng("A_Basic_Info", "B19:B19", lockCell)
    Call lockUnlock_cell_rng("A_Basic_Info", "B20:B20", lockCell) 'Commented till TSS goes live
    
    Call lockUnlock_cell_rng("A_Basic_Info", "C30:C32", lockCell)
    
    'new wife section clear Start
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "C4:C5", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "D4:D5", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "E4:E5", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "F4:F5", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "G4:G5", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "H4:H5", lockCell)
    
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "C9:C11", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "D9:D11", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "E9:E11", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "F9:F11", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "G9:G11", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "H9:H11", lockCell)
    
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "C14:C15", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "D14:D15", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "E14:E15", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "F14:F15", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "G14:G15", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "H14:H15", lockCell)
    
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "C18:C19", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "D18:D19", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "E18:E19", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "F18:F19", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "G18:G19", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "H18:H19", lockCell)
    
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "C22:C27", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "D22:D27", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "E22:E27", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "F22:F27", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "G22:G27", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "H22:H27", lockCell)
    
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "C30:C32", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "D30:D32", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "E30:E32", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "F30:F32", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "G30:G32", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "H30:H32", lockCell)
    
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "C37:C46", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "D37:D46", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "E37:E46", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "F37:F46", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "G37:G46", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "H37:H46", lockCell)
    
    
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "C50:C76", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "D50:D77", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "E50:E76", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "F50:F76", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "G50:G76", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "H50:H76", lockCell)
    
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "C80:C83", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "D80:D83", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "E80:E83", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "F80:F83", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "G80:G83", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "H80:H83", lockCell)
    
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "C86:C94", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "D86:D94", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "E86:E94", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "F86:F94", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "G86:G94", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "H86:H94", lockCell)
    
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "C97:C102", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "D97:D102", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "E97:E102", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "F97:F102", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "G97:G102", lockCell)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "H97:H102", lockCell)
    
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "I108:I108", lockCell)
    
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "PLA.OtherExpensesListW", lockCell, 9)
    Call lockUnlock_cell_rng("B_Profit_Loss_Account_Wife", "PLA.OtherIncomeListW", lockCell, 9)
    
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "C4:C4", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "D4:D4", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "E4:E4", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "F4:F4", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "G4:G4", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "H4:H4", lockCell)
    
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "C9:C10", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "D9:D10", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "E9:E10", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "F9:F10", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "G9:G10", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "H9:H10", lockCell)
    
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "C12:C13", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "D12:D13", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "E12:E13", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "F12:F13", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "G12:G13", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "H12:H13", lockCell)
    
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "C15:C15", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "D15:D15", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "E15:E15", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "F15:F15", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "G15:G15", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "H15:H15", lockCell)
    
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "C17:C19", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "D17:D19", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "E17:E19", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "F17:F19", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "G17:G19", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "H17:H19", lockCell)
    
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "C25:C26", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "D25:D26", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "E25:E26", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "F25:F26", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "G25:G26", lockCell)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "H25:H26", lockCell)
    
    'Call lockUnlock_cell_rng("T_Income_Computation_Wife", "C31:C31", lockCell)
    'Call lockUnlock_cell_rng("T_Income_Computation_Wife", "D31:D31", lockCell)
    'Call lockUnlock_cell_rng("T_Income_Computation_Wife", "E31:E31", lockCell)
    'Call lockUnlock_cell_rng("T_Income_Computation_Wife", "F31:F31", lockCell)
    'Call lockUnlock_cell_rng("T_Income_Computation_Wife", "G31:G31", lockCell)
    'Call lockUnlock_cell_rng("T_Income_Computation_Wife", "H31:H31", lockCell)
    
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "TaxComp.OthExpListW", lockCell, 9)
    Call lockUnlock_cell_rng("T_Income_Computation_Wife", "TaxComp.OthDedListW", lockCell, 9)
    
    'new wife section clear End
    
    Call lockUnlock_cell_rng("A_Basic_Info", "RentPaid.RentListW", lockCell, 17)
    Call lockUnlock_cell_rng("A_Basic_Info", "RentalIncome.ListW", lockCell, 17)
    Call lockUnlock_cell_rng("A_Basic_Info", "ExemptCerti.ListW", True, 4)
    
    'Call lock_cells("A_Basic_Info", "RentPaid.RentListW", "M", True)
    'Call lock_cells("A_Basic_Info", "RentalIncome.ListW", "M", True)
    
    Call lockUnlock_cell_rng("C_Balance_Sheet", "D5:D11", lockCell)
    Call lockUnlock_cell_rng("C_Balance_Sheet", "D14:D18", lockCell)
    Call lockUnlock_cell_rng("C_Balance_Sheet", "D23:D26", lockCell)
    Call lockUnlock_cell_rng("C_Balance_Sheet", "D29:D31", lockCell)
    Call lockUnlock_cell_rng("C_Balance_Sheet", "D34:D35", lockCell)
    Call lockUnlock_cell_rng("C_Balance_Sheet", "D37:D37", lockCell)
    Call lockUnlock_cell_rng("C_Balance_Sheet", "D40:D43", lockCell)
    Call lockUnlock_cell_rng("C_Balance_Sheet", "D45:D45", lockCell)
    Call lockUnlock_cell_rng("C_Balance_Sheet", "D49:D53", lockCell)
    Call lockUnlock_cell_rng("C_Balance_Sheet", "D56:D58", lockCell)
    Call lockUnlock_cell_rng("C_Balance_Sheet", "D63:D68", lockCell)
    Call lockUnlock_cell_rng("C_Balance_Sheet", "D72:D75", lockCell)
    Call lockUnlock_cell_rng("C_Balance_Sheet", "D78:D81", lockCell)
    Call lockUnlock_cell_rng("C_Balance_Sheet", "D83:D83", lockCell)
    
    Call lockUnlock_cell_rng("D_Stock_Analysis", "Inv.QuantDtlsListW", lockCell, 7)
    Call lockUnlock_cell_rng("D_Stock_Analysis", "Inv.QuantDtlsListKIIW", lockCell, 10)
    
    Call lockUnlock_cell_rng("E1_IDA_CA", "IniAllPlanMach.ListPart1W", lockCell, 11)
    Call lockUnlock_cell_rng("E1_IDA_CA", "IniAllIBD.ListPart2W", lockCell, 3) '3 added by atul for skipping locked cells (3 is insignificant, added jst bcoz something was to be put)
    
    Call lock_cells("E1_IDA_CA", "IniAllIBD.ListPart2W", "D", lockCell)
    Call lock_cells("E1_IDA_CA", "IniAllIBD.ListPart2W", "E", lockCell)
    Call lock_cells("E1_IDA_CA", "IniAllIBD.ListPart2W", "G", lockCell)
    Call lock_cells("E1_IDA_CA", "IniAllIBD.ListPart2W", "H", lockCell)
    
    Call lockUnlock_cell_rng("E1_IDA_CA", "AgrLandDed.ListW", lockCell, 7) 'changed
    Call lockUnlock_cell_rng("E1_IDA_CA", "DeprIntengAst.ListW", lockCell, 5) ' changed
    
    Call lock_cells("E1_IDA_CA", "DeprIntengAst.ListW", "F", lockCell)
    Call lock_cells("E1_IDA_CA", "DeprIntengAst.ListW", "G", lockCell)
    
    Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "B17:B19", lockCell)
    Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "C17:C19", lockCell)
    Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "D17:D19", lockCell)
    Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "E17:E19", lockCell)
    
    Call lockUnlock_cell_rng("E2_CA_WTA_SLM", "WAT.ListBW", lockCell, 8)
    
    Call lockUnlock_cell_rng("F_Employment_Income", "EmpIncome.ListW", lockCell, 5)
    Call lock_cells("F_Employment_Income", "EmpIncome.ListW", "F", lockCell)
    Call lock_cells("F_Employment_Income", "EmpIncome.ListW", "G", lockCell)
    
    Call lockUnlock_cell_rng("G_Partnership_Income", "ProfitShare.ListW", lockCell, 9)
    
    Call lockUnlock_cell_rng("H_Estate_Trust_Income", "EstateTrust.ListW", lockCell, 5)
    
    Call lockUnlock_cell_rng("I_Computation_of_Car_Benefit", "CarBenefit.ListW", lockCell, 10)
    Call lock_cells("I_Computation_of_Car_Benefit", "CarBenefit.ListW", "G", True)
    
    Call lockUnlock_cell_rng("J_Computation_of_Mortgage", "MortgageIntDtls.ListW", lockCell, 7)
    
    Call lockUnlock_cell_rng("K_Home_Ownership_Saving_Plan", "HomeOwnershipSavingPlan.ListW", lockCell, 5)
    
    Call lockUnlock_cell_rng("L_Computation_of_Insu_Relief", "InsReliefDtls.ListW", lockCell, 11) 'changed
    Call lock_cells("L_Computation_of_Insu_Relief", "InsReliefDtls.ListW", "F", True)
    
    Call lockUnlock_cell_rng("M_Details_of_PAYE_Deducted", "PayeDed.ListW", lockCell, 6)
    Call lock_cells("M_Details_of_PAYE_Deducted", "PayeDed.ListW", "E", lockCell)
    
    Call lockUnlock_cell_rng("N_Installment_Tax_Credits", "InstallmentTax.ListW", lockCell, 4)
    
    'Call lockUnlock_cell_rng("O_WHT_Credits", "WithHolding.ListW", lockCell, 6)
    
    'Call lockUnlock_cell_rng("P_Advance_Tax_Credits", "VehicleAdvTaxPaid.ListW", lockCell, 5)
    'Call lockUnlock_cell_rng("P_Advance_Tax_Credits", "VehicleAdvTaxPaid.ListW", lockCell, 6)
    'Call lock_cells("P_Advance_Tax_Credits", "VehicleAdvTaxPaid.ListW", "E", True)
    
    Call lockUnlock_cell_rng("Q_IT_Payment_Credits", "DtlIncomePaid.IncomePaidAdvanceListW", lockCell, 4)
    Call lockUnlock_cell_rng("Q_IT_Payment_Credits", "DtlIncomePaid.IncomePaidSelfAssmntListW", lockCell, 4)
    
    Call lockUnlock_cell_rng("R_DTAA_Credits", "DTAACredits.DetailsW", lockCell, 6)
    
    Call lockUnlock_cell_rng("T_Tax_Computation", "D5:D5", lockCell)
    Call lockUnlock_cell_rng("T_Tax_Computation", "TaxComp.PersonalReliefW", lockCell)
    
    'new added code for reset Day and Rate in respective section Start
    Call lock_cells("E1_IDA_CA", "DeprIntengAst.ListW", "F", True)
    Call lock_cells("E1_IDA_CA", "IniAllIBD.ListPart2W", "E", True)
    'new added code for reset Day and Rate in respective section End
    'new added lock cell code for new added sch start
   If (Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value <> "") Then
    If (Worksheets("A_Basic_Info").Range("SecA.RtnYear").value < "2015") Then
                Call lockUnlock_cell_rng("S_Previous_Years_Losses", "C23:H27", True)
    Else
        Call lockUnlock_cell_rng("S_Previous_Years_Losses", "C28:H27", lockCell)
    End If
    End If
    Call lockUnlock_cell_rng("S_Previous_Years_Losses", "C23:H31", lockCell)
    Call lockUnlock_cell_rng("S_Previous_Years_Losses", "H23:H31", lockCell)
    Call lockUnlock_cell_rng("S_Previous_Years_Losses", "C33:C33", lockCell)
    Call lockUnlock_cell_rng("S_Previous_Years_Losses", "D33:D33", lockCell)
    Call lockUnlock_cell_rng("S_Previous_Years_Losses", "E33:E33", lockCell)
    Call lockUnlock_cell_rng("S_Previous_Years_Losses", "F33:F33", lockCell)
    Call lockUnlock_cell_rng("S_Previous_Years_Losses", "G33:G33", lockCell)
    Call lockUnlock_cell_rng("S_Previous_Years_Losses", "H33:H33", lockCell)
    'new added lock cell code for new added sch End
    'reset Code in Wife section Start
    Call lock_cellsHidden("A_Basic_Info", "RentPaid.RentListW", "R", True)
    Call lock_cellsHidden("A_Basic_Info", "RentPaid.RentListW", "T", True)
    Call lock_cellsHidden("A_Basic_Info", "RentPaid.RentListW", "U", True)
    
    Call lock_cellsHidden("A_Basic_Info", "RentalIncome.ListW", "R", True)
    Call lock_cellsHidden("A_Basic_Info", "RentalIncome.ListW", "T", True)
    Call lock_cellsHidden("A_Basic_Info", "RentalIncome.ListW", "U", True)
    
    Call lock_cellsHidden("I_Computation_of_Car_Benefit", "CarBenefit.ListW", "L", True)
    Call lock_cellsHidden("I_Computation_of_Car_Benefit", "CarBenefit.ListW", "M", True)
    
    If lockCell = False Then
        Call lock_cells("E1_IDA_CA", "IniAllIBD.ListPart2W", "F", False)
        Call lock_cells("E1_IDA_CA", "IniAllIBD.ListPart2W", "G", False)
        Call lock_cells("E1_IDA_CA", "IniAllIBD.ListPart2W", "I", False)
        Call lock_cells("E1_IDA_CA", "IniAllIBD.ListPart2W", "J", False)
         Call lock_cells("E1_IDA_CA", "IniAllIBD.ListPart2W", "H", True)
        
        Call lock_cells("E1_IDA_CA", "DeprIntengAst.ListW", "I", True) 'added
        Call lock_cells("E1_IDA_CA", "DeprIntengAst.ListW", "J", True) 'added
        Call lock_cells("E1_IDA_CA", "DeprIntengAst.ListW", "C", True) 'added
        Call lock_cells("E1_IDA_CA", "DeprIntengAst.ListW", "G", False) 'added
        Call lock_cells("E1_IDA_CA", "DeprIntengAst.ListW", "H", False) 'added
        Call lock_cells("E1_IDA_CA", "DeprIntengAst.ListW", "E", False)
    Else
         Call lock_cells("E1_IDA_CA", "IniAllIBD.ListPart2W", "C", True)
        Call lock_cells("E1_IDA_CA", "IniAllIBD.ListPart2W", "F", True)
        Call lock_cells("E1_IDA_CA", "IniAllIBD.ListPart2W", "I", True)
        Call lock_cells("E1_IDA_CA", "IniAllIBD.ListPart2W", "J", True)
        
        Call lock_cells("E1_IDA_CA", "DeprIntengAst.ListW", "E", True) 'added
        Call lock_cells("E1_IDA_CA", "DeprIntengAst.ListW", "H", True) 'added
    End If
    'Call lock_cellsHidden("P_Advance_Tax_Credits", "VehicleAdvTaxPaid.ListW", "H", True)
    
    Call lock_cellsHidden("G_Partnership_Income", "ProfitShare.ListW", "K", True)
    'reset Code in Wife section End
End Sub
'toggleSpouseFields function using for base on selected value from do you want to declare wife income then wife related section and field will be Enable/Disable End

Public Sub toggleSpouseFieldsWithoutClearingContent(lockCell As Boolean)

    Call lockUnlock_cell_rng_without_clearing_contents("A_Basic_Info", "B18:B18", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("A_Basic_Info", "B19:B19", lockCell)
    'Call lockUnlock_cell_rng_without_clearing_contents("A_Basic_Info", "B20:B20", lockCell) Commented till TSS goes live
    
    'new wife section clear Start
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "C4:C5", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "D4:D5", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "E4:E5", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "F4:F5", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "G4:G5", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "H4:H5", lockCell)
    
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "C9:C11", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "D9:D11", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "E9:E11", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "F9:F11", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "G9:G11", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "H9:H11", lockCell)
    
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "C14:C15", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "D14:D15", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "E14:E15", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "F14:F15", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "G14:G15", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "H14:H15", lockCell)
    
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "C18:C19", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "D18:D19", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "E18:E19", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "F18:F19", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "G18:G19", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "H18:H19", lockCell)
    
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "C22:C27", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "D22:D27", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "E22:E27", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "F22:F27", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "G22:G27", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "H22:H27", lockCell)
    
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "C30:C32", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "D30:D32", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "E30:E32", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "F30:F32", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "G30:G32", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "H30:H32", lockCell)
    
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "C37:C46", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "D37:D46", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "E37:E46", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "F37:F46", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "G37:G46", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "H37:H46", lockCell)
    
    
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "C50:C76", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "D50:D77", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "E50:E76", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "F50:F76", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "G50:G76", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "H50:H76", lockCell)
    
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "C80:C83", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "D80:D83", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "E80:E83", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "F80:F83", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "G80:G83", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "H80:H83", lockCell)
    
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "C86:C94", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "D86:D94", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "E86:E94", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "F86:F94", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "G86:G94", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "H86:H94", lockCell)
    
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "C97:C102", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "D97:D102", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "E97:E102", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "F97:F102", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "G97:G102", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "H97:H102", lockCell)
    
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "I108:I108", lockCell)
    
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "PLA.OtherExpensesListW", lockCell, 9)
    Call lockUnlock_cell_rng_without_clearing_contents("B_Profit_Loss_Account_Wife", "PLA.OtherIncomeListW", lockCell, 9)
    
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "C4:C4", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "D4:D4", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "E4:E4", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "F4:F4", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "G4:G4", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "H4:H4", lockCell)
    
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "C9:C10", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "D9:D10", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "E9:E10", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "F9:F10", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "G9:G10", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "H9:H10", lockCell)
    
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "C12:C13", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "D12:D13", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "E12:E13", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "F12:F13", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "G12:G13", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "H12:H13", lockCell)
    
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "C15:C15", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "D15:D15", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "E15:E15", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "F15:F15", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "G15:G15", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "H15:H15", lockCell)
    
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "C17:C19", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "D17:D19", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "E17:E19", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "F17:F19", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "G17:G19", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "H17:H19", lockCell)
    
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "C25:C26", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "D25:D26", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "E25:E26", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "F25:F26", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "G25:G26", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "H25:H26", lockCell)
    
    'Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "C31:C31", lockCell)
    'Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "D31:D31", lockCell)
    'Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "E31:E31", lockCell)
    'Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "F31:F31", lockCell)
    'Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "G31:G31", lockCell)
    'Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "H31:H31", lockCell)
    
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "TaxComp.OthExpListW", lockCell, 9)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Income_Computation_Wife", "TaxComp.OthDedListW", lockCell, 9)
    
    'new wife section clear End
    
    Call lockUnlock_cell_rng_without_clearing_contents("A_Basic_Info", "RentPaid.RentListW", lockCell, 17)
    Call lockUnlock_cell_rng_without_clearing_contents("A_Basic_Info", "RentalIncome.ListW", lockCell, 17)
    Call lockUnlock_cell_rng_without_clearing_contents("A_Basic_Info", "C30:C32", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("A_Basic_Info", "ExemptCerti.ListW", True, 4)
    
    'Call lock_cells("A_Basic_Info", "RentPaid.RentListW", "M", True, False)
    'Call lock_cells("A_Basic_Info", "RentalIncome.ListW", "M", True, False)
    
    Call lockUnlock_cell_rng_without_clearing_contents("C_Balance_Sheet", "D5:D11", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("C_Balance_Sheet", "D14:D18", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("C_Balance_Sheet", "D23:D26", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("C_Balance_Sheet", "D29:D31", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("C_Balance_Sheet", "D34:D35", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("C_Balance_Sheet", "D37:D37", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("C_Balance_Sheet", "D40:D43", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("C_Balance_Sheet", "D45:D45", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("C_Balance_Sheet", "D49:D53", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("C_Balance_Sheet", "D56:D58", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("C_Balance_Sheet", "D63:D68", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("C_Balance_Sheet", "D72:D75", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("C_Balance_Sheet", "D78:D81", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("C_Balance_Sheet", "D83:D83", lockCell)
    
    Call lockUnlock_cell_rng_without_clearing_contents("D_Stock_Analysis", "Inv.QuantDtlsListW", lockCell, 7)
    Call lockUnlock_cell_rng_without_clearing_contents("D_Stock_Analysis", "Inv.QuantDtlsListKIIW", lockCell, 10)
    
    Call lockUnlock_cell_rng_without_clearing_contents("E1_IDA_CA", "IniAllPlanMach.ListPart1W", lockCell, 9)
    Call lockUnlock_cell_rng_without_clearing_contents("E1_IDA_CA", "IniAllIBD.ListPart2W", lockCell, 3)
    
    Call lock_cells("E1_IDA_CA", "IniAllIBD.ListPart2W", "D", lockCell, False)
    Call lock_cells("E1_IDA_CA", "IniAllIBD.ListPart2W", "E", lockCell, False)
    Call lock_cells("E1_IDA_CA", "IniAllIBD.ListPart2W", "G", lockCell, False)
    Call lock_cells("E1_IDA_CA", "IniAllIBD.ListPart2W", "H", lockCell, False)
    
    Call lockUnlock_cell_rng_without_clearing_contents("E1_IDA_CA", "AgrLandDed.ListW", lockCell, 7) 'changed
    Call lockUnlock_cell_rng_without_clearing_contents("E1_IDA_CA", "DeprIntengAst.ListW", lockCell, 5) ' changed
    
    Call lock_cells("E1_IDA_CA", "DeprIntengAst.ListW", "F", lockCell, False)
    Call lock_cells("E1_IDA_CA", "DeprIntengAst.ListW", "G", lockCell, False)
    
    Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_WDV", "B17:B19", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_WDV", "C17:C19", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_WDV", "D17:D19", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_WDV", "E17:E19", lockCell)
    
    Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_SLM", "WAT.ListBW", lockCell, 8)
    
    Call lockUnlock_cell_rng_without_clearing_contents("F_Employment_Income", "EmpIncome.ListW", lockCell, 5)
    Call lock_cells("F_Employment_Income", "EmpIncome.ListW", "F", lockCell, False)
    Call lock_cells("F_Employment_Income", "EmpIncome.ListW", "G", lockCell, False)
    
    Call lockUnlock_cell_rng_without_clearing_contents("G_Partnership_Income", "ProfitShare.ListW", lockCell, 9)
    
    Call lockUnlock_cell_rng_without_clearing_contents("H_Estate_Trust_Income", "EstateTrust.ListW", lockCell, 5)
    
    Call lockUnlock_cell_rng_without_clearing_contents("I_Computation_of_Car_Benefit", "CarBenefit.ListW", lockCell, 10)
    Call lock_cells("I_Computation_of_Car_Benefit", "CarBenefit.ListW", "G", True, False)
    
    Call lockUnlock_cell_rng_without_clearing_contents("J_Computation_of_Mortgage", "MortgageIntDtls.ListW", lockCell, 7)
    
    Call lockUnlock_cell_rng_without_clearing_contents("K_Home_Ownership_Saving_Plan", "HomeOwnershipSavingPlan.ListW", lockCell, 5)
    
    Call lockUnlock_cell_rng_without_clearing_contents("L_Computation_of_Insu_Relief", "InsReliefDtls.ListW", lockCell, 11) 'changed
    Call lock_cells("L_Computation_of_Insu_Relief", "InsReliefDtls.ListW", "F", True, False)
    
    Call lockUnlock_cell_rng_without_clearing_contents("M_Details_of_PAYE_Deducted", "PayeDed.ListW", lockCell, 6)
    Call lock_cells("M_Details_of_PAYE_Deducted", "PayeDed.ListW", "E", lockCell, False)
    
    Call lockUnlock_cell_rng_without_clearing_contents("N_Installment_Tax_Credits", "InstallmentTax.ListW", lockCell, 4)
    
    'Call lockUnlock_cell_rng_without_clearing_contents("O_WHT_Credits", "WithHolding.ListW", lockCell, 6)
    
    'Call lockUnlock_cell_rng_without_clearing_contents("P_Advance_Tax_Credits", "VehicleAdvTaxPaid.ListW", lockCell, 5)
    'Call lockUnlock_cell_rng_without_clearing_contents("P_Advance_Tax_Credits", "VehicleAdvTaxPaid.ListW", lockCell, 6)
    'Call lock_cells("P_Advance_Tax_Credits", "VehicleAdvTaxPaid.ListW", "E", True, False)
    
    Call lockUnlock_cell_rng_without_clearing_contents("Q_IT_Payment_Credits", "DtlIncomePaid.IncomePaidAdvanceListW", lockCell, 4)
    Call lockUnlock_cell_rng_without_clearing_contents("Q_IT_Payment_Credits", "DtlIncomePaid.IncomePaidSelfAssmntListW", lockCell, 4)
    
    Call lockUnlock_cell_rng_without_clearing_contents("R_DTAA_Credits", "DTAACredits.DetailsW", lockCell, 5)
    
    Call lockUnlock_cell_rng_without_clearing_contents("T_Tax_Computation", "D5:D5", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("T_Tax_Computation", "D15:D15", lockCell)
    
    'new added code for reset Day and Rate in respective section Start
    Call lock_cells("E1_IDA_CA", "DeprIntengAst.ListW", "E", True, False)
    Call lock_cells("E1_IDA_CA", "IniAllIBD.ListPart2W", "C", True, False)
    'new added code for reset Day and Rate in respective section End
    'new added lock cell code for new added sch start
    Call lockUnlock_cell_rng_without_clearing_contents("S_Previous_Years_Losses", "C23:C31", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("S_Previous_Years_Losses", "D23:D31", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("S_Previous_Years_Losses", "E23:E31", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("S_Previous_Years_Losses", "F23:F31", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("S_Previous_Years_Losses", "G23:G31", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("S_Previous_Years_Losses", "H23:H31", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("S_Previous_Years_Losses", "C33:C33", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("S_Previous_Years_Losses", "D33:D33", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("S_Previous_Years_Losses", "E33:E33", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("S_Previous_Years_Losses", "F33:F33", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("S_Previous_Years_Losses", "G33:G33", lockCell)
    Call lockUnlock_cell_rng_without_clearing_contents("S_Previous_Years_Losses", "H33:H33", lockCell)
    'new added lock cell code for new added sch End
    'reset Code in Wife section Start
    Call lock_cellsHidden("A_Basic_Info", "RentPaid.RentListW", "R", True, False)
    Call lock_cellsHidden("A_Basic_Info", "RentPaid.RentListW", "T", True, False)
    Call lock_cellsHidden("A_Basic_Info", "RentPaid.RentListW", "U", True, False)
    
    Call lock_cellsHidden("A_Basic_Info", "RentalIncome.ListW", "R", True, False)
    Call lock_cellsHidden("A_Basic_Info", "RentalIncome.ListW", "T", True, False)
    Call lock_cellsHidden("A_Basic_Info", "RentalIncome.ListW", "U", True, False)
    
    Call lock_cellsHidden("I_Computation_of_Car_Benefit", "CarBenefit.ListW", "L", True, False)
    Call lock_cellsHidden("I_Computation_of_Car_Benefit", "CarBenefit.ListW", "M", True, False)
    
    'Call lock_cellsHidden("P_Advance_Tax_Credits", "VehicleAdvTaxPaid.ListW", "H", True, False)
    
    Call lock_cellsHidden("G_Partnership_Income", "ProfitShare.ListW", "K", True, False)
    'reset Code in Wife section End
End Sub




Public Sub getyear(year As String)

                'Added by lawrence 22122020 for the covid tax rates affecting IT1
                'Declare an array of sheets that have the rows to be hidden or unhidden
                Dim workSheetArray() As Variant
                'Initilize the array with the data of the sheets.
                'This variable at this sub only has all the sheets with fields to be hidden,
                'including those without wife income
                workSheetArray = Array("F_Employment_Income", "T_Tax_Computation", "J_Computation_of_Mortgage", "K_Home_Ownership_Saving_Plan")

                If (year <= 2018) Then
                    Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
                    Worksheets("T_Tax_Computation").Activate
                    Worksheets("T_Tax_Computation").Range("B7").value = "Deposit in Home Ownership Saving Plan (Total of ""Total Deposit for the Year"" from K_Home_Ownership_Saving_Plan but subject to Maximum 48,000 Ksh)"
                    Worksheets("T_Tax_Computation").Protect (Pwd)
                    Worksheets("A_Basic_Info").Activate
                End If
                If (year = 2019) Then
                    Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
                    Worksheets("T_Tax_Computation").Activate
                    Worksheets("T_Tax_Computation").Range("B7").value = "Deposit in Home Ownership Saving Plan (Total of ""Total Deposit for the Year"" from K_Home_Ownership_Saving_Plan but subject to Maximum 72,000 Ksh)"
                    Worksheets("T_Tax_Computation").Protect (Pwd)
                    Worksheets("A_Basic_Info").Activate
                End If
                If (year > 2019) Then
                    Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
                    Worksheets("T_Tax_Computation").Activate
                    Worksheets("T_Tax_Computation").Range("B7").value = "Deposit in Home Ownership Saving Plan (Total of ""Total Deposit for the Year"" from K_Home_Ownership_Saving_Plan but subject to Maximum 96,000 Ksh)"
                    Worksheets("T_Tax_Computation").Protect (Pwd)
                    Worksheets("A_Basic_Info").Activate
                End If
                
                'Added by Lawrence 21/12/2020 for the covid tax rates affecting IT1
                'Check the year of income
                'Pass all sheets with fields to be hidden,
                'pass true for fields to be hidden or false for the fields to be unhidden and
                'specify the trigger cell in this case cell B5
                If (year <> 2020) Then
                    Call hideRows(workSheetArray, True, "B5", year)
                End If
                If (year = 2020 Or year = 2023) Then
                    Call hideRows(workSheetArray, False, "B5", year)
                End If
End Sub
'Added by Lawrence 21/12/2020 for the covid tax rates affecting IT1
Public Sub hideRows(workSheetArray() As Variant, hide As Boolean, callerCell As String, year As String)
                   'Declare varibales to hold the first and last rows to be hidden/unhidden
                   Dim firstRowSelf As String
                   Dim lastRowSelf As String
                   Dim firstRowWife As String
                   Dim lastRowWife As String
                   'Loop through the worksheets
                   For i = LBound(workSheetArray) To UBound(workSheetArray)
                        Worksheets(workSheetArray(i)).Unprotect Password:=Pwd
                        Worksheets(workSheetArray(i)).Activate
                        'When worksheet is F_Employment_Income call hideRowsSheetsWithWifeIncome
                        'Pass the sheet, hide or unhide, callerCell(either B5 or B17),
                        'rows to be deactivated for self and rows to be deactivated for wife
                        If (StrComp(workSheetArray(i), "F_Employment_Income") = 0) Then
                            firstRowSelf = Worksheets(workSheetArray(i)).Range("EmpIncome.ListSTOJANMAR").row
                            lastRowSelf = Worksheets(workSheetArray(i)).Range("EmpIncome.ListSPENSAPRDEC").row
                            firstRowWife = Worksheets(workSheetArray(i)).Range("EmpIncome.ListWTOJANMAR").row
                            lastRowWife = Worksheets(workSheetArray(i)).Range("EmpIncome.ListWPENAPRDEC").row
                            Call hideRowsSheetsWithWifeIncome(workSheetArray(i), hide, callerCell, firstRowSelf & ":" & lastRowSelf, firstRowWife & ":" & lastRowWife)
                            
                          
                            'Call hideRowsSheetsWithWifeIncome(workSheetArray(i), hide, callerCell, "6:9", "18:21")
                         'When worksheet is T_Tax_Computation hide/unhide rows 16 and 17
                        ElseIf (StrComp(workSheetArray(i), "T_Tax_Computation") = 0) Then
                            firstRowSelf = Worksheets(workSheetArray(i)).Range("TaxComp.TaxPayableSJanMar").row
                            lastRowSelf = Worksheets(workSheetArray(i)).Range("TaxComp.TaxPayableSAprDec").row
                            Worksheets(workSheetArray(i)).Rows(firstRowSelf & ":" & lastRowSelf).Hidden = hide
                            
                
                        'When worksheet is J_Computation_of_Mortgage call hideRowsSheetsWithWifeIncome
                        'Pass the sheet, hide or unhide, callerCell(either B5 or B17),
                        'rows to be deactivated for self and rows to be deactivated for wife
                        ElseIf (StrComp(workSheetArray(i), "J_Computation_of_Mortgage") = 0) Then
                            firstRowSelf = Worksheets(workSheetArray(i)).Range("MortgageIntDtls.ListSTOJANMAR").row
                            lastRowSelf = Worksheets(workSheetArray(i)).Range("MortgageIntDtls.ListSTOAPRDEC").row
                            firstRowWife = Worksheets(workSheetArray(i)).Range("MortgageIntDtls.ListWTOJANMAR").row
                            lastRowWife = Worksheets(workSheetArray(i)).Range("MortgageIntDtls.ListWTOAPRDEC").row
                            Call hideRowsSheetsWithWifeIncome(workSheetArray(i), hide, callerCell, firstRowSelf & ":" & lastRowSelf, firstRowWife & ":" & lastRowWife)
                
                           
                           ElseIf (StrComp(workSheetArray(i), "K_Home_Ownership_Saving_Plan") = 0) Then
                            firstRowSelf = Worksheets(workSheetArray(i)).Range("HomeOwnershipSavingPlan.ListSTOJANMAR").row
                            lastRowSelf = Worksheets(workSheetArray(i)).Range("HomeOwnershipSavingPlan.ListSTOAPRDEC").row
                            firstRowWife = Worksheets(workSheetArray(i)).Range("HomeOwnershipSavingPlan.ListWTOJANMAR").row
                            lastRowWife = Worksheets(workSheetArray(i)).Range("HomeOwnershipSavingPlan.ListWTOAPRDEC").row
                            
                            Call hideRowsSheetsWithWifeIncome(workSheetArray(i), hide, callerCell, firstRowSelf & ":" & lastRowSelf, firstRowWife & ":" & lastRowWife)
          
                           End If

                        Worksheets(workSheetArray(i)).Protect (Pwd)
                   Next i
                   
                  'Added for 2023 finance act changes on 04/12/2023
                   If (year = 2023) Then
                        Worksheets("F_Employment_Income").Unprotect (Pwd)
                        Worksheets("F_Employment_Income").Range("EmpIncome.ListSTOJANMARLbl").value = "January - June:  Income"
                        Worksheets("F_Employment_Income").Range("EmpIncome.ListSPENSJANMARLbl").value = "January - June:  Pension"
                        Worksheets("F_Employment_Income").Range("EmpIncome.ListSTOAPRDECLbl").value = "July - December:  Income"
                        Worksheets("F_Employment_Income").Range("EmpIncome.ListSPENSAPRDECLbl").value = "July - December:  Pension"
                        Worksheets("F_Employment_Income").Protect (Pwd)
                        
                        Worksheets("J_Computation_of_Mortgage").Unprotect (Pwd)
                        Worksheets("J_Computation_of_Mortgage").Range("MortgageIntDtls.ListSTOJANMARLBL").value = "January - June:  Interest"
                        Worksheets("J_Computation_of_Mortgage").Range("MortgageIntDtls.ListSTOAPRDECLBL").value = "July - December:  Interest"
                        Worksheets("J_Computation_of_Mortgage").Protect (Pwd)
                        
                        Worksheets("T_Tax_Computation").Unprotect (Pwd)
                        Worksheets("T_Tax_Computation").Range("TaxComp.TaxPayableSJanMarLbl").value = "Tax on Taxable Employment Income January -June Period"
                        Worksheets("T_Tax_Computation").Range("TaxComp.TaxPayableSAprDecLbl").value = "Tax on Taxable Employment Income July -December Period"
                        Worksheets("T_Tax_Computation").Protect (Pwd)
                   ElseIf (year = 2020) Then
                        Worksheets("F_Employment_Income").Unprotect (Pwd)
                        Worksheets("F_Employment_Income").Range("EmpIncome.ListSTOJANMARLbl").value = "January - March:  Income"
                        Worksheets("F_Employment_Income").Range("EmpIncome.ListSPENSJANMARLbl").value = "January - March:  Pension"
                        Worksheets("F_Employment_Income").Range("EmpIncome.ListSTOAPRDECLbl").value = "April - December:  Income"
                        Worksheets("F_Employment_Income").Range("EmpIncome.ListSPENSAPRDECLbl").value = "April - December:  Pension"
                        Worksheets("F_Employment_Income").Protect (Pwd)
                        
                        Worksheets("J_Computation_of_Mortgage").Unprotect (Pwd)
                        Worksheets("J_Computation_of_Mortgage").Range("MortgageIntDtls.ListSTOJANMARLBL").value = "January - March:  Interest"
                        Worksheets("J_Computation_of_Mortgage").Range("MortgageIntDtls.ListSTOAPRDECLBL").value = "April - December:  Interest"
                        Worksheets("J_Computation_of_Mortgage").Protect (Pwd)
                        
                        Worksheets("T_Tax_Computation").Unprotect (Pwd)
                        Worksheets("T_Tax_Computation").Range("TaxComp.TaxPayableSJanMarLbl").value = "Tax on Taxable Employment Income January -March Period"
                        Worksheets("T_Tax_Computation").Range("TaxComp.TaxPayableSAprDecLbl").value = "Tax on Taxable Employment Income April -December Period"
                        Worksheets("T_Tax_Computation").Protect (Pwd)
                   End If
                   'End 04/12/2023

                    Worksheets("A_Basic_Info").Activate
End Sub
'Added by Lawrence 23/12/2020 for the covid tax rates affecting IT1
Public Sub hideRowsSheetsWithWifeIncome(workingWorkSheet As Variant, hide As Boolean, callerCell As String, disableRows As String, disableRowsWife As String)
                'When the change was made in cell B5
                If callerCell = "B5" Then
                'Hide/unhide rows for self
                Worksheets(workingWorkSheet).Rows(disableRows).Hidden = hide
                If (hide) Then
                    Call resetValues(workingWorkSheet, "self")
                End If
                'If the year is 2020, check if wife inconme is set to yes/no
                    If Not hide Then
                            If (Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value = "No") Then
                                'When cell B5 is No hide the rows related to wife income
                                Worksheets(workingWorkSheet).Rows(disableRowsWife).Hidden = True
                                Call resetValues(workingWorkSheet, "wife")
                            End If
                            If (Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value = "Yes") Then
                                'When cell B5 is Yes hide/unhide the rows related to wife income
                                'Hide is passed from the caller method depending on year
                                Worksheets(workingWorkSheet).Rows(disableRowsWife).Hidden = hide
                                Call resetValues(workingWorkSheet, "wife")
                            End If
                     Else
                        Worksheets(workingWorkSheet).Unprotect Password:=Pwd
                        Worksheets(workingWorkSheet).Activate
                        Worksheets(workingWorkSheet).Rows(disableRowsWife).Hidden = hide
                        Worksheets(workingWorkSheet).Protect Password:=Pwd
                        Worksheets("A_Basic_Info").Activate
                     End If
                ElseIf callerCell = "B17" Then
                'When the change was made in cell B17 hide/unhide the rows related to wife income
                'Hide is passed from the caller method depending on year
                Worksheets(workingWorkSheet).Rows(disableRowsWife).Hidden = hide
                Call resetValues(workingWorkSheet, "wife")
                End If
End Sub

Public Sub resetValues(workingWorkSheet As Variant, ByVal caller As String)
    Worksheets(workingWorkSheet).Unprotect Password:=Pwd
    Worksheets(workingWorkSheet).Activate
    If (caller = "self") Then
    
       If (workingWorkSheet = "F_Employment_Income") Then
            Worksheets("F_Employment_Income").Range("EmpIncome.ListSTOJANMAR").value = 0#
            Worksheets("F_Employment_Income").Range("EmpIncome.ListSTOAPRDEC").value = 0#
            Worksheets("F_Employment_Income").Range("EmpIncome.ListSPENSAPRDEC").value = 0#
            Worksheets("F_Employment_Income").Range("EmpIncome.ListSPENSJANMAR").value = 0#
        End If
        If (workingWorkSheet = "T_Tax_Computation") Then
            Worksheets("T_Tax_Computation").Range("TaxComp.TaxPayableSJanMar").value = 0#
            Worksheets("T_Tax_Computation").Range("TaxComp.TaxPayableSAprDec").value = 0#
        End If
        If (workingWorkSheet = "J_Computation_of_Mortgage") Then
            Worksheets("J_Computation_of_Mortgage").Range("MortgageIntDtls.ListSTOJANMAR").value = 0#
            Worksheets("J_Computation_of_Mortgage").Range("MortgageIntDtls.ListSTOAPRDEC").value = 0#
        End If
        If (workingWorkSheet = "K_Home_Ownership_Saving_Plan") Then
            Worksheets("K_Home_Ownership_Saving_Plan").Range("HomeOwnershipSavingPlan.ListSTOJANMAR").value = 0#
            Worksheets("K_Home_Ownership_Saving_Plan").Range("HomeOwnershipSavingPlan.ListSTOAPRDEC").value = 0#
        End If
    ElseIf (caller = "wife") Then
        If (workingWorkSheet = "F_Employment_Income") Then
            Worksheets("F_Employment_Income").Range("EmpIncome.ListWTOJANMAR").value = 0#
            Worksheets("F_Employment_Income").Range("EmpIncome.ListWTOAPRDEC").value = 0#
            Worksheets("F_Employment_Income").Range("EmpIncome.ListWPENAPRDEC").value = 0#
            Worksheets("F_Employment_Income").Range("EmpIncome.ListWPENJANMAR").value = 0#
        End If
        If (workingWorkSheet = "T_Tax_Computation") Then
            Worksheets("T_Tax_Computation").Range("TaxComp.TaxPayableWJanMar").value = 0#
            Worksheets("T_Tax_Computation").Range("TaxComp.TaxPayableWAprDec").value = 0#
        End If
        If (workingWorkSheet = "J_Computation_of_Mortgage") Then
            Worksheets("J_Computation_of_Mortgage").Range("MortgageIntDtls.ListWTOJANMAR").value = 0#
            Worksheets("J_Computation_of_Mortgage").Range("MortgageIntDtls.ListWTOAPRDEC").value = 0#
        End If
        If (workingWorkSheet = "K_Home_Ownership_Saving_Plan") Then
            Worksheets("K_Home_Ownership_Saving_Plan").Range("HomeOwnershipSavingPlan.ListWTOJANMAR").value = 0#
            Worksheets("K_Home_Ownership_Saving_Plan").Range("HomeOwnershipSavingPlan.ListWTOAPRDEC").value = 0#
        End If
    End If
Worksheets(workingWorkSheet).Protect Password:=Pwd
Worksheets("A_Basic_Info").Activate

End Sub

Public Sub getyear2(year As String)
If (year <= 2016) Then
                    Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
                    Worksheets("T_Tax_Computation").Activate
                    Worksheets("T_Tax_Computation").Range("TaxComp.Mortgage").value = "Mortgage Interest  (Total of Amount of Interest Paid from J_Computation_of_Mortgage  but subject to Maximum 150,000 Ksh)"
                    Worksheets("T_Tax_Computation").Range("TaxComp.PersonalReliefS").value = 0#
                    Worksheets("T_Tax_Computation").Range("TaxComp.PersonalReliefW").value = 0#
                    Worksheets("T_Tax_Computation").Protect (Pwd)
                    Worksheets("A_Basic_Info").Activate
                End If
                If (year = 2017) Then
                    Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
                    Worksheets("T_Tax_Computation").Activate
                    Worksheets("T_Tax_Computation").Range("TaxComp.Mortgage").value = "Mortgage Interest  (Total of Amount of Interest Paid from J_Computation_of_Mortgage  but subject to Maximum 300,000 Ksh)"
                    Worksheets("T_Tax_Computation").Range("TaxComp.PersonalReliefS").value = 0#
                    Worksheets("T_Tax_Computation").Range("TaxComp.PersonalReliefW").value = 0#
                    Worksheets("T_Tax_Computation").Protect (Pwd)
                    Worksheets("A_Basic_Info").Activate
                End If
End Sub
Public Sub enabledisable(mm As String, year As String)
Dim bizIncome As String
bizIncome = Worksheets("A_Basic_Info").Range("RetInf.DeclareSelfBusIncome").value
    If mm > 3 And year = 2020 Or year > 2020 Then
        ActiveWorkbook.Unprotect (Pwd)
        Worksheets("E2_CA_WTA_SLM").Unprotect (Pwd)
        Worksheets("E2_CA_WTA_SLM").Visible = xlSheetHidden
        Call lockUnlock_cell_rng("E2_CA_WTA_SLM", "WAT.ListBS", True, "10")
        Call lockUnlock_cell_rng("E2_CA_WTA_SLM", "WAT.ListBW", True, "10")
        Worksheets("E2_CA_WTA_SLM").Protect (Pwd)
    Else
        If (bizIncome = "Yes") Then
            ActiveWorkbook.Unprotect (Pwd)
            Worksheets("E2_CA_WTA_SLM").Unprotect (Pwd)
            Worksheets("E2_CA_WTA_SLM").Visible = xlSheetVisible
            Call lockUnlock_cell_rng("E2_CA_WTA_SLM", "WAT.ListBS", False, "10")
            If Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value = "Yes" Then
             Call lockUnlock_cell_rng("E2_CA_WTA_SLM", "WAT.ListBW", False, "10")
            End If
            Worksheets("E2_CA_WTA_SLM").Protect (Pwd)
        End If

    End If
End Sub

Public Sub setmonth(mm As String)

                    If (mm = "01") Then
                        Worksheets("A_Basic_Info").Range("RetInf.MonthDesc").value = "January"
                    ElseIf (mm = "02") Then
                        Worksheets("A_Basic_Info").Range("RetInf.MonthDesc").value = "February"
                    ElseIf (mm = "03") Then
                        Worksheets("A_Basic_Info").Range("RetInf.MonthDesc").value = "March"
                    ElseIf (mm = "04") Then
                        Worksheets("A_Basic_Info").Range("RetInf.MonthDesc").value = "April"
                    ElseIf (mm = "05") Then
                        Worksheets("A_Basic_Info").Range("RetInf.MonthDesc").value = "May"
                    ElseIf (mm = "06") Then
                        Worksheets("A_Basic_Info").Range("RetInf.MonthDesc").value = "June"
                    ElseIf (mm = "07") Then
                        Worksheets("A_Basic_Info").Range("RetInf.MonthDesc").value = "July"
                    ElseIf (mm = "08") Then
                        Worksheets("A_Basic_Info").Range("RetInf.MonthDesc").value = "August"
                    ElseIf (mm = "09") Then
                        Worksheets("A_Basic_Info").Range("RetInf.MonthDesc").value = "September"
                    ElseIf (mm = "10") Then
                        Worksheets("A_Basic_Info").Range("RetInf.MonthDesc").value = "October"
                    ElseIf (mm = "11") Then
                        Worksheets("A_Basic_Info").Range("RetInf.MonthDesc").value = "November"
                    ElseIf (mm = "12") Then
                        Worksheets("A_Basic_Info").Range("RetInf.MonthDesc").value = "December"
                    End If
End Sub

Public Sub b17(ByVal cellB17 As String, workSheetArray() As Variant, ByVal year As String, ByVal mm As String)
                    
If UCase(Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value) = "YES" Then
                    Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value = "Yes"
                    Application.ThisWorkbook.Unprotect (Pwd)
                    lockCell = False
                                        
                    Call lockUnlock_cell_rng_without_clearing_contents("T_Tax_Computation", "TaxComp.WithHoldingListWTO", False)
                    Call lockUnlock_cell_rng_without_clearing_contents("T_Tax_Computation", "TaxComp.VehicleAdvTaxPaidListWTO", False)
                    Call lockUnlock_cell_rng_without_clearing_contents("T_Tax_Computation", "TaxComp.MRIWTO", False)
                    'Lawrence added by Lawrence 02/06/2022
                    Call lockUnlock_cell_rng_without_clearing_contents("T_Tax_Computation", "TaxComp.EstateTrustListTOW", False)
                    Call lockUnlock_cell_rng_without_clearing_contents("T_Tax_Computation", "TaxComp.InstallmentTaxListWTO", False)
                    Call lockUnlock_cell_rng_without_clearing_contents("T_Tax_Computation", "TaxComp.IncomePaidAdvanceListWTO", False)
                    
'                    Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
'                    Worksheets("T_Tax_Computation").Range("TaxComp.WithHoldingListWTO").value = 0#
'                    Worksheets("T_Tax_Computation").Range("TaxComp.VehicleAdvTaxPaidListWTO").value = 0#
'                    Worksheets("T_Tax_Computation").Range("TaxComp.MRIWTO").value = 0#
'                    Worksheets("T_Tax_Computation").Range("TaxComp.EstateTrustListTOW").value = 0#
'                    Worksheets("T_Tax_Computation").Range("TaxComp.InstallmentTaxListWTO").value = 0#
'                    Worksheets("T_Tax_Computation").Range("TaxComp.IncomePaidAdvanceListWTO").value = 0#
'                    Worksheets("T_Tax_Computation").Protect (Pwd)
                    'End 02/06/2022
                   If (Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value <> "") Then
                         If (Worksheets("A_Basic_Info").Range("SecA.RtnYear").value < "2015") Then
                            Call lockUnlock_cell_rng("S_Previous_Years_Losses", "C23:H27", True)
                         Else
                             Call lockUnlock_cell_rng("S_Previous_Years_Losses", "C23:H27", False)
                         End If
                    End If
                    Application.ThisWorkbook.Protect (Pwd)

                'Added by Lawrence 22122020 for the covid tax rates affecting IT1
                'Checking if the year selected is 2020 from after choosing wifes income to yes in B17
                If (Worksheets("A_Basic_Info").Range("SecA.RtnYear").value = "2020") Then
                'when year is 2020 and wife income is yes then set hidden to false for the rows
                'in the affected sheet
                            Call hideRows(workSheetArray, False, "B17")
                End If

                ElseIf UCase(Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value) = "NO" _
                Or Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value = "" Then


                'Added by Lawrence 22122020 for the covid tax rates affecting IT1
                'Checking if the year selected is 2020 from after choosing wifes income to no in B17
                If (Worksheets("A_Basic_Info").Range("SecA.RtnYear").value = "2020") Then
                 'when year is 2020 and wife income is No then set hidden to false for the rows
                'in the affected sheet
                            Call hideRows(workSheetArray, True, "B17")
                End If
                    
                Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
                Worksheets("T_Tax_Computation").Range("TaxComp.HomeOwnershipSavingPlanListWTO").value = 0#
               
                Worksheets("T_Tax_Computation").Protect (Pwd)
                    If Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value <> "" Then
                        Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value = "No"
                    End If
                    Call lockUnlock_cell_rng("S_Previous_Years_Losses", "C23:H31", True)
                    Application.ThisWorkbook.Unprotect (Pwd)
                    lockCell = True
                    Worksheets("B_Profit_Loss_Account_Wife").Visible = xlHidden
                    Worksheets("T_Income_Computation_Wife").Visible = xlHidden
                    Call lockUnlock_cell_rng("T_Tax_Computation", "TaxComp.WithHoldingListWTO", True)
                    Call lockUnlock_cell_rng("T_Tax_Computation", "TaxComp.VehicleAdvTaxPaidListWTO", True)
                    Call lockUnlock_cell_rng("T_Tax_Computation", "TaxComp.MRIWTO", True)
                    
                    'Lawrence added by Lawrence 02/06/2022
                    Call lockUnlock_cell_rng_without_clearing_contents("T_Tax_Computation", "TaxComp.EstateTrustListTOW", True)
                    Call lockUnlock_cell_rng_without_clearing_contents("T_Tax_Computation", "TaxComp.InstallmentTaxListWTO", True)
                    Call lockUnlock_cell_rng_without_clearing_contents("T_Tax_Computation", "TaxComp.IncomePaidAdvanceListWTO", True)
                    
'                    Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
'                    Worksheets("T_Tax_Computation").Range("TaxComp.WithHoldingListWTO").value = 0#
'                    Worksheets("T_Tax_Computation").Range("TaxComp.VehicleAdvTaxPaidListWTO").value = 0#
'                    Worksheets("T_Tax_Computation").Range("TaxComp.MRIWTO").value = 0#
'                    Worksheets("T_Tax_Computation").Range("TaxComp.EstateTrustListTOW").value = 0#
'                    Worksheets("T_Tax_Computation").Range("TaxComp.InstallmentTaxListWTO").value = 0#
'                    Worksheets("T_Tax_Computation").Range("TaxComp.IncomePaidAdvanceListWTO").value = 0#
'                    Worksheets("T_Tax_Computation").Protect (Pwd)
                    'End 02/06/2022
                    
                    
                    Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
                    Worksheets("T_Tax_Computation").Range("TaxComp.WithHoldingListWTO").value = 0#
                    Worksheets("T_Tax_Computation").Range("TaxComp.VehicleAdvTaxPaidListWTO").value = 0#
                    Worksheets("T_Tax_Computation").Range("TaxComp.MRIWTO").value = 0#
                    Worksheets("T_Tax_Computation").Protect Password:=Pwd
                    Application.ThisWorkbook.Protect (Pwd)
                End If
                toggleSpouseFields (lockCell)
            If UCase(Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value) = "YES" Then
             If Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value <> "" Then
                'by Palak for Enh-6 SR2 starts
                Worksheets("E1_IDA_CA").Unprotect (Pwd)
                Worksheets("E2_CA_WTA_WDV").Unprotect (Pwd)
                'year = Format(CDate(Trim(endDate)), "yyyy")
                'mm = Format(CDate(Trim(endDate)), "MM")
                
                'Section E2
                If (year >= 2020) Then
                    If (year = 2020 And mm > 3) Or year > 2020 Then
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "B16" & ":C16", False)
                        Worksheets("E2_CA_WTA_WDV").Range("B16").value = "25"
                        Worksheets("E2_CA_WTA_WDV").Range("C16").value = "10"
                        Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_WDV", "B16:C16", True)
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "D17" & ":D19", True)
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "E17" & ":E19", True)
                    Else
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "D17" & ":D19", False)
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "E17" & ":E19", False)
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "B16" & ":C16", False)
                        Worksheets("E2_CA_WTA_WDV").Range("B16").value = "37.5"
                        Worksheets("E2_CA_WTA_WDV").Range("C16").value = "30"
                        Worksheets("E2_CA_WTA_WDV").Range("D17").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("D18").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("D19").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("E17").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("E18").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("E19").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("B17").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("B18").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("B19").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("C17").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("C18").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("C19").value = "0"
                        Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_WDV", "B16:C16", True)
                    End If
                Else
                    Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "D17" & ":D19", False)
                    Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "E17" & ":E19", False)
                    Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "B16" & ":C16", False)
                    Worksheets("E2_CA_WTA_WDV").Range("B16").value = "37.5"
                    Worksheets("E2_CA_WTA_WDV").Range("C16").value = "30"
                    Worksheets("E2_CA_WTA_WDV").Range("D17").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("D18").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("D19").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("E17").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("E18").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("E19").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("B17").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("B18").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("B19").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("C17").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("C18").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("C19").value = "0"
                    Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_WDV", "B16:C16", True)
                End If
                
                'For Section E1 Part 1 wife
                rangeName = Worksheets("E1_IDA_CA").Range("IniAllPlanMach.ListPart1W").Address
                startRow = Worksheets("E1_IDA_CA").Range(rangeName).row
                endRow = Worksheets("E1_IDA_CA").Range(rangeName).Rows.Count + startRow - 1
                
                
                
                If year >= 2020 Then
                    If (year = 2020 And mm >= 4) Or year > 2020 Then
                        Call lockUnlock_cell_rng("E1_IDA_CA", "A" & startRow & ":A" & endRow, True)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "B" & startRow & ":B" & endRow, False)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "C" & startRow & ":C" & endRow, True)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "D" & startRow & ":D" & endRow, False)
                    Else
                        Call lockUnlock_cell_rng("E1_IDA_CA", "A" & startRow & ":A" & endRow, False)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "B" & startRow & ":B" & endRow, True)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "C" & startRow & ":C" & endRow, False)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "D" & startRow & ":D" & endRow, True)
                    End If
                Else
                    Call lockUnlock_cell_rng("E1_IDA_CA", "A" & startRow & ":A" & endRow, False)
                    Call lockUnlock_cell_rng("E1_IDA_CA", "B" & startRow & ":B" & endRow, True)
                    Call lockUnlock_cell_rng("E1_IDA_CA", "C" & startRow & ":C" & endRow, False)
                    Call lockUnlock_cell_rng("E1_IDA_CA", "D" & startRow & ":D" & endRow, True)
                End If
                
                'For Section E1 Part 2 wife
                rangeName = Worksheets("E1_IDA_CA").Range("IniAllIBD.ListPart2W").Address
                startRow = Worksheets("E1_IDA_CA").Range(rangeName).row
                endRow = Worksheets("E1_IDA_CA").Range(rangeName).Rows.Count + startRow - 1
                
                If year >= 2020 Then
                    If (year = 2020 And mm >= 4) Or year > 2020 Then
                        Call lockUnlock_cell_rng("E1_IDA_CA", "A" & startRow & ":A" & endRow, True)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "B" & startRow & ":B" & endRow, False)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "C" & startRow & ":C" & endRow, False)
                    Else
                        Call lockUnlock_cell_rng("E1_IDA_CA", "A" & startRow & ":A" & endRow, False)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "B" & startRow & ":B" & endRow, True)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "C" & startRow & ":C" & endRow, True)
                    End If
                Else
                    Call lockUnlock_cell_rng("E1_IDA_CA", "A" & startRow & ":A" & endRow, False)
                    Call lockUnlock_cell_rng("E1_IDA_CA", "B" & startRow & ":B" & endRow, True)
                    Call lockUnlock_cell_rng("E1_IDA_CA", "C" & startRow & ":C" & endRow, True)
                End If
                
                'For Section E1 Part 4
                rangeName = Worksheets("E1_IDA_CA").Range("DeprIntengAst.ListW").Address
                startRow = Worksheets("E1_IDA_CA").Range(rangeName).row
                endRow = Worksheets("E1_IDA_CA").Range(rangeName).Rows.Count + startRow - 1
                
                If year >= 2020 Then
                    If (year = 2020 And mm >= 4) Or year > 2020 Then
                        Call lockUnlock_cell_rng("E1_IDA_CA", "A" & startRow & ":A" & endRow, True)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "B" & startRow & ":B" & endRow, False)
                    Else
                        Call lockUnlock_cell_rng("E1_IDA_CA", "A" & startRow & ":A" & endRow, False)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "B" & startRow & ":B" & endRow, True)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "C" & startRow & ":C" & endRow, False)
                    End If
                Else
                    Call lockUnlock_cell_rng("E1_IDA_CA", "A" & startRow & ":A" & endRow, False)
                    Call lockUnlock_cell_rng("E1_IDA_CA", "B" & startRow & ":B" & endRow, True)
                    Call lockUnlock_cell_rng("E1_IDA_CA", "C" & startRow & ":C" & endRow, False)
                End If
                Worksheets("E1_IDA_CA").Protect (Pwd)
                Worksheets("E2_CA_WTA_WDV").Protect (Pwd)
            End If
            End If
                'by Palak for Enh-6 SR2 ends
                Worksheets("A_Basic_Info").Activate
                Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").Select
End Sub
