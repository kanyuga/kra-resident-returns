VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Add Row Functionality for Part 1 Self Start
Sub addRow_IniAllPlanMach_ListPart1S()
    Worksheets("E1_IDA_CA").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("IniAllPlanMach.ListPart1S")
    Worksheets("E1_IDA_CA").Protect (Pwd)
End Sub
'Add Row Functionality for Part 1 Self End

'Add Row Functionality for Part 1 Wife Start
Sub addRow_IniAllPlanMach_ListPart1W()
    Worksheets("E1_IDA_CA").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("IniAllPlanMach.ListPart1W")
    Worksheets("E1_IDA_CA").Protect (Pwd)
End Sub
'Add Row Functionality for Part 1 Wife End

'Add Row Functionality for Part 2 Self Start
Sub addRow_IniAllIBD_ListPart2S()
    Worksheets("E1_IDA_CA").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("IniAllIBD.ListPart2S")
    Worksheets("E1_IDA_CA").Protect (Pwd)
End Sub
'Add Row Functionality for Part 2 Self End

'Add Row Functionality for Part 2 Wife Start
Sub addRow_IniAllIBD_ListPart2W()
    Worksheets("E1_IDA_CA").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("IniAllIBD.ListPart2W")
    Worksheets("E1_IDA_CA").Protect (Pwd)
End Sub
'Add Row Functionality for Part 2 Wife End

'Add Row Functionality for Part 4 Self Start
Sub addRow_AgrLandDed_ListS()
    Worksheets("E1_IDA_CA").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("AgrLandDed.ListS")
    Worksheets("E1_IDA_CA").Protect (Pwd)
End Sub
'Add Row Functionality for Part 4 Self End

'Add Row Functionality for Part 4 Wife Start
Sub addRow_AgrLandDed_ListW()
    Worksheets("E1_IDA_CA").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("AgrLandDed.ListW")
    Worksheets("E1_IDA_CA").Protect (Pwd)
End Sub
'Add Row Functionality for Part 4 Wife End

'Add Row Functionality for Part 5 Self Start
Sub addRow_DeprIntengAst_ListS()
    Worksheets("E1_IDA_CA").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("DeprIntengAst.ListS")
    Worksheets("E1_IDA_CA").Protect (Pwd)
End Sub
'Add Row Functionality for Part 5 Self End

'Add Row Functionality for Part 5 Wife Start
Sub addRow_DeprIntengAst_ListW()
    Worksheets("E1_IDA_CA").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("DeprIntengAst.ListW")
    Worksheets("E1_IDA_CA").Protect (Pwd)
End Sub
'Add Row Functionality for Part 5 Wife End

Private Sub Worksheet_Change(ByVal Target As Range)

    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Worksheets("E1_IDA_CA").Unprotect (Pwd)
    Worksheets("E1_IDA_CA").Activate

    If (Target.Columns.Count = 1) Then
        Dim sheetName As String
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
    
        sheetName = Worksheets("E1_IDA_CA").name
        'code for Part 1 Self Start
        rangeName = Worksheets("E1_IDA_CA").Range("IniAllPlanMach.ListPart1S").Address
        startRow = Worksheets("E1_IDA_CA").Range(rangeName).row
        endRow = Worksheets("E1_IDA_CA").Range(rangeName).Rows.Count + startRow - 1
        
        'Validation for Investment Deduction,Investment Deduction can be greater than Cost for 2 different location and can not be greater than 150 % of cost for 3rd location Start
        If Application.Intersect(Target.Cells, Range("J" & startRow & ":J" & endRow)) Is Nothing Then
            'Do nothing
        Else
            cellRow = Target.row
            
            Dim rateval As Integer
            CostExceed = Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value
            If Worksheets("E1_IDA_CA").Range("J" & cellRow & ":J" & cellRow).value <> "" Then
                If (Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "") Then
                    MsgBox "Please Select Location first"
                    Target.value = ""
                    Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).Select
                End If
                If (Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value = "") Then
                    MsgBox "Please Enter Cost Value first"
                    Target.value = ""
                    Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).Select
                End If
                If (Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value <> "") Then
                    If (Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC1" Or Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC2" Or Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC4") Then
                        rateval = 100
                        deductionExceed = CostExceed * rateval / 100
                        If Target.value > deductionExceed Then
                            MsgBox "Investment Deduction cannot be greater than Cost"
                            Target.value = ""
                            Target.Select
                        End If
                    ElseIf (Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC3") Then
                        rateval = 150
                        deductionExceed = CostExceed * rateval / 100
                        If (Target.value > deductionExceed) Then
                            MsgBox "Investment deduction cannot be greater than 150% of the Cost"
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                End If
                End If
            End If
        
        'Validation for Investment Deduction,Investment Deduction can be greater than Cost for 2 different location and can not be greater than 150 % of cost for 3rd location End

        'Validation for Cost,Cost can not be less than Investment Deduction for 2 different location and can not be less than 150 % of cost for 3rd location Start
        If Application.Intersect(Target.Cells, Range("I" & startRow & ":I" & endRow)) Is Nothing Then
        'Do Nothing
        Else
            cellRow = Target.row
            
            CostExceed = Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value
            If Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value <> "" Then
                If (Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "") Then
                    MsgBox "Please Select Location first"
                    Target.value = ""
                    Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).Select
                End If
                If (Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("J" & cellRow & ":J" & cellRow).value <> "") Then
                    If (Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC1" Or Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC2" Or Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC4") Then
                        If (Target.value < Worksheets("E1_IDA_CA").Range("J" & cellRow & ":J" & cellRow).value) Then
                            MsgBox "Cost cannot be less than Investment Deduction"
                            Target.value = ""
                            Target.Select
                        End If
                    ElseIf (Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC3") Then
                        CostExceed = Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value
                        rateval = 150
                        deductionExceed = CostExceed * rateval / 100
                        If (deductionExceed < Worksheets("E1_IDA_CA").Range("J" & cellRow & ":J" & cellRow).value) Then
                        MsgBox "Investment deduction cannot be greater than 150% of the Cost"
                        Target.value = ""
                        Target.Select
                        End If
                    End If
                End If
            End If
        End If
        
        'Validation for Cost,Cost can not be less than Investment Deduction for 2 different location and can not be less than 150 % of cost for 3rd location End

        For i = startRow To endRow
            Select Case Target.Address
                'Validation for Date of First Use,Date of First Use should be between Return period Start
                Case "$G$" & i
                    
                    Dim tSValue As String
                    If Target.value <> "" Then
                        If TestDate(Target.value) = True Then
                        tSValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                            If (startDate <> "" And endDate <> "") Then
                                If (DateValue(tSValue) < DateValue(startDate) Or DateValue(tSValue) > DateValue(endDate)) Then
                                    MsgBox "Date of First Use in Row " & i & " should be within Return period " & Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value & " to " & Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Activate
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                                GoTo endSub
                            End If
                        Else
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'Validation for Date of First Use,Date of First Use should be between Return period End
                
                Case "$B$" & i
                    cellRow = Target.row
                    Worksheets("E1_IDA_CA").Range("D" & i).value = ""
                
                'Validation for cost,base on selected location cost value will be less than or equal to 200 million Start
                Case "$I$" & i
                    cellRow = Target.row
                    If (Target.value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value <> "") Then
                            If (Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC3" And Target.value < 200000000) Then
                                MsgBox "Cost against the selected Location must be equal or greater than 200 million."
                                Target.value = ""
                                Target.Select
                            ElseIf (Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC4" And Target.value >= 200000000) Then
                                MsgBox "Cost against the selected Location must be less than 200 million."
                                Target.value = ""
                                Target.Select
                            End If
                        End If
                    End If
                    
                'Validation for cost,base on selected location cost value will be less than or equal to 200 million End
                    
                'Validation for cost,base on selected location cost value will be less than or equal to 200 million Start
                Case "$H$" & i
                    cellRow = Target.row
                    
                    If (Target.value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value <> "") Then
                            If (Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value < 200000000 And Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC3") Then
                                MsgBox "Cost is less than 200 million. Please select appropriate location."
                                Target.value = ""
                                Target.Select
                            ElseIf (Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value >= 200000000 And Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC4") Then
                                MsgBox "Cost is greater than or equal to 200 million. Please select appropriate location."
                                Target.value = ""
                                Target.Select
                            End If
                        End If
                    End If
                    
                    'Added on 13/07/2023 ticket 64843
                        Dim foundRng As Range
                        
                        Dim rtnPrdTo As String
                        Dim mm As String
                        Dim year As String
                        
                        rtnPrdTo = CDate(Format(Worksheets("A_Basic_Info").Range("RetInf.RtnPrdToAct").value, "dd/mm/yyyy"))
                        If (rtnPrdTo <> "") Then
                             mm = Format(CDate(Trim(rtnPrdTo)), "MM")
                             year = DatePart("yyyy", rtnPrdTo)
                             
                             If ((year = 2022 And mm > 6) Or (year > 2022)) Then
                                 With Worksheets("Data").Range("LocationJuly2022")
                                     Set foundRng = .Find(Target.value, LookIn:=xlValues)
                                     If Not foundRng Is Nothing Then
                                        Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = foundRng.Offset(0, 1).value
                                     End If
                                 End With
                             ElseIf (year = 2022 And mm < 7) Then
                                 With Worksheets("Data").Range("LocationJan2022")
                                    Set foundRng = .Find(Target.value, LookIn:=xlValues)
                                    If Not foundRng Is Nothing Then
                                        Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = foundRng.Offset(0, 1).value
                                    End If
                                End With
                             ElseIf ((year = 2020 And mm > 4) Or (year > 2020 And year < 2022)) Then
                                 With Worksheets("Data").Range("Location2020")
                                    Set foundRng = .Find(Target.value, LookIn:=xlValues)
                                    If Not foundRng Is Nothing Then
                                        Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = foundRng.Offset(0, 1).value
                                    End If
                                End With
                             ElseIf ((year = 2020 And mm < 5) Or year < 2020) Then
                                 With Worksheets("Data").Range("Location")
                                    Set foundRng = .Find(Target.value, LookIn:=xlValues)
                                    If Not foundRng Is Nothing Then
                                        Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = foundRng.Offset(0, 1).value
                                    End If
                                End With
                             
                             Else
                            
                             End If
                        End If
                    'End on 13/07/2023 ticket 64843
                    
                'Validation for cost,base on selected location cost value will be less than or equal to 200 End
            End Select
        Next
        'code for Part 1 Self End

        'code for Part 1 Wife Start
        rangeName = Worksheets("E1_IDA_CA").Range("IniAllPlanMach.ListPart1W").Address
        startRow = Worksheets("E1_IDA_CA").Range(rangeName).row
        endRow = Worksheets("E1_IDA_CA").Range(rangeName).Rows.Count + startRow - 1
        
        
        'Validation for Investment Deduction,Investment Deduction can be greater than Cost for 2 different location and can not be greater than 150 % of cost for 3rd location Start
        If Application.Intersect(Target.Cells, Range("J" & startRow & ":J" & endRow)) Is Nothing Then
            'Do nothing
        Else
            cellRow = Target.row
            
            CostExceed = Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value
            If Worksheets("E1_IDA_CA").Range("J" & cellRow & ":J" & cellRow).value <> "" Then
                If (Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "") Then
                    MsgBox "Please Select Location first"
                    Target.value = ""
                    Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).Select
                End If
                If (Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value = "") Then
                    MsgBox "Please Enter Cost Value first"
                    Target.value = ""
                    Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).Select
                End If
                If (Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value <> "") Then
                    If (Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC1" Or Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC2" Or Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC4") Then
                        rateval = 100
                        deductionExceed = CostExceed * rateval / 100
                        If Target.value > deductionExceed Then
                            MsgBox "Investment Deduction cannot be greater than Cost"
                            Target.value = ""
                            Target.Select
                        End If
                    ElseIf (Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC3") Then
                        rateval = 150
                        deductionExceed = CostExceed * rateval / 100
                        If (Target.value > deductionExceed) Then
                            MsgBox "Investment deduction cannot be greater than 150% of the Cost"
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                End If
            End If
        End If
        
        'Validation for Investment Deduction,Investment Deduction can be greater than Cost for 2 different location and can not be greater than 150 % of cost for 3rd location End

        'Validation for Cost,Cost can not be less than Investment Deduction for 2 different location and can not be less than 150 % of cost for 3rd location Start
        If Application.Intersect(Target.Cells, Range("I" & startRow & ":I" & endRow)) Is Nothing Then
        'Do Nothing
        Else
            cellRow = Target.row
            
            CostExceed = Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value
            If Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value <> "" Then
                If (Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "") Then
                    MsgBox "Please Select Location first"
                    Target.value = ""
                    Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).Select
                End If
                If (Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("J" & cellRow & ":J" & cellRow).value <> "") Then
                    If (Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC1" Or Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC2") Then
                        If (Target.value < Worksheets("E1_IDA_CA").Range("J" & cellRow & ":J" & cellRow).value) Then
                            MsgBox "Cost cannot be less than Investment Deduction"
                            Target.value = ""
                            Target.Select
                        End If
                    ElseIf (Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC3") Then
                        CostExceed = Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value
                        rateval = 150
                        deductionExceed = CostExceed * rateval / 100
                        If (deductionExceed < Worksheets("E1_IDA_CA").Range("J" & cellRow & ":J" & cellRow).value) Then
                            MsgBox "Investment deduction cannot be greater than 150% of the Cost"
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                End If
            End If
        End If
        
        'Validation for Cost,Cost can not be less than Investment Deduction for 2 different location and can not be less than 150 % of cost for 3rd location End

    
        For i = startRow To endRow
            Select Case Target.Address
                'Validation for Date of First Use,Date of First Use should be between Return period Start
                Case "$G$" & i
                    
                    Dim tWValue As String
                    If Target.value <> "" Then
                        If TestDate(Target.value) = True Then
                            tWValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                            If (startDate <> "" And endDate <> "") Then
                                If (DateValue(tWValue) < DateValue(startDate) Or DateValue(tWValue) > DateValue(endDate)) Then
                                    MsgBox "Date of First Use in Row " & i & " should be within Return period " & Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value & " to " & Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Activate
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                                GoTo endSub
                            End If
                        Else
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'Validation for Date of First Use,Date of First Use should be between Return period End
                    
                Case "$B$" & i
                    cellRow = Target.row
                    Worksheets("E1_IDA_CA").Range("D" & i).value = ""
                    
                'Validation for cost,base on selected location cost value will be less than or equal to 200 million Start
                Case "$I$" & i
                    cellRow = Target.row
                    
                    If (Target.value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value <> "") Then
                            If (Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC3" And Target.value < 200000000) Then
                                MsgBox "Cost against the selected Location must be equal to or greater than 200 million."
                                Target.value = ""
                                Target.Select
                            ElseIf (Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC4" And Target.value >= 200000000) Then
                                MsgBox "Cost against the selected Location must be less than 200 million."
                                Target.value = ""
                                Target.Select
                            End If
                        End If
                    End If
                    
                 'Validation for cost,base on selected location cost value will be less than or equal to 200 million End
                    
                'Validation for cost,base on selected location cost value will be less than or equal to 200 million Start
                Case "$H$" & i
                    cellRow = Target.row
                    
                    If (Target.value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value <> "") Then
                            If (Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value < 200000000 And Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC3") Then
                                MsgBox "Cost is less than 200 million. Please select appropriate location."
                                Target.value = ""
                                Target.Select
                            ElseIf (Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value >= 200000000 And Worksheets("E1_IDA_CA").Range("K" & cellRow & ":K" & cellRow).value = "LC4") Then
                                MsgBox "Cost is greater than or equal to 200 million. Please select appropriate location."
                                Target.value = ""
                                Target.Select
                            End If
                        End If
                    End If
                    
                'Validation for cost,base on selected location cost value will be less than or equal to 200 million End
            End Select
        Next
    
    
        'code for Part 2 Self Start
    
        rangeName = Range("IniAllIBD.ListPart2S").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1
    
        For i = startRow To endRow
            Select Case Target.Address
            
                'by Palak for Enh-6 SR2
                Case "$B$" & i
                    cellRow = Target.row
                    If (Target.value <> "" And Worksheets("E1_IDA_CA").Range("B" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("B" & cellRow).value = "Construction of Bulk Storage and Handling Facilities(SGR)") Then
                            Call lockUnlock_cell_rng("E1_IDA_CA", "C" & cellRow, False)
                        Else
                            Call lockUnlock_cell_rng("E1_IDA_CA", "C" & cellRow, True)
                        End If
                    End If
                
                'Validation for Date of First Use,Date of First Use should be between Return period and base one Date of First use and Return Peiod End Date,Days in use or available for use will be populate Start
                Case "$D$" & i
                    Dim tIBDSValue As String
                    If (Target.value <> "") Then
                        If (TestDate(Target.value) = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            tIBDSValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                            If (startDate <> "" And endDate <> "") Then
                                'If (DateValue(tIBDSValue) < DateValue(startDate) Or DateValue(tIBDSValue) > DateValue(endDate)) Then
                                If (DateValue(tIBDSValue) > DateValue(endDate)) Then
                                    MsgBox "Date Of First Use in Row " & i & " should be before or within Return period " & Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value & " to " & Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                ElseIf Target.value <> "" And Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value <> "" Then
                                    Dim dy As Integer
                                    Dim tdSvalue As String
                                    Dim rSEndDate As String
                                    tdSvalue = CDate(Format(Target.value, "dd/mm/yyyy"))
                                    rSEndDate = CDate(Format(Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value, "dd/mm/yyyy"))
                                    dy = DateDiff("d", tdSvalue, rSEndDate)
                                    If (dy <= 366 And dy >= 0) Then
                                        
                                        Target.Next.value = dy
                                        
                                    End If
                                    If (dy > 366) Then
                                        If (isyLeapYear(Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value) = True Or isyLeapYear(Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value) = True) Then
                                            
                                            Target.Next.value = 366
                                            
                                        Else
                                            
                                            Target.Next.value = 365
                                            
                                        End If
                                    End If
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Activate
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                                GoTo endSub
                            End If
                        End If
                    Else
                        
                        Target.Next.value = ""
                        
                    End If
                'Validation for Date of First Use,Date of First Use should be between Return period and base one Date of First use and Return Peiod End Date,Days in use or available for use will be populate End
        
                'Validation for Cost,Cost should be Greater than or equal to Investment Deduction Start
                Case "$F$" & i
                    cellRow = Target.row
                    If (Target.value <> "" And Worksheets("E1_IDA_CA").Range("B" & cellRow).value = "Construction of Bulk Storage and Handling Facilities(SGR)") Then
                        If (Worksheets("E1_IDA_CA").Range("F" & cellRow).value < 5000000000#) Then
                            MsgBox "Value of Cost should be greater than or equal to 5,000,000,000"
                        End If
                    Else
                        If (Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value <> "") Then
                            If (Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value < Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value) Then
                                MsgBox "Cost should be greater than or equal to Investment Deduction."
                                Target.value = ""
                                Target.Select
                            End If
                        End If
                        If (Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("J" & cellRow & ":J" & cellRow).value <> "") Then
                            If (Worksheets("E1_IDA_CA").Range("J" & cellRow & ":J" & cellRow).value > Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value + Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value) Then
                                    MsgBox "Capital Allowance should be less than or equal to sum of values in C and D."
                                    Target.value = ""
                                    Target.Select
                            End If
                        End If
                    End If
                    
                'Validation for Cost,Cost should be Greater than or equal to Investment Deduction End
                
                'Validation for Investment Deduction,Investment Deduction can not be greater than cost Start
                Case "$G$" & i
                    cellRow = Target.row
                    
                    If (Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value <> "" Or Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value > Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value) Then
                            MsgBox "Investment Deduction cannot be greater than Cost"
                            Target.value = ""
                            Target.Select
                        ElseIf (Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value <> "" And (Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value <> "")) Then
                            If (Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value = "0") Then
                                MsgBox "Investment Deduction should not be allowed for 0 Day."
                                Target.value = ""
                                Target.Select
                            End If
                        End If
                    End If
                    If (Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("J" & cellRow & ":J" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("J" & cellRow & ":J" & cellRow).value > Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value + Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value) Then
                                MsgBox "Capital Allowance should be less than or equal to sum of values in C and D."
                                Target.value = ""
                                Target.Select
                        End If
                    End If
                    
                'Validation for Investment Deduction,Investment Deduction can not be greater than cost End
                
                'Validation for Days in use or Available for Use,not allowed for Investmenr Deduction value Zero Start
                Case "$E$" & i
                    cellRow = Target.row
                    If (Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value = "0" And Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value > 0) Then
                            MsgBox "Days in use or Availabel for use not allowed zero for any Investment Deduction."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
               'Validation for Days in use or Available for Use,not allowed for Investmenr Deduction value Zero End
                
                'Validation for Capital Allowance,Capital Allowance should be less than or equal to sum of Qualifying Expenditure for Capital Allowance and Resident Value B/F Start
                Case "$J$" & i
                    cellRow = Target.row
                    If (Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("J" & cellRow & ":J" & cellRow).value > Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value + Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value) Then
                            MsgBox "Capital Allowance should be less than or equal to sum of values in C and D."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'Validation for Capital Allowance,Capital Allowance should be less than or equal to sum of Qualifying Expenditure for Capital Allowance and Resident Value B/F End
                    
                'Validation for Capital Allowance,Capital Allowance should be less than or equal to sum of Qualifying Expenditure for Capital Allowance and Resident Value B/F Start
                Case "$I$" & i
                    cellRow = Target.row
                    
                    If (Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("J" & cellRow & ":J" & cellRow).value > Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value + Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value) Then
                            MsgBox "Capital Allowance should be less than or equal to sum of values in C and D."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'Validation for Capital Allowance,Capital Allowance should be less than or equal to sum of Qualifying Expenditure for Capital Allowance and Resident Value B/F End
            End Select
        Next
        'code for Part 2 Self End
    
    
        'code for Part 2 Wife Start
        rangeName = Range("IniAllIBD.ListPart2W").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1
    
        For i = startRow To endRow
            Select Case Target.Address
                'Validation for Date of First Use,Date of First Use should be between Return period and base one Date of First use and Return Peiod End Date,Days in use or available for use will be populate Start
                Case "$D$" & i
                    Dim tIBDWValue As String
                    If (Target.value <> "") Then
                        If (TestDate(Target.value) = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            tIBDWValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                            If (startDate <> "" And endDate <> "") Then
                                'If (DateValue(tIBDWValue) < DateValue(startDate) Or DateValue(tIBDWValue) > DateValue(endDate)) Then
                                If (DateValue(tIBDWValue) > DateValue(endDate)) Then
                                    MsgBox "Date Of First Use in Row " & i & " should be before within Return period " & Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value & " to " & Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                ElseIf Target.value <> "" And Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value <> "" Then
                                    Dim dyw As Integer
                                    Dim tdWvalue As String
                                    Dim rWEndDate As String
                                    tdWvalue = CDate(Format(Target.value, "dd/mm/yyyy"))
                                    rWEndDate = CDate(Format(Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value, "dd/mm/yyyy"))
                                    dyw = DateDiff("d", tdWvalue, rWEndDate)
                                    If (dyw <= 366 And dyw >= 0) Then
                                        
                                        Target.Next.value = dyw
                                        
                                    End If
                                    If (dyw > 366) Then
                                        If (isyLeapYear(Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value) = True Or isyLeapYear(Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value) = True) Then
                                            
                                            Target.Next.value = 366
                                            
                                        Else
                                            
                                            Target.Next.value = 365
                                            
                                        End If
                                    End If
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Activate
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                                GoTo endSub
                            End If
                        End If
                    'Else
                        
                       ' Target.Next.value = ""
                        
                    End If
                'Validation for Date of First Use,Date of First Use should be between Return period and base one Date of First use and Return Peiod End Date,Days in use or available for use will be populate End
                
                'Validation for Cost,Cost should be Greater than or equal to Investment Deduction Start
                Case "$F$" & i
                    cellRow = Target.row
                    
                    If (Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value < Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value) Then
                            MsgBox "Cost should be greater than or equal to Investment Deduction."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    If (Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("J" & cellRow & ":J" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("J" & cellRow & ":J" & cellRow).value > Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value + Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value) Then
                                MsgBox "Capital Allowance should be less than or equal to sum of values in C and D."
                                Target.value = ""
                                Target.Select
                        End If
                    End If
                    
                'Validation for Cost,Cost should be Greater than or equal to Investment Deduction End
                
                'Validation for Investment Deduction,Investment Deduction can not be greater than cost Start
                Case "$G$" & i
                    cellRow = Target.row
                    
                    If (Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value <> "" Or Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value > Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value) Then
                            MsgBox "Investment Deduction cannot be greater than Cost"
                            Target.value = ""
                            Target.Select
                        ElseIf (Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value <> "" And (Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value <> "")) Then
                            If (Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value = "0") Then
                                MsgBox "Investment Deduction should not be allowed for 0 Day."
                                Target.value = ""
                                Target.Select
                            End If
                        End If
                    End If
                    If (Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("J" & cellRow & ":J" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("J" & cellRow & ":J" & cellRow).value > Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value + Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value) Then
                                MsgBox "Capital Allowance should be less than or equal to sum of values in C and D."
                                Target.value = ""
                                Target.Select
                        End If
                    End If
                    
                'Validation for Investment Deduction,Investment Deduction can not be greater than cost End
                
                'Validation for Days in use or Available for Use,not allowed for Investmenr Deduction value Zero Start
                Case "$E$" & i
                    cellRow = Target.row
                    
                    If (Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value = "0" And Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value > 0) Then
                            MsgBox "Days in use or Availabel for use not allowed zero for any Investment Deduction."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'Validation for Days in use or Available for Use,not allowed for Investmenr Deduction value Zero End
                
                'Validation for Capital Allowance,Capital Allowance should be less than or equal to sum of Qualifying Expenditure for Capital Allowance and Resident Value B/F Start
                Case "$J$" & i
                    cellRow = Target.row
                    
                    If (Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("J" & cellRow & ":J" & cellRow).value > Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value + Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value) Then
                            MsgBox "Capital Allowance should be less than or equal to sum of values in C and D."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'Validation for Capital Allowance,Capital Allowance should be less than or equal to sum of Qualifying Expenditure for Capital Allowance and Resident Value B/F End
                    
                'Validation for Capital Allowance,Capital Allowance should be less than or equal to sum of Qualifying Expenditure for Capital Allowance and Resident Value B/F Start
                Case "$I$" & i
                    cellRow = Target.row
                    
                    If (Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("J" & cellRow & ":J" & cellRow).value > Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value + Worksheets("E1_IDA_CA").Range("I" & cellRow & ":I" & cellRow).value) Then
                            MsgBox "Capital Allowance should be less than or equal to sum of values in C and D."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'Validation for Capital Allowance,Capital Allowance should be less than or equal to sum of Qualifying Expenditure for Capital Allowance and Resident Value B/F End
            End Select
        Next
         'code for Part 2 Wife End
         
    
    
        'code for Part 4 Self Start
    
        rangeName = Range("AgrLandDed.ListS").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1
    
        For i = startRow To endRow
            Select Case Target.Address
                'Validation for Date of Expenditure,Date of Expenditure should be between Return period Start
                Case "$B$" & i
                    
                    Dim tAGRSValue As String
                    If Target.value <> "" Then
                        If TestDate(Target.value) = True Then
                            tAGRSValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                            If (startDate <> "" And endDate <> "") Then
                                If (DateValue(tAGRSValue) < DateValue(startDate) Or DateValue(tAGRSValue) > DateValue(endDate)) Then
                                    MsgBox "Date of Expenditure in Row " & i & " should be within Return period " & Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value & " to " & Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Activate
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                                GoTo endSub
                            End If
                        Else
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'Validation for Date of Expenditure,Date of Expenditure should be between Return period End
                    
                'Validation for Deductions During the year,Deduction during the year should be less than or equal to sum of Expenditure Incurred and written down value begining of the year and subtract from Sales or Trasfer Start
                Case "$F$" & i
                    cellRow = Target.row
                    
                    If (Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value > Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value - Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value) Then
                            'MsgBox "Deductions during the Year should be less than or equal to A+B-C."
                            MsgBox "The sum of values of A and B should be greater than or equal to sum of values of C and D."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'Validation for Deductions During the year,Deduction during the year should be less than or equal to sum of Expenditure Incurred and written down value begining of the year and subtract from Sales or Trasfer End
                    
                'Validation for sales or Trasfer,Sales or Trasfer should be less than or equal to sum of Expenditure Incurred and written down value begining of the year Start
                Case "$E$" & i
                    cellRow = Target.row
                    
                    If (Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value > Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value) Then
                            'MsgBox "Sale or Transfer should be less than or equal to A+B."
                            MsgBox "The sum of values of A and B should be greater than or equal to sum of values of C and D."
                            Target.value = ""
                            Target.Select
                        Else
                            If (Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value <> "") Then
                                If (Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value > Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value - Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value) Then
                                        'MsgBox "Sale or Transfer should be less than or equal to A+B-D."
                                        MsgBox "The sum of values of A and B should be greater than or equal to sum of values of C and D."
                                        Target.value = ""
                                        Target.Select
                                End If
                            End If
                        End If
                    End If
                    
                'Validation for sales or Trasfer,Sales or Trasfer should be less than or equal to sum of Expenditure Incurred and written down value begining of the year End
                    
                'Validation for sales or Trasfer,Sales or Trasfer should be less than or equal to sum of Expenditure Incurred and written down value begining of the year Start
                Case "$C$" & i
                    cellRow = Target.row
                    
'                    If (Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value <> "") Then
'                        If (Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value > Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value) Then
'                            'MsgBox "Sale or Transfer should be less than or equal to A+B."
'                            MsgBox "The sum of values of A and B should be greater than or equal to sum of values of C and D."
'                            Target.value = ""
'                            Target.Select
'                        End If
'                    End If
                     If (Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value < Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value + Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value) Then
                            MsgBox "The sum of values of A and B should be greater than or equal to sum of values of C and D."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                    
                'Validation for sales or Trasfer,Sales or Trasfer should be less than or equal to sum of Expenditure Incurred and written down value begining of the year End
                    
                'Validation for sales or Trasfer,Sales or Trasfer should be less than or equal to sum of Expenditure Incurred and written down value begining of the year Start
                Case "$D$" & i
                    cellRow = Target.row
                    
'                    If (Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value <> "") Then
'                        If (Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value > Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value) Then
'                            'MsgBox "Sale or Transfer should be less than or equal to A+B."
'                            MsgBox "The sum of values of A and B should be greater than or equal to sum of values of C and D."
'                            Target.value = ""
'                            Target.Select
'                        End If
'                    End If
                    
                     If (Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value < Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value + Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value) Then
                            MsgBox "The sum of values of A and B should be greater than or equal to sum of values of C and D."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                    
                'Validation for sales or Trasfer,Sales or Trasfer should be less than or equal to sum of Expenditure Incurred and written down value begining of the year End
            End Select
        Next
        'code for Part 4 Self End
    
        'code for Part 4 Wife Start
    
        rangeName = Range("AgrLandDed.ListW").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1
    
        For i = startRow To endRow
            Select Case Target.Address
                'Validation for Date of Expenditure,Date of Expenditure should be between Return period Start
                Case "$B$" & i
                    
                    Dim tAGRWValue As String
                    If Target.value <> "" Then
                        If TestDate(Target.value) = True Then
                            tAGRWValue = CDate(Format(Target.value, "dd/mm/yyyy"))
                            If (startDate <> "" And endDate <> "") Then
                                If (DateValue(tAGRWValue) < DateValue(startDate) Or DateValue(tAGRWValue) > DateValue(endDate)) Then
                                    MsgBox "Date of Expenditure in Row " & i & " should be within Return period " & Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value & " to " & Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Activate
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                                GoTo endSub
                            End If
                        Else
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'Validation for Date of Expenditure,Date of Expenditure should be between Return period End
                    
                'Validation for Deductions During the year,Deduction during the year should be less than or equal to sum of Expenditure Incurred and written down value begining of the year and subtract from Sales or Trasfer Start
                Case "$F$" & i
                    cellRow = Target.row
                    
                    If (Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value > Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value - Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value) Then
                            'MsgBox "Deductions during the Year should be less than or equal to A+B-C."
                            MsgBox "The sum of values of A and B should be greater than or equal to sum of values of C and D."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'Validation for Deductions During the year,Deduction during the year should be less than or equal to sum of Expenditure Incurred and written down value begining of the year and subtract from Sales or Trasfer End
                    
                'Validation for sales or Trasfer,Sales or Trasfer should be less than or equal to sum of Expenditure Incurred and written down value begining of the year Start
                Case "$E$" & i
                    cellRow = Target.row
                    
                    If (Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value > Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value) Then
                            'MsgBox "Sale or Transfer should be less than or equal to A+B."
                            MsgBox "The sum of values of A and B should be greater than or equal to sum of values of C and D."
                            Target.value = ""
                            Target.Select
                        Else
                            If (Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value <> "") Then
                                If (Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value > Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value - Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value) Then
                                        'MsgBox "Sale or Transfer should be less than or equal to A+B-D."
                                        MsgBox "The sum of values of A and B should be greater than or equal to sum of values of C and D."
                                        Target.value = ""
                                        Target.Select
                                End If
                            End If
                        End If
                    End If
                    
                'Validation for sales or Trasfer,Sales or Trasfer should be less than or equal to sum of Expenditure Incurred and written down value begining of the year End
                    
                'Validation for sales or Trasfer,Sales or Trasfer should be less than or equal to sum of Expenditure Incurred and written down value begining of the year Start
                Case "$C$" & i
                    cellRow = Target.row
                    
'                    If (Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value <> "") Then
'                        If (Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value > Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value) Then
'                            'MsgBox "Sale or Transfer should be less than or equal to A+B."
'                            MsgBox "The sum of values of A and B should be greater than or equal to sum of values of C and D."
'                            Target.value = ""
'                            Target.Select
'                        End If
'                    End If
                    
                    If (Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value < Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value + Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value) Then
                            MsgBox "The sum of values of A and B should be greater than or equal to sum of values of C and D."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                    
                'Validation for sales or Trasfer,Sales or Trasfer should be less than or equal to sum of Expenditure Incurred and written down value begining of the year End
                    
                'Validation for sales or Trasfer,Sales or Trasfer should be less than or equal to sum of Expenditure Incurred and written down value begining of the year Start
                Case "$D$" & i
                    cellRow = Target.row
                    
'                    If (Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value <> "") Then
'                        If (Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value > Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value) Then
'                            'MsgBox "Sale or Transfer should be less than or equal to A+B."
'                            MsgBox "The sum of values of A and B should be greater than or equal to sum of values of C and D."
'                            Target.value = ""
'                            Target.Select
'                        End If
'                    End If
                    If (Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("E1_IDA_CA").Range("D" & cellRow & ":D" & cellRow).value < Worksheets("E1_IDA_CA").Range("E" & cellRow & ":E" & cellRow).value + Worksheets("E1_IDA_CA").Range("F" & cellRow & ":F" & cellRow).value) Then
                            MsgBox "The sum of values of A and B should be greater than or equal to sum of values of C and D."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                'Validation for sales or Trasfer,Sales or Trasfer should be less than or equal to sum of Expenditure Incurred and written down value begining of the year End
            End Select
        Next
         'code for Part 4 Wife End
    
        'code for Part 5 Self Start
    
        rangeName = Range("DeprIntengAst.ListS").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1
    
        For i = startRow To endRow
            Select Case Target.Address
                'Validation for Date of First use,Date of First use should be before return period End Date Start
                Case "$E$" & i
                    If (Target.value <> "") Then
                        result = TestDate(Target.value)
                        If (result = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            If (startDate <> "" And endDate <> "") Then
                                If TestDateBeforeRtnPeriodEnd(Target.value) = False Then
                                    'MsgBox "Date of First Use should be before Return Period end date."
                                    MsgBox "Date Of First Use in Row " & i & " should be before or within Return period " & Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value & " to " & Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                Else
                                    If TestDateBtwnRtnPeriod(Target.value) = True Then
                                        Target.Next.value = 40
                                     Else
                                        currYr = DatePart("yyyy", startDate)
                                        diffYr = DatePart("yyyy", CDate(Format(Target.value, "dd/mm/yyyy")))
                                        If ((currYr - diffYr) > 6) Then
                                            Target.Next.value = 0
                                        Else
                                            Target.Next.value = 10
                                        End If
                                     End If
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Activate
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                                GoTo endSub
                            End If
                        End If
                    Else
                        
                        Target.Next.value = ""
                        
                    End If
                'Validation for Date of First use,Date of First use should be before return period End Date End
                    
                'Validation for Disposal at cost during the year,Disposal at cost during the year should be less than or equal to Actual Cost Start.
                Case "$H$" & i
                    cellRow = Target.row
                    
                    If (Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value > Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value) Then
                            MsgBox "Disposal at Cost during the Year should be less than or equal to Actual Cost."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'Validation for Disposal at cost during the year,Disposal at cost during the year should be less than or equal to Actual Cost End
                    
                'by Palak for enh-6 SR2 Section E1 Part 4
                Case "$B$" & i
                    cellRow = Target.row
                    If (Worksheets("E1_IDA_CA").Range("B" & cellRow & ":B" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("B" & cellRow & ":B" & cellRow).value = "Other Machinery") Then
                        Call lockUnlock_cell_rng("E1_IDA_CA", "C" & i, False)
                    Else
                        Call lockUnlock_cell_rng("E1_IDA_CA", "C" & i, True)
                    End If
                
                'Validation for Actual cost,Actual Cost should be Greater than or equal to Disposal at Cost during the year Start
                Case "$G$" & i
                    cellRow = Target.row
                    
                    If (Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value < Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value) Then
                            MsgBox " Actual Cost should be greater than or equal to Disposal at Cost during the Year."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'Validation for Actual cost,Actual Cost should be Greater than or equal to Disposal at Cost during the year End
            End Select
        Next
        'code for Part 5 Self End
    
        'code for Part 5 Wife Start
        rangeName = Range("DeprIntengAst.ListW").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1
    
        For i = startRow To endRow
            Select Case Target.Address
                'Validation for Date of First use,Date of First use should be before return period End Date Start
                Case "$E$" & i
                    If (Target.value <> "") Then
                        result = TestDate(Target.value)
                        If (result = False) Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            If (startDate <> "" And endDate <> "") Then
                                If TestDateBeforeRtnPeriodEnd(Target.value) = False Then
                                    'MsgBox "Date of First Use should be before Return Period end date."
                                    MsgBox "Date Of First Use in Row " & i & " should be before or within Return period " & Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").value & " to " & Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
                                    Target.value = ""
                                    Target.Select
                                Else
                                     If TestDateBtwnRtnPeriod(Target.value) = True Then
                                        Target.Next.value = 40
                                     Else
                                        currYr = DatePart("yyyy", startDate)
                                        diffYr = DatePart("yyyy", CDate(Format(Target.value, "dd/mm/yyyy")))
                                        If ((currYr - diffYr) > 6) Then
                                            Target.Next.value = 0
                                        Else
                                            Target.Next.value = 10
                                        End If
                                     End If
                                End If
                            Else
                                MsgBox "Please first enter the Return Period From and Return Period To."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Activate
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                                GoTo endSub
                            End If
                        End If
                    Else
                        Target.Next.value = ""
                    End If
                'Validation for Date of First use,Date of First use should be before return period End Date End
                    
                'Validation for Disposal at cost during the year,Disposal at cost during the year should be less than or equal to Actual Cost Start.
                Case "$H$" & i
                    cellRow = Target.row
                    
                    If (Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value > Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value) Then
                            MsgBox "Disposal at Cost during the Year should be less than or equal to Actual Cost."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'Validation for Disposal at cost during the year,Disposal at cost during the year should be less than or equal to Actual Cost End
                    
                'by Palak for enh-6 SR2 Section E1 Part 4
                Case "$B$" & i
                    cellRow = Target.row
                    If (Worksheets("E1_IDA_CA").Range("B" & cellRow & ":B" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("B" & cellRow & ":B" & cellRow).value = "Other Machinery") Then
                        Call lockUnlock_cell_rng("E1_IDA_CA", "C" & i, False)
                    Else
                        Call lockUnlock_cell_rng("E1_IDA_CA", "C" & i, True)
                    End If
                    
                'Validation for Actual cost,Actual Cost should be Greater than or equal to Disposal at Cost during the year Start
                Case "$G$" & i
                    cellRow = Target.row
                    
                    If (Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value <> "" And Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value <> "") Then
                        If (Worksheets("E1_IDA_CA").Range("G" & cellRow & ":G" & cellRow).value < Worksheets("E1_IDA_CA").Range("H" & cellRow & ":H" & cellRow).value) Then
                            MsgBox " Actual Cost should be greater than or equal to Disposal at Cost during the Year."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'Validation for Actual cost,Actual Cost should be Greater than or equal to Disposal at Cost during the year End
            End Select
        Next
        'code for Part 5 Wife End
    End If
endSub:
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Worksheets("E1_IDA_CA").Protect (Pwd)
End Sub
