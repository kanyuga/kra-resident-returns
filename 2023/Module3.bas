Attribute VB_Name = "Module3"
Dim isFileModified As Boolean
Dim isWifeDec As String
Public Const sLLMarker As String = "¨¦¨©LL©¨¦¨"
Dim rFromDate As String
Dim rToDate As String
Sub importCSV_B_Profit_Loss_Account_Self_Part1()
    Dim ret As Boolean
    Dim Str As String
    Application.ScreenUpdating = False
    Application.EnableEvents = False
   fieldInfoArr = Array(Array(1, 2), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1))
    ret = OpenCSV("B_Profit_Loss_Account_Self", "PLA.OtherExpensesListS", "7", True, 732, 765, "H", False, "", fieldInfoArr, 719, 725)

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Sub importCSV_B_Profit_Loss_Account_Wife_Part1()
    Dim ret As Boolean
    Dim Str As String
    Application.ScreenUpdating = False
    Application.EnableEvents = False
   fieldInfoArr = Array(Array(1, 2), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1))
    ret = OpenCSV("B_Profit_Loss_Account_Wife", "PLA.OtherExpensesListW", "7", True, 732, 765, "H", False, "", fieldInfoArr, 719, 725)

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Sub importCSV_B_Profit_Loss_Account_Self_Part2()
    Dim ret As Boolean
    Dim Str As String
    Application.ScreenUpdating = False
    Application.EnableEvents = False
   fieldInfoArr = Array(Array(1, 2), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1))
    ret = OpenCSV("B_Profit_Loss_Account_Self", "PLA.OtherIncomeListS", "7", True, 783, 816, "H", False, "", fieldInfoArr, 770, 776)
    

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Sub importCSV_B_Profit_Loss_Account_Wife_Part2()
    Dim ret As Boolean
    Dim Str As String
    Application.ScreenUpdating = False
    Application.EnableEvents = False
   fieldInfoArr = Array(Array(1, 2), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1))
    ret = OpenCSV("B_Profit_Loss_Account_Wife", "PLA.OtherIncomeListW", "7", True, 783, 816, "H", False, "", fieldInfoArr, 770, 776)
    

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub ImportCsvSecD1Self()
    Dim ret As Boolean
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1))
    ret = OpenCSV("D_Stock_Analysis", "Inv.QuantDtlsListS", "6", True, 20, 59, "F", False, "", fieldInfoArr, 5, 11)

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Sub ImportCsvSecD1Wife()
    isWifeDec = Sheet14.Range("RetInf.DeclareWifeIncome").value
    If UCase(isWifeDec) = "NO" Then
        MsgBox "You have opted not to declare Wife's income in A_Basic_Info"
    Else
        Dim ret As Boolean
        
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1))
        ret = OpenCSV("D_Stock_Analysis", "Inv.QuantDtlsListW", "6", True, 20, 59, "F", False, "", fieldInfoArr, 5, 11)
    
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
End Sub
Sub ImportCsvSecD2Self()
    Dim ret As Boolean
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1))
    ret = OpenCSV("D_Stock_Analysis", "Inv.QuantDtlsListKIIS", "9", True, 100, 148, "I", False, "", fieldInfoArr, 85, 94)

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Sub ImportCsvSecD2Wife()
    isWifeDec = Sheet14.Range("RetInf.DeclareWifeIncome").value
    If UCase(isWifeDec) = "NO" Then
        MsgBox "You have opted not to declare Wife's income in A_Basic_Info"
    Else
        Dim ret As Boolean
        
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1))
        ret = OpenCSV("D_Stock_Analysis", "Inv.QuantDtlsListKIIW", "9", True, 100, 148, "I", False, "", fieldInfoArr, 85, 94)
    
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
End Sub
Sub ImportCsvSecE1_1Self()
    rFromDate = Sheet14.Range("RetInf.RetStartDate").value
    rToDate = Sheet14.Range("RetInf.RetEndDate").value
   
    Dim year As String
    Dim mm As String
    
    If (rFromDate <> "" And rToDate <> "") Then
        Dim ret As Boolean
        rToDate = CDate(Format(Sheet14.Range("RetInf.RetEndDate").value, "dd/mm/yyyy"))
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        
        mm = Format(CDate(Trim(rToDate)), "MM")
        year = DatePart("yyyy", rToDate)
        
        If year >= 2020 Then
            If (year = 2020 And mm >= 4) Or year > 2020 Then
               fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), Array(9, 1), Array(10, 1))
                ret = OpenCSV("E1_IDA_CA", "tempIniPlanList", "9", True, 840, 882, "J", True, "G", fieldInfoArr, 825, 833)
            Else
                fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), Array(9, 1), Array(10, 1))
                ret = OpenCSV("E1_IDA_CA", "IniAllPlanMach.ListPart1S", "10", True, 198, 240, "J", True, "G", fieldInfoArr, 182, 191)
            End If
        Else
            fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), Array(9, 1), Array(10, 1))
            ret = OpenCSV("E1_IDA_CA", "IniAllPlanMach.ListPart1S", "10", True, 198, 240, "J", True, "G", fieldInfoArr, 182, 191)
        End If
        
        If ret Then
            Application.EnableEvents = True
            Dim startRow As Long, endRow, tempStr, endColumn As Long
            startRow = ActiveSheet.Range("tempIniPlanList").row
            endRow = ActiveSheet.Range("tempIniPlanList").Rows.Count + startRow - 1
            For i = startRow To endRow
            If (Worksheets("E1_IDA_CA").Cells(i, "B").value <> "") Then
                If Worksheets("E1_IDA_CA").Cells(i, "B").value = "Machinery" Then
                    If (Worksheets("E1_IDA_CA").Cells(i, "D").value = "Machinery used for Manufacture" Or Worksheets("E1_IDA_CA").Cells(i, "D").value = "Hospital Equipment" Or Worksheets("E1_IDA_CA").Cells(i, "D").value = "Ships or Aircraft") Then
                    Else
                        MsgBox "Please select Item Description as per the Item Category."
                        Worksheets("E1_IDA_CA").Cells(i, "D").value = ""
                    End If
                ElseIf Worksheets("E1_IDA_CA").Cells(i, "B").value = "Indefeasible Right" Then
                    If (Worksheets("E1_IDA_CA").Cells(i, "D").value <> "Fibre Optic Cable by Telecommunication Operator") Then
                        MsgBox "Please select Item Description as per the Item Category."
                        Worksheets("E1_IDA_CA").Cells(i, "D").value = ""
                    End If
                End If
            End If
            Next i
        End If
        'fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 1), Array(8, 1))
       ' ret = OpenCSV("E1_IDA_CA", "IniAllPlanMach.ListPart1S", "8", True, 198, 240, "H", True, "E", fieldInfoArr, 182, 189)
    
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    Else
        MsgBox "Please first enter the Return Period From and Return Period To."
        Worksheets("A_Basic_Info").Activate
        Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
    End If
End Sub
Sub ImportCsvSecE1_1Wife()
    isWifeDec = Sheet14.Range("RetInf.DeclareWifeIncome").value
    If UCase(isWifeDec) = "NO" Then
        MsgBox "You have opted not to declare Wife's income in A_Basic_Info"
    Else
        rFromDate = Sheet14.Range("RetInf.RetStartDate").value
        rToDate = Sheet14.Range("RetInf.RetEndDate").value
        Dim year As String
        Dim mm As String
        
        If (rFromDate <> "" And rToDate <> "") Then
            Dim ret As Boolean
            rToDate = CDate(Format(Sheet14.Range("RetInf.RetEndDate").value, "dd/mm/yyyy"))
            Application.ScreenUpdating = False
            Application.EnableEvents = False
            mm = Format(CDate(Trim(rToDate)), "MM")
            year = DatePart("yyyy", rToDate)
        
            If year >= 2020 Then
                If (year = 2020 And mm >= 4) Or year > 2020 Then
                   fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), Array(9, 1), Array(10, 1))
                    ret = OpenCSV("E1_IDA_CA", "tempIniPlanListW", "9", True, 840, 882, "J", True, "G", fieldInfoArr, 825, 833)
                Else
                    fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), Array(9, 1), Array(10, 1))
                    ret = OpenCSV("E1_IDA_CA", "IniAllPlanMach.ListPart1W", "10", True, 198, 240, "J", True, "G", fieldInfoArr, 182, 191)
                End If
            Else
                fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), Array(9, 1), Array(10, 1))
                ret = OpenCSV("E1_IDA_CA", "IniAllPlanMach.ListPart1W", "10", True, 198, 240, "J", True, "G", fieldInfoArr, 182, 191)
            End If
            
            
            If ret Then
            Application.EnableEvents = True
            Dim startRow As Long, endRow, tempStr, endColumn As Long
            startRow = ActiveSheet.Range("tempIniPlanListW").row
            endRow = ActiveSheet.Range("tempIniPlanListW").Rows.Count + startRow - 1
            For i = startRow To endRow
            If (Worksheets("E1_IDA_CA").Cells(i, "B").value <> "") Then
                If Worksheets("E1_IDA_CA").Cells(i, "B").value = "Machinery" Then
                    If (Worksheets("E1_IDA_CA").Cells(i, "D").value = "Machinery used for Manufacture" Or Worksheets("E1_IDA_CA").Cells(i, "D").value = "Hospital Equipment" Or Worksheets("E1_IDA_CA").Cells(i, "D").value = "Ships or Aircraft") Then
                    Else
                        MsgBox "Please select Item Description as per the Item Category."
                        Worksheets("E1_IDA_CA").Cells(i, "D").value = ""
                    End If
                ElseIf Worksheets("E1_IDA_CA").Cells(i, "B").value = "Indefeasible Right" Then
                    If (Worksheets("E1_IDA_CA").Cells(i, "D").value <> "Fibre Optic Cable by Telecommunication Operator") Then
                        MsgBox "Please select Item Description as per the Item Category."
                        Worksheets("E1_IDA_CA").Cells(i, "D").value = ""
                    End If
                End If
            End If
            Next i
            End If
            'fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 1), Array(8, 1))
            'ret = OpenCSV("E1_IDA_CA", "IniAllPlanMach.ListPart1W", "8", True, 198, 240, "H", True, "E", fieldInfoArr, 182, 189)
        
            Application.EnableEvents = True
            Application.ScreenUpdating = True
        Else
            MsgBox "Please first enter the Return Period From and Return Period To."
            Worksheets("A_Basic_Info").Activate
            Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
        End If
    End If
End Sub
Sub ImportCsvSecE1_2Self()
    rFromDate = Sheet14.Range("RetInf.RetStartDate").value
    rToDate = Sheet14.Range("RetInf.RetEndDate").value
    
    If (rFromDate <> "" And rToDate <> "") Then
        Dim ret As Boolean
        Dim year As String
        Dim mm As String
        
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        rToDate = CDate(Format(Sheet14.Range("RetInf.RetEndDate").value, "dd/mm/yyyy"))
        'fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1))
        'ret = OpenCSV("E1_IDA_CA", "IniAllIBD.ListPart2S", "8", True, 282, 319, "H", True, "B", fieldInfoArr, 266, 273)
        mm = Format(CDate(Trim(rToDate)), "MM")
        year = DatePart("yyyy", rToDate)
        
        If year >= 2020 Then
            If (year = 2020 And mm >= 4) Or year > 2020 Then
               fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), Array(9, 1), Array(10, 1))
                ret = OpenCSV("E1_IDA_CA", "tempIniAllIBDS", "9", True, 900, 941, "J", True, "D", fieldInfoArr, 889, 896)
            Else
                fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), Array(9, 1), Array(10, 1))
                ret = OpenCSV("E1_IDA_CA", "IniAllIBD.ListPart2S", "10", True, 278, 319, "J", True, "D", fieldInfoArr, 266, 275)
            End If
        Else
            fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), Array(9, 1), Array(10, 1))
            ret = OpenCSV("E1_IDA_CA", "IniAllIBD.ListPart2S", "10", True, 278, 319, "J", True, "D", fieldInfoArr, 266, 275)
        End If
        
        If ret Then
            Application.EnableEvents = True
            Dim startRow As Long, endRow, tempStr, endColumn As Long
            startRow = ActiveSheet.Range("tempIniAllIBDS").row
            endRow = ActiveSheet.Range("tempIniAllIBDS").Rows.Count + startRow - 1
            For i = startRow To endRow
            If (Worksheets("E1_IDA_CA").Cells(i, "B").value <> "") Then
                Dim lockCellFlag As Boolean
                If Worksheets("E1_IDA_CA").Cells(i, "B").value = "Construction of Bulk Storage and Handling Facilities (SGR)" Then
                    If (Worksheets("E1_IDA_CA").Cells(i, "C").value <> "") And Worksheets("E1_IDA_CA").Cells(i, "C").value < 100000 Then
                        MsgBox "Storage Capacity should be greater than or equal to 10000."
                        Worksheets("E1_IDA_CA").Unprotect (Pwd)
                        Worksheets("E1_IDA_CA").Cells(i, "C").value = ""
                        lockCellFlag = False
                        
                    End If
                Else
                    If (Worksheets("E1_IDA_CA").Cells(i, "C").value <> "") Then
                        MsgBox "Storage Capacity field should be blank."
                        Worksheets("E1_IDA_CA").Unprotect (Pwd)
                        Worksheets("E1_IDA_CA").Cells(i, "C").value = ""
                        Worksheets("E1_IDA_CA").Unprotect (Pwd)
                        Worksheets("E1_IDA_CA").Cells(i, "C").Locked = True
                        lockCellFlag = True
                    End If
                    If lockCellFlag = True Then
                            cellColor = RGB(146, 146, 146)
                        ElseIf lockCellFlag = False Then
                            cellColor = RGB(255, 255, 255)
                    End If
                    Worksheets("E1_IDA_CA").Unprotect (Pwd)
                    With Worksheets("E1_IDA_CA").Cells(startRow, "C").Interior
                                       .Color = cellColor
                                       .Pattern = xlSolid
                                       .PatternColorIndex = xlAutomatic
                            End With
                End If
                 
            End If
            Next i
            Application.EnableEvents = False
        End If
        Worksheets("E1_IDA_CA").Protect (Pwd)
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    Else
        MsgBox "Please first enter the Return Period From and Return Period To."
        Worksheets("A_Basic_Info").Activate
        Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
    End If
End Sub
Sub ImportCsvSecE1_2Wife()
    isWifeDec = Sheet14.Range("RetInf.DeclareWifeIncome").value
     Dim year As String
     Dim mm As String
    If UCase(isWifeDec) = "NO" Then
        MsgBox "You have opted not to declare Wife's income in A_Basic_Info"
    Else
        rFromDate = Sheet14.Range("RetInf.RetStartDate").value
        rToDate = Sheet14.Range("RetInf.RetEndDate").value
        
        If (rFromDate <> "" And rToDate <> "") Then
            Dim ret As Boolean
            Application.ScreenUpdating = False
            Application.EnableEvents = False
            rToDate = CDate(Format(Sheet14.Range("RetInf.RetEndDate").value, "dd/mm/yyyy"))
        'fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1))
        'ret = OpenCSV("E1_IDA_CA", "IniAllIBD.ListPart2S", "8", True, 282, 319, "H", True, "B", fieldInfoArr, 266, 273)
        mm = Format(CDate(Trim(rToDate)), "MM")
        year = DatePart("yyyy", rToDate)
        
        If year >= 2020 Then
            If (year = 2020 And mm >= 4) Or year > 2020 Then
               fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), Array(9, 1), Array(10, 1))
                ret = OpenCSV("E1_IDA_CA", "tempIniAllIBDW", "9", True, 900, 941, "J", True, "D", fieldInfoArr, 889, 896)
            Else
                fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), Array(9, 1), Array(10, 1))
                ret = OpenCSV("E1_IDA_CA", "IniAllIBD.ListPart2W", "10", True, 278, 319, "J", True, "D", fieldInfoArr, 266, 275)
            End If
        Else
            fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), Array(9, 1), Array(10, 1))
            ret = OpenCSV("E1_IDA_CA", "IniAllIBD.ListPart2W", "10", True, 278, 319, "J", True, "D", fieldInfoArr, 266, 275)
        End If
        
        If ret Then
            Application.EnableEvents = True
            Dim startRow As Long, endRow, tempStr, endColumn As Long
            startRow = ActiveSheet.Range("tempIniAllIBDW").row
            endRow = ActiveSheet.Range("tempIniAllIBDW").Rows.Count + startRow - 1
            For i = startRow To endRow
            If (Worksheets("E1_IDA_CA").Cells(i, "B").value <> "") Then
                Dim lockCellFlag As Boolean
                If Worksheets("E1_IDA_CA").Cells(i, "B").value = "Construction of Bulk Storage and Handling Facilities (SGR)" Then
                    If (Worksheets("E1_IDA_CA").Cells(i, "C").value <> "") And Worksheets("E1_IDA_CA").Cells(i, "C").value < 100000 Then
                        MsgBox "Storage Capacity should be greater than or equal to 10000."
                        Worksheets("E1_IDA_CA").Unprotect (Pwd)
                        lockCellFlag = False
                        Worksheets("E1_IDA_CA").Cells(i, "C").value = ""
                        Worksheets("E1_IDA_CA").Protect (Pwd)
                    End If
                Else
                    If (Worksheets("E1_IDA_CA").Cells(i, "C").value <> "") Then
                        MsgBox "Storage Capacity field should be blank."
                        Worksheets("E1_IDA_CA").Unprotect (Pwd)
                        Worksheets("E1_IDA_CA").Cells(i, "C").value = ""
                        Worksheets("E1_IDA_CA").Unprotect (Pwd)
                        Worksheets("E1_IDA_CA").Cells(i, "C").Locked = True
                        lockCellFlag = True
                        Worksheets("E1_IDA_CA").Protect (Pwd)
                    End If
                     If lockCellFlag = True Then
                            cellColor = RGB(146, 146, 146)
                        ElseIf lockCellFlag = False Then
                            cellColor = RGB(255, 255, 255)
                    End If
                    Worksheets("E1_IDA_CA").Unprotect (Pwd)
                    With Worksheets("E1_IDA_CA").Cells(startRow, "C").Interior
                                       .Color = cellColor
                                       .Pattern = xlSolid
                                       .PatternColorIndex = xlAutomatic
                            End With
                End If
            End If
            Next i
            Application.EnableEvents = False
        End If
            Application.EnableEvents = True
            Application.ScreenUpdating = True
        Else
            MsgBox "Please first enter the Return Period From and Return Period To."
            Worksheets("A_Basic_Info").Activate
            Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
        End If
    End If
End Sub
Sub ImportCsvSecE1_3Self()
    rFromDate = Sheet14.Range("RetInf.RetStartDate").value
    rToDate = Sheet14.Range("RetInf.RetEndDate").value
    If (rFromDate <> "" And rToDate <> "") Then
        Dim ret As Boolean
        
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1))
        ret = OpenCSV("E1_IDA_CA", "AgrLandDed.ListS", "6", True, 360, 398, "F", True, "B", fieldInfoArr, 343, 348)
    
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    Else
        MsgBox "Please first enter the Return Period From and Return Period To."
        Worksheets("A_Basic_Info").Activate
        Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
    End If
End Sub
Sub ImportCsvSecE1_3Wife()
    isWifeDec = Sheet14.Range("RetInf.DeclareWifeIncome").value
    If UCase(isWifeDec) = "NO" Then
        MsgBox "You have opted not to declare Wife's income in A_Basic_Info"
    Else
        rFromDate = Sheet14.Range("RetInf.RetStartDate").value
        rToDate = Sheet14.Range("RetInf.RetEndDate").value
        If (rFromDate <> "" And rToDate <> "") Then
            Dim ret As Boolean
            
            Application.ScreenUpdating = False
            Application.EnableEvents = False
            fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1))
            ret = OpenCSV("E1_IDA_CA", "AgrLandDed.ListW", "6", True, 360, 398, "F", True, "B", fieldInfoArr, 343, 348)
        
            Application.EnableEvents = True
            Application.ScreenUpdating = True
        Else
            MsgBox "Please first enter the Return Period From and Return Period To."
            Worksheets("A_Basic_Info").Activate
            Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
        End If
    End If
End Sub
Sub ImportCsvSecE1_4Self()
    rFromDate = Sheet14.Range("RetInf.RetStartDate").value
    rToDate = Sheet14.Range("RetInf.RetEndDate").value
    
    If (rFromDate <> "" And rToDate <> "") Then
        Dim ret As Boolean
        rToDate = CDate(Format(Sheet14.Range("RetInf.RetEndDate").value, "dd/mm/yyyy"))
         Dim year As String
        Dim mm As String
        
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        'fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1))
        'ret = OpenCSV("E1_IDA_CA", "DeprIntengAst.ListS", "7", True, 435, 475, "G", True, "D", fieldInfoArr, 421, 427)
        mm = Format(CDate(Trim(rToDate)), "MM")
        year = DatePart("yyyy", rToDate)
        
        If year >= 2020 Then
            If (year = 2020 And mm >= 4) Or year > 2020 Then
               fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1))
                ret = OpenCSV("E1_IDA_CA", "tempDeprS", "7", True, 961, 1001, "H", True, "E", fieldInfoArr, 949, 954)
            Else
                fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1))
                ret = OpenCSV("E1_IDA_CA", "DeprIntengAst.ListS", "8", True, 435, 475, "H", True, "E", fieldInfoArr, 421, 428)
            End If
        Else
            fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1))
            ret = OpenCSV("E1_IDA_CA", "DeprIntengAst.ListS", "8", True, 435, 475, "H", True, "E", fieldInfoArr, 421, 428)
        End If
    
        If ret Then
                Application.EnableEvents = True
                    Dim lockCellFlag As Boolean
                   
                Dim startRow As Long, endRow, tempStr, endColumn As Long
                startRow = ActiveSheet.Range("DeprIntengAst.ListS").row
                endRow = ActiveSheet.Range("DeprIntengAst.ListS").Rows.Count + startRow - 1
                Worksheets("E1_IDA_CA").Unprotect (Pwd)
                For i = startRow To endRow
                    If (Worksheets("E1_IDA_CA").Cells(i, "B").value <> "") Then
                    
                        If Worksheets("E1_IDA_CA").Cells(i, "B").value = "Other Machinery" Then
                             lockCellFlag = False
                             Worksheets("E1_IDA_CA").Cells(i, "C").Locked = False
                        Else
                             lockCellFlag = True
                             If (Worksheets("E1_IDA_CA").Cells(i, "C").value <> "") Then
                                MsgBox "Item Description should be blank."
                             End If
                             Worksheets("E1_IDA_CA").Cells(i, "C").value = ""
                             Worksheets("E1_IDA_CA").Unprotect (Pwd)
                             Worksheets("E1_IDA_CA").Cells(i, "C").Locked = True
                        End If
                        'to change the color of the cell
                        If lockCellFlag = True Then
                            cellColor = RGB(146, 146, 146)
                        ElseIf lockCellFlag = False Then
                            cellColor = RGB(255, 255, 255)
                        End If
                        With Worksheets("E1_IDA_CA").Cells(startRow, "C").Interior
                                   .Color = cellColor
                                   .Pattern = xlSolid
                                   .PatternColorIndex = xlAutomatic
                            End With
                    End If
                Next i
                  Application.EnableEvents = False
                  Worksheets("E1_IDA_CA").Protect (Pwd)
            End If
    
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    Else
        MsgBox "Please first enter the Return Period From and Return Period To."
        Worksheets("A_Basic_Info").Activate
        Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
    End If
End Sub
Sub ImportCsvSecE1_4Wife()
    isWifeDec = Sheet14.Range("RetInf.DeclareWifeIncome").value
    If UCase(isWifeDec) = "NO" Then
        MsgBox "You have opted not to declare Wife's income in A_Basic_Info"
    Else
        rFromDate = Sheet14.Range("RetInf.RetStartDate").value
        rToDate = Sheet14.Range("RetInf.RetEndDate").value
        
        If (rFromDate <> "" And rToDate <> "") Then
        Dim ret As Boolean
        Dim year As String
        Dim mm As String
        rToDate = CDate(Format(Sheet14.Range("RetInf.RetEndDate").value, "dd/mm/yyyy"))
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        'fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1))
        'ret = OpenCSV("E1_IDA_CA", "DeprIntengAst.ListS", "7", True, 435, 475, "G", True, "D", fieldInfoArr, 421, 427)
        mm = Format(CDate(Trim(rToDate)), "MM")
        year = DatePart("yyyy", rToDate)
        
        If year >= 2020 Then
            If (year = 2020 And mm >= 4) Or year > 2020 Then
               fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1))
                ret = OpenCSV("E1_IDA_CA", "tempDeprW", "7", True, 961, 1001, "H", True, "E", fieldInfoArr, 949, 954)
            Else
                fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1))
                ret = OpenCSV("E1_IDA_CA", "DeprIntengAst.ListW", "8", True, 435, 475, "H", True, "E", fieldInfoArr, 421, 428)
            End If
        Else
            fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1))
            ret = OpenCSV("E1_IDA_CA", "DeprIntengAst.ListW", "8", True, 435, 475, "H", True, "E", fieldInfoArr, 421, 428)
        End If
    
        If ret Then
                Application.EnableEvents = True
                    Dim lockCellFlag As Boolean
                   
                Dim startRow As Long, endRow, tempStr, endColumn As Long
                startRow = ActiveSheet.Range("DeprIntengAst.ListW").row
                endRow = ActiveSheet.Range("DeprIntengAst.ListW").Rows.Count + startRow - 1
                Worksheets("E1_IDA_CA").Unprotect (Pwd)
                For i = startRow To endRow
                    If (Worksheets("E1_IDA_CA").Cells(i, "B").value <> "") Then
                    
                        If Worksheets("E1_IDA_CA").Cells(i, "B").value = "Other Machinery" Then
                             lockCellFlag = False
                             Worksheets("E1_IDA_CA").Cells(i, "C").Locked = False
                        Else
                             lockCellFlag = True
                             If (Worksheets("E1_IDA_CA").Cells(i, "C").value <> "") Then
                                MsgBox "Item Description should be blank."
                             End If
                             Worksheets("E1_IDA_CA").Cells(i, "C").value = ""
                              Worksheets("E1_IDA_CA").Unprotect (Pwd)
                             Worksheets("E1_IDA_CA").Cells(i, "C").Locked = True
                        End If
                        'to change the color of the cell
                        If lockCellFlag = True Then
                            cellColor = RGB(146, 146, 146)
                        ElseIf lockCellFlag = False Then
                            cellColor = RGB(255, 255, 255)
                    End If
                        With Worksheets("E1_IDA_CA").Cells(startRow, "C").Interior
                                   .Color = cellColor
                                   .Pattern = xlSolid
                                   .PatternColorIndex = xlAutomatic
                            End With
                    End If
                Next i
                  Application.EnableEvents = False
                  Worksheets("E1_IDA_CA").Protect (Pwd)
            End If
           Application.EnableEvents = True
            Application.ScreenUpdating = True
            
        Else
            MsgBox "Please first enter the Return Period From and Return Period To."
            Worksheets("A_Basic_Info").Activate
            Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
        End If
    End If
End Sub
Sub ImportCsvSecE2Self()
    rFromDate = Sheet14.Range("RetInf.RetStartDate").value
    rToDate = Sheet14.Range("RetInf.RetEndDate").value
    If (rFromDate <> "" And rToDate <> "") Then
        Dim ret As Boolean
        
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1))
        ret = OpenCSV("E2_CA_WTA_SLM", "WAT.ListBS", "7", True, 520, 567, "G", True, "D", fieldInfoArr, 506, 512)
    
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    Else
        MsgBox "Please first enter the Return Period From and Return Period To."
        Worksheets("A_Basic_Info").Activate
        Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
    End If
End Sub
Sub ImportCsvSecE2Wife()
    isWifeDec = Sheet14.Range("RetInf.DeclareWifeIncome").value
    If UCase(isWifeDec) = "NO" Then
        MsgBox "You have opted not to declare Wife's income in A_Basic_Info"
    Else
        rFromDate = Sheet14.Range("RetInf.RetStartDate").value
        rToDate = Sheet14.Range("RetInf.RetEndDate").value
        If (rFromDate <> "" And rToDate <> "") Then
            Dim ret As Boolean
            
            Application.ScreenUpdating = False
            Application.EnableEvents = False
            fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1))
            ret = OpenCSV("E2_CA_WTA_SLM", "WAT.ListBW", "7", True, 520, 567, "G", True, "D", fieldInfoArr, 506, 512)
        
            Application.EnableEvents = True
            Application.ScreenUpdating = True
        Else
            MsgBox "Please first enter the Return Period From and Return Period To."
            Worksheets("A_Basic_Info").Activate
            Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
        End If
    End If
End Sub
'Sub ImportCsvSecOSelf()
'    Dim ret As Boolean
'
'    Application.ScreenUpdating = False
'    Application.EnableEvents = False
'    fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 1), Array(4, 2), Array(5, 1))
'    ret = OpenCSV("O_WHT_Credits", "WithHolding.ListS", "5", True, 608, 638, "E", True, "D", fieldInfoArr, 595, 599)
'
'    Call enableDisableBankDetails
'
'    Application.EnableEvents = True
'    Application.ScreenUpdating = True
'End Sub
'Sub ImportCsvSecOWife()
'    isWifeDec = Sheet14.Range("RetInf.DeclareWifeIncome").value
'    If UCase(isWifeDec) = "NO" Then
'        MsgBox "You have opted not to declare Wife's income in A_Basic_Info"
'    Else
'        Dim ret As Boolean
'
'        Application.ScreenUpdating = False
'        Application.EnableEvents = False
'        fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 1), Array(4, 2), Array(5, 1))
'        ret = OpenCSV("O_WHT_Credits", "WithHolding.ListW", "5", True, 608, 638, "E", True, "D", fieldInfoArr, 595, 599)
'
'        Call enableDisableBankDetails
'
'        Application.EnableEvents = True
'        Application.ScreenUpdating = True
'    End If
'End Sub
'Sub ImportCsvSecPSelf()
'    Dim ret As Boolean
'
'    Application.ScreenUpdating = False
'    Application.EnableEvents = False
'    fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 1), Array(5, 1), Array(6, 1))
'    ret = OpenCSV("P_Advance_Tax_Credits", "VehicleAdvTaxPaid.ListS", "4,5", True, 675, 715, "E", False, "", fieldInfoArr, 661, 667)
'    If ret Then
'        Dim startRow As Long, endRow, startColumn, endColumn As Long
'        startRow = ActiveSheet.Range("VehicleAdvTaxPaid.ListS").row
'        endRow = ActiveSheet.Range("VehicleAdvTaxPaid.ListS").Rows.Count + startRow - 1
'        Call enableDisableAfterImport("P_Advance_Tax_Credits", "A" & startRow & ":F" & endRow, "A", "F", "D,E")
'    End If
'
'    Call enableDisableBankDetails
'
'    Application.EnableEvents = True
'    Application.ScreenUpdating = True
'End Sub
'Sub ImportCsvSecPWife()
'    isWifeDec = Sheet14.Range("RetInf.DeclareWifeIncome").value
'    If UCase(isWifeDec) = "NO" Then
'        MsgBox "You have opted not to declare Wife's income in A_Basic_Info"
'    Else
'        Dim ret As Boolean
'
'        Application.ScreenUpdating = False
'        Application.EnableEvents = False
'        fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 1), Array(5, 1), Array(6, 1))
'        ret = OpenCSV("P_Advance_Tax_Credits", "VehicleAdvTaxPaid.ListW", "4,5", True, 675, 715, "E", False, "", fieldInfoArr, 661, 667)
'        If ret Then
'            Dim startRow As Long, endRow, startColumn, endColumn As Long
'            startRow = ActiveSheet.Range("VehicleAdvTaxPaid.ListW").row
'            endRow = ActiveSheet.Range("VehicleAdvTaxPaid.ListW").Rows.Count + startRow - 1
'            Call enableDisableAfterImport("P_Advance_Tax_Credits", "A" & startRow & ":F" & endRow, "A", "F", "D,E")
'        End If
'
'        Call enableDisableBankDetails
'
'        Application.EnableEvents = True
'        Application.ScreenUpdating = True
'    End If
'End Sub


Sub enableDisableAfterImport(sheetName As String, rangeName As String, startColName As String, endColName As String, edColumsNames As String)
    
        ActiveWorkbook.Worksheets(sheetName).Unprotect Password:=Pwd
        Dim startRow, endRow As Long
        Dim startColumn, endColumn As Long
        startRow = ActiveWorkbook.Worksheets(sheetName).Range(rangeName).row
        endRow = ActiveWorkbook.Worksheets(sheetName).Range(rangeName).Rows.Count + startRow - 1
        startColumn = ActiveWorkbook.Worksheets(sheetName).Range(rangeName).column
        endColumn = ActiveWorkbook.Worksheets(sheetName).Range(rangeName).Columns.Count + startRow - 1
        Dim startColString, endColString As String
        'for input data less than 10000 rows due to go to special selection range constraint
        Dim blockStartRow, blockEndRow As Long
        blockStartRow = startRow
        blockEndRow = endRow
        Dim supposedRange As String
        Dim colNames As Variant
        
        If endRow <= 10000 Then
            supposedRange = ""
            colNames = Split(edColumsNames, ",")
            For intIndex = LBound(colNames) To UBound(colNames)
                If intIndex = LBound(colNames) Then
                    supposedRange = supposedRange + colNames(intIndex) & blockStartRow & ":" & colNames(intIndex) & blockEndRow
                Else
                    supposedRange = supposedRange + "," & colNames(intIndex) & blockStartRow & ":" & colNames(intIndex) & blockEndRow
                End If
            Next
            'ActiveWorkbook.Worksheets(SheetName).Range("B" & blockStartRow & ":B" & blockEndRow & ",H" & blockStartRow & ":H" & blockEndRow & ",I" & blockStartRow & ":I" & blockEndRow).Select
            ActiveWorkbook.Worksheets(sheetName).Range(supposedRange).Select
            'Range("H" & startRow).Activate
            On Error GoTo Next1
            Selection.SpecialCells(xlCellTypeBlanks).Select
            Selection.Locked = True
            Selection.Interior.Color = RGB(146, 146, 146)
            Selection.Interior.Pattern = xlSolid
            Selection.Interior.PatternColorIndex = xlAutomatic
Next1:
             Resume NextLine1
NextLine1:
            'ActiveWorkbook.Worksheets(SheetName).Range("B" & blockStartRow & ":B" & blockEndRow & ",H" & blockStartRow & ":H" & blockEndRow & ",I" & blockStartRow & ":I" & blockEndRow).Select
            ActiveWorkbook.Worksheets(sheetName).Range(supposedRange).Select
            'Range("H" & startRow).Activate
            On Error GoTo Next2
            Selection.SpecialCells(xlCellTypeConstants).Select
            Selection.Locked = False
            Selection.Interior.Color = RGB(255, 255, 255)
            Selection.Interior.Pattern = xlSolid
            Selection.Interior.PatternColorIndex = xlAutomatic
Next2:
        Resume NextLine2
NextLine2:
        Else
            blockStartRow = startRow
            blockEndRow = 10000 + startRow
            If blockEndRow > endRow Then
                blockEndRow = endRow
            End If
            supposedRange = ""
            colNames = Split(edColumsNames, ",")
            Do While blockEndRow <= endRow
                For intIndex = LBound(colNames) To UBound(colNames)
                    If intIndex = LBound(colNames) Then
                        supposedRange = supposedRange + colNames(intIndex) & blockStartRow & ":" & colNames(intIndex) & blockEndRow
                    Else
                        supposedRange = supposedRange + "," & colNames(intIndex) & blockStartRow & ":" & colNames(intIndex) & blockEndRow
                    End If
                Next
                'ActiveWorkbook.Worksheets(SheetName).Range("B" & blockStartRow & ":B" & blockEndRow & ",H" & blockStartRow & ":H" & blockEndRow & ",I" & blockStartRow & ":I" & blockEndRow).Select
                ActiveWorkbook.Worksheets(sheetName).Range(supposedRange).Select
                'Range("H" & startRow).Activate
                On Error GoTo Next3
                Selection.SpecialCells(xlCellTypeBlanks).Select
                Selection.Locked = True
                Selection.Interior.Color = RGB(146, 146, 146)
                Selection.Interior.Pattern = xlSolid
                Selection.Interior.PatternColorIndex = xlAutomatic
Next3:
                Resume NextLine3
NextLine3:
                'ActiveWorkbook.Worksheets(SheetName).Range("B" & blockStartRow & ":B" & blockEndRow & ",H" & blockStartRow & ":H" & blockEndRow & ",I" & blockStartRow & ":I" & blockEndRow).Select
                ActiveWorkbook.Worksheets(sheetName).Range(supposedRange).Select
                'Range("H" & startRow).Activate
                On Error GoTo Next4
                Selection.SpecialCells(xlCellTypeConstants).Select
                Selection.Locked = False
                Selection.Interior.Color = RGB(255, 255, 255)
                Selection.Interior.Pattern = xlSolid
                Selection.Interior.PatternColorIndex = xlAutomatic
Next4:
                Resume NextLine4
NextLine4:
                If blockEndRow = endRow Then
                    Exit Do
                End If
                If blockEndRow + 10000 > endRow Then
                    blockEndRow = endRow
                Else
                blockEndRow = blockEndRow + 10000
                End If
                If blockStartRow + 10000 > blockEndRow Then
                blockStartRow = blockEndRow
                Else
                    blockStartRow = blockStartRow + 10000
                End If
                supposedRange = ""
            Loop
        End If
        ActiveWorkbook.Worksheets(sheetName).Protect Password:=Pwd
    'End If
End Sub


Function OpenCSV(sheetName As String, rangeName As String, enterableColumnsCountString As String, executeCriteria As Boolean, critStartRow As Long, critEndRow As Long, lastColName As String, isDatePresent As Boolean, DateColumns As String, fieldInfoArr As Variant, Optional criteriaRangeS As Long, Optional criteriaRangeE As Long) As Boolean
    
    OpenCSV = True
    Dim PathZipProgram As String, strFileName As String
    Dim ShellStr As String, strDate As String, DefPath As String
    Dim Fld As Object
    Dim d_sales_count, d_export_count As Long
    Dim startPasteColumn As String
    startPasteColumn = "A"
    endPasteColumn = lastColName
    maxRecordsCount = 50000
    'strFileName = Application.GetOpenFilename("CSV (Comma Delimited) (*.csv,*.txt), *.csv,*.txt")
    strFileName = Application.GetOpenFilename("CSV (Comma Delimited) (*.csv), *.csv")
    
    'Added by Atul to put longest row marker if the longest row ends with blank cell
    If Dir(strFileName) <> "" Then
        PutLLMarker (strFileName)
    End If
    'Added by Atul to put longest row marker if the longest row ends with blank cell
    
    Dim i As Long
    'change this next line to reflect the actual directory
    Const strDir = "d:\"
    Dim ThisWB As Workbook
    Dim wb As Workbook
    Dim WS As Worksheet
    Dim strWS As String
    Dim importFunctionCalledOnSheet As String
    Set ThisWB = ActiveWorkbook
    Dim startRow As Long, endRow, startColumn, endColumn As Long
    startRow = ThisWB.Worksheets(sheetName).Range(rangeName).row
    endRow = ThisWB.Worksheets(sheetName).Range(rangeName).Rows.Count + startRow - 1
    startColumn = ThisWB.Worksheets(sheetName).Range(rangeName).column
    endColumn = ThisWB.Worksheets(sheetName).Range(rangeName).Columns.Count + startColumn - 1
    importFunctionCalledOnSheet = ThisWB.Worksheets(sheetName).name
    Dim StartTime, endTime As Double
    
    StartTime = Timer
       
    If Dir(strFileName) <> "" Then
        
        Set wb = Workbooks.Open(strFileName)
        Set wb = Application.Workbooks.Open(fileName:=strFileName, Format:=xlDelimited, Local:=True)
        
        'Workbooks.OpenText(strFileName @ @ xlDelimited xlDoubleQuote 0 -1 0 0 0 0 @ VbsEval("Array(Array(1, 1), Array(2, 2), Array(3, 4))"))
        If wb.Sheets(1).UsedRange.Rows.Count > maxRecordsCount Then
            MsgBox "Csv file contains more than " & maxRecordsCount & " records. Please select proper csv for this schedule."
            wb.Close SaveChanges:=False
            OpenCSV = False
            GoTo exitFunc
            'Exit Function
        End If
        
        'Import csv total records check for Section D
        If (rangeName = "PLA.OtherExpensesListS" Or rangeName = "PLA.OtherExpensesListW" Or rangeName = "PLA.OtherIncomeListS" Or rangeName = "PLA.OtherIncomeListW") Then
          
            If wb.Sheets(1).UsedRange.Columns.Count = 7 And (rangeName = "PLA.OtherExpensesListS" Or rangeName = "PLA.OtherExpensesListW" Or rangeName = "PLA.OtherIncomeListS" Or rangeName = "PLA.OtherIncomeListW") Then
                startPasteColumn = "B"
            End If
        End If
        
        Set wb = Application.ActiveWorkbook
        Dim colsCount As Variant
        Dim colsCsvMatchReqCols As Boolean
        colsCsvMatchReqCols = False
        colscolunt = Split(enterableColumnsCountString, ",")
        For i = LBound(colscolunt) To UBound(colscolunt)
            If wb.Sheets(1).UsedRange.Columns.Count = CLng(colscolunt(i)) Then
                colsCsvMatchReqCols = True
                Exit For
            End If
        Next i
        If colsCsvMatchReqCols = False Then
            MsgBox "Please select a proper csv file for this schedule."
            wb.Close SaveChanges:=False
            OpenCSV = False
            'Exit Function
            GoTo exitFunc
        End If
        
        'Remove LLMarker string before proceeding with the data
        wb.Sheets(1).UsedRange.Replace What:=sLLMarker, Replacement:="", LookAt:=xlWhole
        
        Dim dataValidationErrors As String
        If executeCriteria Then
            dataValidationErrors = validateImportSchedule(critStartRow, critEndRow, lastColName, ThisWB, wb, criteriaRangeS, criteriaRangeE)
        Else
            dataValidationErrors = ""
        End If
        If dataValidationErrors <> "" Then
            MsgBox2 dataValidationErrors
            wb.Close 'SaveChanges:=False
            ThisWB.Activate
            ThisWB.Worksheets(sheetName).Activate
            OpenCSV = False
            GoTo exitFunc
            'Exit Function
        End If
                
        If wb.Sheets(1).UsedRange.Rows.Count > ThisWB.Worksheets(sheetName).Range(rangeName).Rows.Count Then
            ThisWB.Activate
            ThisWB.Worksheets(sheetName).Activate
            Call InsertGivenRowsAndFillFormulas(rangeName, wb.Sheets(1).UsedRange.Rows.Count - ThisWB.Worksheets(sheetName).Range(rangeName).Rows.Count)
            InsertRowsTime = Timer
        End If
    
        startRow = ThisWB.Worksheets(sheetName).Range(rangeName).row
        endRow = ThisWB.Worksheets(sheetName).Range(rangeName).Rows.Count + startRow - 1
        startColumn = ThisWB.Worksheets(sheetName).Range(rangeName).column
        endColumn = ThisWB.Worksheets(sheetName).Range(rangeName).Columns.Count + startColumn - 1
        
        wb.Close SaveChanges:=False
        
        'restore original content of CSV file
        resetFile (strFileName)
        
        sSourcePath = strFileName
        Dim temp As Variant
        temp = Split(strFileName, ".")
        txtFilePath = temp(0) & ".txt"
        FileCopy sSourcePath, txtFilePath
        'fieldInfoArr = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), Array(9, 2), Array(10, 2), Array(11, 2))
        Workbooks.OpenText fileName:=txtFilePath, DataType:=xlDelimited, Comma:=True, FieldInfo:=fieldInfoArr, Local:=True
        
        Set wb = Application.ActiveWorkbook
        wb.Activate
       
        ThisWB.Worksheets(sheetName).Unprotect Password:=Pwd
        
        wb.Worksheets(wb.Sheets(1).name).UsedRange.Copy
        ThisWB.Worksheets(sheetName).Unprotect Password:=Pwd
        'by Vaishali for Enh-6
        If startColumn = "2" Then
            startPasteColumn = "B"
        End If
        
        If wb.Sheets(1).UsedRange.Rows.Count < (endRow - startRow + 1) Then
            ThisWB.Worksheets(sheetName).Range(startPasteColumn & startRow & ":" & endPasteColumn & (startRow + wb.Sheets(1).UsedRange.Rows.Count - 1)).PasteSpecial Paste:=xlValues, SkipBlanks:=True
        Else
            ThisWB.Worksheets(sheetName).Range(startPasteColumn & startRow & ":" & endPasteColumn & endRow).PasteSpecial Paste:=xlValues, SkipBlanks:=True
        End If
        
        
        'commented by maulika as not required now
'        If isDatePresent Then
'            Dim splittedColums As Variant
'            splittedColums = Split(DateColumns, ",")
'            Dim dateRange As String
'            dateRange = ""
'            For i = LBound(splittedColums) To UBound(splittedColums)
'                If i = LBound(splittedColums) Then
'                    dateRange = dateRange & splittedColums(i) & startRow & ":" & splittedColums(i) & endRow
'                Else
'                    dateRange = dateRange + "," & splittedColums(i) & startRow & ":" & splittedColums(i) & endRow
'                End If
'            Next
'
'            ThisWB.Activate
'            ThisWB.Worksheets(SheetName).Range(dateRange).Select
'            'for date
'            'Selection.Replace What:="""", Replacement:="", LookAt:=xlPart, _
'            '    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
'            '    ReplaceFormat:=False
'            'Selection.NumberFormat = "dd\/mm\/yyyy"
'            'Selection.NumberFormat = "mm\/dd\/yyyy"
'        End If
        
        'Application.DisplayAlerts = True
        ThisWB.Activate
        ThisWB.Worksheets(sheetName).Protect Password:=Pwd
        wb.Close SaveChanges:=False
        On Error Resume Next
        Kill txtFilePath

        endTime = Timer
        ThisWB.Worksheets(sheetName).Protect Password:=Pwd
        
        'MsgBox "Total Time Taken : " & (endTime - StartTime)
    Else
        MsgBox "There were no csv files found."
        OpenCSV = False
    End If
exitFunc:
    If Dir(strFileName) <> "" Then
        resetFile (strFileName)
    End If
End Function

Function validateImportSchedule(startRow As Long, endRow As Long, lastColName As String, mainWb As Workbook, csvWB As Workbook, Optional criteriaStartRow As Long, Optional criteriaEndRow As Long) As String
    Application.ScreenUpdating = False
    validateImportSchedule = ""
    Dim numberofInitialRows As Long
    Dim alertMsg As String
    Dim criteriaRow As Long
    Dim finalAlertMsg As String
    Dim errorNumber As Long
    Dim rangeArray() As String
    Dim errorRowNumber As String
    Dim errorCount As Long
    Dim rowNum As Long
    Dim countVisible As Long
    Dim errorColumnLetter As String
    Dim firstRow, lastRow, totRows As Long
    Dim critRange As Range
    Dim startRowNum As Long
    
    errorNumber = 0
    errorRowNumber = vbNullString
    csvWB.Sheets(1).Activate
    numberofInitialRows = ActiveSheet.UsedRange.Rows.Count
    Rows("1:" & (endRow - startRow + 1)).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    Range("A1").value = ""
    mainWb.Activate
    mainWb.Unprotect Password:=Pwd
    mainWb.Sheets("ImportCsv").Visible = xlSheetVisible
    mainWb.Sheets("ImportCsv").Activate
    Rows(startRow & ":" & endRow).Select
    Selection.Copy
      
    endRow = endRow - startRow + 1
    csvWB.Sheets(1).Paste
    csvWB.Activate
    csvWB.Sheets(1).Activate
    criteriaRow = criteriaStartRow
   ' Criteria selected from Import csv
   Do While (criteriaRow <= criteriaEndRow)
        Range("A" & endRow & ":" & lastColName & (endRow + numberofInitialRows)).Select
        Selection.AutoFilter
        Range("A" & endRow & ":" & lastColName & (endRow + numberofInitialRows)).Select
        Set critRange = Range(mainWb.Worksheets("ImportCsv").Range("C" & criteriaRow & ":C" & criteriaRow).value)
        errorColumnLetter = mainWb.Worksheets("ImportCsv").Range("D" & criteriaRow & ":D" & criteriaRow).value
        Selection.AdvancedFilter Action:=xlFilterInPlace, criteriaRange:=critRange, Unique:=False
        alertMsg = mainWb.Worksheets("ImportCsv").Cells(mainWb.Worksheets("ImportCsv").Range("B" & criteriaRow & ":B" & criteriaRow).value + 1, 1).value
        
        countVisible = 0
        errorCount = 0
        If ActiveSheet.UsedRange.Rows.Count > 10000 Then
            lastRow = 10000
        Else
            lastRow = endRow + numberofInitialRows
        End If
        
        errorRowNumber = vbNullString
        firstRow = 1
        totRows = ActiveSheet.UsedRange.Rows.Count
        Do While lastRow <= totRows
            Erase rangeArray
            ActiveSheet.Range("A" & firstRow & ":" & lastColName & lastRow).Select
            On Error Resume Next
            With Selection.SpecialCells(xlVisible)
                countVisible = countVisible + Selection.SpecialCells(xlVisible).Cells.Count
            End With
            If InStr(ActiveSheet.Range("A" & firstRow & ":" & lastColName & lastRow).SpecialCells(xlVisible).Rows.Address, ",") <> 0 Then
                rangeArray = Split(ActiveSheet.Range("A" & firstRow & ":" & lastColName & lastRow).SpecialCells(xlVisible).Rows.Address, ",")
                If (Err.Number = 0) Then
                    For i = 0 To UBound(rangeArray)
                        cellArray = Split(rangeArray(i), ":")
                        columnArray1 = Split(cellArray(0), "$")
                        columnArray2 = Split(cellArray(1), "$")
                        If (CLng(columnArray1(2)) - endRow) > 0 Then
                            startRowNum = CLng(columnArray1(2)) - endRow
                        Else
                            startRowNum = 1
                        End If
                        rowNum = (CLng(columnArray2(2)) - endRow)
                        If (rowNum > 0) Then
                            For j = startRowNum To rowNum
                                highlightCell = j + endRow
                                ActiveSheet.Range(errorColumnLetter & highlightCell & ":" & errorColumnLetter & highlightCell).Cells.Interior.Color = RGB(255, 0, 0)
                                If errorRowNumber = vbNullString Then
                                    errorCount = errorCount + 1
                                    errorRowNumber = j
                                Else
                                    If errorCount > 49 Then
                                        errorRowNumber = errorRowNumber & " ..."
                                        Exit Do
                                    End If
                                    errorCount = errorCount + 1
                                    errorRowNumber = errorRowNumber & "," & j
                                End If
                            Next
                        End If
                    Next
                End If
            Else
                Erase cellArray
                cellArray = Split(Selection.SpecialCells(xlVisible).Rows.Address, ":")
                columnArray2 = Split(cellArray(1), "$")
                columnArray1 = Split(cellArray(0), "$")
                If ((CLng(columnArray2(2)) - endRow) > 1) Then
                    rowNum = CLng(columnArray2(2)) - endRow
                    For j = 1 To rowNum
                        highlightCell = j + endRow
                        ActiveSheet.Range(errorColumnLetter & highlightCell & ":" & errorColumnLetter & highlightCell).Cells.Interior.Color = RGB(255, 0, 0)
                        If errorRowNumber = vbNullString Then
                            errorCount = errorCount + 1
                            errorRowNumber = j
                        Else
                            If errorCount > 49 Then
                                errorRowNumber = errorRowNumber & " ..."
                                Exit Do
                            End If
                            errorCount = errorCount + 1
                            errorRowNumber = errorRowNumber & "," & j
                        End If
                    Next
                Else
                    highlightCell = CLng(columnArray2(2))
                    ActiveSheet.Range(errorColumnLetter & highlightCell & ":" & errorColumnLetter & highlightCell).Cells.Interior.Color = RGB(255, 0, 0)
                    rowNum = (CLng(columnArray2(2)) - endRow)
                    If rowNum > 0 Then
                        errorRowNumber = rowNum
                    End If
                End If
            End If
             
            If lastRow = totRows Then
exitdo:
                Exit Do
            End If
            
            firstRow = lastRow + 1
            If lastRow + 10000 <= totRows Then
                lastRow = lastRow + 10000
            Else
                lastRow = totRows
            End If
        Loop

        If countVisible <> Range("A1:" & lastColName & (endRow)).Cells.Count Then 'And countVisible <> 0 Then
            If (errorNumber = 0) Then
                finalAlertMsg = "Please rectify the following errors found in the CSV file mentioned below :" & Chr(13)
            End If
            errorNumber = errorNumber + 1
            finalAlertMsg = finalAlertMsg & Chr(13) & errorNumber & ". " & alertMsg & errorRowNumber & Chr(13)
        End If
        
        criteriaRow = criteriaRow + 1
        Rows(endRow & ":" & (endRow + numberofInitialRows)).Select
        Selection.AutoFilter
    Loop
    
    If Not IsEmpty(alertMsg) Then
        validateImportSchedule = finalAlertMsg
    End If
    
    Rows("1:" & endRow).Select
    Selection.Delete Shift:=xlUp
    
    mainWb.Sheets("ImportCsv").Protect Password:=Pwd
    mainWb.Sheets("ImportCsv").Visible = xlSheetHidden
    mainWb.Protect Password:=Pwd
    Application.ScreenUpdating = True
    End Function
Function ColumnLetter(ColumnNumber As Long) As String


  If ColumnNumber > 26 Then

    ' 1st character:  Subtract 1 to map the characters to 0-25,
    '                 but you don't have to remap back to 1-26
    '                 after the 'Int' operation since columns
    '                 1-26 have no prefix letter

    ' 2nd character:  Subtract 1 to map the characters to 0-25,
    '                 but then must remap back to 1-26 after
    '                 the 'Mod' operation by adding 1 back in
    '                 (included in the '65')

    ColumnLetter = Chr(Int((ColumnNumber - 1) / 26) + 64) & _
                   Chr(((ColumnNumber - 1) Mod 26) + 65)
  Else
    ' Columns A-Z
    ColumnLetter = Chr(ColumnNumber + 64)
  End If
End Function

Sub ShowInNotepad(sPath)
Shell "notepad.exe """ & sPath & """", vbNormalFocus
End Sub

Function MsgBox2(Prompt As String, _
    Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
    Optional Title As String = "Microsoft Excel", _
    Optional HelpFile As String, _
    Optional Context As Long) As VbMsgBoxResult
     '
     '****************************************************************************************
     '       Title       MsgBox2
     '       Target Application: any
     '       Function:   substitute for standard MsgBox; displays more text than the ~1024 character
     '                   limit of MsgBox.  Displays blocks of approx 900 characters (properly split
     '                   at blanks or line feeds or "returns" and adds some "special text" to suggest
     '                   that more text is coming for each block except the last.  Special text is
     '                   easily changed.
     '
     '                   An EndOfBlack separator is also supported.  If found, MsgBox2 will only
     '                   display the characters through the EndOfBlock separator.  This provides
     '                   complete control over how text is displayed.  The current separator is
     '                   "||".
     '       Limitations:  the optional values for MsgBox display, i.e., Buttons, Title, HelpFile,
     '                     and Context  are the same for each block of text displayed.
     '       Passed Values:  same arguement list and type as standard MsgBox
     '
     '****************************************************************************************
     '
     '
    Dim CurLocn         As Long
    Dim EndOfBlock      As String
    Dim EOBIndex        As Long
    Dim EOBLen          As Long
    Dim Index           As Long
    Dim MaxLen          As Long
    Dim OldIndex        As Long
    Dim strMoreToCome   As String
    Dim strTemp         As String
    Dim ThisChar        As String
    Dim TotLen          As Long
     
     '
     '           set procedure variable that control how/what procedure does:
     '
     '       EndOfBlock is the string variable containing the character or characters
     '           that denote the end of a block of text.  These characters are not displayed.
     '           Do not use a character or characters that might be used in normal text.
     '       MaxLen is the maximum number of characters to be displayed at one time.  The
     '           limit for MsgBox is approx 1024, but that depends on the particular chars
     '           in the prompt string.  900 is a safe number as long as the len(strMoreToCome)
     '           is reasonable.
     '       strMoreToCome is text displayed at the bottom of each block indicating that more
     '           text/data is coming.
     '
    EndOfBlock = "||"
    MaxLen = 875
    strMoreToCome = "... press SPACE to see next errors ..."
     
    EOBLen = Len(EndOfBlock)
    CurLocn = 0
    OldIndex = 1
    TotLen = 0
     
NextBlock:
     '
     '           test for special break and, if found, that it is not the last chars in Prompt
     '
    EOBIndex = InStr(1, Mid(Prompt, OldIndex, MaxLen), EndOfBlock)
    If EOBIndex > 0 And CurLocn < Len(Prompt) - 1 Then
        CurLocn = EOBIndex + OldIndex - 1
        strTemp = Mid(Prompt, OldIndex, CurLocn - OldIndex)
        TotLen = TotLen + Len(strTemp) + EOBLen
        OldIndex = CurLocn + EOBLen
        GoTo MidDisplay
    End If
     '
     '           no special break, handle as normal block
     '
    Index = OldIndex + MaxLen
     '
     '           test for last block
     '
    If Index > Len(Prompt) Then
        strTemp = Mid(Prompt, OldIndex, Len(Prompt) - OldIndex + 1)
LastDisplay:
        MsgBox2 = MsgBox(strTemp, Buttons, Title, HelpFile, Context)
        Exit Function
    End If
     '
     '           not last display; process block
     '
    CurLocn = Index
NextIndex:
    ThisChar = Mid(Prompt, CurLocn, 1)
    If ThisChar = " " Or _
    ThisChar = Chr(10) Or _
    ThisChar = Chr(13) Then
         '
         '           block break found
         '
        strTemp = Mid(Prompt, OldIndex, CurLocn - OldIndex + 1)
        TotLen = TotLen + Len(strTemp)
        OldIndex = CurLocn + 1
MidDisplay:
         '
         '           display current block of text appending string indicating that
         '           more text is to come.  Then test if user hit Cancel button or
         '           equivalent; if so, exit MsgBox2 without further processing
         '
        MsgBox2 = MsgBox(strTemp & vbCrLf & strMoreToCome, _
        Buttons, Title, HelpFile, Context)
        If MsgBox2 = vbCancel Then Exit Function
        GoTo NextBlock
    End If
    CurLocn = CurLocn - 1
    If CurLocn > OldIndex Then GoTo NextIndex
     '
     '           no blanks, CR's, LF's or special breaks found in previous block
     '           display these characters and move on
     '
    strTemp = Mid(Prompt, OldIndex, MaxLen)
    CurLocn = OldIndex + MaxLen
    TotLen = TotLen + Len(strTemp)
    OldIndex = CurLocn + 1
    GoTo MidDisplay
     
End Function
'This function puts a marker at the end of the longest row (in terms of no. of columns) if it ends with a blank cell i.e. comma (,)
Sub PutLLMarker(sFileName As String)

Dim sBuf As String
Dim sTemp As String
Dim iFileNum As Integer

iFileNum = FreeFile
Open sFileName For Input As iFileNum
iMaxComma = -1
Do Until EOF(iFileNum)
    Line Input #iFileNum, sBuf
    
    
    Dim iCountComma As Long
    Dim sParts() As String

    sParts = Split(sBuf, ",")

    iResult = UBound(sParts, 1)

    If (iResult = -1) Then
    iResult = 0
    End If

    iCountComma = iResult
    
    If iMaxComma = -1 Then
        iMaxComma = iCountComma
        
'        If Left(sBuf, 1) = "," Then
'            sBuf = sLLMarker & sBuf
'            isFileModified = True
'        End If
        If Right(sBuf, 1) = "," Then
            sBuf = sBuf & sLLMarker
            isFileModified = True
        End If
    ElseIf iCountComma > iMaxComma Then
        iMaxComma = iCountComma
        If isFileModified Then
            sTemp = Replace(sTemp, sLLMarker, "")
        End If
        If Right(sBuf, 1) = "," Or Left(sBuf, 1) = "," Then
'            If Left(sBuf, 1) = "," Then
'                sBuf = sLLMarker & sBuf
'                isFileModified = True
'            End If
            
            If Right(sBuf, 1) = "," Then
                sBuf = sBuf & sLLMarker
                isFileModified = True
            End If
            
        Else
            isFileModified = False
        End If
    End If
    
    sTemp = sTemp & sBuf & vbCrLf
Loop
Close iFileNum

Dim sLines() As String
sLines = Split(sTemp, vbCrLf)
sTemp = ""
For i = 0 To UBound(sLines) - 2
    sTemp = sTemp & sLines(i) & vbCrLf
Next
sTemp = sTemp & sLines(i)

'sTemp = Trim(sTemp)

iFileNum = FreeFile
Open sFileName For Output As iFileNum
Print #iFileNum, sTemp
Close iFileNum

End Sub

Sub resetFile(sFileName As String)
If isFileModified Then
    Dim sBuf As String
    Dim sTemp As String
    Dim iFileNum As Integer
    
    'sFileName = "C:\Users\765203\Desktop\test.csv"
    
    iFileNum = FreeFile
    Open sFileName For Input As iFileNum
    Do Until EOF(iFileNum)
        Line Input #iFileNum, sBuf
        sTemp = sTemp & sBuf & vbCrLf
    Loop
    Close iFileNum
    sTemp = Trim(sTemp)
    sTemp = Replace(sTemp, sLLMarker, "")
    
    
    Dim sLines() As String
    sLines = Split(sTemp, vbCrLf)
    sTemp = ""
    For i = 0 To UBound(sLines) - 2
        sTemp = sTemp & sLines(i) & vbCrLf
    Next
    sTemp = sTemp & sLines(i)
    
    
    iFileNum = FreeFile
    Open sFileName For Output As iFileNum
    Print #iFileNum, sTemp
    Close iFileNum
    isFileModified = False
End If
End Sub






