VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Add Row Functionality For Self Start
Sub addRow_InsReliefDtls_ListS()
    Worksheets("L_Computation_of_Insu_Relief").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("InsReliefDtls.ListS")
    Worksheets("L_Computation_of_Insu_Relief").Protect (Pwd)
End Sub
'Add Row Functionality For Self End

'Add Row Functionality For Wife Start
Sub addRow_InsReliefDtls_ListW()
    Worksheets("L_Computation_of_Insu_Relief").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("InsReliefDtls.ListW")
    Worksheets("L_Computation_of_Insu_Relief").Protect (Pwd)
End Sub
'Add Row Functionality For Wife End

Private Sub Worksheet_Change(ByVal Target As Range)
   
    
    Worksheets("L_Computation_of_Insu_Relief").Unprotect (Pwd)
    Worksheets("L_Computation_of_Insu_Relief").Activate
    Application.EnableEvents = False
    Application.ScreenUpdating = False
   
    If (Target.Columns.Count = 1) Then
        PINNoS = Sheet14.Range("RetInf.PIN").value
        PINNoW = Sheet14.Range("RetInf.SpousePIN").value
        
        If (Sheet14.Range("RetInf.RetStartDate").value <> "") Then
            If (TestDate(Sheet14.Range("RetInf.RetStartDate").value) = True) Then
                startDate = CDate(Format(Sheet14.Range("RetInf.RetStartDate").value, "dd/mm/yyyy"))
            End If
        End If
    
        If (Sheet14.Range("RetInf.RetEndDate").value <> "") Then
            If (TestDate(Sheet14.Range("RetInf.RetEndDate").value) = True) Then
                endDate = CDate(Format(Sheet14.Range("RetInf.RetEndDate").value, "dd/mm/yyyy"))
            End If
        End If
        
        Dim dateOfComm, dateOfMaturity As String
        Dim CommDate As Date
        Dim MaturityDate As Date
        Dim strIntYY As Integer
    
        'Code for Self Start
        rangeName = Worksheets("L_Computation_of_Insu_Relief").Range("InsReliefDtls.ListS").Address
        startRow = Worksheets("L_Computation_of_Insu_Relief").Range(rangeName).row
        endRow = Worksheets("L_Computation_of_Insu_Relief").Range(rangeName).Rows.Count + startRow - 1
        startColumn = Worksheets("L_Computation_of_Insu_Relief").Range(rangeName).column
        lastColumn = startColumn + Worksheets("L_Computation_of_Insu_Relief").Range(rangeName).Columns.Count - 1
    
                         
        For i = startRow To endRow
            Select Case Target.Address
                'PIN convert in uppar Case Start
                Case "$A$" & i
                    Target.value = StrConv(Target.Text, vbUpperCase)
                    If (Target.value <> "") Then
                        result = TestOtherPINNonId(Target.value)
                        If (result = False) Then
                            MsgBox "Please enter Non Individual Insurance Company's PIN in proper format."
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
                
                    'Insurance Policy Number convert in Uppar Case Start
                    Case "$D$" & i
                        If Target.value <> "" Then
                            If Target.value <> UCase(Target.value) Then
                                Target.value = UCase(Target.value)
                            End If
                        End If
                        'Insurance Policy Number convert in Uppar Case End
                
                        'Base on selected value from Purpose of policy Commencement Date will be come,for Education and Life Commencement Date should be on after 1st June 2003 and before current date,for Health Commencement Date should be on or after 1st Jan 2007 and before current date Start
                    Case "$G$" & i
                        Dim cmHDte As String
                        Dim fcmHDte As String
                        Dim syDate As String
                        Dim cmELDte As String
                        Dim fcmELDte As String
                        Dim tempF As Boolean
                        If Target.value <> "" Then
                            If TestDate(Target.value) = False Then
                                MsgBox "Please enter date value in dd/mm/yyyy format."
                                Target.value = ""
                                Target.Select
                            Else
                                If Worksheets("L_Computation_of_Insu_Relief").Range("C" & Target.row & ":C" & Target.row).value = "" Then
                                    MsgBox "Please first select purpose Of policy."
                                    Target.value = ""
                                    Worksheets("L_Computation_of_Insu_Relief").Range("C" & Target.row & ":C" & Target.row).Select
                                ElseIf Worksheets("L_Computation_of_Insu_Relief").Range("C" & Target.row & ":C" & Target.row).value = "Health" Then
                                    syDate = Date
                                    cmHDte = CDate(Format(Target.value, "dd/mm/yyyy"))
                                    fcmHDte = "1/1/2007"
                                    If (endDate <> "") Then
                                        If (DateValue(endDate) >= DateValue(cmHDte)) Then
                                            If DateValue(cmHDte) < DateValue(fcmHDte) Or DateValue(cmHDte) > DateValue(endDate) Then
                                                MsgBox "Commencement Date should be on or after 1st Jan 2007 and before or within Return Period if Purpose of Policy selected is Health"
                                                Target.value = ""
                                                Target.Select
                                            End If
                                        Else
                                        'MsgBox "Commencement Date should be before or within Return Period"
                                            MsgBox "Commencement Date in Row " & i & " should be before within Return period " & Sheet14.Range("RetInf.RetStartDate").value & " to " & Sheet14.Range("RetInf.RetEndDate").value
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
                            ElseIf Worksheets("L_Computation_of_Insu_Relief").Range("C" & Target.row & ":C" & Target.row).value = "Life" Or Worksheets("L_Computation_of_Insu_Relief").Range("C" & Target.row & ":C" & Target.row).value = "Education" Then
                                syDate = Date
                                cmELDte = CDate(Format(Target.value, "dd/mm/yyyy"))
                                fcmELDte = "6/1/2003"
                                Dim lngDateFormatType As Long
                                lngDateFormatType = Application.International(xlDateOrder)
                                If (lngDateFormatType = 0) Then
                                    fcmELDte = "6/1/2003"
                                ElseIf (lngDateFormatType = 1) Then
                                    fcmELDte = "1/6/2003"
                                ElseIf (lngDateFormatType = 2) Then
                                    fcmELDte = "2003/1/6"
                                End If
                                If (endDate <> "") Then
                                    If (DateValue(endDate) >= DateValue(cmELDte)) Then
                                        If DateValue(cmELDte) < DateValue(fcmELDte) Or DateValue(cmELDte) > DateValue(endDate) Then
                                            MsgBox "Commencement Date should be on or after 1st June 2003 and before or within Return Period if Purpose of Policy selected is Life or Education"
                                            Target.value = ""
                                            Target.Select
                                        End If
                                    Else
                                        'MsgBox "Commencement Date should be before or within Return Period"
                                        MsgBox "Commencement Date in Row " & i & " should be before within Return period " & Sheet14.Range("RetInf.RetStartDate").value & " to " & Sheet14.Range("RetInf.RetEndDate").value
                                        Target.value = ""
                                        Target.Select
                                    End If
                                Else
                                    MsgBox "Please first enter the Return for the Month and Year of Income."
                                    Target.value = ""
                                    Worksheets("A_Basic_Info").Activate
                                    Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                                    GoTo endSub
                                End If
                            End If
                            If Target.value <> "" And Target.Next.value <> "" Then
                            cellRow = Target.row
                                If (Worksheets("L_Computation_of_Insu_Relief").Range("C" & cellRow & ":C" & cellRow).value = "Education") Then
                                    tempF = compareDate(Worksheets("L_Computation_of_Insu_Relief").Range("G" & (Target.row) & ":G" & (Target.row)).value, Worksheets("L_Computation_of_Insu_Relief").Range("H" & (Target.row) & ":H" & (Target.row)).value)
                                    If (tempF = True) Then
                                    'do nthing
                                    Else
                                        MsgBox "The difference between Maturity Date and Commencement Date must be equal to or greater than 10 years."
                                        Target.value = ""
                                        Target.Select
                                    End If
                                End If
                            End If
                            If Target.value <> "" And Target.Next.value <> "" Then
                                    cmTNxtVal = CDate(Format(Target.Next.value, "dd/mm/yyyy"))
                                    cmTVal = CDate(Format(Target.value, "dd/mm/yyyy"))
                                    If (DateValue(cmTNxtVal) > DateValue(cmTVal)) Then
                                     'do nothing
                                     Else
                                         'MsgBox "Date of Commencement should be less than Maturity date."
                                         MsgBox "Commencement date shall be before Maturity date"
                                         Target.value = ""
                                         Target.Select
                                    End If
                            End If
                        End If
                    End If
                    
                'Base on selected value from Purpose of policy Commencement Date will be come,for Education and Life Commencement Date should be on after 1st June 2003 and before current date,for Health Commencement Date should be on or after 1st Jan 2007 and before current date End
                
                'Base on selected item from policy holder is child then Age of child field will be enable Start
                Case "$E$" & i
                    If Target.value <> "" And UCase(Trim(Target.value)) = "CHILD" Then
                        Call lockUnlock_cell_rng(Worksheets("L_Computation_of_Insu_Relief").name, "F" & i & ":F" & i, False)
                    Else
                        Call lockUnlock_cell_rng(Worksheets("L_Computation_of_Insu_Relief").name, "F" & i & ":F" & i, True)
                    End If
                    Target.Next.Select
                'Base on selected item from policy holder is child then Age of child field will be enable End
                
                'Validation for Maturity Date,Maturity Date should be 10year greater than commencement Date Start
                Case "$H$" & i
                    If (Target.value <> "") Then
                         If Target.value <> "" And TestDate(Target.value) = False Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            cellRow = Target.row
                            If (endDate <> "") Then
                                mDTrVal = CDate(Format(Target.value, "dd/mm/yyyy"))
                                If (DateValue(mDTrVal) >= DateValue(startDate)) Then
                                    If Target.value <> "" And Target.Previous.value <> "" Then
                                         If (Worksheets("L_Computation_of_Insu_Relief").Range("C" & cellRow & ":C" & cellRow).value = "Education") Then
                                             tempF = compareDate(Worksheets("L_Computation_of_Insu_Relief").Range("G" & (Target.row) & ":G" & (Target.row)).value, Worksheets("L_Computation_of_Insu_Relief").Range("H" & (Target.row) & ":H" & (Target.row)).value)
                                             If (tempF = True) Then
                                             'do nthing
                                             Else
                                                 MsgBox "The difference between Maturity Date and Commencement Date must be equal to or greater than 10 years."
                                                 Target.value = ""
                                                 Target.Select
                                             End If
                                         End If
                                         If (Worksheets("L_Computation_of_Insu_Relief").Range("C" & cellRow & ":C" & cellRow).value = "Health" Or Worksheets("L_Computation_of_Insu_Relief").Range("C" & cellRow & ":C" & cellRow).value = "Life") Then
                                             cmTprvVal = CDate(Format(Target.Previous.value, "dd/mm/yyyy"))
                                             cmTVal = CDate(Format(Target.value, "dd/mm/yyyy"))
                                         
                                             If (DateValue(cmTprvVal) < DateValue(cmTVal)) Then
                                             'do nothing
                                             Else
                                                 MsgBox "Maturity date shall not be before Commencement date"
                                                 Target.value = ""
                                                 Target.Select
                                             End If
                                         End If
                                     End If
                                Else
                                    MsgBox "Maturity Date in Row " & Target.row & " should be within or after Return Period"
                                    Target.value = ""
                                    Target.Select
                                End If
                            Else
                                MsgBox "Please first enter the Return for the Month and Year of Income."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Activate
                                Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                                GoTo endSub
                            End If
                        End If
                End If
    
    
                'Validation for Maturity Date,Maturity Date should be 10year greater than commencement Date End
                
                'Base on selected value from Purpose of policy Commencement Date will be come,for Education and Life Commencement Date should be on after 1st June 2003 and before current date,for Health Commencement Date should be on or after 1st Jan 2007 and before current date Start
                Case "$C$" & i
                    Dim cmDate As String
                    Dim fcmDate As String
                    'Dim syDate As String
                    Dim cmHDate As String
                    Dim fcmHDate As String
                    Dim tempFE As String
            
                    If Worksheets("L_Computation_of_Insu_Relief").Range("G" & Target.row & ":G" & Target.row).value <> "" And TestDate(Worksheets("L_Computation_of_Insu_Relief").Range("G" & Target.row & ":G" & Target.row).value) Then
                        If Target.value <> "" And Target.value = "Health" Then
                        syDate = Date
                        cmDate = CDate(Format((Worksheets("L_Computation_of_Insu_Relief").Range("G" & Target.row & ":G" & Target.row).value), "dd/mm/yyyy"))
                        fcmDate = "1/1/2007"
                            If DateValue(cmDate) < DateValue(fcmDate) Or DateValue(cmDate) > DateValue(endDate) Then
                                MsgBox "Commencement Date should be on or after 1st Jan 2007 and before or within Return Period if Purpose of Policy selected is Health"
                                Worksheets("L_Computation_of_Insu_Relief").Range("G" & Target.row & ":G" & Target.row).value = ""
                                Worksheets("L_Computation_of_Insu_Relief").Range("H" & Target.row & ":H" & Target.row).value = ""
                            End If
                            ElseIf Target.value <> "" And (Target.value = "Education" Or Target.value = "Life") Then
                                syDate = Date
                                cmHDate = CDate(Format((Worksheets("L_Computation_of_Insu_Relief").Range("G" & Target.row & ":G" & Target.row).value), "dd/mm/yyyy"))
                                fcmHDate = "6/1/2003"
                                lngDateFormatType = Application.International(xlDateOrder)
                                If (lngDateFormatType = 0) Then
                                    fcmHDate = "6/1/2003"
                                ElseIf (lngDateFormatType = 1) Then
                                    fcmHDate = "1/6/2003"
                                ElseIf (lngDateFormatType = 2) Then
                                    fcmHDate = "2003/1/6"
                                End If
                                
                                If DateValue(cmHDate) < DateValue(fcmHDate) Or DateValue(cmHDate) > DateValue(endDate) Then
                                    MsgBox "Commencement Date should be on or after 1st June 2003 and before or within Return Period if Purpose of Policy selected is Life or Education"
                                    Worksheets("L_Computation_of_Insu_Relief").Range("G" & Target.row & ":G" & Target.row).value = ""
                                    Sheet13.Range("H" & Target.row & ":H" & Target.row).value = ""
                                End If
                        End If
                    End If
                    'Added on 03/02/2023 64901
                    Worksheets("L_Computation_of_Insu_Relief").Unprotect (Pwd)
                    If (Target.value <> "" And Target.value = "Nhif") Then
                        Worksheets("L_Computation_of_Insu_Relief").Range("E" & Target.row & ":E" & Target.row).value = "Self"
                        Worksheets("L_Computation_of_Insu_Relief").Range("E" & Target.row & ":E" & Target.row).Locked = True
                        Worksheets("L_Computation_of_Insu_Relief").Range("I" & Target.row & ":I" & Target.row).value = 0#
                        Call lockUnlock_cell_rng(Worksheets("L_Computation_of_Insu_Relief").name, "F" & Target.row & ":I" & Target.row, True)
                    Else
                        Worksheets("L_Computation_of_Insu_Relief").Range("E" & Target.row & ":E" & Target.row).value = ""
                        Worksheets("L_Computation_of_Insu_Relief").Range("E" & Target.row & ":E" & Target.row).Locked = False
                        Call lockUnlock_cell_rng(Worksheets("L_Computation_of_Insu_Relief").name, "G" & Target.row & ":I" & Target.row, False)
                    End If
                     Worksheets("L_Computation_of_Insu_Relief").Protect (Pwd)
                     'End 03/02/2023
                    If (Target.value <> "" And Target.value = "Education") Then
                    End If
    
    
                'Base on selected value from Purpose of policy Commencement Date will be come,for Education and Life Commencement Date should be on after 1st June 2003 and before current date,for Health Commencement Date should be on or after 1st Jan 2007 and before current date End
            End Select
        Next
    
        'Validation for PIN of Insurance company,Policy Holder and Insurance Policy Number,in paire of PIN of Insurance Company,Policy Holder and Insurance Policy Numbe can not be repeable Start
        If (Application.Intersect(Target.Cells, Range("A" & startRow & ":A" & endRow)) Is Nothing) Then
            'Target Range is not in Given Range
        ElseIf Target.Cells.Count = 1 Then
            Dim PINSchPatrn1 As String
            Dim curPINNo1 As String
            Dim refSchedule1 As String
            Dim chkPINNo1 As String
            Dim chkRefSch1 As String
            Dim chkpolicynum1 As String
            Dim policynum1 As String
            PINSchPatrn1 = Worksheets("L_Computation_of_Insu_Relief").Range("O" & Target.row & ":O" & Target.row).value
            curPINNo1 = Worksheets("L_Computation_of_Insu_Relief").Range("A" & Target.row & ":A" & Target.row).value
            refSchedule1 = Worksheets("L_Computation_of_Insu_Relief").Range("C" & Target.row & ":C" & Target.row).value
            policynum1 = Worksheets("L_Computation_of_Insu_Relief").Range("D" & Target.row & ":D" & Target.row).value
            If curPINNo1 <> "" And refSchedule1 <> "" And policynum1 <> "" Then
                For Each r In Range("InsReliefDtls.ListS").Rows
                    For Each c In r.Cells
                            If c.row > endRow Then
                                Exit For
                            End If
                            If c.column = Target.column Then
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                Else
                                    chkPINNo1 = Worksheets("L_Computation_of_Insu_Relief").Range("A" & c.row & ":A" & c.row).value
                                    chkRefSch1 = Worksheets("L_Computation_of_Insu_Relief").Range("C" & c.row & ":C" & c.row).value
                                    chkpolicynum1 = Worksheets("L_Computation_of_Insu_Relief").Range("D" & c.row & ":D" & c.row).value
                                    If UCase(PINSchPatrn1) = UCase(chkPINNo1 & chkRefSch1 & chkpolicynum1) Then
                                        MsgBox "The Insurance Policy Number against the same PIN of Insurance Company and Purpose of Policy cannot be same."
                                        Target.value = ""
                                        Target.Select
                                        Exit For
                                    Exit For
                                    End If
                                End If
                            End If
                    Next
                Next
            End If
        End If
        'Validation for PIN of Insurance company,Policy Holder and Insurance Policy Number,in paire of PIN of Insurance Company,Policy Holder and Insurance Policy Numbe can not be repeable End
        
        
        'Validation for PIN of Insurance company,Policy Holder and Insurance Policy Number,in paire of PIN of Insurance Company,Policy Holder and Insurance Policy Numbe can not be repeable Start
        If (Application.Intersect(Target.Cells, Range("C" & startRow & ":C" & endRow)) Is Nothing) Then
            'Target Range is not in Given Range
        ElseIf Target.Cells.Count = 1 Then
            PINSchPatrn1 = Worksheets("L_Computation_of_Insu_Relief").Range("O" & Target.row & ":O" & Target.row).value
            curPINNo1 = Worksheets("L_Computation_of_Insu_Relief").Range("A" & Target.row & ":A" & Target.row).value
            refSchedule1 = Worksheets("L_Computation_of_Insu_Relief").Range("C" & Target.row & ":C" & Target.row).value
            policynum1 = Worksheets("L_Computation_of_Insu_Relief").Range("D" & Target.row & ":D" & Target.row).value
            If curPINNo1 <> "" And refSchedule1 <> "" And policynum1 <> "" Then
                For Each r In Range("InsReliefDtls.ListS").Rows
                    For Each c In r.Cells
                        If c.row > endRow Then
                            Exit For
                        End If
                        If c.column = Target.column Then
                            If c.Address = Target.Address Then 'Skip comparison with self value
                            Else
                                chkPINNo1 = Worksheets("L_Computation_of_Insu_Relief").Range("A" & c.row & ":A" & c.row).value
                                chkRefSch1 = Worksheets("L_Computation_of_Insu_Relief").Range("C" & c.row & ":C" & c.row).value
                                chkpolicynum1 = Worksheets("L_Computation_of_Insu_Relief").Range("D" & c.row & ":D" & c.row).value
                                If UCase(PINSchPatrn1) = UCase(chkPINNo1 & chkRefSch1 & chkpolicynum1) Then
                                    MsgBox "The Insurance Policy Number against the same PIN of Insurance Company and Purpose of Policy cannot be same."
                                    Target.value = ""
                                    Target.Select
                                    Exit For
                                Exit For
                                End If
                            End If
                        End If
                    Next
                Next
            End If
        End If
        'Validation for PIN of Insurance company,Policy Holder and Insurance Policy Number,in paire of PIN of Insurance Company,Policy Holder and Insurance Policy Numbe can not be repeable End
        
        'Validation for PIN of Insurance company,Policy Holder and Insurance Policy Number,in paire of PIN of Insurance Company,Policy Holder and Insurance Policy Numbe can not be repeable Start
        If (Application.Intersect(Target.Cells, Range("D" & startRow & ":D" & endRow)) Is Nothing) Then
            'Target Range is not in Given Range
        ElseIf Target.Cells.Count = 1 Then
            PINSchPatrn1 = Worksheets("L_Computation_of_Insu_Relief").Range("O" & Target.row & ":O" & Target.row).value
            curPINNo1 = Worksheets("L_Computation_of_Insu_Relief").Range("A" & Target.row & ":A" & Target.row).value
            refSchedule1 = Worksheets("L_Computation_of_Insu_Relief").Range("C" & Target.row & ":C" & Target.row).value
            policynum1 = Worksheets("L_Computation_of_Insu_Relief").Range("D" & Target.row & ":D" & Target.row).value
            If curPINNo1 <> "" And refSchedule1 <> "" And policynum1 <> "" Then
                For Each r In Range("InsReliefDtls.ListS").Rows
                    For Each c In r.Cells
                        If c.row > endRow Then
                            Exit For
                        End If
                        If c.column = Target.column Then
                            If c.Address = Target.Address Then 'Skip comparison with self value
                            Else
                                chkPINNo1 = Worksheets("L_Computation_of_Insu_Relief").Range("A" & c.row & ":A" & c.row).value
                                chkRefSch1 = Worksheets("L_Computation_of_Insu_Relief").Range("C" & c.row & ":C" & c.row).value
                                chkpolicynum1 = Worksheets("L_Computation_of_Insu_Relief").Range("D" & c.row & ":D" & c.row).value
                                If UCase(PINSchPatrn1) = UCase(chkPINNo1 & chkRefSch1 & chkpolicynum1) Then
                                    MsgBox "The Insurance Policy Number against the same PIN of Insurance Company and Purpose of Policy cannot be same."
                                    Target.value = ""
                                    Target.Select
                                    Exit For
                                Exit For
                                End If
                            End If
                        End If
                    Next
                Next
            End If
        End If
        'Validation for PIN of Insurance company,Policy Holder and Insurance Policy Number,in paire of PIN of Insurance Company,Policy Holder and Insurance Policy Numbe can not be repeable End
        'Code for Self End
    
        'Code for Wife Start
        'Worksheets("L_Computation_of_Insu_Relief").Activate
        rangeName = Worksheets("L_Computation_of_Insu_Relief").Range("InsReliefDtls.ListW").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1
    
        For i = startRow To endRow
            Select Case Target.Address
                'PIN convert in uppar Case Start
                Case "$A$" & i
                    Target.value = StrConv(Target.Text, vbUpperCase)
                    If (Target.value <> "") Then
                    result = TestOtherPINNonId(Target.value)
                        If (result = False) Then
                            MsgBox "Please enter Non Individual Insurance Company's PIN in proper format."
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
                
                'Insurance Policy Number convert in Uppar Case Start
                Case "$D$" & i
                    If Target.value <> "" Then
                        If Target.value <> UCase(Target.value) Then
                            Target.value = UCase(Target.value)
                        End If
                    End If
                'Insurance Policy Number convert in Uppar Case End
                
                'Base on selected value from Purpose of policy Commencement Date will be come,for Education and Life Commencement Date should be on after 1st June 2003 and before current date,for Health Commencement Date should be on or after 1st Jan 2007 and before current date Start
                
                 Case "$G$" & i
                    If Target.value <> "" Then
                        If TestDate(Target.value) = False Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            If Worksheets("L_Computation_of_Insu_Relief").Range("C" & Target.row & ":C" & Target.row).value = "" Then
                                MsgBox "Please first select purpose Of policy."
                                Target.value = ""
                                Worksheets("L_Computation_of_Insu_Relief").Range("C" & Target.row & ":C" & Target.row).Select
                            ElseIf Worksheets("L_Computation_of_Insu_Relief").Range("C" & Target.row & ":C" & Target.row).value = "Health" Then
                                syDate = Date
                                cmHDte = CDate(Format(Target.value, "dd/mm/yyyy"))
                                fcmHDte = "1/1/2007"
                                If (endDate <> "") Then
                                
                                    If (DateValue(endDate) >= DateValue(cmHDte)) Then
                                        If DateValue(cmHDte) < DateValue(fcmHDte) Or DateValue(cmHDte) > DateValue(endDate) Then
                                            MsgBox "Commencement Date should be on or after 1st Jan 2007 and before or within Return Period if Purpose of Policy selected is Health"
                                            Target.value = ""
                                            Target.Select
                                        End If
                                    Else
                                        MsgBox "Commencement Date in Row " & i & " should be before within Return period " & Sheet14.Range("RetInf.RetStartDate").value & " to " & Sheet14.Range("RetInf.RetEndDate").value
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
                            ElseIf Worksheets("L_Computation_of_Insu_Relief").Range("C" & Target.row & ":C" & Target.row).value = "Life" Or Worksheets("L_Computation_of_Insu_Relief").Range("C" & Target.row & ":C" & Target.row).value = "Education" Then
                                syDate = Date
                                cmELDte = CDate(Format(Target.value, "dd/mm/yyyy"))
                                fcmELDte = "6/1/2003"
                                lngDateFormatType = Application.International(xlDateOrder)
                                If (lngDateFormatType = 0) Then
                                    fcmELDte = "6/1/2003"
                                ElseIf (lngDateFormatType = 1) Then
                                    fcmELDte = "1/6/2003"
                                ElseIf (lngDateFormatType = 2) Then
                                    fcmELDte = "2003/1/6"
                                End If
                                
                                If (endDate <> "") Then
                                    If (DateValue(endDate) >= DateValue(cmELDte)) Then
                                        If DateValue(cmELDte) < DateValue(fcmELDte) Or DateValue(cmELDte) > DateValue(endDate) Then
                                            MsgBox "Commencement Date should be on or after 1st June 2003 and before or within Return Period if Purpose of Policy selected is Life or Education"
                                            Target.value = ""
                                            Target.Select
                                        End If
                                    Else
                                        MsgBox "Commencement Date in Row " & i & " should be before within Return period " & Sheet14.Range("RetInf.RetStartDate").value & " to " & Sheet14.Range("RetInf.RetEndDate").value
                                        Target.value = ""
                                        Target.Select
                                    End If
                                Else
                                    MsgBox "Please first enter the Return for the Month and Year of Income."
                                    Target.value = ""
                                    Worksheets("A_Basic_Info").Activate
                                    Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
                                    GoTo endSub
                                End If
                            End If
                            If Target.value <> "" And Target.Next.value <> "" Then
                                cellRow = Target.row
                                If (Worksheets("L_Computation_of_Insu_Relief").Range("C" & cellRow & ":C" & cellRow).value = "Education") Then
                                    tempF = compareDate(Worksheets("L_Computation_of_Insu_Relief").Range("G" & (Target.row) & ":G" & (Target.row)).value, Worksheets("L_Computation_of_Insu_Relief").Range("H" & (Target.row) & ":H" & (Target.row)).value)
                                    If (tempF = True) Then
                                    'do nthing
                                    Else
                                        MsgBox "The difference between Maturity Date and Commencement Date must be equal to or greater than 10 years."
                                        Target.value = ""
                                        Target.Select
                                    End If
                                End If
                            End If
                            If Target.value <> "" And Target.Next.value <> "" Then
                                cmTNxtVal = CDate(Format(Target.Next.value, "dd/mm/yyyy"))
                                cmTVal = CDate(Format(Target.value, "dd/mm/yyyy"))
                                If (DateValue(cmTNxtVal) > DateValue(cmTVal)) Then
                                 'do nothing
                                 Else
                                     'MsgBox "Date of Commencement should be less than Maturity date."
                                     MsgBox "Commencement date shall be before Maturity date"
                                     Target.value = ""
                                     Target.Select
                                End If
                            End If
                        End If
                    End If
                
                'Base on selected value from Purpose of policy Commencement Date will be come,for Education and Life Commencement Date should be on after 1st June 2003 and before current date,for Health Commencement Date should be on or after 1st Jan 2007 and before current date End
            
                'Base on selected item from policy holder is child then Age of child field will be enable Start
                Case "$E$" & i
                    If Target.value <> "" And UCase(Trim(Target.value)) = "CHILD" Then
                        Call lockUnlock_cell_rng(Worksheets("L_Computation_of_Insu_Relief").name, "F" & i & ":F" & i, False)
                    Else
                        Call lockUnlock_cell_rng(Worksheets("L_Computation_of_Insu_Relief").name, "F" & i & ":F" & i, True)
                    End If
                    Target.Next.Select
                'Base on selected item from policy holder is child then Age of child field will be enable End
                
                'Validation for Maturity Date,Maturity Date should be 10year greater than commencement Date Start
                Case "$H$" & i
                    If (Target.value <> "") Then
                         If Target.value <> "" And TestDate(Target.value) = False Then
                            MsgBox "Please enter date value in dd/mm/yyyy format."
                            Target.value = ""
                            Target.Select
                        Else
                            cellRow = Target.row
                            
                            If (endDate <> "") Then
                                mDTrVal = CDate(Format(Target.value, "dd/mm/yyyy"))
                                If (DateValue(mDTrVal) >= DateValue(startDate)) Then
                                    If Target.value <> "" And Target.Previous.value <> "" Then
                                         If (Worksheets("L_Computation_of_Insu_Relief").Range("C" & cellRow & ":C" & cellRow).value = "Education") Then
                                             tempF = compareDate(Worksheets("L_Computation_of_Insu_Relief").Range("G" & (Target.row) & ":G" & (Target.row)).value, Worksheets("L_Computation_of_Insu_Relief").Range("H" & (Target.row) & ":H" & (Target.row)).value)
                                             If (tempF = True) Then
                                             'do nthing
                                             Else
                                                 MsgBox "The difference between Maturity Date and Commencement Date must be equal to or greater than 10 years."
                                                 Target.value = ""
                                                 Target.Select
                                             End If
                                         End If
                                         If (Worksheets("L_Computation_of_Insu_Relief").Range("C" & cellRow & ":C" & cellRow).value = "Health" Or Worksheets("L_Computation_of_Insu_Relief").Range("C" & cellRow & ":C" & cellRow).value = "Life") Then
                                             cmTprvVal = CDate(Format(Target.Previous.value, "dd/mm/yyyy"))
                                             cmTVal = CDate(Format(Target.value, "dd/mm/yyyy"))
                                         
                                             If (DateValue(cmTprvVal) < DateValue(cmTVal)) Then
                                             'do nothing
                                             Else
                                                 MsgBox "Maturity date shall not be before Commencement date"
                                                 Target.value = ""
                                                 Target.Select
                    
                                             End If
                                         End If
                                     End If
                                Else
                                    MsgBox "Maturity Date in Row " & Target.row & " should be within or after Return Period"
                                    Target.value = ""
                                    Target.Select
                                    
                                End If
                            Else
                                MsgBox "Please first enter the Return for the Month and Year of Income."
                                Target.value = ""
                                Worksheets("A_Basic_Info").Activate
                                Worksheets("A_Basic_Info").Range("B5:B5").Select
                                GoTo endSub
                            End If
                        End If
                End If
                'Validation for Maturity Date,Maturity Date should be 10year greater than commencement Date End
        
                'Base on selected value from Purpose of policy Commencement Date will be come,for Education and Life Commencement Date should be on after 1st June 2003 and before current date,for Health Commencement Date should be on or after 1st Jan 2007 and before current date Start
                Case "$C$" & i
                    If Worksheets("L_Computation_of_Insu_Relief").Range("G" & Target.row & ":G" & Target.row).value <> "" And TestDate(Worksheets("L_Computation_of_Insu_Relief").Range("G" & Target.row & ":G" & Target.row).value) Then
                        If Target.value <> "" And Target.value = "Health" Then
                        syDate = Date
                        cmDate = CDate(Format((Worksheets("L_Computation_of_Insu_Relief").Range("G" & Target.row & ":G" & Target.row).value), "dd/mm/yyyy"))
                        fcmDate = "1/1/2007"
                            If DateValue(cmDate) < DateValue(fcmDate) Or DateValue(cmDate) > DateValue(endDate) Then
                                MsgBox "Commencement Date should be on or after 1st Jan 2007 and before or within Return Period if Purpose of Policy selected is Health"
                                Worksheets("L_Computation_of_Insu_Relief").Range("G" & Target.row & ":G" & Target.row).value = ""
                                Worksheets("L_Computation_of_Insu_Relief").Range("H" & Target.row & ":H" & Target.row).value = ""
                            End If
                            ElseIf Target.value <> "" And (Target.value = "Education" Or Target.value = "Life") Then
                                syDate = Date
                                cmHDate = CDate(Format((Worksheets("L_Computation_of_Insu_Relief").Range("G" & Target.row & ":G" & Target.row).value), "dd/mm/yyyy"))
                                fcmHDate = "6/1/2003"
                                
                                lngDateFormatType = Application.International(xlDateOrder)
                                If (lngDateFormatType = 0) Then
                                    fcmHDate = "6/1/2003"
                                ElseIf (lngDateFormatType = 1) Then
                                    fcmHDate = "1/6/2003"
                                ElseIf (lngDateFormatType = 2) Then
                                    fcmHDate = "2003/1/6"
                                End If
                                
                                If DateValue(cmHDate) < DateValue(fcmHDate) Or DateValue(cmHDate) > DateValue(endDate) Then
                                    MsgBox "Commencement Date should be on or after 1st June 2003 and before or within Return Period if Purpose of Policy selected is Life or Education"
                                    Worksheets("L_Computation_of_Insu_Relief").Range("G" & Target.row & ":G" & Target.row).value = ""
                                    Sheet13.Range("H" & Target.row & ":H" & Target.row).value = ""
                                End If
                        End If
                    End If
                    If (Target.value <> "" And Target.value = "Education") Then
                        If (Worksheets("L_Computation_of_Insu_Relief").Range("G" & Target.row & ":G" & Target.row).value <> "" And Worksheets("L_Computation_of_Insu_Relief").Range("H" & Target.row & ":H" & Target.row).value <> "") Then
                            tempFE = compareDate(Worksheets("L_Computation_of_Insu_Relief").Range("G" & (Target.row) & ":G" & (Target.row)).value, Worksheets("L_Computation_of_Insu_Relief").Range("H" & (Target.row) & ":H" & (Target.row)).value)
                                If (tempFE = True) Then
                                'do nthing
                                Else
                                    MsgBox "The difference between Maturity Date and Commencement Date must be equal to or greater than 10 years."
                                    Target.value = ""
                                    Target.Select
                                End If
                        End If
                    End If
                'Base on selected value from Purpose of policy Commencement Date will be come,for Education and Life Commencement Date should be on after 1st June 2003 and before current date,for Health Commencement Date should be on or after 1st Jan 2007 and before current date End
            End Select
        Next
    
        'Validation for PIN of Insurance company,Policy Holder and Insurance Policy Number,in paire of PIN of Insurance Company,Policy Holder and Insurance Policy Numbe can not be repeable Start
        If (Application.Intersect(Target.Cells, Range("A" & startRow & ":A" & endRow)) Is Nothing) Then
            'Target Range is not in Given Range
        ElseIf Target.Cells.Count = 1 Then
            
            Dim PINSchPatrn1W As String
            Dim curPINNo1W As String
            Dim refSchedule1W As String
            Dim chkPINNo1W As String
            Dim chkRefSch1W As String
            Dim chkpolicynum1w As String
            Dim policynum1w As String
            PINSchPatrn1W = Worksheets("L_Computation_of_Insu_Relief").Range("O" & Target.row & ":O" & Target.row).value
            curPINNo1W = Worksheets("L_Computation_of_Insu_Relief").Range("A" & Target.row & ":A" & Target.row).value
            refSchedule1W = Worksheets("L_Computation_of_Insu_Relief").Range("C" & Target.row & ":C" & Target.row).value
            policynum1w = Worksheets("L_Computation_of_Insu_Relief").Range("D" & Target.row & ":D" & Target.row).value
            If curPINNo1W <> "" And refSchedule1W <> "" And policynum1w <> "" Then
                For Each r In Range("InsReliefDtls.ListW").Rows
                    For Each c In r.Cells
                            If c.row > endRow Then
                                Exit For
                            End If
                            If c.column = Target.column Then
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                Else
                                    chkPINNo1W = Worksheets("L_Computation_of_Insu_Relief").Range("A" & c.row & ":A" & c.row).value
                                    chkRefSch1W = Worksheets("L_Computation_of_Insu_Relief").Range("C" & c.row & ":C" & c.row).value
                                    chkpolicynum1w = Worksheets("L_Computation_of_Insu_Relief").Range("D" & c.row & ":D" & c.row).value
                                    If UCase(PINSchPatrn1W) = UCase(chkPINNo1W & chkRefSch1W & chkpolicynum1w) Then
                                        MsgBox "The Insurance Policy Number against the same PIN of Insurance Company and Purpose of Policy cannot be same."
                                        Target.value = ""
                                        Target.Select
                                        Exit For
                                    Exit For
                                    End If
                                End If
                            End If
                    Next
                Next
            End If
            
        End If
        'Validation for PIN of Insurance company,Policy Holder and Insurance Policy Number,in paire of PIN of Insurance Company,Policy Holder and Insurance Policy Numbe can not be repeable End
        
        'Validation for PIN of Insurance company,Policy Holder and Insurance Policy Number,in paire of PIN of Insurance Company,Policy Holder and Insurance Policy Numbe can not be repeable Start
        If (Application.Intersect(Target.Cells, Range("C" & startRow & ":C" & endRow)) Is Nothing) Then
            'Target Range is not in Given Range
        ElseIf Target.Cells.Count = 1 Then
            
            PINSchPatrn1W = Worksheets("L_Computation_of_Insu_Relief").Range("O" & Target.row & ":O" & Target.row).value
            curPINNo1W = Worksheets("L_Computation_of_Insu_Relief").Range("A" & Target.row & ":A" & Target.row).value
            refSchedule1W = Worksheets("L_Computation_of_Insu_Relief").Range("C" & Target.row & ":C" & Target.row).value
            policynum1w = Worksheets("L_Computation_of_Insu_Relief").Range("D" & Target.row & ":D" & Target.row).value
            If curPINNo1W <> "" And refSchedule1W <> "" And policynum1w <> "" Then
                For Each r In Range("InsReliefDtls.ListW").Rows
                    For Each c In r.Cells
                            If c.row > endRow Then
                                Exit For
                            End If
                            If c.column = Target.column Then
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                Else
                                    chkPINNo1W = Worksheets("L_Computation_of_Insu_Relief").Range("A" & c.row & ":A" & c.row).value
                                    chkRefSch1W = Worksheets("L_Computation_of_Insu_Relief").Range("C" & c.row & ":C" & c.row).value
                                    chkpolicynum1w = Worksheets("L_Computation_of_Insu_Relief").Range("D" & c.row & ":D" & c.row).value
                                    If UCase(PINSchPatrn1W) = UCase(chkPINNo1W & chkRefSch1W & chkpolicynum1w) Then
                                        MsgBox "The Insurance Policy Number against the same PIN of Insurance Company and Purpose of Policy cannot be same."
                                        Target.value = ""
                                        Target.Select
                                        Exit For
                                    Exit For
                                    End If
                                End If
                            End If
                    Next
                Next
            End If
            
        End If
        'Validation for PIN of Insurance company,Policy Holder and Insurance Policy Number,in paire of PIN of Insurance Company,Policy Holder and Insurance Policy Numbe can not be repeable End
        
        'Validation for PIN of Insurance company,Policy Holder and Insurance Policy Number,in paire of PIN of Insurance Company,Policy Holder and Insurance Policy Numbe can not be repeable Start
        If (Application.Intersect(Target.Cells, Range("D" & startRow & ":D" & endRow)) Is Nothing) Then
            'Target Range is not in Given Range
        ElseIf Target.Cells.Count = 1 Then
            
            PINSchPatrn1W = Worksheets("L_Computation_of_Insu_Relief").Range("O" & Target.row & ":O" & Target.row).value
            curPINNo1W = Worksheets("L_Computation_of_Insu_Relief").Range("A" & Target.row & ":A" & Target.row).value
            refSchedule1W = Worksheets("L_Computation_of_Insu_Relief").Range("C" & Target.row & ":C" & Target.row).value
            policynum1w = Worksheets("L_Computation_of_Insu_Relief").Range("D" & Target.row & ":D" & Target.row).value
            If curPINNo1W <> "" And refSchedule1W <> "" And policynum1w <> "" Then
                For Each r In Range("InsReliefDtls.ListW").Rows
                    For Each c In r.Cells
                            If c.row > endRow Then
                                Exit For
                            End If
                            If c.column = Target.column Then
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                Else
                                    chkPINNo1W = Worksheets("L_Computation_of_Insu_Relief").Range("A" & c.row & ":A" & c.row).value
                                    chkRefSch1W = Worksheets("L_Computation_of_Insu_Relief").Range("C" & c.row & ":C" & c.row).value
                                    chkpolicynum1w = Worksheets("L_Computation_of_Insu_Relief").Range("D" & c.row & ":D" & c.row).value
                                    If UCase(PINSchPatrn1W) = UCase(chkPINNo1W & chkRefSch1W & chkpolicynum1w) Then
                                        MsgBox "The Insurance Policy Number against the same PIN of Insurance Company and Purpose of Policy cannot be same."
                                        Target.value = ""
                                        Target.Select
                                        Exit For
                                    Exit For
                                    End If
                                End If
                            End If
                    Next
                Next
            End If
            
        End If
        'Validation for PIN of Insurance company,Policy Holder and Insurance Policy Number,in paire of PIN of Insurance Company,Policy Holder and Insurance Policy Numbe can not be repeable End
        'Code for Wife End
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
    End If
endSub:
    
    Worksheets("L_Computation_of_Insu_Relief").Protect (Pwd)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
