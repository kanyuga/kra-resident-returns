VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet33"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'Add Row Functionality For Self Start
Sub addRow_ProfitShare_ListS()
    Worksheets("G_Partnership_Income").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("ProfitShare.ListS")
    rangeName = Worksheets("G_Partnership_Income").Range("ProfitShare.ListS").Address
    startRowIndex = Range(rangeName).row
    endRowIndex = Range(rangeName).Rows.Count + startRowIndex - 1
    Dim Str As String
    Str = "C" & startRowIndex & ":C" & endRowIndex
    ActiveWorkbook.Names("TypeOfIncomeS").Delete
    Range(Str).name = "TypeOfIncomeS"
    
    Str = "I" & startRowIndex & ":I" & endRowIndex
    ActiveWorkbook.Names("TotalIncomePartershipS").Delete
    Range(Str).name = "TotalIncomePartershipS"
    Worksheets("G_Partnership_Income").Protect (Pwd)
End Sub
'Add Row Functionality For Self End

'Add Row Functionality For Wife Start
Sub addRow_ProfitShare_ListW()
    Worksheets("G_Partnership_Income").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("ProfitShare.ListW")
    rangeName = Worksheets("G_Partnership_Income").Range("ProfitShare.ListW").Address
    startRowIndex = Range(rangeName).row
    endRowIndex = Range(rangeName).Rows.Count + startRowIndex - 1
    Dim Str As String
    Str = "C" & startRowIndex & ":C" & endRowIndex
    ActiveWorkbook.Names("TypeOfIncomeW").Delete
    Range(Str).name = "TypeOfIncomeW"
    
    Str = "I" & startRowIndex & ":I" & endRowIndex
    ActiveWorkbook.Names("TotalIncomePartershipW").Delete
    Range(Str).name = "TotalIncomePartershipW"
    
    Worksheets("G_Partnership_Income").Protect (Pwd)
End Sub
'Add Row Functionality For Wife End


Private Sub Worksheet_Change(ByVal Target As Range)
    
    
    Worksheets("G_Partnership_Income").Unprotect (Pwd)
    Worksheets("G_Partnership_Income").Activate
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    If (Target.Columns.Count = 1) Then
        PINNoS = Worksheets("A_Basic_Info").Range("RetInf.PIN").value
        PINNoW = Worksheets("A_Basic_Info").Range("RetInf.SpousePIN").value
        'Code for Self Start
        rangeName = Worksheets("G_Partnership_Income").Range("ProfitShare.ListS").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1
        startColumn = Range(rangeName).column
        lastColumn = startColumn + Range(rangeName).Columns.Count - 1
                         
        For i = startRow To endRow
            Select Case Target.Address
                'PIN convert in uppar Case Start
                Case "$A$" & i
                    Target.value = StrConv(Target.Text, vbUpperCase)
                    If (Target.value <> "") Then
                    result = TestOtherPINNonId(Target.value)
                        If (result = False) Then
                            MsgBox "Please enter Non Individual Partnership Firm's PIN in proper format."
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
                'PIN convert in uppar Case End
                
                'Base on selected value from Type of Income Respective code will be set in hidden field Start
                Case "$C$" & i
                    If (Target.value <> "") Then
                        rowvalue = find_NatureOfIncomeCode(Target.value)
                        natureOfIncomeCode = Sheet18.Cells(rowvalue, natureOfIncomeCodeCol).value
                         If (rowvalue <> 0) Then
                                 Range("K" & i & ":K" & i).value = natureOfIncomeCode
                                 Range("K" & i & ":K" & i).Locked = True
                         End If
                         Target.Next.Select
                    Else
                        Range("K" & i & ":K" & i).Locked = True
                        Range("K" & i & ":K" & i).value = ""
                        Target.Select
                    End If
                'Base on selected value from Type of Income Respective code will be set in hidden field End
            End Select
        Next
        
        'new code added for same PIN with Same Type of Income not allowed Start
        If (Application.Intersect(Target.Cells, Range("C" & startRow & ":C" & endRow)) Is Nothing) Then
        'Target Range is not in Given Range
        ElseIf Target.Cells.Count = 1 Then
        
            Dim PINSchPatrn1 As String
            Dim curPINNo1 As String
            Dim refSchedule1 As String
            Dim chkPINNo1 As String
            Dim chkRefSch1 As String
            
            PINSchPatrn1 = Worksheets("G_Partnership_Income").Range("L" & Target.row & ":L" & Target.row).value
            curPINNo1 = Worksheets("G_Partnership_Income").Range("A" & Target.row & ":A" & Target.row).value
            refSchedule1 = Worksheets("G_Partnership_Income").Range("C" & Target.row & ":C" & Target.row).value
            
            If curPINNo1 <> "" And refSchedule1 <> "" Then
                For Each r In Range("ProfitShare.ListS").Rows
                    For Each c In r.Cells
                        If c.row > endRow Then
                            Exit For
                        End If
                        If c.column = Target.column Then
                            If c.Address = Target.Address Then
                                'Skip comparison with self value
                            Else
                                chkPINNo1 = Worksheets("G_Partnership_Income").Range("A" & c.row & ":A" & c.row).value
                                chkRefSch1 = Worksheets("G_Partnership_Income").Range("C" & c.row & ":C" & c.row).value
                                
                                If UCase(PINSchPatrn1) = UCase(chkPINNo1 & chkRefSch1) Then
                                    MsgBox "You have already entered details for " & chkRefSch1 & " for PIN of Partnership Firm " & chkPINNo1 & " in Row " & c.row
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
    If (Application.Intersect(Target.Cells, Range("A" & startRow & ":A" & endRow)) Is Nothing) Then
     'Target Range is not in Given Range
    ElseIf Target.Cells.Count = 1 Then
        Dim PINSchPatrn2 As String
        Dim curPINNo2 As String
        Dim refSchedule2 As String
        Dim chkPINNo2 As String
        Dim chkRefSch2 As String
        PINSchPatrn2 = Worksheets("G_Partnership_Income").Range("L" & Target.row & ":L" & Target.row).value
        curPINNo2 = Worksheets("G_Partnership_Income").Range("A" & Target.row & ":A" & Target.row).value
        refSchedule2 = Worksheets("G_Partnership_Income").Range("C" & Target.row & ":C" & Target.row).value
        If curPINNo2 <> "" And refSchedule2 <> "" Then
            For Each r In Range("ProfitShare.ListS").Rows
            For Each c In r.Cells
                    If c.row > endRow Then
                        Exit For
                    End If
                    If c.column = Target.column Then
                        If c.Address = Target.Address Then
                            'Skip comparison with self value
                        Else
                            chkPINNo2 = Worksheets("G_Partnership_Income").Range("A" & c.row & ":A" & c.row).value
                            chkRefSch2 = Worksheets("G_Partnership_Income").Range("C" & c.row & ":C" & c.row).value
                            If UCase(PINSchPatrn2) = UCase(chkPINNo2 & chkRefSch2) Then
                                MsgBox "You have already entered details for " & chkRefSch2 & " of PIN of Partnership Firm " & chkPINNo2 & " in Row " & c.row
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
    'new code added for same PIN with Same Type of Income not allowed End
    'Code for Self End
  
    
    'Code for Wife Start
    rangeName = Worksheets("G_Partnership_Income").Range("ProfitShare.ListW").Address
    startRow = Range(rangeName).row
    endRow = Range(rangeName).Rows.Count + startRow - 1
    startColumn = Range(rangeName).column
    lastColumn = startColumn + Range(rangeName).Columns.Count - 1
                         
    For i = startRow To endRow
        Select Case Target.Address
            'PIN convert in uppar Case Start
            Case "$A$" & i
                    Target.value = StrConv(Target.Text, vbUpperCase)
                    If (Target.value <> "") Then
                    result = TestOtherPINNonId(Target.value)
                        If (result = False) Then
                            MsgBox "Please enter Non Individual Partnership Firm's PIN in proper format."
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
                'PIN convert in uppar Case End
                
                'Base on selected value from Type of Income Respective code will be set in hidden field Start
                Case "$C$" & i
                    If (Target.value <> "") Then
                        rowvalue = find_NatureOfIncomeCode(Target.value)
                        natureOfIncomeCode = Sheet18.Cells(rowvalue, natureOfIncomeCodeCol).value
                         If (rowvalue <> 0) Then
                                 Range("K" & i & ":K" & i).value = natureOfIncomeCode
                                 Range("K" & i & ":K" & i).Locked = True
                         End If
                         Target.Next.Select
                    Else
                        Range("K" & i & ":K" & i).Locked = True
                        Range("K" & i & ":K" & i).value = ""
                        Target.Select
                    End If
                'Base on selected value from Type of Income Respective code will be set in hidden field End
            End Select
        Next
        'new code added for same PIN with Same Type of Income not allowed Start
        If (Application.Intersect(Target.Cells, Range("C" & startRow & ":C" & endRow)) Is Nothing) Then
        'Target Range is not in Given Range
        ElseIf Target.Cells.Count = 1 Then
            Dim PINSchPatrn1W As String
            Dim curPINNo1W As String
            Dim refSchedule1W As String
            Dim chkPINNo1W As String
            Dim chkRefSch1W As String
            
            PINSchPatrn1W = Worksheets("G_Partnership_Income").Range("L" & Target.row & ":L" & Target.row).value
            curPINNo1W = Worksheets("G_Partnership_Income").Range("A" & Target.row & ":A" & Target.row).value
            refSchedule1W = Worksheets("G_Partnership_Income").Range("C" & Target.row & ":C" & Target.row).value
            
            If curPINNo1W <> "" And refSchedule1W <> "" Then
                For Each r In Range("ProfitShare.ListW").Rows
                    For Each c In r.Cells
                        If c.row > endRow Then
                            Exit For
                        End If
                        If c.column = Target.column Then
                            If c.Address = Target.Address Then
                                'Skip comparison with self value
                            Else
                                chkPINNo1W = Worksheets("G_Partnership_Income").Range("A" & c.row & ":A" & c.row).value
                                chkRefSch1W = Worksheets("G_Partnership_Income").Range("C" & c.row & ":C" & c.row).value
                                If UCase(PINSchPatrn1W) = UCase(chkPINNo1W & chkRefSch1W) Then
                                    MsgBox "You have already entered details for " & chkRefSch1W & " for PIN of Partnership Firm " & chkPINNo1W & " in Row " & c.row
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
            
        If (Application.Intersect(Target.Cells, Range("A" & startRow & ":A" & endRow)) Is Nothing) Then
         'Target Range is not in Given Range
        ElseIf Target.Cells.Count = 1 Then
            Dim PINSchPatrn2W As String
            Dim curPINNo2W As String
            Dim refSchedule2W As String
            Dim chkPINNo2W As String
            Dim chkRefSch2W As String
            
            PINSchPatrn2W = Worksheets("G_Partnership_Income").Range("L" & Target.row & ":L" & Target.row).value
            curPINNo2W = Worksheets("G_Partnership_Income").Range("A" & Target.row & ":A" & Target.row).value
            refSchedule2W = Worksheets("G_Partnership_Income").Range("C" & Target.row & ":C" & Target.row).value
            
            If curPINNo2W <> "" And refSchedule2W <> "" Then
                For Each r In Range("ProfitShare.ListW").Rows
                    For Each c In r.Cells
                        If c.row > endRow Then
                            Exit For
                        End If
                        If c.column = Target.column Then
                            If c.Address = Target.Address Then
                                'Skip comparison with self value
                            Else
                                chkPINNo2W = Worksheets("G_Partnership_Income").Range("A" & c.row & ":A" & c.row).value
                                chkRefSch2W = Worksheets("G_Partnership_Income").Range("C" & c.row & ":C" & c.row).value
                                If UCase(PINSchPatrn2W) = UCase(chkPINNo2W & chkRefSch2W) Then
                                    MsgBox "You have already entered details for " & chkRefSch2W & " of PIN of Partnership Firm " & chkPINNo2W & " in Row " & c.row
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
        'new code added for same PIN with Same Type of Income not allowed End
        'Code for Wife End
    End If
endSub:
    
    Worksheets("G_Partnership_Income").Protect (Pwd)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
