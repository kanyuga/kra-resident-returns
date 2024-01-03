VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Add Row Functionality For Self Start
Sub addRow_CarBenefit_ListS()
    Worksheets("I_Computation_of_Car_Benefit").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("CarBenefit.ListS")
    rangeName = Worksheets("I_Computation_of_Car_Benefit").Range("CarBenefit.ListS").Address
    startRowIndex = Range(rangeName).row
    endRowIndex = Range(rangeName).Rows.Count + startRowIndex - 1
    Dim Str As String
    Str = "A" & startRowIndex & ":A" & endRowIndex
    ActiveWorkbook.Names("PINofEmployerCarS").Delete
    Range(Str).name = "PINofEmployerCarS"
    Str = "J" & startRowIndex & ":J" & endRowIndex
    ActiveWorkbook.Names("valueOfCarS").Delete
    Range(Str).name = "valueOfCarS"
    Worksheets("I_Computation_of_Car_Benefit").Protect (Pwd)
End Sub
'Add Row Functionality For Self End

'Add Row Functionality For Wife Start
Sub addRow_CarBenefit_ListW()
    Worksheets("I_Computation_of_Car_Benefit").Unprotect (Pwd)
    Call InsertRowsAndFillFormulas("CarBenefit.ListW")
    rangeName = Worksheets("I_Computation_of_Car_Benefit").Range("CarBenefit.ListW").Address
    startRowIndex = Range(rangeName).row
    endRowIndex = Range(rangeName).Rows.Count + startRowIndex - 1
    Dim Str As String
    Str = "A" & startRowIndex & ":A" & endRowIndex
    ActiveWorkbook.Names("PINofEmployerCarW").Delete
    Range(Str).name = "PINofEmployerCarW"
    Str = "J" & startRowIndex & ":J" & endRowIndex
    ActiveWorkbook.Names("valueOfCarW").Delete
    Range(Str).name = "valueOfCarW"
    Worksheets("I_Computation_of_Car_Benefit").Protect (Pwd)
End Sub
'Add Row Functionality For Wife End


Private Sub Worksheet_Change(ByVal Target As Range)
    
    Worksheets("I_Computation_of_Car_Benefit").Unprotect (Pwd)
    Worksheets("I_Computation_of_Car_Benefit").Activate
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    If (Target.Columns.Count = 1) Then
        
        'Code for Self Start
        rangeName = Worksheets("I_Computation_of_Car_Benefit").Range("CarBenefit.ListS").Address
        startRow = Worksheets("I_Computation_of_Car_Benefit").Range(rangeName).row
        endRow = Worksheets("I_Computation_of_Car_Benefit").Range(rangeName).Rows.Count + startRow - 1
    
        For i = startRow To endRow
            Select Case Target.Address
                'PIN convert in uppar Case Start
                Case "$B$" & i
                    If Target.value <> "" Then
                        Target.value = UCase(Target.value)
                    End If
                'PIN convert in uppar Case End
                'Base on selected item from BodyType,respecitve code will be set in hidden field Start
                Case "$D$" & i
                    If (Target.value <> "") Then
                        rowvalue = find_BodyTypeCode(Target.value)
                        BodyTypeCode = Sheet18.Cells(rowvalue, BodyTypeCodeCol).value
                        If (rowvalue <> 0) Then
                            Range("L" & i & ":L" & i).value = BodyTypeCode
                        End If
                    Else
                        Range("L" & i & ":L" & i).Locked = True
                        Range("L" & i & ":L" & i).value = ""
                    End If
                        Target.Next.Select
                        
                'Base on selected item from BodyType,respecitve code will be set in hidden field End
                
                'Base one selected item from Type of Car Cost,respective code will be set in hidden field Start
                Case "$F$" & i
                    If (Target.value <> "") Then
                        rowvalue = find_OwnHireCode(Target.value)
                        OwnHireCode = Sheet18.Cells(rowvalue, OwnHireCodeCol).value
                        
                        If (rowvalue <> 0) Then
                            Range("M" & i & ":M" & i).value = OwnHireCode
                        End If
                    Else
                        Range("M" & i & ":M" & i).Locked = True
                        Range("M" & i & ":M" & i).value = ""
                    End If
                    If Target.value <> "" And Target.value = "Own" Then
                        Call lockUnlock_cell_rng(Worksheets("I_Computation_of_Car_Benefit").name, "G" & i & ":G" & i, True) 'disable is hired
                        Call lockUnlock_cell_rng(Worksheets("I_Computation_of_Car_Benefit").name, "H" & i & ":H" & i, False) 'enable is owned
                        Target.Next.Select
                    ElseIf Target.value <> "" And Target.value = "Hired" Then
                        Call lockUnlock_cell_rng(Worksheets("I_Computation_of_Car_Benefit").name, "G" & i & ":G" & i, False) 'enable is hired
                        Call lockUnlock_cell_rng(Worksheets("I_Computation_of_Car_Benefit").name, "H" & i & ":H" & i, True) 'disable is owned
                        Target.Next.Select
                    End If
                    
                'Base one selected item from Type of Car Cost,respective code will be set in hidden field End
            End Select
        Next
        
        'Validation for PIN and Car Registration number,PIN and Car Registration Numbe pair should be Unique Start
        If (Application.Intersect(Target.Cells, Range("B" & startRow & ":B" & endRow)) Is Nothing) Then
            'Target Range is not in Given Range
        ElseIf Target.Cells.Count = 1 Then
            
            Dim PINSchPatrn1 As String
            Dim curPINNo1 As String
            Dim refSchedule1 As String
            Dim chkPINNo1 As String
            Dim chkRefSch1 As String
            PINSchPatrn1 = Worksheets("I_Computation_of_Car_Benefit").Range("N" & Target.row & ":N" & Target.row).value
            curPINNo1 = Worksheets("I_Computation_of_Car_Benefit").Range("A" & Target.row & ":A" & Target.row).value
            refSchedule1 = Worksheets("I_Computation_of_Car_Benefit").Range("B" & Target.row & ":B" & Target.row).value
            If curPINNo1 <> "" And refSchedule1 <> "" Then
                For Each r In Range("CarBenefit.ListS").Rows
                    For Each c In r.Cells
                            If c.row > endRow Then
                                Exit For
                            End If
                            If c.column = Target.column Then
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                Else
                                    chkPINNo1 = Worksheets("I_Computation_of_Car_Benefit").Range("A" & c.row & ":A" & c.row).value
                                    chkRefSch1 = Worksheets("I_Computation_of_Car_Benefit").Range("B" & c.row & ":B" & c.row).value
                                    If UCase(PINSchPatrn1) = UCase(chkPINNo1 & chkRefSch1) Then
                                        MsgBox "The Car Registration Number " & chkRefSch1 & " has already been entered for PIN of Employee " & chkPINNo1 & " in Row" & c.row
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
        'Validation for PIN and Car Registration number,PIN and Car Registration Numbe pair should be Unique End
        
        'Validation for PIN and Car Registration number,PIN and Car Registration Numbe pair should be Unique Start
        If (Application.Intersect(Target.Cells, Range("A" & startRow & ":A" & endRow)) Is Nothing) Then
         'Target Range is not in Given Range
        ElseIf Target.Cells.Count = 1 Then
            
            Dim PINSchPatrn2 As String
            Dim curPINNo2 As String
            Dim refSchedule2 As String
            Dim chkPINNo2 As String
            Dim chkRefSch2 As String
            PINSchPatrn2 = Worksheets("I_Computation_of_Car_Benefit").Range("N" & Target.row & ":N" & Target.row).value
            curPINNo2 = Worksheets("I_Computation_of_Car_Benefit").Range("A" & Target.row & ":A" & Target.row).value
            refSchedule2 = Worksheets("I_Computation_of_Car_Benefit").Range("B" & Target.row & ":B" & Target.row).value
            If curPINNo2 <> "" And refSchedule2 <> "" Then
                For Each r In Range("CarBenefit.ListS").Rows
                    For Each c In r.Cells
                            If c.row > endRow Then
                                Exit For
                            End If
                            If c.column = Target.column Then
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                Else
                                    chkPINNo2 = Worksheets("I_Computation_of_Car_Benefit").Range("A" & c.row & ":A" & c.row).value
                                    chkRefSch2 = Worksheets("I_Computation_of_Car_Benefit").Range("B" & c.row & ":B" & c.row).value
                                    If UCase(PINSchPatrn2) = UCase(chkPINNo2 & chkRefSch2) Then
                                        MsgBox "The Car Registration Number " & chkRefSch2 & " has already been entered for PIN of Employee " & chkPINNo2 & " in Row" & c.row
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
        
        
        
        'Validation for PIN and Car Registration number,PIN and Car Registration Numbe pair should be Unique End
    
        'Code for Self End
        
        'Code for Wife Start
        rangeName = Range("CarBenefit.ListW").Address
        startRow = Range(rangeName).row
        endRow = Range(rangeName).Rows.Count + startRow - 1
    
        For i = startRow To endRow
            Select Case Target.Address
                'PIN convert in uppar Case Start
                Case "$B$" & i
                    If Target.value <> "" Then
                        Target.value = UCase(Target.value)
                    End If
                'PIN convert in uppar Case End
                
                Case "$D$" & i
                'Base on selected item from BodyType,respecitve code will be set in hidden field Start
                    If (Target.value <> "") Then
                        rowvalue = find_BodyTypeCode(Target.value)
                        BodyTypeCode = Sheet18.Cells(rowvalue, BodyTypeCodeCol).value
                        If (rowvalue <> 0) Then
                            Range("L" & i & ":L" & i).value = BodyTypeCode
                        End If
                    Else
                        Range("L" & i & ":L" & i).Locked = True
                        Range("L" & i & ":L" & i).value = ""
                    End If
                    Target.Next.Select
                    
                'Base on selected item from BodyType,respecitve code will be set in hidden field End
                
                'Base one selected item from Type of Car Cost,respective code will be set in hidden field Start
                Case "$F$" & i
                    If (Target.value <> "") Then
                        rowvalue = find_OwnHireCode(Target.value)
                        OwnHireCode = Sheet18.Cells(rowvalue, OwnHireCodeCol).value
                         If (rowvalue <> 0) Then
                                 Range("M" & i & ":M" & i).value = OwnHireCode
                         End If
                    Else
                        Range("M" & i & ":M" & i).Locked = True
                        Range("M" & i & ":M" & i).value = ""
                    End If
                    If Target.value <> "" And Target.value = "Own" Then
                        Call lockUnlock_cell_rng(Worksheets("I_Computation_of_Car_Benefit").name, "G" & i & ":G" & i, True) 'disable is hired
                        Call lockUnlock_cell_rng(Worksheets("I_Computation_of_Car_Benefit").name, "H" & i & ":H" & i, False) 'enable is owned
                        Target.Next.Select
                    ElseIf Target.value <> "" And Target.value = "Hired" Then
                        Call lockUnlock_cell_rng(Worksheets("I_Computation_of_Car_Benefit").name, "G" & i & ":G" & i, False) 'enable is hired
                        Call lockUnlock_cell_rng(Worksheets("I_Computation_of_Car_Benefit").name, "H" & i & ":H" & i, True) 'disable is owned
                        Target.Next.Select
                    End If
                        
                'Base one selected item from Type of Car Cost,respective code will be set in hidden field End
            End Select
        Next
    
        'Validation for PIN and Car Registration number,PIN and Car Registration Numbe pair should be Unique Start
        If (Application.Intersect(Target.Cells, Range("B" & startRow & ":B" & endRow)) Is Nothing) Then
            'Target Range is not in Given Range
        ElseIf Target.Cells.Count = 1 Then
            
            Dim PINSchPatrn1W As String
            Dim curPINNo1W As String
            Dim refSchedule1W As String
            Dim chkPINNo1W As String
            Dim chkRefSch1W As String
            PINSchPatrn1W = Worksheets("I_Computation_of_Car_Benefit").Range("N" & Target.row & ":N" & Target.row).value
            curPINNo1W = Worksheets("I_Computation_of_Car_Benefit").Range("A" & Target.row & ":A" & Target.row).value
            refSchedule1W = Worksheets("I_Computation_of_Car_Benefit").Range("B" & Target.row & ":B" & Target.row).value
            If curPINNo1W <> "" And refSchedule1W <> "" Then
                For Each r In Range("CarBenefit.ListW").Rows
                    For Each c In r.Cells
                            If c.row > endRow Then
                                Exit For
                            End If
                            If c.column = Target.column Then
                                If c.Address = Target.Address Then 'Skip comparison with self value
                                Else
                                    chkPINNo1W = Worksheets("I_Computation_of_Car_Benefit").Range("A" & c.row & ":A" & c.row).value
                                    chkRefSch1W = Worksheets("I_Computation_of_Car_Benefit").Range("B" & c.row & ":B" & c.row).value
                                    If UCase(PINSchPatrn1W) = UCase(chkPINNo1W & chkRefSch1W) Then
                                        MsgBox "The Car Registration Number " & chkRefSch1W & " has already been entered for PIN of Employee " & chkPINNo1W & " in Row" & c.row
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
        'Validation for PIN and Car Registration number,PIN and Car Registration Numbe pair should be Unique End
        
        'Validation for PIN and Car Registration number,PIN and Car Registration Numbe pair should be Unique Start
        If (Application.Intersect(Target.Cells, Range("A" & startRow & ":A" & endRow)) Is Nothing) Then
         'Target Range is not in Given Range
        ElseIf Target.Cells.Count = 1 Then
            
            Dim PINSchPatrn2W As String
            Dim curPINNo2W As String
            Dim refSchedule2W As String
            Dim chkPINNo2W As String
            Dim chkRefSch2W As String
            PINSchPatrn2W = Worksheets("I_Computation_of_Car_Benefit").Range("N" & Target.row & ":N" & Target.row).value
            curPINNo2W = Worksheets("I_Computation_of_Car_Benefit").Range("A" & Target.row & ":A" & Target.row).value
            refSchedule2W = Worksheets("I_Computation_of_Car_Benefit").Range("B" & Target.row & ":B" & Target.row).value
            If curPINNo2W <> "" And refSchedule2W <> "" Then
                For Each r In Range("CarBenefit.ListW").Rows
                    For Each c In r.Cells
                        If c.row > endRow Then
                            Exit For
                        End If
                        If c.column = Target.column Then
                            If c.Address = Target.Address Then 'Skip comparison with self value
                            Else
                                chkPINNo2W = Worksheets("I_Computation_of_Car_Benefit").Range("A" & c.row & ":A" & c.row).value
                                chkRefSch2W = Worksheets("I_Computation_of_Car_Benefit").Range("B" & c.row & ":B" & c.row).value
                                If UCase(PINSchPatrn2W) = UCase(chkPINNo2W & chkRefSch2W) Then
                                    MsgBox "The Car Registration Number " & chkRefSch2W & " has already been entered for PIN of Employee " & chkPINNo2W & " in Row" & c.row
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
        'Validation for PIN and Car Registration number,PIN and Car Registration Numbe pair should be Unique End
        'Code for Wife End
    End If
endSub:
    
    Worksheets("I_Computation_of_Car_Benefit").Protect (Pwd)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
