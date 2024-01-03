VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Add Row Functionality for Part 1 Self Start
Sub addRow_Inv_QuantDtlsListS()
ActiveSheet.Unprotect (Pwd)
Call InsertRowsAndFillFormulas("Inv.QuantDtlsListS")
ActiveSheet.Protect (Pwd)
End Sub
'Add Row Functionality for Part 1 Self end

'Add Row Functionality for Part 1 Wife Start
Sub addRow_Inv_QuantDtlsListW()
ActiveSheet.Unprotect (Pwd)
Call InsertRowsAndFillFormulas("Inv.QuantDtlsListW")
ActiveSheet.Protect (Pwd)
End Sub
'Add Row Functionality for Part 1 Wife End

'Add Row Functionality for Part 2 Self Start
Sub addRow_Inv_QuantDtlsListKIIS()
ActiveSheet.Unprotect (Pwd)
Call InsertRowsAndFillFormulas("Inv.QuantDtlsListKIIS")
ActiveSheet.Protect (Pwd)
End Sub
'Add Row Functionality for Part 2 Self End

'Add Row Functionality for Part 2 Wife Start
Sub addRow_Inv_QuantDtlsListKIIW()
ActiveSheet.Unprotect (Pwd)
Call InsertRowsAndFillFormulas("Inv.QuantDtlsListKIIW")
ActiveSheet.Protect (Pwd)
End Sub
'Add Row Functionality for Part 2 Wife End


Private Sub Worksheet_Change(ByVal Target As Range)
    
    Worksheets("D_Stock_Analysis").Unprotect (Pwd)
    Worksheets("D_Stock_Analysis").Activate
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    If (Target.Columns.Count = 1) Then
        Dim resultMsg As String
        'Code for Part 1 Self Start
        rangeName = Worksheets("D_Stock_Analysis").Range("Inv.QuantDtlsListS").Address
        startRow = Worksheets("D_Stock_Analysis").Range(rangeName).row
        endRow = Worksheets("D_Stock_Analysis").Range(rangeName).Rows.Count + startRow - 1
    
        For i = startRow To endRow
            Select Case Target.Address
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Start
                Case "$C$" & i
                    cellRow = Target.row
                    If (Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value <> "") Then
                        If (Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value > Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value) Then
                            MsgBox "Sale Quantity should be less than or equal to sum of Opening Stock and Purchase."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase End
                
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Start
                Case "$D$" & i
                    cellRow = Target.row
                    If (Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value <> "") Then
                        If (Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value > Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value) Then
                            MsgBox "Sale Quantity should be less than or equal to sum of Opening Stock and Purchase."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase End
    
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Start
                 Case "$E$" & i
                    cellRow = Target.row
                    If (Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value <> "") Then
                        If (Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value > Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value) Then
                            MsgBox "Sale Quantity should be less than or equal to sum of Opening Stock and Purchase."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase End
            End Select
        Next
    'Code for Part 1 Self End
    
    'Code for Part 2 Self Start
        rangeName = Worksheets("D_Stock_Analysis").Range("Inv.QuantDtlsListKIIS").Address
        startRow = Worksheets("D_Stock_Analysis").Range(rangeName).row
        endRow = Worksheets("D_Stock_Analysis").Range(rangeName).Rows.Count + startRow - 1
        For i = startRow To endRow
            Select Case Target.Address
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Start
                Case "$C$" & i
                    cellRow = Target.row
                    If (Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("F" & cellRow & ":F" & cellRow).value <> "") Then
                        If (Worksheets("D_Stock_Analysis").Range("F" & cellRow & ":F" & cellRow).value > Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value - Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value) Then
                            MsgBox "Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Less Consumption.."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase End
                
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Start
                Case "$D$" & i
                    cellRow = Target.row
                    If (Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("F" & cellRow & ":F" & cellRow).value <> "") Then
                        If (Worksheets("D_Stock_Analysis").Range("F" & cellRow & ":F" & cellRow).value > Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value - Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value) Then
                            MsgBox "Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Less Consumption.."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase End
    
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Start
                Case "$F$" & i
                    cellRow = Target.row
                    If (Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("F" & cellRow & ":F" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value <> "") Then
                        If (Worksheets("D_Stock_Analysis").Range("F" & cellRow & ":F" & cellRow).value > Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value - Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value) Then
                            MsgBox "Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Less Consumption."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                    If (Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("F" & cellRow & ":F" & cellRow).value <> "") Then
                        If (Worksheets("D_Stock_Analysis").Range("F" & cellRow & ":F" & cellRow).value > Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value) Then
                            MsgBox "Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Less Consumption."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase End
                 
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Start
                Case "$E$" & i
                    cellRow = Target.row
                    If (Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("F" & cellRow & ":F" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value <> "") Then
                        If (Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value > Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value - Worksheets("D_Stock_Analysis").Range("F" & cellRow & ":F" & cellRow).value) Then
                            MsgBox "Consumption should be less than or equal to sum of Opening Stock and Purchase Less Sales Quantity."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                     If (Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value <> "") Then
                        If (Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value > Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value) Then
                            MsgBox "Consumption should be less than or equal to sum of Opening Stock and Purchase Less Sales Quantity."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
            End Select
        Next
    'Code for Part 2 Self End
    
    'Code for Part 1 wife Start
        rangeName = Worksheets("D_Stock_Analysis").Range("Inv.QuantDtlsListW").Address
        startRow = Worksheets("D_Stock_Analysis").Range(rangeName).row
        endRow = Worksheets("D_Stock_Analysis").Range(rangeName).Rows.Count + startRow - 1
    
        For i = startRow To endRow
            Select Case Target.Address
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Start
                Case "$C$" & i
                    cellRow = Target.row
                    If (Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value <> "") Then
                        If (Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value > Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value) Then
                            MsgBox "Sale Quantity should be less than or equal to sum of Opening Stock and Purchase."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase End
                
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Start
                Case "$D$" & i
                    cellRow = Target.row
                    If (Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value <> "") Then
                        If (Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value > Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value) Then
                            MsgBox "Sale Quantity should be less than or equal to sum of Opening Stock and Purchase."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase End
                
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Start
                Case "$E$" & i
                    cellRow = Target.row
                    If (Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value <> "") Then
                        If (Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value > Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value) Then
                            MsgBox "Sale Quantity should be less than or equal to sum of Opening Stock and Purchase."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase End
            End Select
        Next
    'Code for Part 1 wife End
    
    'Code for Part 2 wife Start
        rangeName = Worksheets("D_Stock_Analysis").Range("Inv.QuantDtlsListKIIW").Address
        startRow = Worksheets("D_Stock_Analysis").Range(rangeName).row
        endRow = Worksheets("D_Stock_Analysis").Range(rangeName).Rows.Count + startRow - 1
    
        For i = startRow To endRow
            Select Case Target.Address
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Start
                Case "$C$" & i
                    cellRow = Target.row
                    If (Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("F" & cellRow & ":F" & cellRow).value <> "") Then
                        If (Worksheets("D_Stock_Analysis").Range("F" & cellRow & ":F" & cellRow).value > Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value - Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value) Then
                            MsgBox "Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Less Consumption.."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase End
                
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Start
                Case "$D$" & i
                    cellRow = Target.row
                    
                    If (Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("F" & cellRow & ":F" & cellRow).value <> "") Then
                        If (Worksheets("D_Stock_Analysis").Range("F" & cellRow & ":F" & cellRow).value > Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value - Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value) Then
                            MsgBox "Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Less Consumption.."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase End
                
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Start
                Case "$F$" & i
                    cellRow = Target.row
                    If (Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("F" & cellRow & ":F" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value <> "") Then
                        If (Worksheets("D_Stock_Analysis").Range("F" & cellRow & ":F" & cellRow).value > Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value - Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value) Then
                            MsgBox "Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Less Consumption."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                    If (Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("F" & cellRow & ":F" & cellRow).value <> "") Then
                        If (Worksheets("D_Stock_Analysis").Range("F" & cellRow & ":F" & cellRow).value > Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value) Then
                            MsgBox "Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Less Consumption."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase End
                
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase Start
                Case "$E$" & i
                    cellRow = Target.row
                    
                    If (Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("F" & cellRow & ":F" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value <> "") Then
                        If (Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value > Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value - Worksheets("D_Stock_Analysis").Range("F" & cellRow & ":F" & cellRow).value) Then
                            MsgBox "Consumption should be less than or equal to sum of Opening Stock and Purchase Less Sales Quantity."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                    If (Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value <> "" And Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value <> "") Then
                        If (Worksheets("D_Stock_Analysis").Range("E" & cellRow & ":E" & cellRow).value > Worksheets("D_Stock_Analysis").Range("C" & cellRow & ":C" & cellRow).value + Worksheets("D_Stock_Analysis").Range("D" & cellRow & ":D" & cellRow).value) Then
                            MsgBox "Consumption should be less than or equal to sum of Opening Stock and Purchase Less Sales Quantity."
                            Target.value = ""
                            Target.Select
                        End If
                    End If
                    
                'check Sale Quantity should be less than or equal to sum of Opening Stock and Purchase End
            End Select
        Next
    'Code for Part 2 wife Start
    End If
endSub:
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Worksheets("D_Stock_Analysis").Protect (Pwd)
End Sub


