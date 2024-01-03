VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

    
    Worksheets("E2_CA_WTA_WDV").Unprotect (Pwd)
    Worksheets("E2_CA_WTA_WDV").Activate
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim columnSelf As Integer
    Dim columnWife As Integer
    columnSelf = 7
    columnWife = 19
    WDVBS = Worksheets("E2_CA_WTA_WDV").Range("B5:B5").value
    InDecBS = Worksheets("E2_CA_WTA_WDV").Range("B6:B6").value
    DisValueBS = Worksheets("E2_CA_WTA_WDV").Range("B7:B7").value
    
    WDVCS = Worksheets("E2_CA_WTA_WDV").Range("C5:C5").value
    InDecCS = Worksheets("E2_CA_WTA_WDV").Range("C6:C6").value
    DisValueCS = Worksheets("E2_CA_WTA_WDV").Range("C7:C7").value
    
    WDVDS = Worksheets("E2_CA_WTA_WDV").Range("D5:D5").value
    InDecDS = Worksheets("E2_CA_WTA_WDV").Range("D6:D6").value
    DisValueDS = Worksheets("E2_CA_WTA_WDV").Range("D7:D7").value
    
    WDVES = Worksheets("E2_CA_WTA_WDV").Range("E5:E5").value
    InDecES = Worksheets("E2_CA_WTA_WDV").Range("E6:E6").value
    DisValueES = Worksheets("E2_CA_WTA_WDV").Range("E7:E7").value
    
    
    WDVBW = Worksheets("E2_CA_WTA_WDV").Range("B17:B17").value
    InDecBW = Worksheets("E2_CA_WTA_WDV").Range("B18:B18").value
    DisValueBW = Worksheets("E2_CA_WTA_WDV").Range("B19:B19").value
    
    WDVCW = Worksheets("E2_CA_WTA_WDV").Range("C17:C17").value
    InDecCW = Worksheets("E2_CA_WTA_WDV").Range("C18:C18").value
    DisValueCW = Worksheets("E2_CA_WTA_WDV").Range("C19:C19").value
    
    WDVDW = Worksheets("E2_CA_WTA_WDV").Range("D17:D17").value
    InDecDW = Worksheets("E2_CA_WTA_WDV").Range("D18:D18").value
    DisValueDW = Worksheets("E2_CA_WTA_WDV").Range("D19:D19").value
    
    WDVEW = Worksheets("E2_CA_WTA_WDV").Range("E17:E17").value
    InDecEW = Worksheets("E2_CA_WTA_WDV").Range("E18:E18").value
    DisValueEW = Worksheets("E2_CA_WTA_WDV").Range("E19:E19").value

    If (Target.Columns.Count = 1) Then
        Select Case Target.Address
            'Validation for Disposal During the year,Disposal During the year should be less than or equal to sum of written down value begining of the year and Additions during the year net of Investment Deduction for Class A Self Start
            Case "$B$" & columnSelf
                If (WDVBS <> "" And InDecBS <> "" And DisValueBS <> "") Then
                   If (DisValueBS > WDVBS + InDecBS) Then
                       MsgBox "Disposal during the Year should be less than or equal to A+B for Class A"
                       Worksheets("E2_CA_WTA_WDV").Range("B7:B7").value = ""
                       Worksheets("E2_CA_WTA_WDV").Range("B7:B7").Select
                   End If
                End If
                
                If (DisValueBS <> "") Then
                    If (WDVBS = "" Or InDecBS = "" And DisValueBS <> "") Then
                       MsgBox "First Enter Written Down Value (WDV) at the Beginning of the Year or Additions during the Year net of Investment Deduction."
                       Worksheets("E2_CA_WTA_WDV").Range("B7:B7").value = ""
                       Worksheets("E2_CA_WTA_WDV").Range("B7:B7").Select
                    End If
                 End If
            'Validation for Disposal During the year,Disposal During the year should be less than or equal to sum of written down value begining of the year and Additions during the year net of Investment Deduction for Class A Self End
                 
            'Validation for Disposal During the year,Disposal During the year should be less than or equal to sum of written down value begining of the year and Additions during the year net of Investment Deduction for Class B Self Start
            Case "$C$" & columnSelf
                If (WDVCS <> "" And InDecCS <> "" And DisValueCS <> "") Then
                   If (DisValueCS > WDVCS + InDecCS) Then
                       MsgBox "Disposal during the Year should be less than or equal to A+B for Class B"
                       Worksheets("E2_CA_WTA_WDV").Range("C7:C7").value = ""
                       Worksheets("E2_CA_WTA_WDV").Range("C7:C7").Select
                   End If
                End If
                
                If (DisValueCS <> "") Then
                    If (WDVCS = "" Or InDecCS = "" And DisValueCS <> "") Then
                       MsgBox "First Enter Written Down Value (WDV) at the Beginning of the Year or Additions during the Year net of Investment Deduction."
                       Worksheets("E2_CA_WTA_WDV").Range("C7:C7").value = ""
                       Worksheets("E2_CA_WTA_WDV").Range("C7:C7").Select
                    End If
                 End If
            'Validation for Disposal During the year,Disposal During the year should be less than or equal to sum of written down value begining of the year and Additions during the year net of Investment Deduction for Class B Self End
                 
            'Validation for Disposal During the year,Disposal During the year should be less than or equal to sum of written down value begining of the year and Additions during the year net of Investment Deduction for Class C Self Start
            Case "$D$" & columnSelf
                If (WDVDS <> "" And InDecDS <> "" And DisValueDS <> "") Then
                   If (DisValueDS > WDVDS + InDecDS) Then
                       MsgBox "Disposal during the Year should be less than or equal to A+B for Class C"
                       Worksheets("E2_CA_WTA_WDV").Range("D7:D7").value = ""
                       Worksheets("E2_CA_WTA_WDV").Range("D7:D7").Select
                   End If
                End If
                
                If (DisValueDS <> "") Then
                    If (WDVDS = "" Or InDecDS = "" And DisValueDS <> "") Then
                       MsgBox "First Enter Written Down Value (WDV) at the Beginning of the Year or Additions during the Year net of Investment Deduction."
                       Worksheets("E2_CA_WTA_WDV").Range("D7:D7").value = ""
                       Worksheets("E2_CA_WTA_WDV").Range("D7:D7").Select
                    End If
                 End If
            'Validation for Disposal During the year,Disposal During the year should be less than or equal to sum of written down value begining of the year and Additions during the year net of Investment Deduction for Class C Self End
                 
            'Validation for Disposal During the year,Disposal During the year should be less than or equal to sum of written down value begining of the year and Additions during the year net of Investment Deduction for Class D Self Start
            Case "$E$" & columnSelf
                If (WDVES <> "" And InDecES <> "" And DisValueES <> "") Then
                   If (DisValueES > WDVES + InDecES) Then
                       MsgBox "Disposal during the Year should be less than or equal to A+B for Class D"
                       Worksheets("E2_CA_WTA_WDV").Range("E7:E7").value = ""
                       Worksheets("E2_CA_WTA_WDV").Range("E7:E7").Select
                   End If
                End If
                
                If (DisValueES <> "") Then
                    If (WDVES = "" Or InDecES = "" And DisValueES <> "") Then
                       MsgBox "First Enter Written Down Value (WDV) at the Beginning of the Year or Additions during the Year net of Investment Deduction."
                       Worksheets("E2_CA_WTA_WDV").Range("E7:E7").value = ""
                       Worksheets("E2_CA_WTA_WDV").Range("E7:E7").Select
                    End If
                 End If
            'Validation for Disposal During the year,Disposal During the year should be less than or equal to sum of written down value begining of the year and Additions during the year net of Investment Deduction for Class D Self End
            
            
            'Validation for Disposal During the year,Disposal During the year should be less than or equal to sum of written down value begining of the year and Additions during the year net of Investment Deduction for Class A Wife Start
            Case "$B$" & columnWife
                If (WDVBW <> "" And InDecBW <> "" And DisValueBW <> "") Then
                    If (DisValueBW > WDVBW + InDecBW) Then
                        MsgBox "Disposal during the Year should be less than or equal to A+B for Class A"
                        Worksheets("E2_CA_WTA_WDV").Range("B19:B19").value = ""
                        Worksheets("E2_CA_WTA_WDV").Range("B19:B19").Select
                    End If
                End If
                
                If (DisValueBW <> "") Then
                    If (WDVBW = "" Or InDecBW = "" And DisValueBW <> "") Then
                       MsgBox "First Enter Written Down Value (WDV) at the Beginning of the Year or Additions during the Year net of Investment Deduction."
                       Worksheets("E2_CA_WTA_WDV").Range("B19:B19").value = ""
                        Worksheets("E2_CA_WTA_WDV").Range("B19:B19").Select
                    End If
                 End If
                
            'Validation for Disposal During the year,Disposal During the year should be less than or equal to sum of written down value begining of the year and Additions during the year net of Investment Deduction for Class A Wife End
             
            'Validation for Disposal During the year,Disposal During the year should be less than or equal to sum of written down value begining of the year and Additions during the year net of Investment Deduction for Class B Wife Start
            Case "$C$" & columnWife
                If (WDVCW <> "" And InDecCW <> "" And DisValueCW <> "") Then
                   If (DisValueCW > WDVCW + InDecCW) Then
                       MsgBox "Disposal during the Year should be less than or equal to A+B for Class B"
                       Worksheets("E2_CA_WTA_WDV").Range("C19:C19").value = ""
                       Worksheets("E2_CA_WTA_WDV").Range("C19:C19").Select
                   End If
                End If
                
                 If (DisValueCW <> "") Then
                    If (WDVCW = "" Or InDecCW = "" And DisValueCW <> "") Then
                       MsgBox "First Enter Written Down Value (WDV) at the Beginning of the Year or Additions during the Year net of Investment Deduction."
                       Worksheets("E2_CA_WTA_WDV").Range("C19:C19").value = ""
                        Worksheets("E2_CA_WTA_WDV").Range("C19:C19").Select
                    End If
                 End If
            'Validation for Disposal During the year,Disposal During the year should be less than or equal to sum of written down value begining of the year and Additions during the year net of Investment Deduction for Class B Wife End
             
            'Validation for Disposal During the year,Disposal During the year should be less than or equal to sum of written down value begining of the year and Additions during the year net of Investment Deduction for Class C Wife Start
            Case "$D$" & columnWife
                If (WDVDW <> "" And InDecDW <> "" And DisValueDW <> "") Then
                   If (DisValueDW > WDVDW + InDecDW) Then
                       MsgBox "Disposal during the Year should be less than or equal to A+B for Class C"
                       Worksheets("E2_CA_WTA_WDV").Range("D19:D19").value = ""
                       Worksheets("E2_CA_WTA_WDV").Range("D19:D19").Select
                   End If
                End If
                
                If (DisValueDW <> "") Then
                    If (WDVDW = "" Or InDecDW = "" And DisValueDW <> "") Then
                       MsgBox "First Enter Written Down Value (WDV) at the Beginning of the Year or Additions during the Year net of Investment Deduction."
                       Worksheets("E2_CA_WTA_WDV").Range("D19:D19").value = ""
                        Worksheets("E2_CA_WTA_WDV").Range("D19:D19").Select
                    End If
                 End If
                
            'Validation for Disposal During the year,Disposal During the year should be less than or equal to sum of written down value begining of the year and Additions during the year net of Investment Deduction for Class C Wife End
                 
            'Validation for Disposal During the year,Disposal During the year should be less than or equal to sum of written down value begining of the year and Additions during the year net of Investment Deduction for Class D Wife Start
            Case "$E$" & columnWife
               If (WDVEW <> "" And InDecEW <> "" And DisValueEW <> "") Then
                  If (DisValueEW > WDVEW + InDecEW) Then
                      MsgBox "Disposal during the Year should be less than or equal to A+B for Class D"
                      Worksheets("E2_CA_WTA_WDV").Range("E19:E19").value = ""
                      Worksheets("E2_CA_WTA_WDV").Range("E19:E19").Select
                  End If
               End If
               
               If (DisValueEW <> "") Then
                    If (WDVEW = "" Or InDecEW = "" And DisValueEW <> "") Then
                       MsgBox "First Enter Written Down Value (WDV) at the Beginning of the Year or Additions during the Year net of Investment Deduction."
                       Worksheets("E2_CA_WTA_WDV").Range("D19:D19").value = ""
                        Worksheets("E2_CA_WTA_WDV").Range("D19:D19").Select
                    End If
                 End If
               
            'Validation for Disposal During the year,Disposal During the year should be less than or equal to sum of written down value begining of the year and Additions during the year net of Investment Deduction for Class D Wife End
        End Select
    End If
endSub:
    
    Worksheets("E2_CA_WTA_WDV").Protect (Pwd)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

