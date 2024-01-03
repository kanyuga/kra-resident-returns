Attribute VB_Name = "Module1"
Global counter As Long
Global tableStr As String
Global cellCounter As Long
Global isGenUpload As Boolean
Global FolderName As String
Global Pwd
Public Const BankIdCol = 44
Public Const BranchIdCol = 47
Public Const TaxRateCol = 2
Public Const natureOfPaymentCodeCol = 3
Public Const natureOfIncomeCodeCol = 4
Public Const usageCodeCol = 4
Public Const PolicyHolderCodeCol = 4
Public Const BodyTypeCodeCol = 6
Public Const OwnHireCodeCol = 2
Public Const RetTypeCodeCol = 5

'variables for forceful macro enabling
Global isCalledOnClose As Boolean
Global isHideUnhidePerformed As Boolean

Global startRowIndex As Integer
Global endRowIndex As Integer

Dim cellBreak As Boolean
Dim prevCellBreakCounterEndRow As Integer
Dim previousLogicalName As String
Dim bufferRow As Integer
'Added by Ruth
'To monitor the level to start taxing business income for the period 2020
Dim taxBandJanMar As Integer
Dim taxBandAprDec As Integer
Dim TotalTaxableIncomeAprDec As Double


Public Sub Generate_upload()
    Worksheets("Data").Unprotect (Pwd)
    Dim generateXML As Boolean
    generateXML = False
    If UCase(Worksheets("Data").Range("genXML").value) = "Y" Then
        generateXML = True
    End If
    Worksheets("Data").Protect (Pwd)
    
    If generateXML Then
        Call Generate_upload_xml
    Else
        Call Generate_upload_xls
    End If
    
End Sub

'Generate_upload_xls Function to generate sheet after completed validation and not found any error Start
Public Sub Generate_upload_xls()
    On Error GoTo ErrorHandle
    isGenUpload = True
    Dim FName, curbook, cursheet, totr, totc, bigmsg
    Dim filesavename As String, Encrypt As String
    Dim fileName As String
    Dim path As String
    Dim uploadFileNm As String
    Dim i As Integer
    Dim newbook, newbook1 As Object
    
    Worksheets("Sheet1").Unprotect (Pwd)
    row_count = row_count + 4
    curbook = ActiveWorkbook.name
    cursheet = "Sheet1"
    Worksheets("Sheet1").Visible = xlSheetVeryHidden
    
    
    
    'File Name Format : DATE_SYSTEMTIME_PIN_OBLIGATIONNAME (Recieved 21/06/2013)  Old : 'IT1_Individual_Resident_Return_Upload.zip
    
    Dim taxPayerPin As String
    taxPayerPin = Worksheets("A_Basic_Info").Range("RetInf.PIN").value
    uploadFileNm = Format(Now, "dd-mm-yyyy") & "_" & Format(Now, "h-mm-ss") & "_" & taxPayerPin & "_ITR"
    'uploadFileNm = Left(curbook, Len(curbook) - 4) & "_Upload"
    
    'generate folder to save upload sheets
    DefPath = Application.DefaultFilePath
    If Right(DefPath, 1) <> "\" Then
        DefPath = DefPath & "\"
    End If

    'strDate = Format(Now, "dd-mm-yyyy")
    'path = DefPath & strDate
    path = DefPath & uploadFileNm
    MkDir path
    FolderName = path
    
    Set newbook = CreateObject("Excel.sheet")
    fileName = uploadFileNm & ".xls"
    
    
    'ChDir path
    ChDir FolderName
        If Application.Version >= 12 Then
            
            newbook.SaveAs fileName:=FolderName & "\" & fileName, FileFormat:=56
        Else
            newbook.SaveAs fileName:=fileName
        End If
        isGenUpload = True
        Set newbook = Nothing
        Set newbook1 = GetObject(FolderName & "\" & fileName)
        Workbooks(curbook).Activate
        Worksheets(cursheet).Activate
    
        isGenUpload = True
        'code to get multi cell upload String
        'Range increased from 3 to 5 By Janhavi
        Range(Cells(1, 1), Cells(5, counter)).Select
        Selection.Copy
        Worksheets("Sheet1").Visible = xlSheetHidden
        Worksheets("Sheet1").Activate
        ActiveWindow.WindowState = xlMaximized
        Windows(fileName).Activate
        Selection.PasteSpecial Paste:=xlValues
        Selection.ColumnWidth = 0
    
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
            :=xlBetween
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
        ActiveSheet.Columns("A:A").ColumnWidth = 0
        ActiveSheet.Columns("K:K").ColumnWidth = 0
        ActiveSheet.name = "Sheet1"
        ' encryption function
        For i = 1 To row_count
            Encrypt = ActiveSheet.Cells(i, 1).value
            ActiveSheet.Cells(i, 1).value = Encrypt
        Next i
        ActiveSheet.Cells.Locked = True
        ActiveSheet.Cells.WrapText = False
        ActiveSheet.Protect Password:=Pwd
        ActiveWorkbook.Close SaveChanges:=True
        Worksheets("Sheet1").Protect (Pwd)
        
        'generate zip file - added by mitali
'        test = A_Zip_Folder_And_SubFolders_Browse(folderName, uploadFileNm) ' return false in case zip is not installed
'        If test = True Then
'            Kill folderName & "\" & fileName
'        Else
'            MsgBox ("No Error Found, Upload file is saved : """ & folderName & "\" & fileName & """.")
'        End If

        'Code Changed By Sameer (379755) to resolve Machine Dependent Issue
        Call Zip_All_Files_in_Folder_Browse(FolderName, fileName)
        Call Delete_Folder(FolderName)
         
        isGenUpload = False
    Exit Sub
      
ErrorHandle:
        MsgBox ("Modifications Are Not Saved,Upload File Not Generated" & Err.Description)
        ActiveWindow.WindowState = xlMaximized
        Exit Sub
        Resume
End Sub
'Generate_upload_xls Function to generate sheet after completed validation and not found any error End

'Generate_upload_xml Function to generate sheet after completed validation and not found any error Start
Public Sub Generate_upload_xml()
    On Error GoTo ErrorHandle
    isGenUpload = True
    Dim fileName As String
    Dim path As String
    Dim uploadFileNm As String
    Dim final_xml As String
    Dim taxPayerPin As String
    
    Worksheets("Sheet1").Unprotect (Pwd)
    row_count = row_count + 4
    Worksheets("Sheet1").Visible = xlSheetVeryHidden
    
    final_xml = getFinalXml()
    
    'File Name Format : DATE_SYSTEMTIME_PIN_OBLIGATIONNAME (Recieved 21/06/2013)  Old : 'IT1_Individual_Resident_Return_Upload.zip
    
    taxPayerPin = Worksheets("A_Basic_Info").Range("RetInf.PIN").value
    uploadFileNm = Format(Now, "dd-mm-yyyy") & "_" & Format(Now, "h-mm-ss") & "_" & taxPayerPin & "_ITR"
    
    'generate folder to save upload sheets
    DefPath = Application.DefaultFilePath
    If Right(DefPath, 1) <> "\" Then
        DefPath = DefPath & "\"
    End If
    
    path = DefPath & uploadFileNm
    MkDir path
    FolderName = path
    
    sFileSaveName = path & "\" & uploadFileNm & ".xml"
    If sFileSaveName <> False Then
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set xmlFile = FSO.CreateTextFile(sFileSaveName, True)
        xmlFile.Write final_xml
        xmlFile.Close
        'MsgBox ("XML Upload file is created at location: " & sFileSaveName)
        'excelObj.Application.Quit
    Else
        MsgBox ("XML Upload file not created")
    End If
    
    fileName = sFileSaveName
    
    ChDir FolderName
        
    Worksheets("Sheet1").Protect (Pwd)
        
    Call Zip_All_Files_in_Folder_Browse(FolderName, fileName)
    Call Delete_Folder(FolderName)
         
    isGenUpload = False
    Exit Sub
      
ErrorHandle:
    MsgBox ("Modifications Are Not Saved,Upload File Not Generated" & Err.Description)
    ActiveWindow.WindowState = xlMaximized
    Exit Sub
    Resume
End Sub
'Generate_upload_xml Function to generate xml after completed validation and not found any error End


Public Function getFinalXml() As String
    getFinalXml = ""
    Dim temp As String
    temp = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf _
            & "<Sheet>" & vbCrLf _
            & "<SingleCellValue>" _
            & "${sc}" _
            & "</SingleCellValue>" & vbCrLf _
            & "<MultiCellValue>" _
            & "${mc}" _
            & "</MultiCellValue>" & vbCrLf _
            & "<SingleCellHash>" _
            & "${sch}" _
            & "</SingleCellHash>" & vbCrLf _
            & "<MultiCellHash>" _
            & "${mch}" _
            & "</MultiCellHash>" & vbCrLf _
            & "<SheetCode>" _
            & "${scode}" _
            & "</SheetCode>" & vbCrLf _
            & "</Sheet>"
    
    Dim singleCell, multiCell, singleCellHash, multiCellHash, sheetCode As String
    singleCell = ""
    multiCell = ""
    'singleCell = Worksheets("Sheet1").Range("$A$1").value
    'multiCell = Worksheets("Sheet1").Range("$A$2").value
    For i = 1 To 255
        If Worksheets("Sheet1").Cells(1, i).value <> "" Then
            singleCell = singleCell & Worksheets("Sheet1").Cells(1, i).value
        Else
         Exit For
        End If
    Next
    For i = 1 To 255
        If Worksheets("Sheet1").Cells(2, i).value <> "" Then
            multiCell = multiCell & Worksheets("Sheet1").Cells(2, i).value
        Else
         Exit For
        End If
    Next
    
    singleCellHash = Worksheets("Sheet1").Range("$A$10").value
    multiCellHash = Worksheets("Sheet1").Range("$A$11").value
    sheetCode = Worksheets("Sheet1").Range("$A$3").value
    
    singleCell = Replace(singleCell, "&", "&amp;")
    singleCell = Replace(singleCell, "<", "&lt;")
    singleCell = Replace(singleCell, ">", "&gt;")
    multiCell = Replace(multiCell, "&", "&amp;")
    multiCell = Replace(multiCell, "<", "&lt;")
    multiCell = Replace(multiCell, ">", "&gt;")
    
    temp = Replace(temp, "${sc}", singleCell)
    temp = Replace(temp, "${mc}", multiCell)
    temp = Replace(temp, "${sch}", singleCellHash)
    temp = Replace(temp, "${mch}", multiCellHash)
    temp = Replace(temp, "${scode}", sheetCode)
    
    getFinalXml = temp
End Function

'Date Validation Check Date Format Start
Public Function TestDate(strDate As String) As Boolean
    Dim strMonth As String
    Dim strDay As String
    Dim strYear As String
    Dim validDay As Boolean
    Dim validMonth As Boolean
    Dim validYear As Boolean
    validDay = False
    validMonth = False
    validYear = False
    Dim validFormat As Boolean
    Dim i As Integer
    Dim Length As Integer
    Dim temp As String
    Dim index1 As Integer
    Dim index2 As Integer
    Dim flag As Integer
    
    TestDate = False
    If IsDate(strDate) = True Then
       Length = Len(strDate)
       flag = 0
       
       If (Length > 10 Or Length < 10) Then
            validFormat = False
       Else
            validFormat = True
            For i = 1 To Length
                If (Mid(strDate, i, 1) = "/") Then 'Or Mid(strDate, i, 1) = "-"
                    flag = flag + 1
                    
                    If (flag = 1) Then
                        index1 = i
                    End If
                    If (flag = 2) Then
                        index2 = i
                    End If
                    
                End If
            Next i
            If (index1 > 0 And index2 > 0) Then
            strDay = Mid(strDate, 1, index1 - 1)
            If (Len(strDay) > 2) Then
                validDay = False
            Else
                If IsNumeric(strDay) = True Then
                    If Val(strDay) > 32 Then
                        validDay = False
                    Else
                        validDay = True
                    End If
                Else
                    validDay = True
                End If
            End If
            
            
            strMonth = Mid(strDate, index1 + 1, (index2 - index1) - 1)
            If (Len(strMonth) > 2) Then
                validDay = False
            Else
                If IsNumeric(strMonth) = True Then
                    If Val(strMonth) > 12 Then
                        validMonth = False
                    Else
                        validMonth = True
                    End If
                Else
                    validMonth = True
                End If
            End If
            
            strYear = Mid(strDate, index2 + 1, Len(strDate))
            If (Len(strYear) > 4) Then
                validYear = False
            Else
            
                If IsNumeric(strYear) = True Then
                    validYear = True
                Else
                    validYear = False
                End If
            End If
       End If
       End If
        If (validFormat = False Or validDay = False Or validMonth = False Or validYear = False) Then
               TestDate = False
        Else
               TestDate = True
        End If
    End If
End Function
'Date Validation Check Date Format End

'Check Numeric Field,allow only Numeric value Start
Public Function TestNumber(lstr_check As String) As Boolean
'allowed characters 0 to 9 and .
Dim i As Integer
Dim ia As Integer
Dim ina As Integer
Dim stlen As Integer

stlen = Len(lstr_check)
ia = 0
ina = 0
For i = 1 To stlen
    Select Case (Mid(lstr_check, i, 1))
       Case "."
        ia = ia + 1
        Case "0" To "9"          '0 to 9
        ia = ia + 1
        Case Else
            ina = ina + 1
    End Select
Next i
If ina = 0 Then
    If InStr(1, lstr_check, ".") = 0 Then
        TestNumber = True
    Else
        TestNumber = True
    End If
Else
    TestNumber = False
End If

End Function
'Check Numeric Field,allow only Numeric value End

'Check Account related Numeric Field,allow only Account Related Numeric value Start
Public Function TestAccountNumber(lstr_check As String) As Boolean
'allowed characters 0 to 9 and .
Dim i As Integer
Dim ia As Integer
Dim ina As Integer
Dim stlen As Integer

stlen = Len(lstr_check)
ia = 0
ina = 0
For i = 1 To stlen
    Select Case (Mid(lstr_check, i, 1))
       
       Case "0" To "9"          '0 to 9
        ia = ia + 1
        Case Else
            ina = ina + 1
    End Select
Next i
If ina = 0 Then
    If InStr(1, lstr_check, ".") = 0 Then
        TestAccountNumber = True
    Else
        TestAccountNumber = True
    End If
Else
    TestAccountNumber = False
End If

End Function
'Check Account related Numeric Field,allow only Account Related Numeric value End

'check entered date between specefic period Start
Public Function TestBfrDate(strDate As String) As Boolean
Dim myDate As Date
Dim sysDate As Date
sysDate = Date
If strDate <> "" Then
    myDate = Format(CDate(strDate))
    If (myDate < sysDate) Then
        TestBfrDate = True
    Else
        TestBfrDate = False
    End If
Else
  TestBfrDate = True
End If

End Function
'check entered date between specefic period End

'Check alphaNumeric Field,Allow only AlphaNumeric Value Start
Public Function TestAlphanumericOnly(lstr_check As String) As Boolean
'allowed characters A to Z, a to z And Blank
Dim i As Integer
Dim ia As Integer
Dim ina As Integer
Dim stlen As Integer

stlen = Len(lstr_check)
ia = 0
ina = 0
For i = 1 To stlen
    Select Case (Mid(lstr_check, i, 1))
        Case "A" To "Z"          'A to Z
            ia = ia + 1
        Case "a" To "z"          'a to z
            ia = ia + 1
        Case "0" To "9"          '0 to 9
        ia = ia + 1
        Case Else
            ina = ina + 1
    End Select
Next i
If ina = 0 Then
    TestAlphanumericOnly = True
Else
    TestAlphanumericOnly = False
End If

End Function
'Check alphaNumeric Field,Allow only AlphaNumeric Value End

'Check AlphaNumeric Field with special character ,allow only AlphaNumeric with special Character value Start
Public Function TestAlphanumeric(lstr_check As String) As Boolean
'allowed characters A to Z, a to z And Blank
Dim i As Integer
Dim ia As Integer
Dim ina As Integer
Dim stlen As Integer

stlen = Len(lstr_check)
ia = 0
ina = 0
For i = 1 To stlen
    Select Case (Mid(lstr_check, i, 1))
            Case "A" To "Z"          'A to Z
                ia = ia + 1
            Case "a" To "z"          'a to z
                ia = ia + 1
            Case " "                 ' Blank
                ia = ia + 1
            Case "0" To "9"          '0 to 9
                ia = ia + 1
            Case ":", "-", ",", "/", ".", "\n", "\", "'", "&", "(", ")", "`", "$"
                ia = ia + 1
            Case "%", "}", "{", "!", "|", "#", "'", ";"
                ia = ia + 1
            Case Else
                ina = ina + 1
        End Select
Next i
If ina = 0 Then
    TestAlphanumeric = True
Else
    TestAlphanumeric = False
End If

End Function
'Check AlphaNumeric Field with special character ,allow only AlphaNumeric with special Character value End

'Check Curreny  Field ,allow only Currency value Start
Public Function curency_neg(lstr_check As String) As Boolean
'allowed characters 0 to 9 and .
Dim i As Integer
Dim ia As Integer
Dim ina As Integer
Dim stlen As Integer
Dim dotCount As Integer
stlen = Len(lstr_check)
ia = 0
ina = 0
For i = 1 To stlen
    Select Case (Mid(lstr_check, i, 1))
        Case "-"
        ia = ia + 1
       Case "."
        ia = ia + 1
        dotCount = dotCount + 1
        Case "0" To "9"          '0 to 9
        ia = ia + 1
        Case Else
            ina = ina + 1
    End Select
Next i
If ina = 0 Then
    If InStr(1, lstr_check, ".") = 0 Then
        curency_neg = True

    Else
        curency_neg = True
    End If
    If dotCount > 1 Then
        curency_neg = False
    End If
Else
    curency_neg = False
End If
End Function
'Check Curreny  Field ,allow only Currency value End

'printErrorStack Function using for printErrorStack Start
Public Function printErrorStack(errline_index As Integer, col_no As Double, field As String, error As String, err_disc As String, start_index As Integer) As Integer
     With Worksheets("Errors")
      .Cells(errline_index, 1) = errline_index - start_index
      .Cells(errline_index, 2) = col_no
      .Cells(errline_index, 3) = field
      .Cells(errline_index, 4) = error
      .Cells(errline_index, 5) = err_disc
      printErrorStack = errline_index + 1
      End With
End Function
'printErrorStack Function using for printErrorStack End

'this function using for find out row Start
Public Function find_row(col) As Double

sel_row = ActiveCell.row
sel_col = ActiveCell.column
Columns("AF:AF").Select
    Selection.Find(What:=col, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate

find_row = ActiveCell.row
Cells(sel_row, sel_col).Select
End Function
'this function using for find out row End

'Add row functionality insert new row Start
Public Sub InsertRowsAndFillFormulas(rangeName As String)
    Application.EnableEvents = False
    Dim vrows As Long
    vrows = 0
    On Error GoTo Errorcatch
    ' ActiveCell.EntireRow.Select  'So you do not have to preselect entire row
    If vrows = 0 Then
        vrows = Application.InputBox(Prompt:= _
            "How many rows do you want to add?", Title:="Add Rows", _
            Default:=1, Type:=1) 'Default for 1 row, type 1 is number
        If vrows = False Then GoTo endSub
    End If
    ' @Author   Kandarp Doshi (194714) and Kalpesh Chaniyara
    ' This Function/Sub routine workes by row selection based on Range Name
    
    Dim startRow As Long, endRow As Long
    'get the start row number from the range name
    startRow = Range(rangeName).row
    ' get the total Number of rows present in the given Range Name
    endRow = Range(rangeName).Rows.Count + startRow - 1
    Range(rangeName).Select
    ActiveSheet.Rows(endRow).Select
    Selection.Resize(rowsize:=2).Rows(2).EntireRow. _
        Resize(rowsize:=vrows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize( _
        rowsize:=vrows + 1), xlFillDefault
     
    Dim endRowNew As Long
    ' get the total Number of rows present in the given Range Name
    endRowNew = Range(rangeName).Rows.Count + startRow - 1
    If endRowNew = endRow Then
        With Range(rangeName)
            .Resize(.Rows.Count + vrows, .Columns.Count).name = rangeName
        End With
    End If
        
        On Error Resume Next    'to handle no constants in range -- John McKee 2000/02/01
        ' to remove the non-formulas -- 1998/03/11 Bill Manville
        Selection.Offset(1).Resize(vrows).EntireRow. _
            SpecialCells(xlConstants).ClearContents
    ActiveSheet.Range("A3:A3").Select
    startRowIndex = startRow
        endRowIndex = endRow + vrows
    GoTo endSub
Errorcatch:
MsgBox Err.Description
endSub:
Application.EnableEvents = True
End Sub
'Add row functionality insert new row End

'Check Alphabet Field ,allow only Alphabet value Start
Public Function TestAlphabet(lstr_check As String) As Boolean
'allowed characters A to Z, a to z And Blank
Dim i As Integer
Dim ia As Integer
Dim ina As Integer
Dim stlen As Integer

stlen = Len(lstr_check)
ia = 0
ina = 0
For i = 1 To stlen
    Select Case (Mid(lstr_check, i, 1))
        Case "A" To "Z"          'A to Z
            ia = ia + 1
        Case "a" To "z"          'a to z
            ia = ia + 1
        Case " "                 ' Blank
            ia = ia + 1
        Case Else
            ina = ina + 1
    End Select
Next i
If ina = 0 Then
    TestAlphabet = True
Else
    TestAlphabet = False
End If

End Function
'Check Alphabet Field ,allow only Alphabet value End

'Check Alphabet with specail character Field ,allow only Alphabet with specail Character value Start
Public Function TestAlphabetNumSpl(lstr_check As String) As Boolean
'allowed characters A to Z, a to z And Blank
Dim i As Integer
Dim ia As Integer
Dim ina As Integer
Dim stlen As Integer

stlen = Len(lstr_check)
ia = 0
ina = 0
For i = 1 To stlen
    Select Case (Mid(lstr_check, i, 1))
        Case "A" To "Z"          'A to Z
            ia = ia + 1
        Case "a" To "z"          'a to z
            ia = ia + 1
        Case "0" To "9"          '0 to 9
            ia = ia + 1
        Case " ", ",", ".", ";", "-", "+", "*", "/", "=", "@", "$", "&", "%"    ' Blank
            ia = ia + 1
        Case Else
            ina = ina + 1
    End Select
Next i
If ina = 0 Then
    TestAlphabetNumSpl = True
Else
    TestAlphabetNumSpl = False
End If

End Function
'Check Alphabet with specail character Field ,allow only Alphabet with specail Character value End

'Check Alphabet with some specail character Field ,allow only Alphabet with some specail Character value Start
Public Function TestName(lstr_check As String) As Boolean
'allowed characters A to Z, a to z And Blank
Dim i As Integer
Dim ia As Integer
Dim ina As Integer
Dim stlen As Integer
Dim ascChr As Integer

stlen = Len(lstr_check)
ia = 0
ina = 0
For i = 1 To stlen
    Select Case (Mid(lstr_check, i, 1))
        Case "A" To "Z"          'A to Z
            ia = ia + 1
        Case "a" To "z"          'a to z
            ia = ia + 1
        Case " "       ' Blank
            ia = ia + 1
        Case "."  'allow dot only when it follows a alphabet or a space
            If (i > 1) Then
                ascChr = asc(Mid(lstr_check, i - 1, 1))
                If (ascChr >= asc("A") And ascChr <= asc("Z")) Or (ascChr >= asc("a") And ascChr <= asc("z")) Or ascChr = asc(" ") Then
                    ia = ia + 1
                Else
                    ina = ina + 1
                End If
           Else
            ina = ina + 1
           End If
           
        Case Else
            ina = ina + 1
    End Select
Next i
If ina = 0 Then
    TestName = True
Else
    TestName = False
End If

End Function
'Check Alphabet with some specail character Field ,allow only Alphabet with some specail Character value End

'ToggleCutCopyAndPaste Function using for disable cut,copy and past Functionality in entire sheet Start
Sub ToggleCutCopyAndPaste(Allow As Boolean)
Dim i As Long, j As Long
Dim sMacro As String, sKey As String
Dim vArr
'Call EnableMenuItem(19, Allow)
vArr = Array("^", "+^")  ' (ctrl, shift-ctrl)
'sMacro = "'" & ThisWorkbook.name & "'!DummyMacro"
Application.CellDragAndDrop = Allow
Dim StartKeyCombination As Variant
On Error Resume Next
    With Application
       ' Select Case Allow
            'Case Is = False
              '  For i = asc("a") To asc("z") ' 97 to 122
              '      If i <> 115 Then ' skip in case of save
                 '       For j = 0 To 1
                 '           sKey = vArr(j) & Chr(i)
                 '           .OnKey sKey, "CutCopyPasteDisabled"
                 '       Next
                '    End If
            '    Next
           ' Case Is = True
              '  For i = asc("a") To asc("z") ' 97 to 122
              '      For j = 0 To 1
                '        sKey = vArr(j) & Chr(i)
                 '       .OnKey sKey
                '    Next
               ' Next
           ' End Select
    End With
    'Application.OnKey "+^{PGUP}", "CutCopyPasteDisabled"
    'Application.OnKey "+^{PGDN}", "CutCopyPasteDisabled"
End Sub
'ToggleCutCopyAndPaste Function using for disable cut,copy and past Functionality in entire sheet End

'EnableMenuItem Function using for Enable Menu Item base on passing control id Start
Sub EnableMenuItem(ctlId As Integer, Enabled As Boolean)
    Dim cBar As CommandBar
    Dim cBarCtrl As CommandBarControl
    For Each cBar In Application.CommandBars
        If cBar.name <> "Clipboard" Then
            Set cBarCtrl = cBar.FindControl(ID:=ctlId, recursive:=True)
            If Not cBarCtrl Is Nothing Then cBarCtrl.Enabled = Enabled
        End If
    Next
End Sub
'EnableMenuItem Function using for Enable Menu Item base on passing control id End

'CutCopyPasteDisabled Function inform to user through alert message Ctrl key disabled in current workbook Start
Sub CutCopyPasteDisabled()
    'MsgBox "Functions with Ctrl keys are disabled in current workbook."
End Sub
'CutCopyPasteDisabled Function inform to user through alert message Ctrl key disabled in current workbook End

'fillUploadSheet function use for fill the upload sheet through data Start
Public Sub fillUploadSheet()
Dim sheetCount As Integer
Dim finalSheet As String
Dim i As Integer
Dim rowCount As Integer
Dim propValuePair As String
Dim PROP_SEP As String
PROP_SEP = "@P_@"
Dim CLASS_SEP As String
CLASS_SEP = "#C_@"
Dim VALUE_SEP As String
VALUE_SEP = "%V_@"
Dim MAIN_PROP_START As String
MAIN_PROP_START = "#"
Dim LIST_PROP_START As String
LIST_PROP_START = "@L_@"
Dim tempString As String
Dim row As Long
Dim column As Long
Dim cellName As String
Dim lastColumn As Long
Dim lastRow As Long
Dim currentWorkSheet As Worksheet
Dim cellRange As Range
 Dim nameCell As name
Dim str1 As String
Dim str2 As String
Dim str3 As String
Dim str4 As String
Dim isSpouseAvail As String
isSpouseAvail = "Yes"
sheetCount = ThisWorkbook.Worksheets.Count
If (sheetCount > 0) Then
    Worksheets(sheetCount).Activate
    ActiveSheet.Unprotect (Pwd)
    finalSheet = Worksheets(sheetCount).name
    If finalSheet = "Amendment" Then
        finalSheet = "Sheet1"
        sheetCount = sheetCount - 1
    End If
    Worksheets(finalSheet).Activate
    ActiveSheet.Unprotect Password:=Pwd
    
    ActiveSheet.Cells(1, 1).Select
    Selection.value = ""
    
    i = 6
    Do While i < sheetCount
        
        lastColumn = Worksheets(i).Cells.Find(What:="*", After:=[A1], _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious).column
         lastRow = Worksheets(i).Cells.Find(What:="*", After:=[A1], _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious).row
        Set currentWorkSheet = Worksheets(i)
        
        For row = 1 To lastRow
            For column = 1 To lastColumn
            cellName = ""
            Set nameCell = Nothing
            Set cellRange = currentWorkSheet.Cells(row, column)
            If (currentWorkSheet.name = "A_Basic_Info") Then
                If (currentWorkSheet.Range("RetInf.DeclareWifeIncome").value = "No") Then
                    isSpouseAvail = "No"
                End If
            End If
            
            On Error Resume Next
                    Set nameCell = cellRange.name
                cellName = nameCell.name
                If Trim(cellName) <> "" Then
                    If Trim(Selection.value) = "" Then
                        Selection.value = Selection.value & cellName & VALUE_SEP & currentWorkSheet.Cells(row, column).value
                    Else
                        Selection.value = Selection.value & PROP_SEP & cellName & VALUE_SEP & currentWorkSheet.Cells(row, column).value
                    End If
                End If
           
            Next
        Next
        i = i + 1
    Loop
    
    Worksheets(finalSheet).Activate
     For i = 1 To counter
         ActiveSheet.Cells(2, i).Select
         Selection.value = ""
     Next
       
End If
counter = 1
tableStr = ""

Call GenerateSheetForList("RentPaid.RentListS")
Call GenerateSheetForList("RentPaid.RentListPartIIS")
Call GenerateSheetForList("ExemptCerti.ListS")
Call GenerateSheetForList("Inv.QuantDtlsListS")
Call GenerateSheetForList("Inv.QuantDtlsListKIIS")
Call GenerateSheetForList("RentalIncome.ListS")
Call GenerateSheetForList("IniAllPlanMach.ListPart1S")
Call GenerateSheetForList("IniAllIBD.ListPart2S")
Call GenerateSheetForList("AgrLandDed.ListS")
Call GenerateSheetForList("DeprIntengAst.ListS")
Call GenerateSheetForList("WAT.ListS")
Call GenerateSheetForList("WAT.ListBS")
Call GenerateSheetForList("EmpIncome.ListS")
Call GenerateSheetForList("CarBenefit.ListS")
Call GenerateSheetForList("MortgageIntDtls.ListS")
Call GenerateSheetForList("HomeOwnershipSavingPlan.ListS")
Call GenerateSheetForList("InsReliefDtls.ListS")
Call GenerateSheetForList("PayeDed.ListS")
Call GenerateSheetForList("VehicleAdvTaxPaid.ListS")
Call GenerateSheetForList("WithHolding.ListS")
Call GenerateSheetForList("InstallmentTax.ListS")
Call GenerateSheetForList("ProfitShare.ListS")
Call GenerateSheetForList("DtlIncomePaid.IncomePaidAdvanceListS")

Call GenerateSheetForList("PLA.OtherExpensesListS")

Call GenerateSheetForList("PLA.BussIncomeDataS", True)
Call GenerateSheetForList("PLA.FarmIncomeDataS", True)
Call GenerateSheetForList("PLA.RentIncomeDataS", True)
Call GenerateSheetForList("PLA.IntIncomeDataS", True)
Call GenerateSheetForList("PLA.CommIncomeDataS", True)
Call GenerateSheetForList("PLA.OthIncomeDataS", True)
Call GenerateSheetForList("PLA.ConsolidateDataS", True)

' Tax Computation
Call GenerateSheetForList("TaxComp.BussinessListS", True)
Call GenerateSheetForList("TaxComp.CnslListS", True)
Call GenerateSheetForList("TaxComp.CommListS", True)
Call GenerateSheetForList("TaxComp.FarmListS", True)
Call GenerateSheetForList("TaxComp.IntListS", True)
Call GenerateSheetForList("TaxComp.OthListS", True)
Call GenerateSheetForList("TaxComp.RentListS", True)
Call GenerateSheetForList("TaxComp.OthDedListS")
Call GenerateSheetForList("TaxComp.OthExpListS")

Call GenerateSheetForList("EstateTrust.ListS")

Call GenerateSheetForList("DtlLossFrwd.BussinessS", True)
Call GenerateSheetForList("DtlLossFrwd.FarmingS", True)
Call GenerateSheetForList("DtlLossFrwd.RentalS", True)
Call GenerateSheetForList("DtlLossFrwd.InterestS", True)
Call GenerateSheetForList("DtlLossFrwd.CommissionS", True)
Call GenerateSheetForList("DtlLossFrwd.OtherS", True)
Call GenerateSheetForList("DtlLossFrwd.TotalS", True)

Call GenerateSheetForList("DtlIncomePaid.IncomePaidSelfAssmntListS")
Call GenerateSheetForList("DTAACredits.DetailsS")

Call GenerateSheetForList("PLA.OtherIncomeListS")

If (isSpouseAvail = "Yes") Then
    Call GenerateSheetForList("RentPaid.RentListW")
    Call GenerateSheetForList("RentPaid.RentListPartIIW") 'added
    Call GenerateSheetForList("ExemptCerti.ListW")
    Call GenerateSheetForList("Inv.QuantDtlsListW")
    Call GenerateSheetForList("Inv.QuantDtlsListKIIW")
    Call GenerateSheetForList("RentalIncome.ListW")
    Call GenerateSheetForList("IniAllPlanMach.ListPart1W")
    Call GenerateSheetForList("IniAllIBD.ListPart2W")
    Call GenerateSheetForList("AgrLandDed.ListW")
    Call GenerateSheetForList("DeprIntengAst.ListW")
    Call GenerateSheetForList("WAT.ListW")
    Call GenerateSheetForList("WAT.ListBW")
    Call GenerateSheetForList("EmpIncome.ListW")
    Call GenerateSheetForList("CarBenefit.ListW")
    Call GenerateSheetForList("MortgageIntDtls.ListW")
    Call GenerateSheetForList("HomeOwnershipSavingPlan.ListW")
    Call GenerateSheetForList("InsReliefDtls.ListW")
    Call GenerateSheetForList("PayeDed.ListW")
    Call GenerateSheetForList("VehicleAdvTaxPaid.ListW")
    Call GenerateSheetForList("WithHolding.ListW")
    Call GenerateSheetForList("InstallmentTax.ListW")
    Call GenerateSheetForList("ProfitShare.ListW")
    Call GenerateSheetForList("DtlIncomePaid.IncomePaidAdvanceListW")
    
    ' For PLA
    Call GenerateSheetForList("PLA.OtherExpensesListW")
    Call GenerateSheetForList("PLA.BussIncomeDataW", True)
    Call GenerateSheetForList("PLA.FarmIncomeDataW", True)
    Call GenerateSheetForList("PLA.RentIncomeDataW", True)
    Call GenerateSheetForList("PLA.IntIncomeDataW", True)
    Call GenerateSheetForList("PLA.CommIncomeDataW", True)
    Call GenerateSheetForList("PLA.OthIncomeDataW", True)
    Call GenerateSheetForList("PLA.ConsolidateDataW", True)
    
    ' Tax Computation
    Call GenerateSheetForList("TaxComp.BussinessListW", True)
    Call GenerateSheetForList("TaxComp.CnslListW", True)
    Call GenerateSheetForList("TaxComp.CommListW", True)
    Call GenerateSheetForList("TaxComp.FarmListW", True)
    Call GenerateSheetForList("TaxComp.IntListW", True)
    Call GenerateSheetForList("TaxComp.OthListW", True)
    Call GenerateSheetForList("TaxComp.RentListW", True)
    Call GenerateSheetForList("TaxComp.OthDedListW")
    Call GenerateSheetForList("TaxComp.OthExpListW")
    
    Call GenerateSheetForList("EstateTrust.ListW")
    
    'new PL
    Call GenerateSheetForList("DtlLossFrwd.BussinessW", True)
    Call GenerateSheetForList("DtlLossFrwd.FarmingW", True)
    Call GenerateSheetForList("DtlLossFrwd.RentalW", True)
    Call GenerateSheetForList("DtlLossFrwd.InterestW", True)
    Call GenerateSheetForList("DtlLossFrwd.CommissionW", True)
    Call GenerateSheetForList("DtlLossFrwd.OtherW", True)
    Call GenerateSheetForList("DtlLossFrwd.TotalW", True)
    
    Call GenerateSheetForList("DtlIncomePaid.IncomePaidSelfAssmntListW")
    Call GenerateSheetForList("DTAACredits.DetailsW")
    Call GenerateSheetForList("PLA.OtherIncomeListW")
    
    
End If
Worksheets(finalSheet).Activate
ActiveSheet.Cells(3, 1).value = "ITR_RET"

' ####### Generating SHA256 Hashcode and setting its value in cell #########
        'Added by Atul
        Dim singleCellAll As String
        Dim multiCellAll As String
        singleCellAll = ""
        multiCellAll = ""
        For i = 1 To 255
            If ActiveSheet.Cells(1, i).value <> "" Then
                'Code Added by Janhavi to generate Hash Code Start
                ActiveSheet.Cells(4, i).value = SHA256(ActiveSheet.Cells(1, i))
                singleCellAll = singleCellAll & ActiveSheet.Cells(1, i).value
                'Code Added by Janhavi to generate Hash Code End
            Else
                Exit For
            End If
        Next

        'Added by Atul
        ActiveSheet.Cells(10, 1).value = SHA256(singleCellAll)
    
        For i = 1 To 255
            If ActiveSheet.Cells(2, i).value <> "" Then
                'Code Added by Janhavi to generate Hash Code Start
                ActiveSheet.Cells(5, i).value = SHA256(ActiveSheet.Cells(2, i))
                multiCellAll = multiCellAll & ActiveSheet.Cells(2, i)
            'Code Added by Janhavi to generate Hash Code End
            Else
                Exit For
            End If
        Next

        'Added by Atul
        ActiveSheet.Cells(11, 1).value = SHA256(multiCellAll)

' ####### Generating SHA256 Hashcode and setting its value in cell #########



Worksheets("Sheet1").Activate
str1 = Worksheets("Sheet1").Range("A1:A1").value
str2 = Worksheets("Sheet1").Range("A2:A2").value
str3 = str1 & str2

End Sub
'fillUploadSheet function use for fill the upload sheet through data End


'GenerateSheetForList function use for fill all list data in sheet Start
Private Sub GenerateSheetForList(ByVal listName As String, Optional generateListColumnWiseFlag As Boolean)

    Dim sheetCount As Integer
    Dim finalSheet As String
    Dim row As Long
    Dim column As Long
    Dim cellName As String
    Dim lastColumn As Long
    Dim lastRow As Long
    Dim currentWorkSheet As Worksheet
    Dim cellRange As Range
    Dim nameCell As name
    Dim startRow As Long
    Dim startColumn As Long
    Dim PROP_SEP As String
    PROP_SEP = "@PL@"
    Dim LIST_SEP As String
    LIST_SEP = "@L_@"
    Dim VALUE_SEP As String
    VALUE_SEP = "%VL@"
    column = 0
    Dim appendListNameFlag As Boolean
    appendListNameFlag = False
    
    Dim listNameAppendFlag As Boolean
    listNameAppendFlag = True
    sheetCount = ThisWorkbook.Worksheets.Count
    finalSheet = Worksheets(sheetCount).name
    If finalSheet = "Amendment" Then
        finalSheet = "Sheet1"
        sheetCount = sheetCount - 1
    End If
    
    If (sheetCount > 0) Then
    
        Worksheets(sheetCount).Activate
        ActiveSheet.Unprotect (Pwd)
        If counter = 0 Then
            counter = 1
        End If
        ActiveSheet.Cells(2, counter).Select
        tableStr = Selection.value
       
        startRow = Range(listName).row
        startColumn = Range(listName).column
        lastColumn = startColumn + Range(listName).Columns.Count - 1
        lastRow = Range(listName).Cells.Find(What:="*", _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious).row
        blnFlag = False
                If generateListColumnWiseFlag = True Then
                    For Each r In Range(listName).Columns
                         For Each c In r.Cells
                                  If c.column > lastColumn Then
                                     Exit For
                                 End If
                                     If c <> "" Then
                                         blnFlag = True
                                         Exit For
                                     End If
                         Next
                         
                         If blnFlag Then
                         blnFlag = False
                             
                             For Each c In r.Cells
                                    
                                    If listNameAppendFlag = True Then  'C.column = startColumn And
                                        If Trim(tableStr) = "" Then
                                            tableStr = listName
                                            
                                        Else
                                            tableStr = tableStr & LIST_SEP & listName
                                            
                                        End If
                                        listNameAppendFlag = False
                                    End If
                                  If c.row > lastRow Then
                                                  Exit For
                                     End If
                                     
                                     If c.Locked = True And c.HasFormula = False And c.value = "" Then
                                         'do nothing
                                         'new Code added for marge cell in Sec B suggestion by riddhi Start
                                         If (listName = "PLA.BussIncomeDataS" Or listName = "PLA.FarmIncomeDataS" Or listName = "PLA.RentIncomeDataS" Or listName = "PLA.IntIncomeDataS" Or listName = "PLA.CommIncomeDataS" Or listName = "PLA.OthIncomeDataS" Or listName = "PLA.BussIncomeDataW" Or listName = "PLA.FarmIncomeDataW" Or listName = "PLA.RentIncomeDataW" Or listName = "PLA.IntIncomeDataW" Or listName = "PLA.CommIncomeDataW" Or listName = "PLA.OthIncomeDataW") Then
                                                If (c.row = 108 Or c.row = 109) Then
                                                    If c.row = startRow Then
                                                        tableStr = tableStr & PROP_SEP
                                                    End If
                                                    If c.row = lastRow Then
                                                        tableStr = tableStr & c
                                                    Else
                                                        tableStr = tableStr & c & VALUE_SEP
                                                    End If
                                                End If
                                            End If
                                         'new Code added for Marge cell in Sec B suggestion by riddhi End
                                     Else
                                         If c.row = startRow Then
                                              tableStr = tableStr & PROP_SEP
                                             
                                         End If
                                             If c.row = lastRow Then
                                                 tableStr = tableStr & c
                                                
                                             Else
                                                 tableStr = tableStr & c & VALUE_SEP
                                                 
                                             End If
                                     End If
                             Next
                             
                             End If
                    Next
                Else
        
                     For Each r In Range(listName).Rows
                         For Each c In r.Cells
                                  If c.row > lastRow Then
                                     Exit For
                                 End If
                                     If c <> "" Then
                                         blnFlag = True
                                         Exit For
                                     End If
                         Next
                         
                         If blnFlag Then
                         blnFlag = False
                             
                             For Each c In r.Cells
                             
                             
                                    If listNameAppendFlag = True Then
                                        If Trim(tableStr) = "" Then
                                            tableStr = listName
                                            
                                        Else
                                            tableStr = tableStr & LIST_SEP & listName
                                            
                                        End If
                                        listNameAppendFlag = False
                                    End If
                                  If c.row > lastRow Then
                                                  Exit For
                                     End If
                                     If c.column = startColumn Then
                                          tableStr = tableStr & PROP_SEP
                                         
                                     End If
                                        If c.column = lastColumn Then
                                            tableStr = tableStr & c
                                            ActiveSheet.Cells(2, counter).Select
                                            
                                            If Left(tableStr, 1) = "@" Then
                                                tableStr = "'" & tableStr
                                            End If
                                            
                                            Selection.value = tableStr
                                            If Len(tableStr) > 28000 And Len(tableStr) < 32000 Then
                                                counter = counter + 1
                                                tableStr = ""
                                                'listNameAppendFlag = True
                                            End If
                                         Else
                                             tableStr = tableStr & c & VALUE_SEP
                                         End If
                             Next
                             
                             End If
                    Next
                End If
     End If
     Selection.value = tableStr
End Sub
'GenerateSheetForList function use for fill all list data in sheet End

'writeRowNumber function using for row Number Start
Public Sub writeRowNumber(rowNumbers As String)
    Worksheets("Errors").Unprotect (Pwd)
    Worksheets("Errors").Rows.ClearContents
    On Error GoTo ErrorHan
    With Worksheets("Errors")
        .Cells(6, 2).value = rowNumbers
    End With
ErrorHan:
    MsgBox Err.Description
End Sub
'writeRowNumber function using for row Number End

'isMandatory function using for check field isMandatory or not Start
Public Function isMandatory(ByRef value As Range, ByRef flagAdd As Range, Optional ByVal colName As String, Optional ByVal isNonZero As String) As String
Dim flag As String
flag = flagAdd.value
Dim column As Range
Dim columnRange As Range
Dim blnFlag As Boolean
blnFlag = True
Dim sName As String
Dim singleRow As Range
Dim rowNumbers As String
sName = value.Cells.Parent.name
Dim sheetName As String

If isNonZero <> "" And isNonZero = "Y" Then
    If UCase(flag) = "Y" And Trim(value) <> "0" Then
            rowNumbers = "NE"
        ElseIf UCase(flag) = "N" Then
            rowNumbers = "NE"
        Else
           rowNumbers = value.Address
        End If
Else
    If Trim(colName) = "" Then
        If UCase(flag) = "Y" And Trim(value) <> "" Then
            rowNumbers = "NE"
        ElseIf UCase(flag) = "N" Then
            rowNumbers = "NE"
        Else
           rowNumbers = value.Address
        End If
    Else
        Set columnRange = Worksheets(sName).Range(colName & "1", colName & Worksheets(sName).UsedRange.Rows.Count)
        Set column = Intersect(value, columnRange)
    
        With Worksheets(sName)
        If Not column Is Nothing Then
            
            For Each c In column.Cells
                On Error Resume Next
                
                If Trim(c) = "" And UCase(flag) = "Y" And c.Locked = False Then
                    blnFlag = False
                    Set singleRow = Intersect(value, Worksheets(sName).Range(.Cells(c.row, 1), .Cells(c.row, Worksheets(sName).UsedRange.Columns.Count)))
                    For Each singleCell In singleRow.Cells
                        If Trim(singleCell) <> "" Then ' And Trim(singleCell) <> 0 Then
                            blnFlag = True
                            Exit For
                        End If
                    Next
                    If blnFlag Then
                        collAddress = colName & singleRow.Cells.row
                         If rowNumbers <> "" Then
                            rowNumbers = rowNumbers & "," & collAddress
                         Else
                            rowNumbers = rowNumbers & collAddress
                         End If
                         
                    End If
                End If
            Next
        End If
        End With
    End If
End If
If rowNumbers = "" Then
    rowNumbers = "NE"
End If

isMandatory = rowNumbers
End Function
'isMandatory function using for check field isMandatory or not End

'isMandatoryOtherExpense function using for check field isMandatory or not Start
Public Function isMandatoryOtherExpense(ByRef value As Range, ByRef flagAdd As Range, Optional ByVal colName As String, Optional ByVal prevColName As String) As String


Dim flag As String
flag = flagAdd.value
Dim column As Range
Dim columnRange As Range
Dim blnFlag As Boolean
blnFlag = True
Dim sName As String
Dim singleRow As Range
Dim rowNumbers As String
sName = value.Cells.Parent.name
Dim sheetName As String

    Set columnRange = Worksheets(sName).Range(colName & "1", colName & Worksheets(sName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)

    With Worksheets(sName)
    If Not column Is Nothing Then
        
        For Each c In column.Cells
            On Error Resume Next
            
            If Trim(c) = "" And UCase(flag) = "Y" And c.Locked = False Then
                If InStr(1, c.Address, "$C") <> 0 Then
                    If c.Previous <> "" Then
                        If c.Next = "" Then
                            blnFlag = True
                        End If
                    End If
                End If
            End If
        Next
    End If
    End With
If rowNumbers = "" Then
    rowNumbers = "NE"
End If

isMandatoryOtherExpense = rowNumbers
End Function
'isMandatoryOtherExpense function using for check field isMandatory or not End

'enableDisableBankFields function to enable/disable bank details fields based on whether the return is credit return Start
Public Function enableDisableBankFields() As String
'    If Worksheets("T_Tax_Computation").Range("FinalTax.TaxRefundDueS").value <> "" And Worksheets("T_Tax_Computation").Range("FinalTax.TaxRefundDueS").value < 0 And Worksheets("A_Basic_Info").Range("BankS").Locked = True Then
'        Call lockUnlock_cell_rng("A_Basic_Info", "BankS", False)
'        Call lockUnlock_cell_rng("A_Basic_Info", "BranchS", False)
'        Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.CityS", False)
'        Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNameS", False)
'        Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNumberS", False)
'    ElseIf Worksheets("T_Tax_Computation").Range("FinalTax.TaxRefundDueS").value <> "" And Worksheets("T_Tax_Computation").Range("FinalTax.TaxRefundDueS").value >= 0 And Worksheets("A_Basic_Info").Range("BankS").Locked = False Then
'        Call lockUnlock_cell_rng("A_Basic_Info", "BankS", True)
'        Call lockUnlock_cell_rng("A_Basic_Info", "BranchS", True)
'        Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.CityS", True)
'        Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNameS", True)
'        Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNumberS", True)
'    End If
End Function
'enableDisableBankFields function to enable/disable bank details fields based on whether the return is credit return End


'isTodaysDate function check current Date Start
Public Function isTodaysDate(ByVal value As String) As Boolean
Dim sheetName As String
value = Format(CDate(Trim(value)), "dd/MM/yyyy")
If Trim(value) <> "" And Trim(value) <> Format(Now(), "dd/MM/yyyy") Then
    isTodaysDate = False
Else
    isTodaysDate = True
End If
End Function
'isTodaysDate function check current Date End

'validateDate function check date in proper format or not Start
Public Function validateDate(ByRef value As Range, Optional ByVal colName As String) As String
Dim sheetName As String
Dim rowNumber As String
If Trim(colName) = "" Then
    If Trim(value) <> "" Then
        If TestDate(value.value) = False Then
         rowNumber = value.Address
        End If
     End If
Else
    
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    sheetName = value.Parent.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)
    
    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If Trim(c) <> "" Then
                If TestDate(c.value) = False Then
                    collAddress = colName & c.row
                    If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
                End If
            End If
        Next
    End If
End With
End If
sheetName = value.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
validateDate = rowNumber
End Function
'validateDate function check date in proper format or not End

Public Function validateNumeric(ByRef value As Range, Optional ByVal colName As String) As String
Dim sheetName As String
Dim rowNumber As String


If Trim(colName) = "" Then
    If TestNumber(value.value) = False Then
                rowNumber = rowNumber & value.Address
    End If
Else
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    sheetName = value.Parent.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)
    
    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If TestNumber(c.value) = False Then
                collAddress = colName & c.row
                If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
            End If
        Next
    End If
End With

End If
sheetName = value.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
validateNumeric = rowNumber
End Function
'validateNumeric function check entered value numeric or not End

'validateNumericAccount function check entered value valid Account numer or not Start
Public Function validateNumericAccount(ByRef value As Range, Optional ByVal colName As String) As String
Dim sheetName As String
Dim rowNumber As String


If Trim(colName) = "" Then
    If TestAccountNumber(value.value) = False Then
        rowNumber = rowNumber & value.Address
    End If
Else
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    sheetName = value.Parent.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)
    
    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If TestAccountNumber(c.value) = False Then
                collAddress = colName & c.row
                If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
            End If
        Next
    End If
End With

End If
sheetName = value.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
validateNumericAccount = rowNumber
End Function
'validateNumericAccount function check entered value valid Account numer or not End

'validatePINDuplication function check entered PIN value Unique or not Start
Public Function validatePINDuplication(ByVal listName As String, ByRef value As Range, Optional ByVal colName As String) As String
Dim sheetName As String
Dim rowNumber As String
    Dim ina As Integer
    ina = 0
    startRow = Range(listName).row
     lastRow = Range(listName).Rows.Count + startRow - 1
     blnFlag = False

     If colName <> "" Then
         For Each r In Range(listName).Rows
             For Each c In r.Cells
                 If c.row > lastRow Then
                     Exit For
                 End If
                 If c.row <> value.row Then
                    If c.Address = "$" & colName & "$" & c.row Then
                    MsgBox value
                    MsgBox Selection.value
                     If value.value = c.value Then
                        collAddress = colName & c.row
                        If rowNumber <> "" Then
                           rowNumber = rowNumber & "," & collAddress
                        Else
                           rowNumber = rowNumber & collAddress
                        End If
                     End If
                    End If
                 End If
             Next
        Next
    End If

    If listName <> "RetInf.PIN" Then
        TPIN = Range("RetInf.PIN").value
        If TPIN = value.value Then
            collAddress = value.Address
            If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
        End If
    End If
    If listName <> "RetInf.SpousePIN" Then
        spousePIN = Range("RetInf.SpousePIN").value
        If spousePIN = value.value Then
            collAddress = value.Address
            If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
        End If
    End If
If rowNumber = "" Then
    rowNumber = "NE"
End If
validatePINDuplication = rowNumber
End Function
'validatePINDuplication function check entered PIN value Unique or not End

'validateAlphabetOnly functuin check enterd value alphabet only or not Start
Public Function validateAlphabetOnly(ByRef value As Range, Optional ByVal colName As String) As String
Dim sheetName As String
If Trim(colName) = "" Then
    If Trim(value) <> "" Then
        If TestAlphabet(value.value) = False Then
            rowNumber = value.Address
        End If
    End If
Else
   
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    sheetName = value.Parent.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)
    
    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If Trim(c) <> "" Then
                If TestAlphabet(c.value) = False Then
                    collAddress = colName & c.row
                    If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
                End If
            End If
        Next
    End If
End With

End If
sheetName = value.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
validateAlphabetOnly = rowNumber
End Function
'validateAlphabetOnly functuin check enterd value alphabet only or not End

'validateAlphaNumeric function check entered value Alphanumeric or not Start
Public Function validateAlphaNumeric(ByRef value As Range, Optional ByVal colName As String) As String
Dim rowNumber As String
Dim sheetName As String

If Trim(colName) = "" Then
    If TestAlphanumeric(value.value) = False Then
       rowNumber = value.Address
     End If
Else
 
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    sheetName = value.Parent.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)
    
    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If TestAlphanumeric(c.value) = False Then
                collAddress = colName & c.row
                If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
            End If
        Next
    End If
End With

End If
sheetName = value.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
validateAlphaNumeric = rowNumber
End Function
'validateAlphaNumeric function check entered value Alphanumeric or not End

'validateAlphaNumericOnly Function Check entered value AlphaNumericOnly or not Start
Public Function validateAlphaNumericOnly(ByRef value As Range, Optional ByVal colName As String) As String
Dim rowNumber As String
Dim sheetName As String

If Trim(colName) = "" Then
    If TestAlphanumericOnly(value.value) = False Then
       rowNumber = value.Address
     End If
Else

    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    sheetName = value.Parent.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)

    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If TestAlphanumericOnly(c.value) = False Then
                collAddress = colName & c.row
                If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
            End If
        Next
    End If
End With

End If
sheetName = value.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
validateAlphaNumericOnly = rowNumber
End Function
'validateAlphaNumericOnly Function Check entered value AlphaNumericOnly or not End


'validateAlphaNumericSpl Function Check entered value Alphanumeric with Specail Character or not Start
Public Function validateAlphaNumericSpl(ByRef value As Range, Optional ByVal colName As String) As Boolean
Dim sheetName As String
validateAlphaNumericSpl = True
If Trim(colName) = "" Then
    If Trim(value) <> "" Then
     validateAlphaNumericSpl = TestAlphabetNumSpl(value.value)
    End If
Else
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    sheetName = value.Parent.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)
    
    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If Trim(c) <> "" Then
                validateAlphaNumericSpl = validateAlphaNumericSpl And TestAlphabetNumSpl(c.value)
            End If
        Next
    End If
End With

End If
sheetName = value.Worksheet.name
End Function
'validateAlphaNumericSpl Function Check entered value Alphanumeric with Specail Character or not End

'validateBfrDate Function check entered date any past date or not Start
Public Function validateBfrDate(ByRef value As Range, Optional ByVal colName As String) As String
Dim sheetName As String
Dim rowNumber As String

If Trim(colName) = "" Then
    If Trim(value) <> "" Then
     
     If Trim(value.value) <> "" And TestDate(value.value) Then
                If TestBfrDate(value.value) = False Then
                    rowNumber = value.Address
                End If
            End If
     
    End If
Else
    
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    sheetName = value.Parent.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)
    
    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If Trim(c) <> "" And TestDate(c.value) Then
                If TestBfrDate(c.value) = False Then
                    collAddress = colName & c.row
                    If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
                End If
            End If
        Next
    End If
End With

End If
sheetName = value.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
validateBfrDate = rowNumber
End Function
'validateBfrDate Function check entered date any past date or not End

'validateTodaysDate Function check entered date value can not be greater than System Date Start
Public Function validateTodaysDate(ByRef value As Range, Optional ByVal colName As String) As Boolean
Dim sheetName As String
validateTodaysDate = True
If Trim(colName) = "" Then
    If Trim(value) <> "" Then
     validateTodaysDate = isTodaysDate(value.value)
    End If
Else
    
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    sheetName = value.Parent.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)
    
    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If Trim(c) <> "" And TestDate(c.value) Then
                validateTodaysDate = validateTodaysDate And isTodaysDate(c.value)
            End If
        Next
    End If
End With

End If
sheetName = value.Worksheet.name
End Function
'validateTodaysDate Function check entered date value can not be greater than System Date End

'createErrorSheet Function call during Validate that time check in sheet if any error found then that error will be display in error Sheet Start
Public Sub createErrorSheet()
        
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull
    Worksheets("Validations").EnableCalculation = True
    Worksheets("ValidationList").EnableCalculation = True
    Dim sectionName As String
    Dim fieldName As String
    Dim row As Long
    Dim column As Long
    Dim srNo As Integer
    Dim rowNumber As String
    Dim sectionDesc As String
    
    'This funciton will enable/disable bank details fields based on whether the return is credit return. For ITR only
    Call enableDisableBankFields
    
    Worksheets("Errors").Unprotect (Pwd)
    Worksheets("Errors").Rows.ClearContents
    
    Dim startIndex As Integer
    startIndex = 2
    
    Dim errInd As Integer
    Dim errDesc As String
    errInd = 1

    With Worksheets("Errors")
        .Cells(errInd, 1) = "Sr. No."
        .Cells(errInd, 2) = "Section Name"
        .Cells(errInd, 3) = "Field"
        .Cells(errInd, 4) = "Error Description"
        .Cells(errInd, 5) = "Reference Cell"
    End With
    errInd = errInd + 1
    
    Worksheets("Validations").Unprotect (Pwd)
    With Worksheets("Validations")
        For row = 2 To 65535
            sectionName = .Cells(row, 1)
            If Trim(sectionName) = "" Then
                Exit For
            End If
            For column = 6 To 65535 Step 2
                If Trim(.Cells(row, column).value) = "" Then
                    Exit For
                End If
                temp = .Cells(row, column).value
                .Cells(row, column).Formula = "=" & temp
                If .Cells(row, column) <> "NE" Then
                    fieldName = .Cells(row, 2)
                    srNo = srNo + 1
                    errDesc = .Cells(row, column + 1)
                    rowNumber = .Cells(row, column)
                    sectionDesc = .Cells(row, 13) 'Description of Section to display in Error Sheet as Section Name (Added by 671255)
                    Call printError(errInd, sectionName, fieldName, errDesc, rowNumber, sectionDesc)
                    errInd = errInd + 1
                End If
                .Cells(row, column).value = temp
            Next
        Next
    End With
    
    Call createErrorSheetForList(errInd, srNo)
    Worksheets("Validations").EnableCalculation = False
    Worksheets("ValidationList").EnableCalculation = False
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
   
End Sub

'printError Function print error in Error Sheet Start
Public Sub printError(errline_index As Integer, secionName As String, fieldName As String, error As String, rowNumber As String, Optional sectionDesc As String)
     With Worksheets("Errors")
      .Cells(errline_index, 1) = errline_index - 1
      .Cells(errline_index, 2) = sectionDesc 'To display section description in error sheet - Added by 671255
      cellAddress = Replace(rowNumber, "$", "")
       .Cells(errline_index, 5) = cellAddress
       If InStr(cellAddress, ",") <> 0 Then
            cellAddress = Mid(cellAddress, 1, InStr(cellAddress, ",") - 1)
        End If
      .Hyperlinks.Add Anchor:=.Cells(errline_index, 3), Address:="", SubAddress:="'" & secionName & "'!" & cellAddress, TextToDisplay:=fieldName, ScreenTip:="Click here to go to the reference cell"
      .EnableSelection = xlNoRestrictions
      .Cells(errline_index, 4) = error
      
      
      
      End With
End Sub
'printError Function print error in Error Sheet End


'createErrorSheetForList Function Create Error Sheet for List Item in Sheet Start
Public Sub createErrorSheetForList(ByVal errInd As Integer, ByVal srNo As Integer)
    
   
    Dim sectionName As String
    Dim fieldName As String
    Dim previousSectionName As String
    Dim row As Long
    Dim column As Long
    Dim maxRows As Long
    Dim startIndex As Integer
    startIndex = 2
    Dim errDesc As String
    Dim innerRows As Long
    Dim rowNumber As String
    Dim sectionDesc As String
    
    Worksheets("ValidationList").Unprotect (Pwd)
    With Worksheets("ValidationList")
    For row = 2 To 65535
        sectionName = .Cells(row, 1)
        sectionDesc = .Cells(row, 18) 'Description of Section to display in Error Sheet as Section Name (Added by 671255)
        If Trim(sectionName) = "" Then
            Exit For
        End If
        
        For column = 2 To 65535
            
            If Trim(.Cells(row, column)) = "" Then
                Exit For
            End If
           fieldName = .Cells(row, column)

           For innerRows = (row + 2) To 65520 Step 2
            If Trim(.Cells(innerRows, column)) = "" Then
                Exit For
            End If
            temp = .Cells(innerRows, column).value
            .Cells(innerRows, column).Formula = "=" & temp
                If .Cells(innerRows, column) <> "NE" Then
                    srNo = srNo + 1
                    errDesc = .Cells(innerRows + 1, column)
                    rowNumber = .Cells(innerRows, column)
                    Call printError(errInd, sectionName, fieldName, errDesc, rowNumber, sectionDesc)
                    errInd = errInd + 1
                End If
                .Cells(innerRows, column).value = temp
            Next
           If maxRows < innerRows Then
                maxRows = innerRows
           End If
        Next
        row = maxRows
  Next
    End With
    Worksheets("Errors").Protect (Pwd)
    If errInd > 2 Then
        MsgBox "Error Found in the sheet"
        Worksheets("Errors").Activate

    Else
    ActiveWorkbook.Unprotect (Pwd)
        Call resetSectionAHiddenDetails
        Call fillUploadSheet
        Msg = "Sheets are ready to be uploaded.Do You want to generate upload file?"
        msg1 = MsgBox(Msg, vbQuestion + vbYesNo, "Generate Upload File")
        If msg1 = vbYes Then
            Generate_upload
        End If
    ActiveWorkbook.Protect (Pwd)
    End If
End Sub
'createErrorSheetForList Function Create Error Sheet for List Item in Sheet End

'validateCurrencyFormat Function check entered value Currency value or not Start
Public Function validateCurrencyFormat(ByRef value As Range, Optional ByVal colName As String) As String

Dim sheetName As String
Dim rowNumber As String


If Trim(colName) = "" Then
    If curency_neg(value.value) = False Then
        rowNumber = value.Address
    End If
Else
    
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    sheetName = value.Parent.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)
    
    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If curency_neg(c.value) = False Then
                collAddress = colName & c.row
                If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
            End If
        Next
    End If
End With

End If
sheetName = value.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
validateCurrencyFormat = rowNumber
End Function
'validateCurrencyFormat Function check entered value Currency value or not End

'validateName Function check entered name value alphabet with some specail character contain or not Start
Public Function validateName(ByRef value As Range, Optional ByVal colName As String) As Boolean
Dim sheetName As String
validateName = True
If Trim(colName) = "" Then
    If Trim(value) <> "" Then
     validateName = TestName(value.value)
    End If
Else
    
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    sheetName = value.Parent.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)
    
    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If Trim(c) <> "" Then
                validateName = validateName And TestName(c.value)
            End If
        Next
    End If
End With

End If
sheetName = value.Worksheet.name

End Function
'validateName Function check entered name value alphabet with some specail character contain or not End

'validatePIN Function check entered PIN in proper formate or not Start
Public Function validatePIN(ByRef value As Range, Optional ByVal colName As String) As String

Dim rowNumber As String
Dim sheetName As String

If Trim(colName) = "" Then
    If TestPIN(value.value) = False Then
       rowNumber = value.Address
    End If
Else
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    sheetName = value.Parent.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)
    
    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If TestPIN(c.value) = False Then
                collAddress = colName & c.row
                If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
            End If
        Next
    End If
End With

End If
sheetName = value.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
validatePIN = rowNumber

End Function
'validatePIN Function check entered PIN in proper formate or not End

'TestPIN Function check entered PIN proper format or not and return Boolean Result Start
Public Function TestPIN(lstr_check As String) As Boolean
'allowed characters A to Z, a to z And Blank
Dim i As Integer
Dim stlen As Integer
Dim alphabates As String
Dim numbers As String
TestPIN = True

numbers = "0123456789"
alphabates = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"

stlen = Len(lstr_check)

For i = 1 To stlen
    
    If i = 1 Then
        If InStr(alphabates, Mid(lstr_check, 1, 1)) = 0 Then
            TestPIN = False
            Exit Function
        ElseIf Mid(lstr_check, 1, 1) <> "P" And Mid(lstr_check, 1, 1) <> "p" And _
                Mid(lstr_check, 1, 1) <> "A" And Mid(lstr_check, 1, 1) <> "a" Then
            TestPIN = False
            Exit Function
        End If
    ElseIf i = 11 Then
        If InStr(alphabates, Mid(lstr_check, i, 1)) = 0 Then
            TestPIN = False
            Exit Function
        End If
    Else
        If InStr(numbers, Mid(lstr_check, i, 1)) = 0 Then
            TestPIN = False
            Exit Function
        End If
    End If
Next i
End Function
'TestPIN Function check entered PIN proper format or not and return Boolean Result End

'Zip_File_Or_Files Function using for Compress file or zip file Start
Sub Zip_File_Or_Files()
    Dim strDate As String, DefPath As String, sFName As String
    Dim oApp As Object, iCtr As Long, i As Integer
    Dim FName, vArr, FileNameZip

    DefPath = Application.DefaultFilePath
    If Right(DefPath, 1) <> "\" Then
        DefPath = DefPath & "\"
    End If

    strDate = Format(Now, " dd-mmm-yy h-mm-ss")
    FileNameZip = DefPath & "MyFilesZip " & strDate & ".zip"
    FName = Application.GetOpenFilename(filefilter:="Excel Files (*.xl*), *.xl*", _
                    MultiSelect:=True, Title:="Select the files you want to zip")
    If IsArray(FName) = False Then
        'do nothing
    Else
        NewZip (FileNameZip)
        Set oApp = CreateObject("Shell.Application")
        i = 0
        For iCtr = LBound(FName) To UBound(FName)
            vArr = Split97(FName(iCtr), "\")
            sFName = vArr(UBound(vArr))
            If bIsBookOpen(sFName) Then
                MsgBox "You can't zip a file that is open!" & vbLf & _
                       "Please close it and try again: " & FName(iCtr)
            Else
                'Copy the file to the compressed folder
                i = i + 1
                oApp.Namespace(FileNameZip).CopyHere FName(iCtr)

                'Keep script waiting until Compressing is done
                On Error Resume Next
                Do Until oApp.Namespace(FileNameZip).items.Count = i
                    Application.Wait (Now + TimeValue("0:00:01"))
                Loop
                On Error GoTo 0
            End If
        Next iCtr

        MsgBox "You find the zip file here: " & FileNameZip
    End If
End Sub
'Zip_File_Or_Files Function using for Compress file or zip file End

'unlock_cell_rng Function Unlock the cell in spcified range base on Sheet name Start
Public Sub lockUnlock_cell_rng(cursheet As String, rangeName As String, lockCellFlag As Boolean, Optional skipColumnFrom As String)
    Dim cellColor
    Dim activeSheetName
    Dim protectedStatus
    
    activeSheetName = ActiveSheet.name
    protectedStatus = ActiveSheet.ProtectContents
        
    Worksheets(cursheet).Activate
    ActiveSheet.Unprotect Password:=Pwd
    
    startRow = Range(rangeName).row
    lastRow = Range(rangeName).Rows.Count + startRow - 1
    
    startColumn = Range(rangeName).column
    lastColumn = startColumn + Range(rangeName).Columns.Count - 1 ' get actual last column
    
    If lockCellFlag = True Then
        cellColor = RGB(146, 146, 146)
    ElseIf lockCellFlag = False Then
        cellColor = RGB(255, 255, 255)
    End If
    If skipColumnFrom = "" Then
        Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).value = Null
        Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).Locked = lockCellFlag
        
        With Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).Interior
               .Color = cellColor
               .Pattern = xlSolid
               .PatternColorIndex = xlAutomatic
        End With
    Else
        If rangeName = "IniAllIBD.ListPart2W" Then 'there are three columns to be skipped for this, hence separate code is written
            Dim rng1 As Range: Set rng1 = ActiveSheet.Range("A" & startRow & ":B" & lastRow)
            Dim rng2 As Range: Set rng2 = ActiveSheet.Range("D" & startRow & ":E" & lastRow)
            Dim rng3 As Range: Set rng3 = ActiveSheet.Range("G" & startRow & ":H" & lastRow)
            
            Dim lockRange As Range: Set lockRange = Union(rng1, rng2, rng3)
            
            
            lockRange.Locked = lockCellFlag
            lockRange.value = Null
            With lockRange.Interior
               .Color = cellColor
               .Pattern = xlSolid
               .PatternColorIndex = xlAutomatic
            End With
        Else
            'skip columns and define new range
            lastColumn = skipColumnFrom - 1
                Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).Locked = lockCellFlag
                Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).value = Null

                With Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).Interior
                   .Color = cellColor
                   .Pattern = xlSolid
                   .PatternColorIndex = xlAutomatic
                End With
        End If
    End If
catch:
    If Err.Description <> "" Then
    End If
    ActiveSheet.Protect Password:=Pwd
    
    Worksheets(activeSheetName).Activate
    If protectedStatus Then
        ActiveSheet.Protect Password:=Pwd
    Else
        ActiveSheet.Unprotect Password:=Pwd
    End If
End Sub
'unlock_cell_rng Function Unlock the cell in spcified range base on Sheet name End

'lockUnlock_cell_rng_without_clearing_contents function local/unlock cell without clearing contents base on passing sheetname,Range Name and lockCellFlag Start
Public Sub lockUnlock_cell_rng_without_clearing_contents(cursheet As String, rangeName As String, lockCellFlag As Boolean, Optional skipColumnFrom As String)

    Dim cellColor
    
    activeSheetName = ActiveSheet.name
    protectedStatus = ActiveSheet.ProtectContents
    
    Worksheets(cursheet).Activate
    ActiveSheet.Unprotect Password:=Pwd
    
    startRow = Range(rangeName).row
    lastRow = Range(rangeName).Rows.Count + startRow - 1
    
    startColumn = Range(rangeName).column
    lastColumn = startColumn + Range(rangeName).Columns.Count - 1 ' get actual last column
    
    If lockCellFlag = True Then
        cellColor = RGB(146, 146, 146)
    ElseIf lockCellFlag = False Then
        cellColor = RGB(255, 255, 255)
    End If
    If skipColumnFrom = "" Then
        Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).Locked = lockCellFlag
        With Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).Interior
           .Color = cellColor
           .Pattern = xlSolid
           .PatternColorIndex = xlAutomatic
        End With
    Else
        'skip columns and define new range
        lastColumn = skipColumnFrom - 1
        Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).Locked = lockCellFlag
        With Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).Interior
           .Color = cellColor
           .Pattern = xlSolid
           .PatternColorIndex = xlAutomatic
        End With
        ActiveSheet.Cells(3, 1).Select
    End If
catch:
    If Err.Description <> "" Then
    End If
    ActiveSheet.Protect Password:=Pwd
    
    Worksheets(activeSheetName).Activate
    
    If protectedStatus Then
        ActiveSheet.Protect Password:=Pwd
    Else
        ActiveSheet.Unprotect Password:=Pwd
    End If
End Sub
'lockUnlock_cell_rng_without_clearing_contents function local/unlock cell without clearing contents base on passing sheetname,Range Name and lockCellFlag End

'lockUnlockOwnedHired Function using for lock/unlock cell base on selected value from Type of Car Cost,selected value Owned then owned respective  cell enable/Disable Start
Public Sub lockUnlockOwnedHired(sheetName As String, rangeName As String, listNameCol As String)
    
    Dim Str As String
    On Error GoTo Errorcatch
  
   
    Dim startRow As Long, endRow As Long
    startRow = Range(rangeName).row
    endRow = Range(rangeName).Rows.Count + startRow - 1
    Worksheets(sheetName).Activate
    Worksheets(sheetName).Unprotect (Pwd)
        
     On Error Resume Next
     Str = listNameCol & startRow & ":" & listNameCol & endRow
     
     For i = startRow To endRow
        If Range(listNameCol & i & ":" & listNameCol & i).value = "Own" Then
            
            Range(ActiveSheet.Range("G" & i & ":G" & i)).Select
            ActiveSheet.Unprotect Password:=Pwd
            Selection.Locked = True
            ActiveSheet.Unprotect Password:=Pwd
            With Selection.Interior
               .Color = RGB(146, 146, 146)
               .Pattern = xlSolid
               .PatternColorIndex = xlAutomatic
            End With
            
            Range(ActiveSheet.Range("H" & i & ":H" & i)).Select
            ActiveSheet.Unprotect Password:=Pwd
            Selection.Locked = False
            ActiveSheet.Unprotect Password:=Pwd
            With Selection.Interior
               .Color = RGB(255, 255, 255)
               .Pattern = xlSolid
               .PatternColorIndex = xlAutomatic
            End With
            ActiveSheet.Protect Password:=Pwd
            
        ElseIf Range(listNameCol & i & ":" & listNameCol & i).value = "Hired" Then
            Range(ActiveSheet.Range("G" & i & ":G" & i)).Select
            ActiveSheet.Unprotect Password:=Pwd
            Selection.Locked = False
            ActiveSheet.Unprotect Password:=Pwd
            With Selection.Interior
               .Color = RGB(255, 255, 255)
               .Pattern = xlSolid
               .PatternColorIndex = xlAutomatic
            End With
            
            Range(ActiveSheet.Range("H" & i & ":H" & i)).Select
            ActiveSheet.Unprotect Password:=Pwd
            Selection.Locked = True
            ActiveSheet.Unprotect Password:=Pwd
            With Selection.Interior
               .Color = RGB(146, 146, 146)
               .Pattern = xlSolid
               .PatternColorIndex = xlAutomatic
            End With
            ActiveSheet.Protect Password:=Pwd
        End If
     Next
    Exit Sub
    
Errorcatch:
MsgBox Err.Description
End Sub
'lockUnlockOwnedHired Function using for lock/unlock cell base on selected value from Type of Car Cost,selected value Owned then owned respective  cell enable/Disable End

 'lock_cells function using for lock cell base on passing Sheet Name,List Name,Col name and lockValue Start
 Public Sub lock_cells(cursheet As String, listName As String, colName As String, lockValue As Boolean, Optional clearFlag As Boolean)
 
    activeSheetName = ActiveSheet.name
    protectedStatus = ActiveSheet.ProtectContents
    
    Worksheets(cursheet).Activate
    ActiveSheet.Unprotect Password:=Pwd
         startRow = Range(listName).row
         lastRow = Range(listName).Rows.Count + startRow - 1
         blnFlag = False
         For Each r In Range(listName).Rows
             For Each c In r.Cells
                 If c.row > lastRow Then
                     Exit For
                 End If
                 If clearFlag = False And c.Address = "$" & colName & "$" & c.row Then
                     Call lockUnlock_cell_rng_without_clearing_contents(cursheet, c.Address, lockValue)
                 ElseIf c.Address = "$" & colName & "$" & c.row Then
                     Call lockUnlock_cell_rng(cursheet, c.Address, lockValue)
                 End If
             Next
        Next
    ActiveSheet.Protect Password:=Pwd
    Worksheets(activeSheetName).Activate
    If protectedStatus Then
        ActiveSheet.Protect Password:=Pwd
    Else
        ActiveSheet.Unprotect Password:=Pwd
    End If
 End Sub
 'lock_cells function using for lock cell base on passing Sheet Name,List Name,Col name and lockValue End
 
 'lock_branchNames function using for lock Branch name base on passing SheetName,List Name,col name and Lock value Start
 Public Sub lock_branchNames(cursheet As String, listName As String, colName As String, lockValue As Boolean)
 Worksheets(cursheet).Activate
    ActiveSheet.Unprotect Password:=Pwd
         startRow = Range(listName).row
         lastRow = Range(listName).Rows.Count + startRow - 1
         blnFlag = False
         For Each r In Range(listName).Rows
             For Each c In r.Cells
                 If c.row > lastRow Then
                     Exit For
                 End If
                 If c.Address = "$" & colName & "$" & c.row Then
                     ActiveSheet.Unprotect Password:=Pwd
                    ActiveSheet.Unprotect Password:=Pwd
                        c.Locked = lockValue
                        c.value = ""
                        ActiveSheet.Unprotect Password:=Pwd
                        With Selection.Interior
                           .Color = RGB(146, 146, 146)  'lock
                           .Pattern = xlSolid
                           .PatternColorIndex = xlAutomatic
                        End With
                        ActiveSheet.Protect Password:=Pwd
                 End If
             Next
        Next
    ActiveSheet.Protect Password:=Pwd
 End Sub
 'lock_branchNames function using for lock Branch name base on passing SheetName,List Name,col name and Lock value End
 
'find_BankID Function find BankName from data sheet passing base on specific column Start
Public Function find_BankID(col) As Double
Dim row As Integer
Dim rngFound As Range
        
Set rngFound = Sheet18.Range("AS2:AS47").Cells.Find(What:=col, After:=Sheet18.Range("AS2:AS47").Cells.Cells(1, 1), LookIn:=xlValues, _
       LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
       MatchCase:=False)

If rngFound Is Nothing Then
    row = -1
Else
    row = rngFound.row
End If

find_BankID = row

End Function
'find_BankID Function find BankName from data sheet passing base on specific column End

'find_BranchFromBankID Function find Branch from data sheet passing base on BankId Start
Public Sub find_BranchFromBankID(bankId)
    Dim row As Integer
    Dim rngFound As Range
    Dim endCount As Integer
    endCount = 1
        
    Sheet18.Unprotect (Pwd)
    Sheet18.Range("selectedBranchNames").Delete
    
    For Each r In Sheet18.Range("AW2:AW565").Rows
        For Each c In r.Cells
            If Trim(c) = Trim(bankId) Then
                endCount = endCount + 1
                Sheet18.Range("AY" & endCount & ":AY" & endCount).value = Sheet18.Range("AV" & c.row & ":AV" & c.row).value
            End If
        Next
    Next
    
    If endCount = 1 Then
        endCount = 2
    End If
    
    Sheet18.Range("AY2" & ":AY" & endCount).name = "selectedBranchNames"
    Sheet18.Protect (Pwd)
End Sub
'find_BranchFromBankID Function find Branch from data sheet passing base on BankId End

'find_BranchID Function find BranchName from data sheet passing base on specific column Start
Public Function find_BranchID(col) As Double
Dim row As Integer
Dim rngFound As Range
        
Set rngFound = Sheet18.Range("BranchName").Cells.Find(What:=col, _
                After:=Sheet18.Range("BranchName").Cells.Cells(1, 1), LookIn:=xlValues, _
                LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                MatchCase:=False)


If rngFound Is Nothing Then
    row = -1
Else
    row = rngFound.row
End If

find_BranchID = row

End Function
'find_BranchID Function find BranchName from data sheet passing base on specific column End

'find_CountyID function find CountyId from Data sheet and set in hidden field in the sheet Start
Public Function find_CountyID(col) As Double
Dim row As Integer
Dim rngFound As Range

Set rngFound = Sheet18.Range("CountyName").Cells.Find(What:=col, After:=Sheet18.Range("CountyName").Cells.Cells(1, 1), LookIn:=xlValues, _
       LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
       MatchCase:=False)
       
If rngFound Is Nothing Then
    row = -1
Else
    row = rngFound.row
End If
find_CountyID = row

End Function
'find_CountyID function find CountyId from Data sheet and set in hidden field in the sheet End

'find_DistrictID function find DistrictId from Data sheet and set in hidden field in the sheet Start
Public Function find_DistrictID(col) As Double
    Dim row As Integer
    Dim rngFound As Range
    
    Set rngFound = Sheet18.Range("districtName").Cells.Find(What:=col, After:=Sheet18.Range("districtName").Cells.Cells(1, 1), LookIn:=xlValues, _
           LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
           MatchCase:=False)
    
    If rngFound Is Nothing Then
        row = -1
    Else
        row = rngFound.row
    End If
    find_DistrictID = row

End Function
'find_DistrictID function find DistrictId from Data sheet and set in hidden field in the sheet End

'find_LocalID function find LocationId from Data sheet and set in hidden field in the sheet Start
Public Function find_LocalID(col) As Double
Dim row As Integer
Dim rngFound As Range

Set rngFound = Sheet18.Range("localityName").Cells.Find(What:=col, After:=Sheet18.Range("localityName").Cells.Cells(1, 1), LookIn:=xlValues, _
       LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
       MatchCase:=False)

If rngFound Is Nothing Then
    row = -1
Else
    row = rngFound.row
End If
find_LocalID = row

End Function
'find_LocalID function find LocationId from Data sheet and set in hidden field in the sheet End

'find_PostCode Function find postCode from Data sheet base on selected towan and set in hidden field in sheet Start
Public Function find_PostCode(col) As Double
Dim row As Integer
Dim rngFound As Range
       
Set rngFound = Sheet18.Range("PostalCode").Cells.Find(What:=col, After:=Sheet18.Range("PostalCode").Cells.Cells(1, 1), LookIn:=xlValues, _
       LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
       MatchCase:=False)

If rngFound Is Nothing Then
    row = -1
Else
    row = rngFound.row
End If
find_PostCode = row

End Function
'find_PostCode Function find postCode from Data sheet base on selected towan and set in hidden field in sheet End

'find_Town Function using for find town from Data Sheet Start
Public Function find_Town(col) As Double
Dim row As Integer
Dim rngFound As Range
      
Set rngFound = Sheet18.Range("Town").Cells.Find(What:=col, _
                                After:=Sheet18.Range("Town").Cells.Cells(1, 1), LookIn:=xlValues, _
                                LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                                MatchCase:=False _
                            )

If rngFound Is Nothing Then
    row = -1
Else
    row = rngFound.row
End If
find_Town = row

End Function
'find_Town Function using for find town from Data Sheet End

Public Function find_PostalCode(col) As Double

    Dim row As Integer
    Dim rngFound As Range
            
    Set rngFound = Worksheets("Data").Range("PostalCode").Cells.Find(What:=col, After:= _
                               Worksheets("Data").Range("PostalCode").Cells.Cells(1, 1), LookIn:=xlValues, _
                               LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                               MatchCase:=False)
    
    If rngFound Is Nothing Then
        row = -1
    Else
        row = rngFound.row
    End If
    find_PostalCode = row
End Function


'find_TaxRATE Function calculate tax Rate base on passing reference col Start
Public Function find_TaxRATE(col) As Double
    Dim row As Integer
    Dim rngFound As Range
            
    Set rngFound = Sheet18.Range("NatureOfPayment").Cells.Find(What:=col, After:= _
                    Sheet18.Range("NatureOfPayment").Cells.Cells(1, 1), LookIn:=xlValues, _
                    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                    MatchCase:=False)
    
    If rngFound Is Nothing Then
        row = -1
    Else
        row = rngFound.row
    End If
    
    find_TaxRATE = row
End Function
'find_TaxRATE Function calculate tax Rate base on passing reference col End

'getReturnEndDate Function find Return period End Date Start
Public Function getReturnEndDate(strDate As String) As String
    Dim strYear As String
    strYear = Mid(strDate, 7, 4)
    getReturnEndDate = "31/12/" + strYear
End Function
'getReturnEndDate Function find Return period End Date End

'checkPINDuplication function check entered PIN number can not be Duplicate Start
Public Function checkPINDuplication(ByVal listNameParam As String, ByRef value As Range, Optional ByVal colName As String) As String
Dim rowNumber As String
Dim sheetName As String
TPIN = Range("RetInf.PIN").value
spousePIN = Range("RetInf.SpousePIN").value

    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name

    Set sName = value.name
    sheetName = Mid(sName, 2, InStrRev(sName, "!") - 2)
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    If colName <> "" Then
        Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
        Set column = Intersect(value, columnRange)
        With Worksheets(sheetName)
            If Not column Is Nothing Then
                For Each c In column.Cells
                    If Trim(c) <> "" Then
                        rangeAddress = Range(listNameParam).Address
                        startRow = Range(listNameParam).row
                        endRow = Mid(rangeAddress, InStrRev(rangeAddress, "$") + 1, Len(rangeAddress))
                            For i = startRow To endRow
                                If i <> c.row Then
                                    If c.value <> "" And Sheet14.Range("A" & (i) & ":A" & i).value <> "" Then
                                        If c.value = Sheet14.Range("A" & (i) & ":A" & i).value Then
                                            collAddress = colName & c.row
                                                If rowNumber <> "" Then
                                                    rowNumber = rowNumber & "," & collAddress
                                                Else
                                                    rowNumber = rowNumber & collAddress
                                                End If
                                        End If

                                    End If
                                End If
                            Next
                    End If
                    If rowNumber = "" Then
                        If TPIN <> "" And Trim(c) <> "" Then
                            If TPIN = Trim(c) Then
                                blnFlag = True
                            End If
                        End If
                        If spousePIN <> "" And Trim(c) <> "" Then
                            If spousePIN = Trim(c) Then
                                blnFlag = True
                            End If
                        End If
                             If blnFlag Then
                                collAddress = colName & c.row
                                    If rowNumber <> "" Then
                                       rowNumber = rowNumber & "," & collAddress
                                    Else
                                       rowNumber = rowNumber & collAddress
                                    End If
                            
                             End If
                     End If
                Next
            End If
        End With
    Else
        If Trim(value.value) <> "" Then
            selRange = Mid(value.name, InStrRev(value.name, "!") + 1, Len(value.name))

            If (Trim(selRange) = "$B$3") And (value.value) = spousePIN Then
                    rowNumber = rowNumber & value.Address
            Else
                If (Trim(selRange) = "$B$9") And (value.value) = TPIN Then
                        rowNumber = rowNumber & value.Address
                End If
            End If
        End If
   End If

    sheetName = value.Worksheet.name
    If rowNumber = "" Then
        rowNumber = "NE"
    End If
checkPINDuplication = rowNumber
End Function
'checkPINDuplication function check entered PIN number can not be Duplicate End

'getTotalTaxPayble function calculate total Tax Payable base on total taxable Income and different tax slab start
Function getTotalTaxPayble(ByVal TotalTaxableIncome As Double, ByVal RtnYear As Integer) As Double
Application.Volatile
Dim taxPayble As Double
Dim taxPayble1 As Double
Dim taxPayble2 As Double
Dim taxPayble3 As Double
Dim taxPayble4 As Double
Dim taxPayble5 As Double
Dim TotalTaxableIncomeBiz As Double
taxPayble1 = 0
taxPayble2 = 0
taxPayble3 = 0
taxPayble4 = 0
taxPayble5 = 0
taxPayble = 0
If TotalTaxableIncome < 0 Then
    taxPayble = 0
Else
    If (RtnYear <= 2016) Then
        If TotalTaxableIncome > 466704 Then
            taxPayble1 = (TotalTaxableIncome - 466704) * 30 / 100
            taxPayble2 = (466704 - 351793) + 1
            taxPayble2 = taxPayble2 * 25 / 100
            taxPayble3 = (351792 - 236881) + 1
            taxPayble3 = taxPayble3 * 20 / 100
            taxPayble4 = (236880 - 121969) + 1
            taxPayble4 = taxPayble4 * 15 / 100
            taxPayble5 = 121968 * 10 / 100
        End If
        If TotalTaxableIncome > 351792 And TotalTaxableIncome < 466705 Then
            taxPayble1 = (TotalTaxableIncome - 351792) * 25 / 100
            taxPayble2 = (351792 - 236881) + 1
            taxPayble2 = taxPayble2 * 20 / 100
            taxPayble3 = (236880 - 121969) + 1
            taxPayble3 = taxPayble3 * 15 / 100
            taxPayble4 = 121968 * 10 / 100
        End If
        If TotalTaxableIncome > 236880 And TotalTaxableIncome < 351793 Then
            taxPayble1 = (TotalTaxableIncome - 236881 + 1) * 20 / 100
            taxPayble2 = (236880 - 121969) + 1
            taxPayble2 = taxPayble2 * 15 / 100
            taxPayble3 = 121968 * 10 / 100
        End If
        If TotalTaxableIncome > 121968 And TotalTaxableIncome < 236881 Then
            taxPayble1 = (TotalTaxableIncome - 121969 + 1) * 15 / 100
            taxPayble2 = 121968 * 10 / 100
        End If
        If TotalTaxableIncome > 0 And TotalTaxableIncome < 121969 Then
            taxPayble1 = TotalTaxableIncome * 10 / 100
        End If
        taxPayble = taxPayble1 + taxPayble2 + taxPayble3 + taxPayble4 + taxPayble5
    End If
    If (RtnYear = 2017) Then
            If TotalTaxableIncome > 513373 Then
                taxPayble1 = (TotalTaxableIncome - 513373) * 30 / 100
                taxPayble2 = ((513373 - 386971) + 1)
                taxPayble2 = taxPayble2 * 25 / 100
                taxPayble3 = ((386970 - 260568) + 1)
                taxPayble3 = taxPayble3 * 20 / 100
                taxPayble4 = ((260567 - 134165) + 1)
                taxPayble4 = taxPayble4 * 15 / 100
                taxPayble5 = 1341640 / 100
            End If
            If TotalTaxableIncome >= 386971 And TotalTaxableIncome <= 513373 Then
                taxPayble1 = (TotalTaxableIncome - 386970) * 25 / 100
                taxPayble2 = ((386970 - 260568) + 1)
                taxPayble2 = taxPayble2 * 20 / 100
                taxPayble3 = ((260567 - 134165) + 1)
                taxPayble3 = taxPayble3 * 15 / 100
                taxPayble4 = 1341640 / 100
            End If
            If TotalTaxableIncome >= 260568 And TotalTaxableIncome <= 386970 Then
                taxPayble1 = (TotalTaxableIncome - 260567) * 20 / 100
                taxPayble2 = ((260567 - 134165) + 1)
                taxPayble2 = taxPayble2 * 15 / 100
                taxPayble3 = 1341640 / 100
            End If
            If TotalTaxableIncome >= 134165 And TotalTaxableIncome <= 260567 Then
                taxPayble1 = (TotalTaxableIncome - 134164) * 15 / 100
                taxPayble2 = 1341640 / 100
            End If
            If TotalTaxableIncome >= 0 And TotalTaxableIncome <= 134164 Then
                taxPayble1 = TotalTaxableIncome / 10
            End If
            taxPayble = taxPayble1 + taxPayble2 + taxPayble3 + taxPayble4 + taxPayble5
    End If
    'Added by Lawrence and Ruth on 24/12/2020
    'Excluded 2020 from the IF
    'If (RtnYear >= 2018 And RtnYear <> 2020) Then
     If (RtnYear >= 2018 And RtnYear <= 2019) Then
            If TotalTaxableIncome > 564709 Then
                taxPayble1 = (TotalTaxableIncome - 564709) * 30 / 100
                taxPayble2 = ((564709 - 425667) + 1)
                taxPayble2 = taxPayble2 * 25 / 100
                taxPayble3 = ((425666 - 286624) + 1)
                taxPayble3 = taxPayble3 * 20 / 100
                taxPayble4 = ((286623 - 147581) + 1)
                taxPayble4 = taxPayble4 * 15 / 100
                taxPayble5 = 1475800 / 100
            End If
            If TotalTaxableIncome >= 425667 And TotalTaxableIncome <= 564709 Then
                taxPayble1 = (TotalTaxableIncome - 425666) * 25 / 100
                taxPayble2 = ((425666 - 286624) + 1)
                taxPayble2 = taxPayble2 * 20 / 100
                taxPayble3 = ((286623 - 147581) + 1)
                taxPayble3 = taxPayble3 * 15 / 100
                taxPayble4 = 1475800 / 100
            End If
            If TotalTaxableIncome >= 286624 And TotalTaxableIncome <= 425666 Then
                taxPayble1 = (TotalTaxableIncome - 286623) * 20 / 100
                taxPayble2 = ((286623 - 147581) + 1)
                taxPayble2 = taxPayble2 * 15 / 100
                taxPayble3 = 1475800 / 100
            End If
            If TotalTaxableIncome >= 147581 And TotalTaxableIncome <= 286623 Then
                taxPayble1 = (TotalTaxableIncome - 147580) * 15 / 100
                taxPayble2 = 1475800 / 100
            End If
            If TotalTaxableIncome >= 0 And TotalTaxableIncome <= 147580 Then
                taxPayble1 = TotalTaxableIncome / 10
            End If
            taxPayble = taxPayble1 + taxPayble2 + taxPayble3 + taxPayble4 + taxPayble5
    End If
    
    'Added by Ruth and Lawrence
    'For 2020
    'Separate Taxable Employment income from TotalTaxable income
    'Separate the Taxable Employment income into Jan -Mar and then Apr -Dec for both husaband and wife
     If (RtnYear = 2020 Or RtnYear = 2023) Then
            TotalTaxableIncomeBiz = 0
         If (TotalTaxableIncome = Range("TaxComp.NetTaxableIncomeW").value) Then
            TotalIncomeJanMar = Range("EmpIncome.ListWTOJANMAR").value
            TotalDeductionWJanMar = Range("DedDtl.TotalDeductionWJanMar").value
            totalEmpIncomeWJanDec = Range("EmpIncome.ListWTO").value
            ExemptedAmtWJanMar = Range("TaxComp.ExemptedAmtWJanMar").value
            TotalDeductionWJanDec = Range("DedDtl.TotalDeductionW").value
            ExemptedAmtWJanDec = Range("TaxComp.ExemptedAmtW").value
            
    
            
            
            TotalTaxableIncomeJanMar = TotalIncomeJanMar - (TotalDeductionWJanMar + ExemptedAmtWJanMar)
            'If biz employment income for wife is nill factor deductions in the biz income
            TotalTaxableIncomeBiz = Range("TaxComp.ChargeableIncomeW").value
            'If (TotalTaxableIncome > totalEmpIncomeWJanDec) Then
                If (TotalTaxableIncomeBiz > 0) Then
                        If (totalEmpIncomeWJanDec = 0) Then
                            TotalTaxableIncomeBiz = Range("TaxComp.ChargeableIncomeW").value - (TotalDeductionWJanDec + ExemptedAmtWJanDec)
                        ElseIf (totalEmpIncomeWJanDec > 0) Then
                             TotalTaxableIncomeBiz = Range("TaxComp.ChargeableIncomeW").value
                        End If
                End If
            
                If (TotalTaxableIncomeJanMar < 0) Then
                    TotalTaxableIncomeAprDec = 0
                Else
                    TotalTaxableIncomeAprDec = TotalTaxableIncome - TotalTaxableIncomeJanMar
                End If
         ElseIf (TotalTaxableIncome = Range("TaxComp.NetTaxableIncomeS").value) Then
            totalEmpIncomeSJanDec = Range("EmpIncome.ListSTO").value
            TotalIncomeJanMar = Range("EmpIncome.ListSTOJANMAR").value
            TotalDeductionSJanMar = Range("DedDtl.TotalDeductionSJanMar").value
            ExemptedAmtSJanMar = Range("TaxComp.ExemptedAmtSJanMar").value
            
            ExemptedAmtSJanMar = Range("TaxComp.ExemptedAmtSJanMar").value
            TotalDeductionSJanDec = Range("DedDtl.TotalDeductionS").value
            ExemptedAmtSJanDec = Range("TaxComp.ExemptedAmtS").value
            
            TotalTaxableIncomeJanMar = TotalIncomeJanMar - (TotalDeductionSJanMar + ExemptedAmtSJanMar)
            TotalTaxableIncomeBiz = Range("TaxComp.ChargeableIncomeS").value
            'If biz employment income for husband is nill factor deductions in the biz income
            'If (TotalTaxableIncome > totalEmpIncomeSJanDec) Then
                If (TotalTaxableIncomeBiz > 0) Then
                        If (totalEmpIncomeSJanDec = 0) Then
                            TotalTaxableIncomeBiz = Range("TaxComp.ChargeableIncomeS").value - (TotalDeductionSJanDec + ExemptedAmtSJanDec)
                        ElseIf (totalEmpIncomeSJanDec > 0) Then
                            TotalTaxableIncomeBiz = Range("TaxComp.ChargeableIncomeS").value
                        End If
                End If
            
            
                If (TotalTaxableIncomeJanMar < 0) Then
                    TotalTaxableIncomeAprDec = 0
                Else
                    TotalTaxableIncomeAprDec = TotalTaxableIncome - TotalTaxableIncomeJanMar
                End If
         End If


            If (TotalTaxableIncomeBiz > 0) Then
            TotalTaxableIncomeAprDec = TotalTaxableIncome - TotalTaxableIncomeJanMar - TotalTaxableIncomeBiz
             taxPayble3 = getTotalTaxPayable2020Biz(TotalTaxableIncome, RtnYear, TotalTaxableIncomeBiz)
             taxPayble1 = getTotalTaxPayble2020JanMar(TotalTaxableIncome, RtnYear, TotalTaxableIncomeJanMar)
             taxPayble2 = getTotalTaxPayble2020AprDec(TotalTaxableIncome, RtnYear, TotalTaxableIncomeAprDec)
             taxPayble = taxPayble1 + taxPayble2 + taxPayble3
            Else
             taxPayble1 = getTotalTaxPayble2020JanMar(TotalTaxableIncome, RtnYear, TotalTaxableIncomeJanMar)
             taxPayble2 = getTotalTaxPayble2020AprDec(TotalTaxableIncome, RtnYear, TotalTaxableIncomeAprDec)
             taxPayble = taxPayble1 + taxPayble2
            End If
    End If
    'Added on 13/01/2022 for 2021 tax rate changes
    If (RtnYear >= 2021 And RtnYear <= 2022) Then
            If TotalTaxableIncome > 388000 Then
                taxPayble1 = (TotalTaxableIncome - 388000) * (30 / 100)
                taxPayble2 = (388000 - 288000)
                taxPayble2 = taxPayble2 * (25 / 100)
                taxPayble3 = 288000 * (10 / 100)
            End If
            If TotalTaxableIncome > 288000 And TotalTaxableIncome <= 388000 Then
                taxPayble1 = (TotalTaxableIncome - 288000) * (25 / 100)
                taxPayble2 = 288000 * (10 / 100)
            End If
            
            If TotalTaxableIncome >= 0 And TotalTaxableIncome <= 288000 Then
                taxPayble1 = TotalTaxableIncome / 10
            End If
            taxPayble = taxPayble1 + taxPayble2 + taxPayble3
    End If
    'End on 13/01/2022
End If
getTotalTaxPayble = taxPayble
End Function

'Added by Ruth and Lawrence
'Calculate Employment Income tax for the period Jan to Mar 2020
Function getTotalTaxPayble2020JanMar(ByVal TotalTaxableIncome As Double, _
ByVal RtnYear As Integer, ByVal TotalTaxableIncomeJanMar As Double) As Double
Application.Volatile
Dim taxPayble As Double
Dim taxPaybleJanMar As Double
Dim taxPayble1 As Double
Dim taxPayble2 As Double
Dim taxPayble3 As Double
Dim taxPayble4 As Double
Dim taxPayble5 As Double



taxPayble1 = 0
taxPayble2 = 0
taxPayble3 = 0
taxPayble4 = 0
taxPayble5 = 0
taxPayble = 0
taxBandJanMar = 0
taxpayebleJanMar = 0
If TotalTaxableIncomeJanMar < 0 Then
    taxPaybleJanMar = 0
Else
    'Added by Lawrence and Ruth on 28/12/2020
    'Excluded 2020 from the IF
    If (RtnYear = 2020) Then
    'TotalTaxableIncomeJanMar = TotalIncomeJanMar - TotalDeductionSJanMar - ExemptedAmtSJanMar
            If TotalTaxableIncomeJanMar > (564709 / 4) Then
                taxPayble1 = (TotalTaxableIncomeJanMar - (564709 / 4)) * 0.3
                taxPayble2 = ((564709 - 425667) / 4) + 1
                taxPayble2 = taxPayble2 * 0.25
                taxPayble3 = ((425666 - 286624) / 4) + 1
                taxPayble3 = taxPayble3 * 0.2
                taxPayble4 = ((286623 - 147581) / 4) + 1
                taxPayble4 = taxPayble4 * 0.15
                taxPayble5 = (1475800 / 4) / 100
                taxBandJanMar = 5
            End If
            If TotalTaxableIncomeJanMar >= (425667 / 4) And TotalTaxableIncomeJanMar <= (564709 / 4) Then
                taxPayble1 = (TotalTaxableIncomeJanMar - (425666 / 4)) * 0.25
                taxPayble2 = ((425666 - 286624) / 4) + 1
                taxPayble2 = taxPayble2 * 0.2
                taxPayble3 = ((286623 - 147581) / 4) + 1
                taxPayble3 = taxPayble3 * 0.15
                taxPayble4 = (1475800 / 4) / 100
                taxBandJanMar = 4
            End If
            If TotalTaxableIncomeJanMar >= (286624 / 4) And TotalTaxableIncomeJanMar <= (425666 / 4) Then
                taxPayble1 = (TotalTaxableIncomeJanMar - (286623 / 4)) * 20 / 100
                taxPayble2 = ((286623 - 147581) / 4) + 1
                taxPayble2 = taxPayble2 * 0.15
                taxPayble3 = (1475800 / 4) / 100
                taxBandJanMar = 3
            End If
            If TotalTaxableIncomeJanMar >= (147581 / 4) And TotalTaxableIncomeJanMar <= (286623 / 4) Then
                taxPayble1 = (TotalTaxableIncomeJanMar - (147580 / 4)) * 0.15
                taxPayble2 = (1475800 / 4) / 100
                taxBandJanMar = 2
            End If
            If TotalTaxableIncomeJanMar >= 0 And TotalTaxableIncomeJanMar <= (147580 / 4) Then
                taxPayble1 = TotalTaxableIncomeJanMar / 10
                taxBandJanMar = 1
            End If
            taxPaybleJanMar = taxPayble1 + taxPayble2 + taxPayble3 + taxPayble4 + taxPayble5
         End If
         
        If (RtnYear = 2023) Then
        'TotalTaxableIncomeJanMar = TotalIncomeJanMar - TotalDeductionSJanMar - ExemptedAmtSJanMar
            If TotalTaxableIncomeJanMar > 388000 Then
                taxPayble1 = (TotalTaxableIncomeJanMar - 388000) * (30 / 100)
                taxPayble2 = (388000 - 288000)
                taxPayble2 = taxPayble2 * (25 / 100)
                taxPayble3 = 288000 * (10 / 100)
            End If
            If TotalTaxableIncomeJanMar > 288000 And TotalTaxableIncomeJanMar <= 388000 Then
                taxPayble1 = (TotalTaxableIncomeJanMar - 288000) * (25 / 100)
                taxPayble2 = 288000 * (10 / 100)
            End If
            
            If TotalTaxableIncomeJanMar >= 0 And TotalTaxableIncomeJanMar <= 288000 Then
                taxPayble1 = TotalTaxableIncomeJanMar / 10
            End If
            taxPaybleJanMar = taxPayble1 + taxPayble2 + taxPayble3 + taxPayble4 + taxPayble5
         End If
End If
taxPayble = taxPaybleJanMar
getTotalTaxPayble2020JanMar = taxPayble
End Function


'Added by Ruth and Lawrence
'Calculate Employment Income tax for the period Apr to Dec 2020
Function getTotalTaxPayble2020AprDec(ByVal TotalTaxableIncome As Double, _
ByVal RtnYear As Integer, ByVal TotalTaxableIncomeAprDec As Double) As Double
Application.Volatile
Dim taxPayble As Double
Dim taxPaybleAprDec As Double
Dim taxPayble1 As Double
Dim taxPayble2 As Double
Dim taxPayble3 As Double
Dim taxPayble4 As Double
Dim taxPayble5 As Double

taxPayble1 = 0
taxPayble2 = 0
taxPayble3 = 0
taxPayble4 = 0
taxPayble5 = 0
taxPayble = 0
If TotalTaxableIncomeAprDec < 0 Then
    taxPaybleAprDec = 0
Else

    If (RtnYear = 2020) Then
    
            Dim diff As Double
            taxPayble1 = TotalTaxableIncomeAprDec - 216000
                If (taxPayble1 < 0) Then
                taxPayble1 = TotalTaxableIncomeAprDec / 10
                ElseIf (taxPayble1 = 0) Then
                taxPayble1 = 24000# * 9#
                taxPayble1 = taxPayble1 * 10 / 100
                ElseIf (taxPayble1 > 0) Then
                taxPayble1 = 24000# * 9#
                taxPayble1 = taxPayble1 * 0.1
                diff = TotalTaxableIncomeAprDec - (24000# * 9#)
                
                    If (diff < (16667# * 9#)) Then
                    taxPayble2 = diff * 0.15
                    ElseIf (diff = (16667# * 9#)) Then
                    taxPayble2 = 16667# * 9# * 0.15
                    taxBandAprDec = 2
                    ElseIf (diff > 16667# * 9#) Then
                    taxPayble2 = 16667# * 9# * 0.15
                    diff = TotalTaxableIncomeAprDec - ((24000# + 16667#) * 9#)
                        

                        If (diff < (16667# * 9#)) Then
                        taxPayble3 = diff * 0.2
                        ElseIf (diff = (16667# * 9#)) Then
                        taxPayble3 = (16667# * 9#) * 0.2
                        ElseIf (diff > (16667# * 9#)) Then
                        taxPayble3 = (16667# * 9#) * 0.2
                        diff = TotalTaxableIncomeAprDec - ((57334#) * 9#)
                            If diff > 0 Then
                            taxPayble4 = diff * 0.25
                            End If
                        End If
                    
                    End If
                End If
         End If
         
          If (RtnYear = 2023) Then
        'TotalTaxableIncomeJanMar = TotalIncomeJanMar - TotalDeductionSJanMar - ExemptedAmtSJanMar
            If TotalTaxableIncomeAprDec > 9600000 Then
                taxPayble1 = (TotalTaxableIncomeAprDec - 9600000) * (35 / 100)
                taxPayble2 = (9600000 - 6000000)
                taxPayble2 = taxPayble2 * (32.5 / 100)
                taxPayble3 = (6000000 - 388000)
                taxPayble3 = taxPayble3 * (30 / 100)
                taxPayble4 = (388000 - 288000)
                taxPayble4 = taxPayble4 * (25 / 100)
                taxPayble5 = 288000 * (10 / 100)
            End If
            If TotalTaxableIncomeAprDec > 6000000 And TotalTaxableIncomeAprDec <= 9600000 Then
                taxPayble1 = (TotalTaxableIncomeAprDec - 6000000) * (32.5 / 100)
                taxPayble2 = (6000000 - 388000)
                taxPayble2 = taxPayble2 * (30 / 100)
                taxPayble3 = (388000 - 288000)
                taxPayble3 = taxPayble3 * (25 / 100)
                taxPayble4 = 288000 * (10 / 100)
            End If
            If TotalTaxableIncomeAprDec > 388000 And TotalTaxableIncomeAprDec <= 6000000 Then
                taxPayble1 = (TotalTaxableIncomeAprDec - 388000) * (30 / 100)
                taxPayble2 = (388000 - 288000)
                taxPayble2 = taxPayble2 * (25 / 100)
                taxPayble3 = 288000 * (10 / 100)
            End If
            If TotalTaxableIncomeAprDec > 288000 And TotalTaxableIncomeAprDec <= 388000 Then
                taxPayble1 = (TotalTaxableIncomeAprDec - 288000) * (25 / 100)
                taxPayble2 = 288000 * (10 / 100)
            End If
            
            If TotalTaxableIncomeAprDec >= 0 And TotalTaxableIncomeAprDec <= 288000 Then
                taxPayble1 = TotalTaxableIncomeAprDec / 10
            End If
         End If
         
    taxPaybleAprDec = taxPayble1 + taxPayble2 + taxPayble3 + taxPayble4 + taxPayble5
End If
taxPayble = taxPaybleAprDec
getTotalTaxPayble2020AprDec = taxPayble
End Function
'Added by Ruth and Lawrence
'Calculate Business Income tax for the period Jan to Dec 2020
Function getTotalTaxPayable2020Biz(ByVal TotalTaxableIncome As Double, _
ByVal RtnYear As Integer, ByVal TotalTaxableIncomeBiz As Double) As Double
Application.Volatile
Dim taxPayble As Double
Dim taxPaybleAprDec As Double
Dim taxPayble1 As Double
Dim taxPayble2 As Double
Dim taxPayble3 As Double
Dim taxPayble4 As Double
Dim taxPayble5 As Double

taxPayble1 = 0
taxPayble2 = 0
taxPayble3 = 0
taxPayble4 = 0
taxPayble5 = 0
taxPayble = 0
If TotalTaxableIncomeBiz < 0 Then
    taxPayble = 0
Else
taxBandAprDec = 0
 Dim empTaxableincome2020 As Double
 Dim usedEmpTaxableincome2020 As Double
 Dim TotalTaxableIncomeBizMinusUsedEmp2020 As Double
 TotalTaxableIncomeBizMinusUsedEmp2020 = 0
 usedEmpTaxableincome2020 = 0
 empTaxableincome2020 = 0
 'empTaxableincome2020 = Range("TaxComp.EmpIncomeListSTO").value
 empTaxableincome2020 = TotalTaxableIncome - TotalTaxableIncomeBiz
If (RtnYear = 2020) Then
        If (empTaxableincome2020 / 688000) >= 1 Then
            taxBandAprDec = 4
            ElseIf (empTaxableincome2020 / 688000) < 1 Then
                If (empTaxableincome2020 / 488000) >= 1 Then
                        taxBandAprDec = 3
                        usedEmpTaxableincome2020 = empTaxableincome2020 - 488000
                ElseIf (empTaxableincome2020 / 488000) < 1 Then
                    If (empTaxableincome2020 / 288000) >= 1 Then
                        taxBandAprDec = 2
                        usedEmpTaxableincome2020 = empTaxableincome2020 - 288000
                    Else
                        taxBandAprDec = 1
                        usedEmpTaxableincome2020 = empTaxableincome2020
                    End If
                End If
        End If



        Select Case taxBandAprDec
        Case Is = 0
                If TotalTaxableIncomeBiz > 688000 Then
                    taxPayble1 = (TotalTaxableIncomeBiz - 688000) * 25 / 100
                    taxPayble2 = ((688000 - 488001) + 1)
                    taxPayble2 = taxPayble2 * 20 / 100
                    taxPayble3 = ((488000 - 288001) + 1)
                    taxPayble3 = taxPayble3 * 15 / 100
                    taxPayble4 = 2880000 / 100
                End If
                If TotalTaxableIncomeBiz >= 488001 And TotalTaxableIncomeBiz <= 688000 Then
                    taxPayble1 = (TotalTaxableIncomeBiz - 488000) * 20 / 100
                    taxPayble2 = ((488000 - 288001) + 1)
                    taxPayble2 = taxPayble2 * 15 / 100
                    taxPayble3 = 2880000 / 100
                End If
                If TotalTaxableIncomeBiz >= 288001 And TotalTaxableIncomeBiz <= 488000 Then
                    taxPayble1 = (TotalTaxableIncomeBiz - 288000) * 15 / 100
                    taxPayble2 = 2880000 / 100
                End If
                If TotalTaxableIncomeBiz >= 0 And TotalTaxableIncomeBiz <= 288000 Then
                    taxPayble1 = TotalTaxableIncomeBiz / 10
                End If
            Case Is = 1
                If (288000 >= TotalTaxableIncomeBiz) Then
                    'Added by Ruth and Lawrence on 08/03/2021
                    If (TotalTaxableIncomeBiz > (288000 - usedEmpTaxableincome2020)) Then
                    taxPayble1 = (288000 - usedEmpTaxableincome2020) * 10 / 100
                    TotalTaxableIncomeBizMinusUsedEmp2020 = TotalTaxableIncomeBiz - (288000 - usedEmpTaxableincome2020)
                    taxPayble2 = TotalTaxableIncomeBizMinusUsedEmp2020 * 15 / 100
                    Else
                    taxPayble1 = TotalTaxableIncomeBiz * 10 / 100
                    End If
        
        
                Else
                TotalTaxableIncomeBizMinusUsedEmp2020 = TotalTaxableIncomeBiz - (288000 - usedEmpTaxableincome2020)
                taxPayble1 = (288000 - usedEmpTaxableincome2020) * 10 / 100
                taxPayble2 = TotalTaxableIncomeBizMinusUsedEmp2020
                    If (taxPayble2 <= 200000) Then
                        taxPayble2 = taxPayble2 * 15 / 100
                    Else
                        taxPayble3 = TotalTaxableIncomeBizMinusUsedEmp2020 - 200000
                        taxPayble2 = 200000 * 15 / 100
                        If (taxPayble3 <= 200000) Then
                                 taxPayble3 = taxPayble3 * 20 / 100
                             Else
                                 taxPayble4 = TotalTaxableIncomeBizMinusUsedEmp2020 - 200000 - 200000
                                 taxPayble3 = 200000 * 20 / 100
                                 taxPayble4 = taxPayble4 * 25 / 100
                        End If
                    End If
                End If
            Case Is = 2
                If (488000 >= TotalTaxableIncomeBiz) Then
        '        taxPayble1 = TotalTaxableIncomeBiz * 15 / 100
                 'Added by Ruth and Lawrence on 08/03/2021
                    If TotalTaxableIncomeBiz > 200000 Then
                    taxPayble1 = (200000 - usedEmpTaxableincome2020) * 15 / 100
                    TotalTaxableIncomeBizMinusUsedEmp2020 = TotalTaxableIncomeBiz - (200000 - usedEmpTaxableincome2020)
                        If TotalTaxableIncomeBizMinusUsedEmp2020 > 200000 Then
                        taxPayble2 = 200000 * 20 / 100
                        taxPayble3 = TotalTaxableIncomeBizMinusUsedEmp2020 - 200000
                        taxPayble3 = taxPayble3 * 25 / 100
                        Else
                        taxPayble2 = TotalTaxableIncomeBizMinusUsedEmp2020 * 20 / 100
                        End If
                    Else
                    
                        If TotalTaxableIncomeBiz > (200000 - usedEmpTaxableincome2020) Then
                        taxPayble1 = (200000 - usedEmpTaxableincome2020) * 15 / 100
                        taxPayble2 = TotalTaxableIncomeBiz - (200000 - usedEmpTaxableincome2020)
                        taxPayble2 = taxPayble2 * 20 / 100
                        Else
                        taxPayble1 = TotalTaxableIncomeBiz * 15 / 100
                        End If
                    
                    End If
        
                 
                Else
                'taxPayble1 = 150003 * 15 / 100
                TotalTaxableIncomeBizMinusUsedEmp2020 = TotalTaxableIncomeBiz - (200000 - usedEmpTaxableincome2020)
                taxPayble1 = (200000 - usedEmpTaxableincome2020) * 15 / 100
                taxPayble2 = TotalTaxableIncomeBizMinusUsedEmp2020
                    If (taxPayble2 <= 200000) Then
                        taxPayble2 = taxPayble2 * 20 / 100
                    Else
                        taxPayble3 = TotalTaxableIncomeBizMinusUsedEmp2020 - 200000
                        taxPayble2 = 200000 * 20 / 100
                        taxPayble3 = taxPayble3 * 25 / 100
                    End If
                End If
            Case Is = 3
                If (688000 >= TotalTaxableIncomeBiz) Then
                 
          'Added by Ruth and Lawrence on 08/03/2021
                    If TotalTaxableIncomeBiz > 200000 Then
                    taxPayble1 = (200000 - usedEmpTaxableincome2020) * 20 / 100
                    TotalTaxableIncomeBizMinusUsedEmp2020 = TotalTaxableIncomeBiz - (200000 - usedEmpTaxableincome2020)
                    taxPayble2 = TotalTaxableIncomeBizMinusUsedEmp2020 * 25 / 100
                    Else
        
                        If TotalTaxableIncomeBiz > (200000 - usedEmpTaxableincome2020) Then
                        taxPayble1 = (200000 - usedEmpTaxableincome2020) * 20 / 100
                        taxPayble2 = TotalTaxableIncomeBiz - (200000 - usedEmpTaxableincome2020)
                        taxPayble2 = taxPayble2 * 25 / 100
                        Else
                        taxPayble1 = TotalTaxableIncomeBiz * 20 / 100
                        End If
        
                    End If
                Else
                'taxPayble1 = 150003 * 15 / 100
                TotalTaxableIncomeBizMinusUsedEmp2020 = TotalTaxableIncomeBiz - (200000 - usedEmpTaxableincome2020)
                taxPayble1 = (200000 - usedEmpTaxableincome2020) * 20 / 100
                taxPayble2 = TotalTaxableIncomeBizMinusUsedEmp2020
                taxPayble2 = taxPayble2 * 25 / 100
                End If
            Case Is = 4
                taxPayble1 = TotalTaxableIncomeBiz * 25 / 100
        End Select
End If


If (RtnYear = 2023) Then


            If (empTaxableincome2020 / 9600000) >= 1 Then
                taxBandAprDec = 5
            ElseIf (empTaxableincome2020 / 9600000) < 1 Then
                If (empTaxableincome2020 / 6000000) >= 1 Then
                            taxBandAprDec = 4
                            usedEmpTaxableincome2020 = empTaxableincome2020 - 6000000
                ElseIf (empTaxableincome2020 / 6000000) < 1 Then
                    If (empTaxableincome2020 / 388000) >= 1 Then
                        taxBandAprDec = 3
                        usedEmpTaxableincome2020 = empTaxableincome2020 - 388000
                    ElseIf (empTaxableincome2020 / 388000) < 1 Then
                        If (empTaxableincome2020 / 288000) >= 1 Then
                            taxBandAprDec = 2
                             usedEmpTaxableincome2020 = empTaxableincome2020 - 288000
                         Else
                         If (usedEmpTaxableincome2020 = 0) Then
                            taxBandAprDec = 0
                          'usedEmpTaxableincome2020 = empTaxableincome2020
                         Else
                            taxBandAprDec = 1
                            usedEmpTaxableincome2020 = empTaxableincome2020
                         End If
                            
                         End If
                    End If
                End If
            End If
 

        Select Case taxBandAprDec
        Case Is = 0
                If TotalTaxableIncomeBiz > 9600000 Then
                    taxPayble1 = (TotalTaxableIncomeBiz - 9600000) * 35 / 100
                    taxPayble2 = ((9600000 - 6000001) + 1)
                    taxPayble2 = taxPayble2 * 32.5 / 100
                    taxPayble3 = ((6000000 - 388001) + 1)
                    taxPayble3 = taxPayble3 * 30 / 100
                    taxPayble4 = ((388000 - 288001) + 1)
                    taxPayble4 = taxPayble4 * 25 / 100
                    taxPayble5 = 2880000 / 100
                End If
                If TotalTaxableIncomeBiz >= 6000001 And TotalTaxableIncomeBiz <= 9600000 Then
                    taxPayble1 = (TotalTaxableIncomeBiz - 6000000) * 32.5 / 100
                    taxPayble2 = ((6000000 - 388001) + 1)
                    taxPayble2 = taxPayble2 * 30 / 100
                    taxPayble3 = ((388000 - 288001) + 1)
                    taxPayble3 = taxPayble3 * 25 / 100
                    taxPayble4 = 2880000 / 100
                End If
                If TotalTaxableIncomeBiz >= 388001 And TotalTaxableIncomeBiz <= 6000000 Then
                    taxPayble1 = (TotalTaxableIncomeBiz - 388000) * 30 / 100
                    taxPayble2 = ((388000 - 288001) + 1)
                    taxPayble2 = taxPayble2 * 25 / 100
                    taxPayble3 = 2880000 / 100
                End If
                If TotalTaxableIncomeBiz >= 288001 And TotalTaxableIncomeBiz <= 388000 Then
                    taxPayble1 = (TotalTaxableIncomeBiz - 288000) * 25 / 100
                    taxPayble2 = 2880000 / 100
                End If
                If TotalTaxableIncomeBiz >= 0 And TotalTaxableIncomeBiz <= 288000 Then
                    taxPayble1 = TotalTaxableIncomeBiz / 10
                End If
            Case Is = 1
                If (288000 >= TotalTaxableIncomeBiz) Then
                    'Added by Ruth and Lawrence on 08/03/2021
                    If (TotalTaxableIncomeBiz > (288000 - usedEmpTaxableincome2020)) Then
                    taxPayble1 = (288000 - usedEmpTaxableincome2020) * 10 / 100
                    TotalTaxableIncomeBizMinusUsedEmp2020 = TotalTaxableIncomeBiz - (288000 - usedEmpTaxableincome2020)
                    taxPayble2 = TotalTaxableIncomeBizMinusUsedEmp2020 * 25 / 100
                    Else
                    taxPayble1 = TotalTaxableIncomeBiz * 10 / 100
                    End If
                Else
                TotalTaxableIncomeBizMinusUsedEmp2020 = TotalTaxableIncomeBiz - (288000 - usedEmpTaxableincome2020)
                taxPayble1 = (288000 - usedEmpTaxableincome2020) * 10 / 100
                taxPayble2 = TotalTaxableIncomeBizMinusUsedEmp2020
                    If (taxPayble2 <= 100000) Then
                        taxPayble2 = taxPayble2 * 25 / 100
                    Else
                        taxPayble3 = TotalTaxableIncomeBizMinusUsedEmp2020 - 388000
                        taxPayble2 = 100000 * 25 / 100
                         If (taxPayble3 <= 5612000) Then
                                 taxPayble3 = taxPayble3 * 30 / 100
                         Else
                                taxPayble4 = TotalTaxableIncomeBizMinusUsedEmp2020 - 6000000
                                taxPayble3 = 5612000 * 30 / 100
                                If (taxPayble4 <= 3600000) Then
                                    taxPayble4 = taxPayble4 * 32.5 / 100
                                Else
                                    taxPayble5 = TotalTaxableIncomeBizMinusUsedEmp2020 - 9600000
                                    taxPayble5 = taxPayble5 * 35 / 100
                                    taxPayble4 = 3600000 * 32.5 / 100
                                End If
                        End If
                    End If
                End If
            Case Is = 2
                If (388000 >= TotalTaxableIncomeBiz) Then
                
                
                
                    If (TotalTaxableIncomeBiz > (388000 - usedEmpTaxableincome2020)) Then
                        taxPayble1 = (388000 - usedEmpTaxableincome2020) * 25 / 100
                        TotalTaxableIncomeBizMinusUsedEmp2020 = TotalTaxableIncomeBiz - (388000 - usedEmpTaxableincome2020)
                        taxPayble2 = TotalTaxableIncomeBizMinusUsedEmp2020 * 30 / 100
                    Else
                        taxPayble1 = TotalTaxableIncomeBiz * 25 / 100
                    End If
                Else
                TotalTaxableIncomeBizMinusUsedEmp2020 = TotalTaxableIncomeBiz - (100000 - usedEmpTaxableincome2020)
                taxPayble1 = (100000 - usedEmpTaxableincome2020) * 25 / 100
                taxPayble2 = TotalTaxableIncomeBizMinusUsedEmp2020
                    If (taxPayble2 <= 5612000) Then
                                 taxPayble2 = taxPayble2 * 30 / 100
                    Else
                                taxPayble3 = TotalTaxableIncomeBizMinusUsedEmp2020 - 6000000
                                taxPayble2 = 5612000 * 30 / 100
                                If (taxPayble3 <= 3600000) Then
                                    taxPayble3 = taxPayble3 * 32.5 / 100
                                Else
                                    taxPayble4 = TotalTaxableIncomeBizMinusUsedEmp2020 - 9600000
                                    taxPayble3 = 3600000 * 32.5 / 100
                                    taxPayble4 = taxPayble4 * 35 / 100
                                End If
                    End If
                End If
            Case Is = 3
                If (6000000 >= TotalTaxableIncomeBiz) Then
                
                
                
                    If (TotalTaxableIncomeBiz > (6000000 - usedEmpTaxableincome2020)) Then
                        taxPayble1 = (6000000 - usedEmpTaxableincome2020) * 30 / 100
                        TotalTaxableIncomeBizMinusUsedEmp2020 = TotalTaxableIncomeBiz - (6000000 - usedEmpTaxableincome2020)
                        taxPayble2 = TotalTaxableIncomeBizMinusUsedEmp2020 * 32.5 / 100
                    Else
                        taxPayble1 = TotalTaxableIncomeBiz * 30 / 100
                    End If
                Else
                TotalTaxableIncomeBizMinusUsedEmp2020 = TotalTaxableIncomeBiz - (5612000 - usedEmpTaxableincome2020)
                taxPayble1 = (5612000 - usedEmpTaxableincome2020) * 30 / 100
                taxPayble2 = TotalTaxableIncomeBizMinusUsedEmp2020
                    If (taxPayble2 <= 3600000) Then
                                 taxPayble2 = taxPayble2 * 32.5 / 100
                    Else
                                taxPayble3 = TotalTaxableIncomeBizMinusUsedEmp2020 - 3600000
                                taxPayble2 = 3600000 * 32.5 / 100
                                taxPayble3 = taxPayble3 * 35 / 100
                                
                    End If
                End If
                
            Case Is = 4
                If (9600000 >= TotalTaxableIncomeBiz) Then
                    If (TotalTaxableIncomeBiz > (9600000 - usedEmpTaxableincome2020)) Then
                        taxPayble1 = (9600000 - usedEmpTaxableincome2020) * 32.5 / 100
                        TotalTaxableIncomeBizMinusUsedEmp2020 = TotalTaxableIncomeBiz - (9600000 - usedEmpTaxableincome2020)
                        taxPayble2 = TotalTaxableIncomeBizMinusUsedEmp2020 * 35 / 100
                    Else
                        taxPayble1 = TotalTaxableIncomeBiz * 32.5 / 100
                    End If
                Else
                TotalTaxableIncomeBizMinusUsedEmp2020 = TotalTaxableIncomeBiz - (3600000 - usedEmpTaxableincome2020)
                taxPayble1 = (3600000 - usedEmpTaxableincome2020) * 32.5 / 100
                'taxPayble2 = TotalTaxableIncomeBizMinusUsedEmp2020
                taxPayble2 = TotalTaxableIncomeBizMinusUsedEmp2020
                taxPayble2 = taxPayble2 * 35 / 100
                End If
            Case Is = 5
                taxPayble1 = TotalTaxableIncomeBiz * 35 / 100
        End Select
 End If


taxPayble = taxPayble1 + taxPayble2 + taxPayble3 + taxPayble4 + taxPayble5
End If

getTotalTaxPayable2020Biz = taxPayble
End Function


'getInterestOnPenaltyAmount function calculate Interest on Penalty Amount Start
Function getInterestOnPenaltyAmount(ByVal balOfTaxDue As Double, ByVal rtnEndDate As String) As Double
Dim overMonths, yearDiff As Integer
Dim penalty As Double
penalty = 0
If (rtnEndDate <> "") Then
    yearDiff = year(Now) - year(DateValue(rtnEndDate))
    If Month(DateValue(Now)) > 7 Then
        overMonths = Month(DateValue(Now)) - 7
        If (overMonths >= 0) Then
            penalty = (balOfTaxDue * 2 * (overMonths + ((yearDiff - 1) * 12))) / 100
        End If
    End If
End If
getInterestOnPenaltyAmount = penalty
End Function
'getInterestOnPenaltyAmount function calculate Interest on Penalty Amount End


'getPenaltyOfInstallmentTax function check penalty on installment tax base on Return period End Date Start
Function getPenaltyOfInstallmentTax(ByVal rtnEndDate As String) As Double
    Dim penalty As Double
    penalty = 0
    penalty = penalty * 20 / 100
    getPenaltyOfInstallmentTax = penalty
End Function
'getPenaltyOfInstallmentTax function check penalty on installment tax base on Return period End Date End

'************************************************
' CODE ADDED FOR AMENDMENT
'************************************************

'fillDataInFields  parse the string and fills the single cell data values to the corresponding cells and internally calls the fillListDataInFields which fills data in list Start
Public Sub fillDataInFields(amendmentSheet As String)

    'Declare Seperators
    Dim PROP_SEP As String
    PROP_SEP = "@P_@"
    
    Dim CLASS_SEP As String
    CLASS_SEP = "#C_@"
    
    Dim VALUE_SEP As String
    VALUE_SEP = "%V_@"

    ' For Bank Details
    Dim bankId As String
    Dim branchId As String
    

    Set wSheet = Worksheets(amendmentSheet)
    If wSheet.Cells(1, 1).value = "" And wSheet.Cells(2, 1).value = "" Then
         Exit Sub
    Else

    'wSheet.Activate
    wSheet.Unprotect (Pwd)
    
    Dim singleCellDataStr As String
    Dim cellPropDataPairArr As Variant
    Dim cellPropDataPair As Variant
    Dim noData As Integer
    'Worksheets(amendmentSheet).Activate
    For colLength = 1 To 256
        If (Worksheets(amendmentSheet).Cells(1, colLength) = "") Then
                Exit For
            End If
        singleCellDataStr = Worksheets(amendmentSheet).Cells(1, colLength)
        
        'new code added for clear value start
        Worksheets(amendmentSheet).Cells(1, colLength).ClearContents
        'new code added for clear value end
    
        If singleCellDataStr = "" Then
         'do Nothing
        Else
                    
          For i = 1 To Worksheets.Count
            Worksheets(i).Unprotect (Pwd)
          Next
          
          cellPropDataPairArr = Split(singleCellDataStr, PROP_SEP) 'This gives Array of property name Value pair
            For i = 0 To UBound(cellPropDataPairArr) 'Loop of prop Value pairs starts
                   If cellPropDataPairArr(i) <> "" Then
                     cellPropDataPair = Split(cellPropDataPairArr(i), VALUE_SEP) 'This gives array of size 2 where 0 index has logical name and 1 has its corresponding value
                       If UBound(cellPropDataPair) = 1 Then
                            Dim Rng As Range
                            If cellPropDataPair(0) <> "" And cellPropDataPair(1) <> "" Then
                                
                                'Code to activate Sheet where given range is referenced
                                Dim rngTest As Range
                                On Error Resume Next
                                Dim sheetNo As Integer
                                
                                For k = 6 To ActiveWorkbook.Sheets.Count Step 1
                                     'Try to set our variable as the named range.
                                    Set rngTest = ActiveWorkbook.Sheets(k).Range(cellPropDataPair(0))
                                     'If there is no error then the name exists.
                                    If Err = 0 Then
                                         Worksheets(k).Activate
                                         sheetNo = k
                                        Exit For
                                    Else
                                         'Clear the error
                                        Err.Clear
                                    End If
                                Next k
                                
                                ' Code for Bank Id and Branch Id - Starts
                                If cellPropDataPair(0) = "BankDtl.BankNameS" Then
                                    bankId = cellPropDataPair(1)
                                ElseIf cellPropDataPair(0) = "BankDtl.BranchNameS" Then
                                    branchId = cellPropDataPair(1)
                                End If
                                ' Code for Bank Id and Branch Id - Ends
                                
                                If Range(cellPropDataPair(0)).Locked = True And InStr(1, Range(cellPropDataPair(0)).Formula, "=") = 1 Then
                                    'do Nothing
                                Else
                                    Range(cellPropDataPair(0)).value = cellPropDataPair(1)
                               End If
                            End If
                       End If
                   End If
            Next i 'Loop of prop Value pairs Ends here
        
            Call setBankBranchNameFromId("A_Basic_Info", "BankS", bankId, "BranchS", branchId)
            
            'new code added start
            If Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value = "Yes" Then
                    Worksheets("B_Profit_Loss_Account_Wife").Visible = xlSheetVisible
                    Worksheets("T_Income_Computation_Wife").Visible = xlSheetVisible
            ElseIf Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value = "No" Then
                    Worksheets("B_Profit_Loss_Account_Wife").Visible = xlHidden
                    Worksheets("T_Income_Computation_Wife").Visible = xlHidden
            End If
            'new code added End
        
            'new code added for Set Date Start
            Dim rtnPrdFrom As String
            Dim rtnPrdTo As String
            Dim rtnDepositDate As String
            Dim mon As String
            Dim mmU As String
            Dim year As String
            Dim Depstyear As String
            Dim DepstDate As String
            rtnPrdFrom = Target.value
            Sheet14.Unprotect (Pwd)
            If (Sheet14.Range("RetInf.RetStartDate").value <> "") Then
                If (TestDate(Sheet14.Range("RetInf.RetStartDate").value) = True) Then
                    startDate = CDate(Format(Sheet14.Range("RetInf.RetStartDate").value, "dd/mm/yyyy"))
                End If
            End If
            
            If (startDate <> "") Then
                dd = Format(CDate(Trim(startDate)), "DD")
                mm = Format(CDate(Trim(startDate)), "MM")
                mmU = Format(CDate(Trim(startDate)), "MM")
                year = DatePart("yyyy", startDate)
                Dim rtnTo  As String
                rtnTo = mm & "/" & year
                Depstyear = mm & "/" & year + 1
                DepstDate = dd & "/" & mm & "/" & year - 1
                
                Sheet14.Activate
                Sheet14.Range("RetInf.RtnPrdToAct").value = "31/12/" & year
                Sheet14.Range("SecA.RtnYear").value = year
                Sheet14.Range("RetInf.DepositStartDate").value = DepstDate
            End If
            If (Sheet14.Range("RetInf.RetEndDate").value <> "") Then
                If (TestDate(Sheet14.Range("RetInf.RetEndDate").value) = True) Then
                    endDate = CDate(Format(Sheet14.Range("RetInf.RetEndDate").value, "dd/mm/yyyy"))
                End If
            End If
            If (endDate <> "") Then
                mm = Format(CDate(Trim(endDate)), "MM")
                dd = Format(CDate(Trim(endDate)), "dd")
                year = Format(CDate(Trim(endDate)), "yyyy")
                yearU = Format(CDate(Trim(endDate)), "yyyy")
                mmU = Format(CDate(Trim(endDate)), "MM")
                rtnTo = mm & "/" & year
                
                'by Palak for Enh-6 SR2 starts amendment
                    'For Section E1 Part 1
                    Worksheets("E1_IDA_CA").Unprotect (Pwd)
                    Worksheets("E2_CA_WTA_WDV").Unprotect (Pwd)
                    rangeName = Worksheets("E1_IDA_CA").Range("IniAllPlanMach.ListPart1S").Address
                    startRow = Worksheets("E1_IDA_CA").Range(rangeName).row
                    endRow = Worksheets("E1_IDA_CA").Range(rangeName).Rows.Count + startRow - 1
                    
                    If year >= 2020 Then
                        If (year = 2020 And mm >= 4) Or year > 2020 Then
                            Call lockUnlock_cell_rng("E1_IDA_CA", "A" & startRow & ":A" & endRow, True)
                            Call lockUnlock_cell_rng("E1_IDA_CA", "B" & startRow & ":B" & endRow, False)
                            Call lockUnlock_cell_rng("E1_IDA_CA", "C" & startRow & ":C" & endRow, True)
                            Call lockUnlock_cell_rng("E1_IDA_CA", "D" & startRow & ":D" & endRow, False)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "D5" & ":D7", True)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "E5" & ":E7", True)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "B4" & ":C4", False)
                            Worksheets("E2_CA_WTA_WDV").Range("B4").value = "25"
                            Worksheets("E2_CA_WTA_WDV").Range("C4").value = "10"
                            Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_WDV", "B4:C4", True)
                        Else
                            Call lockUnlock_cell_rng("E1_IDA_CA", "A" & startRow & ":A" & endRow, False)
                            Call lockUnlock_cell_rng("E1_IDA_CA", "B" & startRow & ":B" & endRow, True)
                            Call lockUnlock_cell_rng("E1_IDA_CA", "C" & startRow & ":C" & endRow, False)
                            Call lockUnlock_cell_rng("E1_IDA_CA", "D" & startRow & ":D" & endRow, True)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "D5" & ":D7", False)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "E5" & ":E7", False)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "B4" & ":C4", False)
                            Worksheets("E2_CA_WTA_WDV").Range("B4").value = "37.5"
                            Worksheets("E2_CA_WTA_WDV").Range("C4").value = "30"
                            Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_WDV", "B4:C4", True)
                        End If
                    Else
                        Call lockUnlock_cell_rng("E1_IDA_CA", "A" & startRow & ":A" & endRow, False)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "B" & startRow & ":B" & endRow, True)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "C" & startRow & ":C" & endRow, False)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "D" & startRow & ":D" & endRow, True)
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "D5" & ":D7", False)
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "E5" & ":E7", False)
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "B4" & ":C4", False)
                        Worksheets("E2_CA_WTA_WDV").Range("B4").value = "37.5"
                        Worksheets("E2_CA_WTA_WDV").Range("C4").value = "30"
                        Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_WDV", "B4:C4", True)
                    End If
                    
                    'For Section E1 Part 2
                    rangeName = Worksheets("E1_IDA_CA").Range("IniAllIBD.ListPart2S").Address
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
                    rangeName = Worksheets("E1_IDA_CA").Range("DeprIntengAst.ListS").Address
                    startRow = Worksheets("E1_IDA_CA").Range(rangeName).row
                    endRow = Worksheets("E1_IDA_CA").Range(rangeName).Rows.Count + startRow - 1
                    
                    If year >= 2020 Then
                        If (year = 2020 And mm >= 4) Or year > 2020 Then
                            Call lockUnlock_cell_rng("E1_IDA_CA", "A" & startRow & ":A" & endRow, True)
                            Call lockUnlock_cell_rng("E1_IDA_CA", "B" & startRow & ":B" & endRow, False)
                            Call lockUnlock_cell_rng("E1_IDA_CA", "C" & startRow & ":C" & endRow, True)
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
                   
                'Wife
                If UCase(Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value) = "YES" Then
                    rangeName = Worksheets("E1_IDA_CA").Range("IniAllPlanMach.ListPart1W").Address
                    startRow = Worksheets("E1_IDA_CA").Range(rangeName).row
                    endRow = Worksheets("E1_IDA_CA").Range(rangeName).Rows.Count + startRow - 1
                    
                    If year >= 2020 Then
                        If (year = 2020 And mm >= 4) Or year > 2020 Then
                            Call lockUnlock_cell_rng("E1_IDA_CA", "A" & startRow & ":A" & endRow, True)
                            Call lockUnlock_cell_rng("E1_IDA_CA", "B" & startRow & ":B" & endRow, False)
                            Call lockUnlock_cell_rng("E1_IDA_CA", "C" & startRow & ":C" & endRow, True)
                            Call lockUnlock_cell_rng("E1_IDA_CA", "D" & startRow & ":D" & endRow, False)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "D17" & ":D19", True)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "E17" & ":E19", True)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "B17" & ":C19", False)
                            Worksheets("E2_CA_WTA_WDV").Range("B16").value = "25"
                            Worksheets("E2_CA_WTA_WDV").Range("C16").value = "10"
                            Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_WDV", "B16:C16", True)
                        Else
                            Call lockUnlock_cell_rng("E1_IDA_CA", "A" & startRow & ":A" & endRow, False)
                            Call lockUnlock_cell_rng("E1_IDA_CA", "B" & startRow & ":B" & endRow, True)
                            Call lockUnlock_cell_rng("E1_IDA_CA", "C" & startRow & ":C" & endRow, False)
                            Call lockUnlock_cell_rng("E1_IDA_CA", "D" & startRow & ":D" & endRow, True)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "D17" & ":D19", False)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "E17" & ":E19", False)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "B17" & ":C19", False)
                            Worksheets("E2_CA_WTA_WDV").Range("B16").value = "37.5"
                            Worksheets("E2_CA_WTA_WDV").Range("C16").value = "30"
                            Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_WDV", "B16:C16", True)
                        End If
                    Else
                        Call lockUnlock_cell_rng("E1_IDA_CA", "A" & startRow & ":A" & endRow, False)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "B" & startRow & ":B" & endRow, True)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "C" & startRow & ":C" & endRow, False)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "D" & startRow & ":D" & endRow, True)
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "D17" & ":D19", False)
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "E17" & ":E19", False)
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "B17" & ":C19", False)
                        Worksheets("E2_CA_WTA_WDV").Range("B16").value = "37.5"
                        Worksheets("E2_CA_WTA_WDV").Range("C16").value = "30"
                        Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_WDV", "B16:C16", True)
                    End If
                    
                    'For Section E1 Part 2
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
                            Call lockUnlock_cell_rng("E1_IDA_CA", "C" & startRow & ":C" & endRow, True)
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
                End If
                
                Worksheets("E1_IDA_CA").Protect (Pwd)
                Worksheets("E2_CA_WTA_WDV").Protect (Pwd)
                'by Palak for Enh-6 SR2 ends
                
                Sheet14.Activate
                Sheet14.Range("RetInf.RtnPrdToActStart").value = "01/01/" & year
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
                        Sheet14.Range("RetInf.dateforAuditCertificate").value = dd & "/" & rtnTo
                Else
                    If (mon = "02") And IsDate(29 & "/" & rtnTo) Then
                        Sheet14.Range("RetInf.dateforAuditCertificate").value = 29 & "/" & rtnTo
                    ElseIf (mon = "02") And IsDate(28 & "/" & rtnTo) Then
                        Sheet14.Range("RetInf.dateforAuditCertificate").value = 28 & "/" & rtnTo
                    ElseIf (mon = "04" Or mon = "06" Or mon = "09" Or mon = "11") And IsDate(30 & "/" & rtnTo) Then
                        Sheet14.Range("RetInf.dateforAuditCertificate").value = 30 & "/" & rtnTo
                    End If
                    
                End If
                'new code added for Audit Date Start
                Dim firstDate As String, secondDate As String
                Dim auditStDt As String
                If (Sheet14.Range("RetInf.RetEndDate").value <> "") Then
                    firstDate = DateValue(Sheet14.Range("RetInf.RetEndDate").value)
                    secondDate = DateAdd("d", 1, firstDate)
                    auditStDt = Format(secondDate, "dd/mm/yyyy")
                End If
                Sheet14.Range("RetInf.auditStartDate").value = auditStDt
            End If
        End If
    Next

    Dim listDataStr As String
    'Worksheets(amendmentSheet).Activate
    For colLength = 1 To 256
        If (Worksheets(amendmentSheet).Cells(2, colLength) = "") Then
                Exit For
        End If
        listDataStr = Worksheets(amendmentSheet).Cells(2, colLength)
        
         'new code added for clear value start
          Worksheets(amendmentSheet).Cells(2, colLength).ClearContents
         'new code added for clear value end
    
        If Worksheets("A_Basic_Info").Range("RetInf.DeclareExemptionCerti").value = "Yes" Then
            Call lockUnlock_cell_rng("A_Basic_Info", "ExemptCerti.ListS", False, 4)
        End If
        
        If Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value = "No" Then
            Sheet14.toggleSpouseFields (True)
        Else
            'Sheet14.toggleSpouseFieldsWithoutClearingContent (False)
            If Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeExemptionCerti").value = "Yes" Then
                Call lockUnlock_cell_rng("A_Basic_Info", "ExemptCerti.ListW", False, 4)
            End If
        End If
    
        If listDataStr = "" Then
           'do Nothing
        Else
           Call fillListDataInFields(listDataStr)
        
       '****************************************
       'CODE SPECIFIC TO IT RESIDENT
        'Reset DATA FOR LIST Using IDS
        Call ResetDataFromIds
        'Call resetFields
       '****************************************
        
        End If
    Next colLength
    
      For i = 1 To Worksheets.Count
        Worksheets(i).Protect (Pwd)
      Next
    End If
End Sub
'fillDataInFields  parse the string and fills the single cell data values to the corresponding cells and internally calls the fillListDataInFields which fills data in list End

'ResetDataFromIds Function reset all data in entire sheet using passing Sheet name and in range Start
Public Sub ResetDataFromIds()

    Call setCountyNamesFromIds("A_Basic_Info", "RentPaid.RentListS", "G", "R")
    Call setDistrictNamesFromIds("A_Basic_Info", "RentPaid.RentListS", "H", "T")
    Call setLocalityNamesFromIds("A_Basic_Info", "RentPaid.RentListS", "I", "U")
    
    Call setCountyNamesFromIds("A_Basic_Info", "RentalIncome.ListS", "G", "R")
    Call setDistrictNamesFromIds("A_Basic_Info", "RentalIncome.ListS", "H", "T")
    Call setLocalityNamesFromIds("A_Basic_Info", "RentalIncome.ListS", "I", "U")
     
    'added on 29.12.2011
    Call setCapacityOnUsages("P_Advance_Tax_Credits", "VehicleAdvTaxPaid.ListS", "C", "D", "E")
    
    Call resetPINRange("F_Employment_Income", "EmpIncome.ListS", "PINofEmployerS", "A")
    Call resetPINRange("I_Computation_of_Car_Benefit", "CarBenefit.ListS", "PINofEmployerCarS", "A")
    Call resetPINRange("I_Computation_of_Car_Benefit", "CarBenefit.ListS", "valueOfCarS", "J")
    
    'new code added
    Call ToggleDataOnChangeOwnHire("I_Computation_of_Car_Benefit", "CarBenefit.ListS", "F", "G", "H")
    Call ToggleDataOnChangePolicyHolder("L_Computation_of_Insu_Relief", "InsReliefDtls.ListS", "E", "F")

    
    If Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value <> "" And Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value = "Yes" Then
            
        Call setCountyNamesFromIds("A_Basic_Info", "RentPaid.RentListW", "G", "R")
        Call setDistrictNamesFromIds("A_Basic_Info", "RentPaid.RentListW", "H", "T")
        Call setLocalityNamesFromIds("A_Basic_Info", "RentPaid.RentListW", "I", "U")
    
        Call setCountyNamesFromIds("A_Basic_Info", "RentalIncome.ListW", "G", "R")
        Call setDistrictNamesFromIds("A_Basic_Info", "RentalIncome.ListW", "H", "T")
        Call setLocalityNamesFromIds("A_Basic_Info", "RentalIncome.ListW", "I", "U")
       
        'added on 29.12.2011
        Call setCapacityOnUsages("P_Advance_Tax_Credits", "VehicleAdvTaxPaid.ListW", "C", "D", "E")
        
        Call resetPINRange("F_Employment_Income", "EmpIncome.ListW", "PINofEmployerW", "A")
        
        Call resetPINRange("I_Computation_of_Car_Benefit", "CarBenefit.ListW", "PINofEmployerCarW", "A")
        Call resetPINRange("I_Computation_of_Car_Benefit", "CarBenefit.ListW", "valueOfCarW", "J")
        
        'new code added
        Call ToggleDataOnChangeOwnHire("I_Computation_of_Car_Benefit", "CarBenefit.ListW", "F", "G", "H")
        Call ToggleDataOnChangePolicyHolder("L_Computation_of_Insu_Relief", "InsReliefDtls.ListW", "E", "F")
        
    End If
End Sub
'ResetDataFromIds Function reset all data in entire sheet using passing Sheet name and in range End


'resetFields function using for reset field conditionally Start
Public Sub resetFields()
'Code to reset fields for Spouse Details

If Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value <> "" Then
        
        If Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value = "Yes" Then
            lockCell = False
            Worksheets("B_Profit_Loss_Account_Wife").Visible = xlSheetVisible
            Worksheets("T_Income_Computation_Wife").Visible = xlSheetVisible
        ElseIf Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value = "No" Then
            lockCell = True
            Worksheets("B_Profit_Loss_Account_Wife").Visible = xlHidden
            Worksheets("T_Income_Computation_Wife").Visible = xlHidden
        End If
        Sheet14.toggleSpouseFields (lockCell)
End If

End Sub
'resetFields function using for reset field conditionally End


'fillListDataInFields parse the string and fills the list data values to the corresponding cells Start
Public Sub fillListDataInFields(listDataStr As String)

    Dim lookupCodeFlag As Boolean
    Dim lastColumn As Long
    Dim lastRow As Long
    Dim currentWorkSheet As Worksheet
    Dim cellRange As Range
    Dim nameCell As name
    Dim startRow As Long
    Dim startColumn As Long
    
    'Declare Seperators
    Dim PROP_SEP As String
    PROP_SEP = "@PL@"
    Dim LIST_SEP As String
    LIST_SEP = "@L_@"
    Dim VALUE_SEP As String
    VALUE_SEP = "%VL@"
    Dim listArr As Variant
    Dim singleListDataArr As Variant
    Dim singleRowData As Variant
    Dim lstName As String
    Dim lstRange As Range
    Dim rngSheet As String
    Dim startCnt As Integer
    Dim newLastRow As Integer
    Dim rowListindex As Integer
    
    listArr = Split(listDataStr, LIST_SEP)
    For i = 0 To UBound(listArr) 'Loop of all lists starts here
        If listArr(i) <> "" Then 'This is always the list's Data
            singleListDataArr = Split(listArr(i), PROP_SEP) 'Gets Array of row Data of single List
            lstName = singleListDataArr(0)
                 
            'added for string read from mutlicolumn start
            If previousLogicalName = lstName Then
                cellBreak = True
                bufferRow = prevCellBreakCounterEndRow
                prevCellBreakCounterEndRow = prevCellBreakCounterEndRow + UBound(singleListDataArr)
            Else
                cellBreak = False
                prevCellBreakCounterEndRow = UBound(singleListDataArr)
            End If
                 
            previousLogicalName = lstName
            'added for string read from mutlicolumn End
                 
            'Code to activate Sheet where given range is referenced
            Dim rngTest As Range
            On Error Resume Next
            For l = 6 To ActiveWorkbook.Sheets.Count Step 1
            'Try to set our variable as the named range.
                Set rngTest = ActiveWorkbook.Sheets(l).Range(lstName)
                'If there is no error then the name exists.
                If Err = 0 Then
                    Worksheets(l).Activate
                    Exit For
                Else
                'Clear the error
                    Err.Clear
                End If
            Next l
                 
            '*****************************
            'get positions of given list Range
            startRow = Range(lstName).row
                 
            'added for string read from mutlicolumn start
                 
            If (cellBreak <> True) Then
                startCnt = Range(lstName).row
            Else
                startCnt = bufferRow + startRow
            End If
            'added for string read from mutlicolumn End
                 
            startColumn = Range(lstName).column
            lastColumn = startColumn + Range(lstName).Columns.Count - 1
            lastRow = startRow + Range(lstName).Rows.Count - 1
            newLastRow = startRow + prevCellBreakCounterEndRow - 1
                 
            If (newLastRow - lastRow) > 0 And (lstName <> "WAT.ListS" And lstName <> "WAT.ListW" And lstName <> "WAT.ListBS" And lstName <> "WAT.ListBW") Then
                Application.ScreenUpdating = True
                Call InsertGivenRowsAndFillFormulas(lstName, (newLastRow - lastRow))
                Application.ScreenUpdating = False
            End If
            
            If (UBound(singleListDataArr) >= 1) Then
                rowListindex = 1
                    
                For j = startCnt To startCnt + UBound(singleListDataArr) - 1 'Loop of single list's multiple rows starts here
                    If singleListDataArr(rowListindex) <> "" Then
                        singleRowData = Split(singleListDataArr(rowListindex), VALUE_SEP) 'gets all column data of single row
                        'Code Specific To IT RESIDENT  RETURNS
                        If (lstName = "WAT.ListS" Or lstName = "WAT.ListW") Then
                            lookUpcode = singleRowData(UBound(singleRowData) - 1)
                            Dim lookUpCodeLast As String
                            lookUpCodeLast = Mid(lookUpcode, Len(lookUpcode), Len(lookUpcode))
                            If (InStr(1, lookUpcode, "rate") = 1 Or InStr(1, lookUpcode, "balBD") = 1 Or InStr(1, lookUpcode, "curAstV") = 1 Or InStr(1, lookUpcode, "sales") = 1) And (lookUpCodeLast = "H" Or lookUpCodeLast = "W") Then
                                If (lstName = "WAT.ListS") Then
                                    If (lookUpcode = "rateH") Then
                                        startCnt = startRow
                                        lookupCodeFlag = True
                                    ElseIf (lookUpcode = "balBDH") Then
                                        startCnt = startRow + 1
                                        lookupCodeFlag = True
                                    ElseIf (lookUpcode = "curAstVH") Then
                                        startCnt = startRow + 2
                                        lookupCodeFlag = True
                                    ElseIf (lookUpcode = "salesH") Then
                                        startCnt = startRow + 3
                                        lookupCodeFlag = True
                                    Else
                                        lookupCodeFlag = False
                                    End If
                                ElseIf (lstName = "WAT.ListW") Then
                                    If (lookUpcode = "rateW") Then
                                        startCnt = startRow
                                        lookupCodeFlag = True
                                    ElseIf (lookUpcode = "balBDW") Then
                                        startCnt = startRow + 1
                                        lookupCodeFlag = True
                                    ElseIf (lookUpcode = "curAstVW") Then
                                        startCnt = startRow + 2
                                        lookupCodeFlag = True
                                    ElseIf (lookUpcode = "salesW") Then
                                        startCnt = startRow + 3
                                        lookupCodeFlag = True
                                    Else
                                       lookupCodeFlag = False
                                    End If
                                End If
                                                 
                                If lookupCodeFlag = True Then
                                   For k = 0 To UBound(singleRowData) 'Loop of columns data filling for single row starts here
                                        If ActiveSheet.Cells(startCnt, startColumn + k).Locked = True And InStr(1, ActiveSheet.Cells(startCnt, startColumn + k).Formula, "=") = 1 Then ' If the cell has formula  or its locked then data should not be filled
                                              'do nothing
                                        Else
                                           ActiveSheet.Unprotect (Pwd)
                                           ActiveSheet.Cells(startCnt, startColumn + k).value = singleRowData(k)
                                           ActiveSheet.Protect (Pwd)
                                        End If
                                   Next k 'single row's column filling ends here
                                End If
                            End If
                        Else
                        ' if condition added by mitali to fill data in Profit Loss account row wise
                            If lstName = "PLA.ConsolidateDataS" Or lstName = "PLA.ConsolidateDataW" Or _
                               lstName = "PLA.BussIncomeDataS" Or lstName = "PLA.FarmIncomeDataS" Or _
                               lstName = "PLA.RentIncomeDataS" Or lstName = "PLA.IntIncomeDataS" Or _
                               lstName = "PLA.CommIncomeDataS" Or lstName = "PLA.OthIncomeDataS" Or _
                               lstName = "PLA.BussIncomeDataW" Or lstName = "PLA.FarmIncomeDataW" Or _
                               lstName = "PLA.RentIncomeDataW" Or lstName = "PLA.IntIncomeDataW" Or _
                               lstName = "PLA.CommIncomeDataW" Or lstName = "PLA.OthIncomeDataW" Or _
                               lstName = "DtlLossFrwd.BussinessS" Or lstName = "DtlLossFrwd.FarmingS" Or _
                               lstName = "DtlLossFrwd.RentalS" Or lstName = "DtlLossFrwd.InterestS" Or _
                               lstName = "DtlLossFrwd.CommissionS" Or lstName = "DtlLossFrwd.OtherS" Or _
                               lstName = "DtlLossFrwd.TotalS" Or lstName = "DtlLossFrwd.BussinessW" Or _
                               lstName = "DtlLossFrwd.FarmingW" Or lstName = "DtlLossFrwd.RentalW" Or _
                               lstName = "DtlLossFrwd.InterestW" Or lstName = "DtlLossFrwd.CommissionW" Or _
                               lstName = "DtlLossFrwd.OtherW" Or lstName = "DtlLossFrwd.TotalW" Then
                                For k = 0 To UBound(singleRowData) - 1 'Loop of rows data filling for single column starts here
                                                
                                ' If the cell has formula  or its locked or has formula then data should not be filled
                                'temp code added start'
                                    If ((startRow = 108 Or startRow = 109) And startColumn < 9) Then
                                        'do nothing
                                        'startRow = startRow + 1
                                        k = k + 1
                                    End If
                                    'temp code added End
                                    If ActiveSheet.Cells(startRow, startColumn).Locked = True And InStr(1, ActiveSheet.Cells(startRow, startColumn).Formula, "=") = 1 Then
                                        'do nothing
                                        startRow = startRow + 1
                                    ElseIf ActiveSheet.Cells(startRow, startColumn).Locked = True And ActiveSheet.Cells(startRow, startColumn).HasFormula = False Then
                                        startRow = startRow + 1
                                        k = k - 1
                                        'do nothing
                                    Else
                                        'General Code for all
                                        ActiveSheet.Unprotect (Pwd)
                                        ActiveSheet.Cells(startRow, startColumn).value = singleRowData(k)
                                        ActiveSheet.Protect (Pwd)
                                        startRow = startRow + 1
                                    End If
                                Next k 'single row's column filling ends here
                            ElseIf lstName = "TaxComp.BussinessListS" Or lstName = "TaxComp.CnslListS" Or _
                                   lstName = "TaxComp.CommListS" Or lstName = "TaxComp.FarmListS" Or _
                                   lstName = "TaxComp.IntListS" Or lstName = "TaxComp.OthListS" Or _
                                   lstName = "TaxComp.RentListS" Or lstName = "TaxComp.BussinessListW" Or _
                                   lstName = "TaxComp.CnslListW" Or lstName = "TaxComp.CommListW" Or _
                                   lstName = "TaxComp.FarmListW" Or lstName = "TaxComp.IntListW" Or _
                                   lstName = "TaxComp.OthListW" Or lstName = "TaxComp.RentListW" Then
                                For k = 0 To UBound(singleRowData) - 1 'Loop of rows data filling for single column starts here
                                    ' If the cell has formula  or its locked or has formula then data should not be filled
                                                
                                    If ActiveSheet.Cells(startRow, startColumn).Locked = True And InStr(1, ActiveSheet.Cells(startRow, startColumn).Formula, "=") = 1 Then
                                                    'do nothing
                                        startRow = startRow + 1
                                    ElseIf ActiveSheet.Cells(startRow, startColumn).Locked = True And ActiveSheet.Cells(startRow, startColumn).HasFormula = False Then
                                        startRow = startRow + 1
                                        k = k - 1
                                        'do nothing
                                    Else
                                        'General Code for all
                                        ActiveSheet.Unprotect (Pwd)
                                        ActiveSheet.Cells(startRow, startColumn).value = singleRowData(k)
                                        ActiveSheet.Protect (Pwd)
                                        startRow = startRow + 1
                                    End If
                                Next k 'single row's column filling ends here
                            Else
                                For k = 0 To UBound(singleRowData) 'Loop of columns data filling for single row starts here
                                            
                                ' If the cell has formula  or its locked or has formula then data should not be filled
                                    If ActiveSheet.Cells(startCnt, startColumn + k).Locked = True And InStr(1, ActiveSheet.Cells(startCnt, startColumn + k).Formula, "=") = 1 Then
                                        'do nothing
                                    Else
                                    'General Code for all
                                        ActiveSheet.Unprotect (Pwd)
                                        ActiveSheet.Cells(startCnt, startColumn + k).value = singleRowData(k)
                                        ActiveSheet.Protect (Pwd)
                                    End If
                                                 
                                Next k 'single row's column filling ends here
                            End If
                        End If
                    End If
                    startCnt = startCnt + 1
                    rowListindex = rowListindex + 1
                Next j 'all rows filling of single list ends here
            End If
        End If
    Next i 'All Lists data filling ends here
End Sub
'fillListDataInFields parse the string and fills the list data values to the corresponding cells End


'InsertGivenRowsAndFillFormulas Function insert new row base on passing RangeName and Row Start
Public Sub InsertGivenRowsAndFillFormulas(rangeName As String, vrows As Integer)
    
   On Error GoTo Errorcatch
  
    Dim startRow As Long, endRow As Long
    'get the start row number from the range name
    startRow = Range(rangeName).row
    ' get the total Number of rows present in the given Range Name
    endRow = Range(rangeName).Rows.Count + startRow - 1
    Range(rangeName).Select
    ActiveSheet.Rows(endRow).Select
    ActiveSheet.Unprotect (Pwd)
    Selection.Resize(rowsize:=2).Rows(2).EntireRow. _
    Resize(rowsize:=vrows).Insert Shift:=xlDown
    
    ActiveSheet.Unprotect (Pwd)
    Selection.AutoFill Selection.Resize( _
    rowsize:=vrows + 1), xlFillDefault
     
    Dim endRowNew As Long
    ' get the total Number of rows present in the given Range Name
    endRowNew = Range(rangeName).Rows.Count + startRow - 1
    If endRowNew = endRow Then
        With Range(rangeName)
            .Resize(.Rows.Count + vrows, .Columns.Count).name = rangeName
        End With
    End If
        On Error Resume Next    'to handle no constants in range -- John McKee 2000/02/01
        ' to remove the non-formulas -- 1998/03/11 Bill Manville
    ActiveSheet.Unprotect (Pwd)
    Selection.Offset(1).Resize(vrows).EntireRow. _
    SpecialCells(xlConstants).ClearContents
    ActiveSheet.Protect (Pwd)
    Exit Sub
    
Errorcatch:
MsgBox Err.Description
End Sub
'InsertGivenRowsAndFillFormulas Function insert new row base on passing RangeName and Row End

'setBankBranchNameFromId Function set BankName and BranchName base on Passing Sheet Name and BankId,BranchId Start
'setBankBranchNameFromId Function set BankName and BranchName base on Passing Sheet Name and BankId,BranchId Start
Public Sub setBankBranchNameFromId(sheetName As String, bankName As String, bankId As String, BranchName As String, branchId As String)

    Dim row As Integer
    Dim rngFound As Range
    Dim name As String
    
    Worksheets(sheetName).Unprotect (Pwd)

    If bankId <> "" And branchId <> "0" And TestNumber(bankId) Then
                
        Set rngFound = Worksheets("Data").Range("BankId").Cells.Find(What:=bankId, After:= _
                        Worksheets("Data").Range("BankId").Cells.Cells(1, 1), LookIn:=xlValues, _
                        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                        MatchCase:=False)
        If rngFound Is Nothing Then
            'do nothing
        Else
            'Code to reset branch
            row = rngFound.row
            name = Worksheets("Data").Cells(row, 2).value
            Worksheets(sheetName).Range(bankName).value = name
            Worksheets(sheetName).Range(bankName).Locked = False
            'Worksheets(sheetName).Range(bankId).Locked = True
        End If
    End If
    
    If branchId <> "" And branchId <> "0" And TestNumber(branchId) Then
        name = ""
        Set rngFound = Worksheets("Data").Range("BranchId").Cells.Find(What:=branchId, After:= _
                       Worksheets("Data").Range("BranchId").Cells.Cells(1, 1), LookIn:=xlValues, _
                        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                        MatchCase:=False)
        
        If rngFound Is Nothing Then
            'do nothing
        Else
            row = rngFound.row
            name = Worksheets("Data").Cells(row, 5).value
            Worksheets(sheetName).Range(BranchName).value = name
            Worksheets(sheetName).Range(BranchName).Locked = False
            'Worksheets(sheetName).Range(branchId).Locked = True
        End If
    End If
End Sub
'setBankBranchNameFromId Function set BankName and BranchName base on Passing Sheet Name and BankId,BranchId End

'find_CountyNameFromId function using for find county name from countyID Start
Public Function find_CountyNameFromId(col) As String
Dim name As String
Dim rngFound As Range

    Set rngFound = Sheet18.Range("countyId").Cells.Find(What:=col, After:= _
                    Sheet18.Range("countyId").Cells.Cells(1, 1), LookIn:=xlValues, _
           LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
           MatchCase:=False)
    If rngFound Is Nothing Then
        name = ""
    Else
        name = Sheet18.Cells(rngFound.row, 16).value
    End If
    find_CountyNameFromId = name
End Function
'find_CountyNameFromId function using for find county name from countyID End

'setCountyNamesFromIds function using for set County Name from County Id Start
Public Sub setCountyNamesFromIds(sheetName As String, listRange As String, countyNameCol As String, CountyIdCol As String)

    Worksheets(sheetName).Activate
    Worksheets(sheetName).Unprotect (Pwd)
    
    Dim countyId As String
    Dim countyName As String
    Dim startRow As Integer
    Dim startColumn As Integer
    Dim lastRow As Integer
    
    startRow = Range(listRange).row
    startColumn = Range(listRange).column
    lastRow = startRow + Range(listRange).Rows.Count - 1
        
    For i = startRow To lastRow
        Worksheets(sheetName).Activate
        countyId = ActiveSheet.Range(CountyIdCol & i & ":" & CountyIdCol & i).value
        countyName = find_CountyNameFromId(countyId)
        Worksheets(sheetName).Activate
        ActiveSheet.Range(countyNameCol & i & ":" & countyNameCol & i).Select
        Selection.value = countyName
    Next i
    
    Worksheets(sheetName).Protect (Pwd)
End Sub
'setCountyNamesFromIds function using for set County Name from County Id End

'setBankNamesFromIds function using for set Bank Name from BankId Start
Public Sub setBankNamesFromIds(sheetName As String, listRange As String, BankNameCol As String, BankIdCol As String)

    Worksheets(sheetName).Activate
    Worksheets(sheetName).Unprotect (Pwd)
    Application.ScreenUpdating = False
    Dim bankId As String
    Dim bankName As String
    Dim startRow As Integer
    Dim startColumn As Integer
    Dim lastRow As Integer
    
    startRow = Range(listRange).row
    startColumn = Range(listRange).column
    lastRow = startRow + Range(listRange).Rows.Count - 1
        
    For i = startRow To lastRow
        Worksheets(sheetName).Activate
        bankId = ActiveSheet.Range(BankIdCol & i & ":" & BankIdCol & i).value
        bankName = find_bankNameFromId(bankId)
        Worksheets(sheetName).Activate
        ActiveSheet.Range(BankNameCol & i & ":" & BankNameCol & i).Select
        Worksheets(sheetName).Unprotect (Pwd)
        Selection.value = bankName
    
    Next i
    
    If bankId <> "" Then
         find_BranchFromBankID (bankId)
    End If
    Worksheets(sheetName).Protect (Pwd)
End Sub
'setBankNamesFromIds function using for set Bank Name from BankId End

'find_bankNameFromId function using for find bank name from bank Id Start
Public Function find_bankNameFromId(col) As String
    
    Dim name As String
    Dim rngFound As Range
    
    Set rngFound = Sheet18.Range("BankId").Cells.Find(What:=col, After:=Sheet18.Range("BankId").Cells.Cells(1, 1), LookIn:=xlValues, _
           LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
           MatchCase:=False)
    If rngFound Is Nothing Then
        name = ""
    Else
        name = Sheet18.Cells(rngFound.row, BankIdCol + 1).value
    End If
    
    find_bankNameFromId = name

End Function
'find_bankNameFromId function using for find bank name from bank Id End

'setBranchNamesFromIds function using for set branch name from Branch Id Start
Public Sub setBranchNamesFromIds(sheetName As String, listRange As String, BranchNameCol As String, BranchIdCol As String)

    Worksheets(sheetName).Activate
    Worksheets(sheetName).Unprotect (Pwd)
    Dim branchId As String
    Dim BranchName As String
    Dim startRow As Integer
    Dim startColumn As Integer
    Dim lastRow As Integer
    
    startRow = Range(listRange).row
    startColumn = Range(listRange).column
    lastRow = startRow + Range(listRange).Rows.Count - 1
        
    For i = startRow To lastRow
        Worksheets(sheetName).Activate
        branchId = ActiveSheet.Range(BranchIdCol & i & ":" & BranchIdCol & i).value
        BranchName = find_BranchNameFromId(branchId)
        Worksheets(sheetName).Activate
        ActiveSheet.Range(BranchNameCol & i & ":" & BranchNameCol & i).Select
        Worksheets(sheetName).Unprotect (Pwd)
        Selection.value = BranchName
        Worksheets(sheetName).Unprotect (Pwd)
        Selection.Locked = True
        
    Next i
    
    Worksheets(sheetName).Protect (Pwd)
End Sub
'setBranchNamesFromIds function using for set branch name from Branch Id End

'find_BranchNameFromId function using for find branch name from data sheet and set branch name related code in hidden field in the sheet Start
Public Function find_BranchNameFromId(col) As String
    Dim name As String
    Dim rngFound As Range
    
    
    Set rngFound = Sheet18.Range("BranchId").Cells.Find(What:=col, After:= _
                    Sheet18.Range("BranchId").Cells.Cells(1, 1), LookIn:=xlValues, _
                    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                    MatchCase:=False)
    
    If rngFound Is Nothing Then
        name = ""
    Else
        name = Sheet18.Cells(rngFound.row, BranchIdCol + 1).value
    End If

    find_BranchNameFromId = name

End Function
'find_BranchNameFromId function using for find branch name from data sheet and set branch name related code in hidden field in the sheet End

Public Sub setTaxRateFrmNatureOfPayments(sheetName As String, listRange As String, NatPymtCol As String, TaxRateColNo As String)

    Worksheets(sheetName).Activate
    Worksheets(sheetName).Unprotect (Pwd)
    Application.ScreenUpdating = False
    Dim TaxRate As String
    Dim NtrPymt As String
    Dim startRow As Integer
    Dim startColumn As Integer
    Dim lastRow As Integer
    Dim rowvalue As Double
    
    startRow = Range(listRange).row
    startColumn = Range(listRange).column
    lastRow = startRow + Range(listRange).Rows.Count - 1
        
    For i = startRow To lastRow
        Worksheets(sheetName).Activate
        NtrPymt = ActiveSheet.Range(NatPymtCol & i & ":" & NatPymtCol & i).value
        rowvalue = find_TaxRATE(NtrPymt)
        If rowvalue <> -1 Then
            TaxRate = Sheet18.Cells(rowvalue, TaxRateCol).value
            Worksheets(sheetName).Unprotect (Pwd)
            ActiveSheet.Range(TaxRateColNo & i & ":" & TaxRateColNo & i).value = TaxRate
        End If
        Worksheets(sheetName).Activate
    Next
    Worksheets(sheetName).Protect (Pwd)
End Sub
           
Public Sub setCapacityOnUsages(sheetName As String, listRange As String, UsageCol As String, LoadCapCol As String, SeatCapCol As String)
    
    Dim usage As String
    Dim loadCap As String
    Dim SeatCap As String
    
    Worksheets(sheetName).Activate
    Worksheets(sheetName).Unprotect (Pwd)
    rangeName = ActiveSheet.Range(listRange).Address
    startRow = Range(rangeName).row
    endRow = Range(rangeName).Rows.Count + startRow - 1
    
    startColumn = Range(rangeName).column
    lastColumn = startColumn + Range(rangeName).Columns.Count - 1
                         
    For i = startRow To endRow
          usage = ActiveSheet.Range(UsageCol & i & ":" & UsageCol & i).value
          loadCap = ActiveSheet.Range(LoadCapCol & i & ":" & LoadCapCol & i).value
          SeatCap = ActiveSheet.Range(SeatCapCol & i & ":" & SeatCapCol & i).value
                If (usage = "Van, Pick-ups, Trucks, Lorries" Or usage = "") Then
                    Call lockUnlock_cell_rng_Hidden_without_clearing_contents(ActiveSheet.name, LoadCapCol & i & ":" & LoadCapCol & i, False)
                    ActiveSheet.Range(LoadCapCol & i & ":" & LoadCapCol & i).value = loadCap
                    Call lockUnlock_cell_rng_Hidden_without_clearing_contents(ActiveSheet.name, SeatCapCol & i & ":" & SeatCapCol & i, True)
                ElseIf (usage = "Saloons, Station-Wagons, Minibuses, Buses, Coaches") Then
                    Call lockUnlock_cell_rng_Hidden_without_clearing_contents(ActiveSheet.name, LoadCapCol & i & ":" & LoadCapCol & i, True)
                    Call lockUnlock_cell_rng_Hidden_without_clearing_contents(ActiveSheet.name, SeatCapCol & i & ":" & SeatCapCol & i, False)
                    ActiveSheet.Range(SeatCapCol & i & ":" & SeatCapCol & i).value = SeatCap
               End If
    Next
    Worksheets(sheetName).Protect (Pwd)

End Sub

Public Sub setCostOnOwnedOrHired(sheetName As String, listRange As String, TypeCol As String, HireCostCol As String, OwnCostCol As String)
    
    Dim TypeCar As String
    Dim OwnCost As String
    Dim HireCol As String
    Worksheets(sheetName).Activate
    Worksheets(sheetName).Unprotect (Pwd)
    rangeName = ActiveSheet.Range(listRange).Address
    startRow = Range(rangeName).row
    endRow = Range(rangeName).Rows.Count + startRow - 1
    
    startColumn = Range(rangeName).column
    lastColumn = startColumn + Range(rangeName).Columns.Count - 1
                         
    For i = startRow To endRow
          TypeCar = ActiveSheet.Range(TypeCol & i & ":" & TypeCol & i).value
          OwnCost = ActiveSheet.Range(OwnCostCol & i & ":" & OwnCostCol & i).value
          HireCol = ActiveSheet.Range(HireCostCol & i & ":" & HireCostCol & i).value
                If (TypeCar = "Own" Or TypeCar = "") Then
                    Call lockUnlock_cell_rng_Hidden_without_clearing_contents(ActiveSheet.name, OwnCostCol & i & ":" & OwnCostCol & i, False)
                    ActiveSheet.Range(OwnCostCol & i & ":" & OwnCostCol & i).value = OwnCost
                    Call lockUnlock_cell_rng_Hidden_without_clearing_contents(ActiveSheet.name, HireCostCol & i & ":" & HireCostCol & i, True)
               
                ElseIf (TypeCar = "Hired") Then
                    Call lockUnlock_cell_rng_Hidden_without_clearing_contents(ActiveSheet.name, OwnCostCol & i & ":" & OwnCostCol & i, True)
                    Call lockUnlock_cell_rng_Hidden_without_clearing_contents(ActiveSheet.name, HireCostCol & i & ":" & HireCostCol & i, False)
                    ActiveSheet.Range(HireCostCol & i & ":" & HireCostCol & i).value = HireCol
               End If
    Next
    
    Worksheets(sheetName).Protect (Pwd)

End Sub

'********************************Methods to fetch codes for combo boxes *********************
'                                         start

'** For schedule 14 in this sheet*******************
Public Function find_NatureOfPaymentCode(col) As Double
    Dim row As Integer
    Dim rngFound As Range
            
    Set rngFound = Sheet18.Range("NatureOfPayment").Cells.Find(What:=col, After:= _
                    Sheet18.Range("NatureOfPayment").Cells.Cells(1, 1), LookIn:=xlValues, _
                    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                    MatchCase:=False)
    
    If rngFound Is Nothing Then
        row = -1
    Else
        row = rngFound.row
    End If
    find_NatureOfPaymentCode = row
End Function

'** For schedule 16 in this sheet*******************
Public Function find_NatureOfIncomeCode(col) As Double
Dim row As Integer
Dim rngFound As Range
        
Set rngFound = Sheet18.Range("natureOfIncome").Cells.Find(What:=col, _
               After:=Sheet18.Range("natureOfIncome").Cells.Cells(1, 1), LookIn:=xlValues, _
               LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
               MatchCase:=False)

If rngFound Is Nothing Then
    row = -1
Else
    row = rngFound.row
End If

find_NatureOfIncomeCode = row

End Function

'** For schedule 13 in this sheet*******************
Public Function find_UsageCode(col) As Double
Dim row As Integer
Dim rngFound As Range
        
Set rngFound = Sheet18.Range("Usage").Cells.Find(What:=col, After:=Sheet18.Range("Usage").Cells.Cells(1, 1), LookIn:=xlValues, _
       LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
       MatchCase:=False)

If rngFound Is Nothing Then
    row = -1
Else
    row = rngFound.row
End If

find_UsageCode = row

End Function

'** For schedule 11 in this sheet*******************
Public Function find_PolicyHolderCode(col) As Double
Dim row As Integer
Dim rngFound As Range
        
Set rngFound = Sheet18.Range("PolicyHolder").Cells.Find(What:=col, After:=Sheet18.Range("PolicyHolder").Cells.Cells(1, 1), LookIn:=xlValues, _
       LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
       MatchCase:=False)

If rngFound Is Nothing Then
    row = -1
Else
    row = rngFound.row
End If

find_PolicyHolderCode = row

End Function

'** For schedule 8 in this sheet*******************
Public Function find_BodyTypeCode(col) As Double
Dim row As Integer
Dim rngFound As Range
        
Set rngFound = Sheet18.Range("BodyTypes").Cells.Find(What:=col, _
                After:=Sheet18.Range("BodyTypes").Cells.Cells(1, 1), LookIn:=xlValues, _
                LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                MatchCase:=False)

If rngFound Is Nothing Then
    row = -1
Else
    row = rngFound.row
End If

find_BodyTypeCode = row

End Function

'** For schedule 8 in this sheet*******************
Public Function find_OwnHireCode(col) As Double
Dim row As Integer
Dim rngFound As Range
       
Set rngFound = Sheet18.Range("CarType").Cells.Find(What:=col, _
                After:=Sheet18.Range("CarType").Cells.Cells(1, 1) _
                , LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
       MatchCase:=False)

If rngFound Is Nothing Then
    row = -1
Else
    row = rngFound.row
End If

find_OwnHireCode = row

End Function


'** For A_Basic_Info in this sheet*******************
Public Function find_RetTypeCode(col) As Double
Dim row As Integer
Dim rngFound As Range
        
Set rngFound = Sheet18.Range("ReturnType").Cells.Find(What:=col, After:=Sheet18.Range("ReturnType").Cells.Cells(1, 1), LookIn:=xlValues, _
       LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
       MatchCase:=False)

If rngFound Is Nothing Then
    row = -1
Else
    row = rngFound.row
End If

find_RetTypeCode = row

End Function


'ADDED BY MITALI ON 15/02/2011
Public Function validateDateBtwnRtnPeriod(ByRef value As Range, Optional ByVal colName As String) As String
Dim sheetName As String
Dim rowNumber As String
sheetName = value.Worksheet.name


If Trim(colName) = "" Then
    If TestDateBtwnRtnPeriod(value.value) = False Then
     rowNumber = value.Address
    End If
Else
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)
    
    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If c.value <> "" And TestDateBtwnRtnPeriod(c.value) = False Then
                'rowNumber = rowNumber & C.Address & ","
                collAddress = colName & c.row
                If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
            End If
        Next
    End If
End With

End If
sheetName = value.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
validateDateBtwnRtnPeriod = rowNumber
End Function

Public Function TestDateBtwnRtnPeriod(ByVal value As String) As Boolean
Dim mm As Integer
Dim yr As Integer
Dim startDate As String
Dim endDate As String
Dim sheet As Worksheet

startDate = Range("RetInf.RetStartDate").value
endDate = Range("RetInf.RetEndDate").value

If TestDate(value) = True And TestDate(startDate) = True And TestDate(endDate) = True Then
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
    tValue = CDate(Format(value, "dd/mm/yyyy"))
    If DateValue(startDate) <= DateValue(tValue) And DateValue(tValue) <= DateValue(endDate) Then
          TestDateBtwnRtnPeriod = True
    Else
          TestDateBtwnRtnPeriod = False
    End If
Else
        TestDateBtwnRtnPeriod = False
End If


End Function
'ENDED BY MITALI ON 15/02/2011

'Code To Reset the Return Year Combo Box
Sub resetRtnYrComboBox(shtName As String, colName As String)
    Dim yrRange As String
    Dim startRow As Integer
    Dim endRow As Integer
    Dim currYr As Long
    Dim vrows As Integer
         
    Worksheets(shtName).Unprotect (Pwd)
    currYr = Format(Now(), "yyyy")
    startRow = Worksheets(shtName).Range("RtnYear").row
    endRow = Worksheets(shtName).Range("RtnYear").Rows.Count + startRow - 1
    
   If Worksheets(shtName).Range(colName & endRow & ":" & colName & endRow).value < currYr Then
    For i = Worksheets(shtName).Range(colName & endRow & ":" & colName & endRow).value To currYr - 2
        vrows = vrows + 1
        Worksheets(shtName).Range(colName & (endRow + vrows) & ":" & colName & (endRow + vrows)).value = Worksheets(shtName).Range(colName & (endRow + vrows - 1) & ":" & colName & (endRow + vrows - 1)).value + 1
        
    Next
   End If
    
    yrRange = colName & startRow & ":" & colName & endRow + vrows
    ActiveWorkbook.Names("RtnYear").Delete
    Worksheets(shtName).Range(yrRange).name = "RtnYear"
    Worksheets(shtName).Protect (Pwd)
End Sub
Function validatePrnAmount(ByRef value As Range, Optional ByVal colName As String) As String

Dim column As Range
Dim columnRange As Range
Dim blnFlag As Boolean
blnFlag = True
Dim sName As String
Dim singleRow As Range
Dim rowNumbers As String
sName = value.Cells.Parent.name
'MsgBox flagAdd.row

Dim sheetName As String
If Trim(colName) = "" Then
    If value <> 0 Then
        rowNumbers = "NE"
    Else
       'rowNumbers = value.row
       'line commented and new line added by mitali to display cell address in
       'validation sheet to add hyperlink in the error sheet on 08 Dec,2011
       rowNumbers = value.Address
    End If
Else
Set columnRange = Worksheets(sName).Range(colName & "1", colName & Worksheets(sName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)

    With Worksheets(sName)
    If Not column Is Nothing Then

        For Each c In column.Cells
            On Error Resume Next
            'MsgBox Application.ActiveCell.Address

            If Trim(c) = 0 Then
                blnFlag = False
                'check if there is content in the entire row
                Set singleRow = Intersect(value, Worksheets(sName).Range(.Cells(c.row, 1), .Cells(c.row, Worksheets(sName).UsedRange.Columns.Count)))
                For Each singleCell In singleRow.Cells
                    collAddress = colName & singleRow.Cells.row
'                    MsgBox collAddress & Worksheets(sName).Range(collAddress & ":" & collAddress).value
                    If Worksheets(sName).Range(collAddress & ":" & collAddress).value = 0 Then
                        blnFlag = True
                        Exit For
                    End If
                Next
                If blnFlag Then

                     'rowNumbers = rowNumbers & singleRow.Cells.row & ","
                     'line commented and new line added by mitali to display cell address in
                     'validation sheet to add hyperlink in the error sheet on 08 Dec,2011
                     collAddress = colName & singleRow.Cells.row
                     If rowNumbers <> "" Then
                        rowNumbers = rowNumbers & "," & collAddress
                     Else
                        rowNumbers = rowNumbers & collAddress
                     End If
                End If
            End If
            'Application.ScreenUpdating = False
        Next
    End If
    End With
End If
If rowNumbers <> "" Then
    validatePrnAmount = rowNumbers
Else
    validatePrnAmount = "NE"
End If
End Function




Public Function compareValues(ByRef value1 As Range, ByRef value2 As Range) As String
Dim rowNumber As String
    If value1.value <> "" Or value2.value <> "" Then
        If Round(value1.value) <> Round(value2.value) Then
            'rowNumber = value.row
           'line commented and new line added by mitali to display cell address in
           'validation sheet to add hyperlink in the error sheet on 08 Dec,2011
           rowNumber = value1.Address
         End If
    End If
sheetName = value1.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
compareValues = rowNumber
End Function

Public Function validateExpenseList(ByRef value As Range, ByVal rangeName As String, ByRef flagAdd As Range, ByVal currCol As String, ByVal othCol As String) As String
Dim flag As String
flag = flagAdd.value
Dim column As Range
Dim columnRange As Range
Dim blnFlag As Boolean
blnFlag = True
Dim sName As String
Dim singleRow As Range
Dim rowNumbers As String
sName = value.Cells.Parent.name
'MsgBox flagAdd.row

Dim sheetName As String

    startRow = Worksheets(sName).Range(rangeName).row
    endRow = Worksheets(sName).Range(rangeName).Rows.Count + startRow - 1
    
    With Worksheets(sName)
        For i = startRow To endRow
            On Error Resume Next
                blnFlag = False
                'check if there is content in the entire row
                If Worksheets(sName).Range("B" & i & ":" & "B" & i).value <> "" Then
                    If Worksheets(sName).Range(currCol & i & ":" & currCol & i).value = "" And Worksheets(sName).Range(othCol & i & ":" & othCol & i).value = "" Then
                        blnFlag = True
                    End If
                End If
                If blnFlag Then
                     'rowNumbers = rowNumbers & singleRow.Cells.row & ","
                     'line commented and new line added by mitali to display cell address in
                     'validation sheet to add hyperlink in the error sheet on 08 Dec,2011
                     collAddress = currCol & i
                     If rowNumbers <> "" Then
                        rowNumbers = rowNumbers & "," & collAddress
                     Else
                        rowNumbers = rowNumbers & collAddress
                     End If
                     
                End If
            'Application.ScreenUpdating = False
        Next
    End With
If rowNumbers = "" Then
    rowNumbers = "NE"
End If
validateExpenseList = rowNumbers
End Function

Sub findPrev_sheet()
Dim flag As Boolean
flag = True
    Dim i As Integer
    For i = 1 To Worksheets.Count
        If flag = True Then
            If Trim(Worksheets(i).name) = Trim(ActiveSheet.name) Then
                    For k = i - 1 To 1 Step -1
                        If Worksheets(k).Visible = xlHidden Or Worksheets(k).Visible = xlVeryHidden Then
                            'do nothing
                        Else
                            Worksheets(k).Activate
                            Exit For
                        End If
                    Next
                
                'Set cursor on the first position
                For Each Rng In ActiveSheet.UsedRange.Rows
                  For Each Cell In Rng.Cells
                    If Not Cell.EntireRow.Hidden Then
                       If Not Cell.Locked Then
                          Cell.Select
                          Exit Sub
                       End If
                    End If
                  Next Cell
                Next Rng
                
                flag = False
            End If
        End If
    Next
End Sub

Sub findNext_sheet()
Dim flag As Boolean
flag = True
    Dim i As Integer
    For i = 1 To Worksheets.Count
        If flag = True Then
            If Trim(Worksheets(i).name) = Trim(ActiveSheet.name) Then
                For k = i + 1 To Worksheets.Count
                        If Worksheets(k).Visible = xlHidden Or Worksheets(k).Visible = xlVeryHidden Then
                            'do nothing
                        Else
                            Worksheets(k).Activate
                            Exit For
                        End If
                    Next
                
                'Set cursor on the first position
                For Each Rng In ActiveSheet.UsedRange.Rows
                  For Each Cell In Rng.Cells
                    If Not Cell.EntireRow.Hidden Then
                       If Not Cell.Locked Then
                          Cell.Select
                          Exit Sub
                       End If
                    End If
                  Next Cell
                Next Rng
                flag = False
            End If
        End If
    Next
End Sub
Public Sub resetPINRange(sheetName As String, rangeName As String, listName As String, listNameCol As String)
    
    Dim Str As String
    On Error GoTo Errorcatch
  
   
    Dim startRow As Long, endRow As Long
    'get the start row number from the range name
    startRow = Range(rangeName).row
    ' get the total Number of rows present in the given Range Name
    endRow = Range(rangeName).Rows.Count + startRow - 1
    Worksheets(sheetName).Activate
    Worksheets(sheetName).Unprotect (Pwd)
        
     On Error Resume Next    'to handle no constants in range -- John McKee 2000/02/01
     Str = listNameCol & startRow & ":" & listNameCol & endRow
     ActiveWorkbook.Names(listName).Delete
     Range(Str).name = listName
     Worksheets(sheetName).Protect (Pwd)
    Exit Sub
    
Errorcatch:
MsgBox Err.Description
End Sub

Public Sub setDistrictNamesFromIds(sheetName As String, listRange As String, districtNameCol As String, districtIdCol As String)

    Worksheets(sheetName).Activate
    Worksheets(sheetName).Unprotect (Pwd)
    
    Dim districtId As String
    Dim districtName As String
    Dim startRow As Integer
    Dim startColumn As Integer
    Dim lastRow As Integer
    
    startRow = Range(listRange).row
    startColumn = Range(listRange).column
    lastRow = startRow + Range(listRange).Rows.Count - 1
        
    For i = startRow To lastRow
        districtId = ActiveSheet.Range(districtIdCol & i & ":" & districtIdCol & i).value
        districtName = find_DistrictNameFromId(districtId)
        ActiveSheet.Range(districtNameCol & i & ":" & districtNameCol & i).value = districtName
    Next i
    
    Worksheets(sheetName).Protect (Pwd)
End Sub

Public Sub setLocalityNamesFromIds(sheetName As String, listRange As String, LocalityNameCol As String, LocalityIdCol As String)

    Worksheets(sheetName).Activate
    Worksheets(sheetName).Unprotect (Pwd)
    Dim LocalityId As String
    Dim LocalityName As String
    Dim startRow As Integer
    Dim startColumn As Integer
    Dim lastRow As Integer
    
    startRow = Range(listRange).row
    startColumn = Range(listRange).column
    lastRow = startRow + Range(listRange).Rows.Count - 1
        
    For i = startRow To lastRow
        LocalityId = ActiveSheet.Range(LocalityIdCol & i & ":" & LocalityIdCol & i).value
        LocalityName = find_LocalityNameFromId(LocalityId)
        ActiveSheet.Range(LocalityNameCol & i & ":" & LocalityNameCol & i).value = LocalityName
    Next i
    Worksheets(sheetName).Protect (Pwd)
End Sub

Public Function find_DistrictNameFromId(col) As String
    Dim name As String
    Dim rngFound As Range
    
    Set rngFound = Sheet18.Range("districtId").Cells.Find(What:=col, After:= _
                    Sheet18.Range("districtId").Cells.Cells(1, 1), LookIn:=xlValues, _
                    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                    MatchCase:=False)
    If rngFound Is Nothing Then
        name = ""
    Else
        name = Sheet18.Cells(rngFound.row, 28).value
    End If
    find_DistrictNameFromId = name
End Function
Public Function find_LocalityNameFromId(col) As String

    Dim name As String
    Dim rngFound As Range
    
    Set rngFound = Sheet18.Range("localityId").Cells.Find(What:=col, After:= _
                    Sheet18.Range("localityId").Cells.Cells(1, 1), LookIn:=xlValues, _
                    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                    MatchCase:=False)
    
    If rngFound Is Nothing Then
        name = ""
    Else
        name = Sheet18.Cells(rngFound.row, 32).value
    End If
    find_LocalityNameFromId = name
End Function
Public Function validateDateBeforeRtnPeriodEnd(ByRef value As Range, Optional ByVal colName As String) As String
Dim sheetName As String
Dim rowNumber As String
sheetName = value.Worksheet.name


If Trim(colName) = "" Then
    If TestDateBeforeRtnPeriodEnd(value.value) = False Then
     rowNumber = value.Address
    End If
Else
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)
    
    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If c.value <> "" And TestDateBeforeRtnPeriodEnd(c.value) = False Then
                'rowNumber = rowNumber & C.Address & ","
                collAddress = colName & c.row
                If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
            End If
        Next
    End If
End With

End If
sheetName = value.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
validateDateBeforeRtnPeriodEnd = rowNumber
End Function
Public Function TestDateBeforeRtnPeriodEnd(ByVal value As String) As Boolean
Dim mm As Integer
Dim yr As Integer
Dim startDate As String
Dim endDate As String
Dim sheet As Worksheet
Dim tValue As String
Dim nEndDate As String

endDate = Range("RetInf.RetEndDate").value

If TestDate(value) = True And TestDate(endDate) = True Then

nEndDate = CDate(Format((Sheet14.Range("RetInf.RetEndDate").value), "dd/mm/yyyy"))
tValue = CDate(Format(value, "dd/mm/yyyy"))

    If DateValue(tValue) <= DateValue(nEndDate) Then
          TestDateBeforeRtnPeriodEnd = True
    Else
          TestDateBeforeRtnPeriodEnd = False
    End If
Else
        TestDateBeforeRtnPeriodEnd = False
End If


End Function

'added function for compare sum of values Start
Public Function compareValuesSum(ByRef value1 As Range, ByRef value2 As Range, ByRef value3 As Range) As String
Dim rowNumber As String
    If value1.value <> "" Or value2.value <> "" Or value3.value <> "" Then
        Dim sum As Double
        sum = (Round(CDbl(value2.value), 2) + Round(CDbl(value3.value), 2))
        absoluteDiff = Abs(sum - Round(CDbl(value1.value), 2))

        If absoluteDiff > 1E-15 Then
            'rowNumber = value.row
            'line commented and new line added by mitali to display cell address in
            'validation sheet to add hyperlink in the error sheet on 08 Dec,2011
           rowNumber = value1.Address
         End If
    End If
sheetName = value1.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
compareValuesSum = rowNumber
End Function

Public Sub ToggleDataOnChangeOwnHire(sheetName As String, listRange As String, refCol As String, targetCol1 As String, targetCol2 As String)
    
    Dim refColVal As String
    
    Worksheets(sheetName).Activate
    Worksheets(sheetName).Unprotect (Pwd)
    rangeName = ActiveSheet.Range(listRange).Address
    startRow = Range(rangeName).row
    endRow = Range(rangeName).Rows.Count + startRow - 1
    
    startColumn = Range(rangeName).column
    lastColumn = startColumn + Range(rangeName).Columns.Count - 1
                         
    For i = startRow To endRow
          refColVal = ActiveSheet.Range(refCol & i & ":" & refCol & i).value
          
                If (refColVal = "Own" Or refColVal = "") Then
                    Call lockUnlock_cell_rng_Hidden_without_clearing_contents(ActiveSheet.name, targetCol1 & i & ":" & targetCol1 & i, True)
                    Call lockUnlock_cell_rng_Hidden_without_clearing_contents(ActiveSheet.name, targetCol2 & i & ":" & targetCol2 & i, False)
                ElseIf (refColVal = "Hired") Then
                    Call lockUnlock_cell_rng_Hidden_without_clearing_contents(ActiveSheet.name, targetCol1 & i & ":" & targetCol1 & i, False)
                    Call lockUnlock_cell_rng_Hidden_without_clearing_contents(ActiveSheet.name, targetCol2 & i & ":" & targetCol2 & i, True)
                End If
    Next
    
    Worksheets(sheetName).Protect (Pwd)

End Sub

Public Sub ToggleDataOnChangePolicyHolder(sheetName As String, listRange As String, refCol As String, targetCol As String)

    Dim refColVal As String
    Worksheets(sheetName).Activate
    Worksheets(sheetName).Unprotect (Pwd)
    rangeName = ActiveSheet.Range(listRange).Address
    startRow = Range(rangeName).row
    endRow = Range(rangeName).Rows.Count + startRow - 1
    
    startColumn = Range(rangeName).column
    lastColumn = startColumn + Range(rangeName).Columns.Count - 1
                         
    For i = startRow To endRow
          refColVal = ActiveSheet.Range(refCol & i & ":" & refCol & i).value
          
                If (refColVal = "Child") Then
                    Call lockUnlock_cell_rng_Hidden_without_clearing_contents(ActiveSheet.name, targetCol & i & ":" & targetCol & i, False)
                ElseIf (refColVal = "Wife" Or refColVal = "Self") Then
                    Call lockUnlock_cell_rng_Hidden_without_clearing_contents(ActiveSheet.name, targetCol & i & ":" & targetCol & i, True)
                ElseIf refColVal = "" Then
                    Call lockUnlock_cell_rng_Hidden_without_clearing_contents(ActiveSheet.name, targetCol & i & ":" & targetCol & i, False)
                End If
    Next
    
    Worksheets(sheetName).Protect (Pwd)

End Sub
Public Sub lockUnlock_cell_rng_reset(cursheet As String, rangeName As String, lockCellFlag As Boolean, Optional skipColumnFrom As String)

    Application.EnableEvents = False
    
    activeSheetName = ActiveSheet.name
    protectedStatus = ActiveSheet.ProtectContents
    
    Dim cellColor
    'On Error GoTo catch
    Worksheets(cursheet).Activate
    ActiveSheet.Unprotect Password:=Pwd
    
    startRow = Range(rangeName).row
    lastRow = Range(rangeName).Rows.Count + startRow - 1
    
    startColumn = Range(rangeName).column
    lastColumn = startColumn + Range(rangeName).Columns.Count - 1 ' get actual last column
    
    If lockCellFlag = True Then
        cellColor = RGB(146, 146, 146) 'lock
    ElseIf lockCellFlag = False Then
        cellColor = RGB(255, 255, 255) 'unlock
    End If
    If skipColumnFrom = "" Then
    ' locked whole range
        Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).Locked = lockCellFlag
        If lockCellFlag = False Then
            ' Do nothing
        Else
            Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).value = ""
        End If
        Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).Select
        With Selection.Interior
               .Color = cellColor
               .Pattern = xlSolid
               .PatternColorIndex = xlAutomatic
        End With
    Else
        'skip columns and define new range
        lastColumn = skipColumnFrom - 1
        Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).Select
        Selection.Locked = lockCellFlag
        If lockCellFlag = False Then
            'Do nothing
        Else
            Selection.value = ""
        End If
           
        With Selection.Interior
           .Color = cellColor
           .Pattern = xlSolid
           .PatternColorIndex = xlAutomatic
        End With
        
        ActiveSheet.Cells(3, 1).Select
    End If
    
catch:
    If Err.Description <> "" Then
    '    MsgBox Err.Description
    End If
    
    ActiveSheet.Protect Password:=Pwd
    Worksheets(activeSheetName).Activate
    If protectedStatus Then
        ActiveSheet.Protect Password:=Pwd
    Else
        ActiveSheet.Unprotect Password:=Pwd
    End If
End Sub


Public Function TestAuditDate(ByVal value As String) As Boolean
Dim mm As Integer
Dim yr As Integer
Dim startDate As String
Dim endDate As String
Dim auditEndDate As String
Dim sheet As Worksheet
Dim nEndDate As String
Dim auditDate As String
Dim tValue As String



endDate = Range("RetInf.RetEndDate").value
auditEndDate = Range("RetInf.dateforAuditCertificate").value



If TestDate(value) = True And TestDate(endDate) = True And TestDate(auditEndDate) = True Then


nEndDate = CDate(Format((Sheet14.Range("RetInf.RetEndDate").value), "dd/mm/yyyy"))
auditDate = CDate(Format((Sheet14.Range("RetInf.dateforAuditCertificate").value), "dd/mm/yyyy"))
tValue = CDate(Format(value, "dd/mm/yyyy"))

    If DateValue(tValue) >= DateValue(nEndDate) And DateValue(tValue) <= DateValue(auditDate) Then
          TestAuditDate = True
    Else
          TestAuditDate = False
    End If
Else
        TestAuditDate = False
End If


End Function



Public Function validateAuditDate(ByRef value As Range, Optional ByVal colName As String) As String
Dim sheetName As String
Dim rowNumber As String
sheetName = value.Worksheet.name


If Trim(colName) = "" Then
    If TestAuditDate(value.value) = False Then
     rowNumber = value.Address
    End If
Else
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)
    
    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If c.value <> "" And TestAuditDate(c.value) = False Then
                'rowNumber = rowNumber & C.Address & ","
                collAddress = colName & c.row
                If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
            End If
        Next
    End If
End With

End If
sheetName = value.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
validateAuditDate = rowNumber
End Function


Public Function compareAmntValues(ByRef value1 As Range, ByRef value2 As Range) As String
Dim rowNumber As String
    If value1.value <> "" Or value2.value <> "" Then
        If value1.value < value2.value Then
            'rowNumber = value.row
           'line commented and new line added by mitali to display cell address in
           'validation sheet to add hyperlink in the error sheet on 08 Dec,2011
           rowNumber = value1.Address
         End If
    End If
sheetName = value1.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
compareAmntValues = rowNumber
End Function

Public Function checkFutureDate(ByVal value As String) As Boolean
Dim mm As Integer
Dim yr As Integer
Dim tDvalue As String
Dim nwValue As String

If (value <> "") Then
    tValue = CDate(Format(value, "dd/mm/yyyy"))
End If
nwValue = CDate(Now())
If IsDate(value) Then
'If DateValue(value) <= DateValue(Format(Now(), "dd/MM/yyyy")) Then
    If DateValue(tValue) <= DateValue(nwValue) Then
        checkFutureDate = True
    Else
        checkFutureDate = False
End If
Else
    checkFutureDate = False
End If

End Function


Public Function compareDate(str1 As String, str2 As String) As Boolean
    Dim dt1 As Date
    Dim dt2 As Date
    Dim dt3 As Date
    Dim diff As Double
    Dim diff1 As Double
    If (str1 <> "" And str2 <> "") Then
        dt1 = CDate((Format(str1, "DD/MM/YYYY")))
        dt2 = CDate((Format(str2, "DD/MM/YYYY")))
        diff = DateDiff("yyyy", dt1, dt2)
        If (diff >= 10) Then
             diff1 = DateDiff("d", dt1, dt2)
             If (diff1 >= 3652) Then
                compareDate = True
             End If
        Else
             compareDate = False
        End If
    End If
End Function

'checkPINDuplicationAuditorWithSelfPIN function check entered PIN number can not be Duplicate Start
Public Function checkPINDuplicationAuditorWithSelfPIN(ByVal listNameParam As String, ByRef value As Range, Optional ByVal colName As String) As String
Dim rowNumber As String
Dim sheetName As String
TPIN = Range("RetInf.PIN").value
spousePIN = Range("RetInf.SpousePIN").value
AdtrPINS = Range("audit.PINOfAuditorS").value
AdtrPINW = Range("audit.PINOfAuditorW").value

    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name

    Set sName = value.name
    sheetName = Mid(sName, 2, InStrRev(sName, "!") - 2)
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    If colName <> "" Then
        Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
        Set column = Intersect(value, columnRange)
        With Worksheets(sheetName)
            If Not column Is Nothing Then
                For Each c In column.Cells
                    If Trim(c) <> "" Then
                        rangeAddress = Range(listNameParam).Address
                        startRow = Range(listNameParam).row
                        endRow = Mid(rangeAddress, InStrRev(rangeAddress, "$") + 1, Len(rangeAddress))
                            For i = startRow To endRow
                                If i <> c.row Then
                                    If c.value <> "" And Sheet14.Range("A" & (i) & ":A" & i).value <> "" Then
                                        If c.value = Sheet14.Range("A" & (i) & ":A" & i).value Then
                                            collAddress = colName & c.row
                                                If rowNumber <> "" Then
                                                    rowNumber = rowNumber & "," & collAddress
                                                Else
                                                    rowNumber = rowNumber & collAddress
                                                End If
                                        End If

                                    End If
                                End If
                            Next
                    End If
                    If rowNumber = "" Then
                        If TPIN <> "" And Trim(c) <> "" Then
                            If TPIN = Trim(c) Then
                                blnFlag = True
                            End If
                        End If
                        If spousePIN <> "" And Trim(c) <> "" Then
                            If spousePIN = Trim(c) Then
                                blnFlag = True
                            End If
                        End If
                        'added new code for Auditor PIN Start
                        If AdtrPINS <> "" And Trim(c) <> "" Then
                            If AdtrPINS = Trim(c) Then
                                blnFlag = True
                            End If
                        End If
                        If AdtrPINW <> "" And Trim(c) <> "" Then
                            If AdtrPINW = Trim(c) Then
                                blnFlag = True
                            End If
                        End If
                        'added new code for Auditor PIN End
                             If blnFlag Then
                                collAddress = colName & c.row
                                    If rowNumber <> "" Then
                                       rowNumber = rowNumber & "," & collAddress
                                    Else
                                       rowNumber = rowNumber & collAddress
                                    End If
                            
                             End If
                     End If
                Next
            End If
        End With
    Else
        If Trim(value.value) <> "" Then
            selRange = Mid(value.name, InStrRev(value.name, "!") + 1, Len(value.name))

            If (Trim(selRange) = "$B$3") And (value.value) = spousePIN Then
                    rowNumber = rowNumber & value.Address
            Else
                If (Trim(selRange) = "$B$9") And (value.value) = TPIN Then
                        rowNumber = rowNumber & value.Address
                End If
            End If
            'added new code for Auditor PIN Start
            If (Trim(selRange) = "$B$20") And (value.value) = TPIN Or (Trim(selRange) = "$B$20") And (value.value) = spousePIN Then
                    rowNumber = rowNumber & value.Address
            Else
                If (Trim(selRange) = "$C$20") And (value.value) = TPIN Or (Trim(selRange) = "$C$20") And (value.value) = spousePIN Then
                        rowNumber = rowNumber & value.Address
                End If
            End If
            'added new code for Auditor PIN End
        End If
   End If

    sheetName = value.Worksheet.name
    If rowNumber = "" Then
        rowNumber = "NE"
    End If
checkPINDuplicationAuditorWithSelfPIN = rowNumber
End Function
'checkPINDuplicationAuditorWithSelfPIN function check entered PIN number can not be Duplicate End

'compareValuesZeroCheck function check when Net Asset and Total laibility vaule is zero that time share capital value must be grater than Zero Start
Public Function compareValuesZeroCheck(ByRef value1 As Range, ByRef value2 As Range, ByRef value3 As Range) As String
    Dim rowNumber As String
        If value1.value <> "" Or value2.value <> "" Or value3.value <> "" Then
            If value1.value = 0 And value2.value = 0 Then
                If (value3.value = 0) Then
                    rowNumber = value3.Address
                End If
            End If
        End If
    sheetName = value1.Worksheet.name
    If rowNumber = "" Then
        rowNumber = "NE"
    End If
    compareValuesZeroCheck = rowNumber
End Function
'compareValuesZeroCheck function check when Net Asset and Total laibility vaule is zero that time share capital value must be grater than Zero End

Public Function validateInstalmentTaxDate(ByRef value As Range, Optional ByVal colName As String) As String
    Dim sheetName As String
    If Trim(colName) = "" Then
        If Trim(value) <> "" Then
            If TestAlphabet(value.value) = False Then
                rowNumber = value.Address
            End If
        End If
    Else
       
        Dim column As Range
        Dim columnRange As Range
        Dim blnFlag As Boolean
        Dim sName As name
        Set sName = value.name
        sheetName = value.Parent.name
        Dim singleRow As Range
        Worksheets(sheetName).Activate
        Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
        Set column = Intersect(value, columnRange)
        
        With Worksheets(sheetName)
        If Not column Is Nothing Then
            For Each c In column.Cells
                If Trim(c) <> "" Then
                    If TestInstalmentTaxDate(c.value) = False Then
                        collAddress = colName & c.row
                        If rowNumber <> "" Then
                            rowNumber = rowNumber & "," & collAddress
                         Else
                            rowNumber = rowNumber & collAddress
                         End If
                    End If
                End If
            Next
        End If
    End With
    
    End If
    sheetName = value.Worksheet.name
    If rowNumber = "" Then
        rowNumber = "NE"
    End If
    validateInstalmentTaxDate = rowNumber
End Function


Public Function validatePaymentCreditDate(ByRef value As Range, Optional ByVal colName As String) As String
    Dim sheetName As String
    If Trim(colName) = "" Then
        If Trim(value) <> "" Then
            If TestAlphabet(value.value) = False Then
                rowNumber = value.Address
            End If
        End If
    Else
       
        Dim column As Range
        Dim columnRange As Range
        Dim blnFlag As Boolean
        Dim sName As name
        Set sName = value.name
        sheetName = value.Parent.name
        Dim singleRow As Range
        Worksheets(sheetName).Activate
        Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
        Set column = Intersect(value, columnRange)
        
        With Worksheets(sheetName)
        If Not column Is Nothing Then
            For Each c In column.Cells
                If Trim(c) <> "" Then
                    If TestPaymentCreditDate(c.value) = False Then
                        collAddress = colName & c.row
                        If rowNumber <> "" Then
                            rowNumber = rowNumber & "," & collAddress
                         Else
                            rowNumber = rowNumber & collAddress
                         End If
                    End If
                End If
            Next
        End If
    End With
    
    End If
    sheetName = value.Worksheet.name
    If rowNumber = "" Then
        rowNumber = "NE"
    End If
    validatePaymentCreditDate = rowNumber
End Function
' Check instalment tax paid date
Public Function TestInstalmentTaxDate(lstr_date As String) As Boolean

    Dim startDate, endDate, StartDepositDate, DsysDate As String
    
    startDate = CDate(Format((Sheet14.Range("RetInf.RetStartDate").value), "dd/mm/yyyy"))
    endDate = CDate(Format((Sheet14.Range("RetInf.RetEndDate").value), "dd/mm/yyyy"))
    StartDepositDate = CDate(Format((Sheet14.Range("RetInf.DepositStartDate").value), "dd/mm/yyyy"))
    DsysDate = Date
    
    If (TestDate(lstr_date) = False) Then
        TestInstalmentTaxDate = False
    Else
        tSValue = CDate(Format(lstr_date, "dd/mm/yyyy"))
        If (DsysDate <> "" And StartDepositDate <> "") Then
            If (DateValue(tSValue) >= DateValue(StartDepositDate) And DateValue(tSValue) < DateValue(startDate)) Then
                TestInstalmentTaxDate = True
            Else
                TestInstalmentTaxDate = False
            End If
        End If
    End If

End Function

' Check instalment tax paid date
Public Function TestPaymentCreditDate(lstr_date As String) As Boolean

    Dim startDate, endDate, StartDepositDate, DsysDate As String
    
    startDate = CDate(Format((Sheet14.Range("RetInf.RetStartDate").value), "dd/mm/yyyy"))
    endDate = CDate(Format((Sheet14.Range("RetInf.RetEndDate").value), "dd/mm/yyyy"))
    StartDepositDate = CDate(Format((Sheet14.Range("RetInf.DepositStartDate").value), "dd/mm/yyyy"))
    DsysDate = Date
    
    If (TestDate(lstr_date) = False) Then
        TestPaymentCreditDate = False
    Else
        tSValue = CDate(Format(lstr_date, "dd/mm/yyyy"))
        If (DsysDate <> "" And StartDepositDate <> "") Then
            If (DateValue(tSValue) >= DateValue(StartDepositDate) And DateValue(tSValue) <= DateValue(DsysDate)) Then
                TestPaymentCreditDate = True
            Else
                TestPaymentCreditDate = False
            End If
        End If
    End If

End Function




' ########## new code added for SHA256 Hash Code Start #############

Public Function Base64_HMACSHA1(ByVal sTextToHash As String, ByVal sSharedSecretKey As String)

    Dim asc As Object, enc As Object
    Dim TextToHash() As Byte
    Dim SharedSecretKey() As Byte
    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA1")

    TextToHash = asc.Getbytes_4(sTextToHash)
    SharedSecretKey = asc.Getbytes_4(sSharedSecretKey)
    enc.Key = SharedSecretKey

    Dim bytes() As Byte
    bytes = enc.ComputeHash_2((TextToHash))
    Base64_HMACSHA1 = EncodeBase64(bytes)
    Set asc = Nothing
    Set enc = Nothing

End Function

Private Function EncodeBase64(ByRef arrData() As Byte) As String

    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement

    Set objXML = New MSXML2.DOMDocument

    ' byte array to base64
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.Text

    Set objNode = Nothing
    Set objXML = Nothing

End Function



Public Function SHA256(sMessage As String)
     
    Dim clsX As CSHA256
    Set clsX = New CSHA256
     
    SHA256 = clsX.SHATCS256(sMessage)
     
    Set clsX = Nothing
     
End Function



'############# new code added for SHA256 Hash Code End ###########

Public Sub lock_cellsHidden(cursheet As String, listName As String, colName As String, lockValue As Boolean, Optional clearFlag As Boolean)
    Worksheets(cursheet).Activate
    ActiveSheet.Unprotect Password:=Pwd
         startRow = Range(listName).row
         lastRow = Range(listName).Rows.Count + startRow - 1
         blnFlag = False
         For Each r In Range(listName).Rows
             For Each c In r.Cells
                 If c.row > lastRow Then
                     Exit For
                 End If
                 If clearFlag = False And c.Address = "$" & colName & "$" & c.row Then
                     Call lockUnlock_cell_rng_Hidden_without_clearing_contents(cursheet, c.Address, lockValue)
                 ElseIf c.Address = "$" & colName & "$" & c.row Then
                     Call lockUnlock_cell_rng_Hidden(cursheet, c.Address, lockValue)
                 End If
             Next
        Next
    ActiveSheet.Protect Password:=Pwd
 End Sub


Public Sub lockUnlock_cell_rng_Hidden(cursheet As String, rangeName As String, lockCellFlag As Boolean, Optional skipColumnFrom As String)

    Dim cellColor
    'On Error GoTo catch
    activeSheetName = ActiveSheet.name
    protectedStatus = ActiveSheet.ProtectContents
    
    Worksheets(cursheet).Activate
    ActiveSheet.Unprotect Password:=Pwd
    
    startRow = Range(rangeName).row
    lastRow = Range(rangeName).Rows.Count + startRow - 1
    
    startColumn = Range(rangeName).column
    lastColumn = startColumn + Range(rangeName).Columns.Count - 1 ' get actual last column
    
    If lockCellFlag = True Then
        cellColor = RGB(215, 215, 215) 'lock
    ElseIf lockCellFlag = False Then
        cellColor = RGB(255, 255, 255) 'unlock
    End If
    If skipColumnFrom = "" Then
    ' locked whole range
        Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).value = ""
        Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).Locked = lockCellFlag
        
        With Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).Interior
            .Color = cellColor
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
        End With
    Else
        'skip columns and define new range
        lastColumn = skipColumnFrom - 1
        Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).Locked = lockCellFlag
        Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).value = ""
            
        
        With Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).Interior
           .Color = cellColor
           .Pattern = xlSolid
           .PatternColorIndex = xlAutomatic
        End With
        ActiveSheet.Cells(3, 1).Select
    End If
    
catch:
    If Err.Description <> "" Then
    '    MsgBox Err.Description
    End If
    
    ActiveSheet.Protect Password:=Pwd
    
    Worksheets(activeSheetName).Activate
    
    If protectedStatus Then
        ActiveSheet.Protect Password:=Pwd
    Else
        ActiveSheet.Unprotect Password:=Pwd
    End If
    
End Sub



'validateRespctiveSection Function using for check entered PIN value and selected Respective Section valid or not Start
Function validateRespctiveSectionS(ByRef value As Range, Optional ByVal colName As String) As String
    
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    blnFlag = True
    Dim sName As String
    Dim statusFlag As Boolean
    statusFlag = False
    Dim singleRow As Range
    Dim rowNumbers As String
    sName = value.Cells.Parent.name
    'MsgBox flagAdd.row
    Dim counter As Integer
    Dim counterF As Integer
    counter = 0
    counterF = 0
    'Dim counter As Integer
    
    Dim sheetName As String
    If Trim(colName) = "" Then
        If value <> 0 Then
            rowNumbers = "NE"
        Else
           rowNumbers = value.Address
        End If
    Else
    Set columnRange = Worksheets(sName).Range(colName & "1", colName & Worksheets(sName).UsedRange.Rows.Count)
        Set column = Intersect(value, columnRange)
        With Worksheets(sName)
        If Not column Is Nothing Then
            For Each c In column.Cells
                On Error Resume Next
                     If (c.value <> "") Then
                        counter = 0
                        For Each r In Sheet3.Range("PINofEmployerS")
                            If (r.value <> "") Then
                                If (c.value = r.value) Then
                                    counter = 0
                                    Exit For
                                Else
                                   counter = 1
                                End If
                            Else
                                'do nothing
                            End If
                        Next
                    Else
                        'do nothing
                    End If
                    
                    If (counter > 0) Then
                        Set singleRow = Intersect(value, Worksheets(sName).Range(.Cells(c.row, 1), .Cells(c.row, Worksheets(sName).UsedRange.Columns.Count)))
                        collAddress = colName & singleRow.Cells.row
                         If rowNumbers <> "" Then
                            rowNumbers = rowNumbers & "," & collAddress
                         Else
                            rowNumbers = rowNumbers & collAddress
                         End If
                        counter = 0
                    End If
                Next
            End If
        End With
    End If
    sheetName = value.Worksheet.name
    If rowNumbers = "" Then
        rowNumbers = "NE"
    End If
    validateRespctiveSectionS = rowNumbers
End Function
'validateRespctiveSection Function using for check entered PIN value and selected Respective Section valid or not End

'validateRespctiveSection Function using for check entered PIN value and selected Respective Section valid or not Start
Function validateRespctiveSectionW(ByRef value As Range, Optional ByVal colName As String) As String

    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    blnFlag = True
    Dim sName As String
    Dim statusFlag As Boolean
    statusFlag = False
    Dim singleRow As Range
    Dim rowNumbers As String
    sName = value.Cells.Parent.name
    'MsgBox flagAdd.row
    Dim counter As Integer
    Dim counterF As Integer
    counter = 0
    counterF = 0
    'Dim counter As Integer
    
    Dim sheetName As String
    If Trim(colName) = "" Then
        If value <> 0 Then
            rowNumbers = "NE"
        Else
           rowNumbers = value.Address
        End If
    Else
    Set columnRange = Worksheets(sName).Range(colName & "1", colName & Worksheets(sName).UsedRange.Rows.Count)
        Set column = Intersect(value, columnRange)
        With Worksheets(sName)
        If Not column Is Nothing Then
            For Each c In column.Cells
                On Error Resume Next
                     If (c.value <> "") Then
                        counter = 0
                        For Each r In Sheet3.Range("PINofEmployerW")
                            If (r.value <> "") Then
                                If (c.value = r.value) Then
                                    counter = 0
                                    Exit For
                                Else
                                   counter = 1
                                End If
                            Else
                                'do nothing
                            End If
                        Next
                    Else
                        'do nothing
                    End If
                    
                    If (counter > 0) Then
                        Set singleRow = Intersect(value, Worksheets(sName).Range(.Cells(c.row, 1), .Cells(c.row, Worksheets(sName).UsedRange.Columns.Count)))
                        collAddress = colName & singleRow.Cells.row
                         If rowNumbers <> "" Then
                            rowNumbers = rowNumbers & "," & collAddress
                         Else
                            rowNumbers = rowNumbers & collAddress
                         End If
                        counter = 0
                    End If
                Next
            End If
        End With
    End If
    sheetName = value.Worksheet.name
    If rowNumbers = "" Then
        rowNumbers = "NE"
    End If
    validateRespctiveSectionW = rowNumbers
End Function
'validateRespctiveSection Function using for check entered PIN value and selected Respective Section valid or not End


Public Function validateRatioS(ByRef value As Range, Optional ByVal colName As String) As String

    Dim rowNumber As String
    Dim sheetName As String
    
    If Trim(colName) = "" Then
        If TestRatioS(value.value) = False Then
            rowNumber = value.Address
         End If
    Else
     
        Dim column As Range
        Dim columnRange As Range
        Dim blnFlag As Boolean
        Dim sName As name
        Set sName = value.name
    '    sheetName = Mid(sName, 2, InStrRev(sName, "!") - 2)
        sheetName = value.Parent.name
        Dim singleRow As Range
        Worksheets(sheetName).Activate
        Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
        Set column = Intersect(value, columnRange)
        
        With Worksheets(sheetName)
        If Not column Is Nothing Then
            For Each c In column.Cells
                If TestRatioS(c.value) = False Then
                    'rowNumber = rowNumber & C.Address & ","
                    collAddress = colName & c.row
                        If rowNumber <> "" Then
                            rowNumber = rowNumber & "," & collAddress
                         Else
                            rowNumber = rowNumber & collAddress
                         End If
                End If
            Next
        End If
    End With
    
    End If
    sheetName = value.Worksheet.name
    If rowNumber = "" Then
        rowNumber = "NE"
    End If
    validateRatioS = rowNumber
End Function

Public Function TestRatioS(ByVal value As String) As Boolean

    'Newly added Code...
    startRow = Worksheets("G_Partnership_Income").Range("ProfitShare.ListS").row
    lastRow = startRow + Worksheets("G_Partnership_Income").Range("ProfitShare.ListS").Rows.Count - 1
    cntRatioB = 0
    cntRatioI = 0
    cntRatioF = 0
    cntRatioR = 0
    cntRatioC = 0
    cntRatioOS = 0
    
       
    For Each c In Sheet33.Range("C" & (startRow) & ":C" & (lastRow)).Cells
        For Each r In Sheet33.Range("D" & (c.row) & ":D" & (c.row)).Cells
          If c.row > lastRow Then
               Exit For
          End If

          If r.row > lastRow Then
               Exit For
          End If
          'Business.
          If (Sheet33.Cells(c.row, c.column) = "Business" And c.column = 3 And value = "Business") Then
             cntRatioB = cntRatioB + r.value
          Else
             Exit For
          End If
        Next
    Next
        
    If (cntRatioB <> 100 And cntRatioB <> 0) Then
        TestRatioS = False
        Exit Function
    Else
        'Interest.
        For Each c In Sheet33.Range("C" & (startRow) & ":C" & (lastRow)).Cells
            For Each r In Sheet33.Range("D" & (c.row) & ":D" & (c.row)).Cells
                If (Sheet33.Cells(c.row, c.column) = "Interest" And c.column = 3) Then
                    cntRatioI = cntRatioI + r.value
                Else
                    Exit For
                End If
            Next
        Next
        
        If (cntRatioI <> 100 And cntRatioI <> 0) Then
            TestRatioS = False
            Exit Function
        Else
            For Each c In Sheet33.Range("C" & (startRow) & ":C" & (lastRow)).Cells
                For Each r In Sheet33.Range("D" & (c.row) & ":D" & (c.row)).Cells
                    'Farming.
                    If (Sheet33.Cells(c.row, c.column) = "Farming" And c.column = 3 And value = "Farming") Then
                       cntRatioF = cntRatioF + r.value
                    Else
                       Exit For
                    End If
                Next
            Next
        
            If (cntRatioF <> 100 And cntRatioF <> 0) Then
                TestRatioS = False
                    Exit Function
            Else
                For Each c In Sheet33.Range("C" & (startRow) & ":C" & (lastRow)).Cells
                    For Each r In Sheet33.Range("D" & (c.row) & ":D" & (c.row)).Cells
                      'Rental.
                        If (Sheet33.Cells(c.row, c.column) = "Rental" And c.column = 3 And value = "Rental") Then
                            cntRatioR = cntRatioR + r.value
                        Else
                            Exit For
                        End If
                    Next
                Next
        
                If (cntRatioR <> 100 And cntRatioR <> 0) Then
                    TestRatioS = False
                    Exit Function
                Else
                    For Each c In Sheet33.Range("C" & (startRow) & ":C" & (lastRow)).Cells
                        For Each r In Sheet33.Range("D" & (c.row) & ":D" & (c.row)).Cells
                            'Commission.
                            If (Sheet33.Cells(c.row, c.column) = "Commission" And c.column = 3 And value = "Commission") Then
                                cntRatioC = cntRatioC + r.value
                            Else
                                Exit For
                            End If
                        Next
                    Next
                    If (cntRatioC <> 100 And cntRatioC <> 0) Then
                        TestRatioS = False
                        Exit Function
                    Else
                        For Each c In Sheet33.Range("C" & (startRow) & ":C" & (lastRow)).Cells
                            For Each r In Sheet33.Range("D" & (c.row) & ":D" & (c.row)).Cells
                                'Other Source.
                                If (Sheet33.Cells(c.row, c.column) = "Other Source" And c.column = 3 And value = "Other Source") Then
                                    cntRatioOS = cntRatioOS + r.value
                                Else
                                    Exit For
                                End If
                            Next
                        Next
                        If (cntRatioOS <> 100 And cntRatioOS <> 0) Then
                            TestRatioS = False
                            Exit Function
                        Else
                            TestRatioS = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Public Function validateRatioW(ByRef value As Range, Optional ByVal colName As String) As String

    Dim rowNumber As String
    Dim sheetName As String
    
    If Trim(colName) = "" Then
        If TestRatioW(value.value) = False Then
            rowNumber = value.Address
         End If
    Else
        Dim column As Range
        Dim columnRange As Range
        Dim blnFlag As Boolean
        Dim sName As name
        Set sName = value.name
    '    sheetName = Mid(sName, 2, InStrRev(sName, "!") - 2)
        sheetName = value.Parent.name
        Dim singleRow As Range
        Worksheets(sheetName).Activate
        Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
        Set column = Intersect(value, columnRange)
        
        With Worksheets(sheetName)
            If Not column Is Nothing Then
                For Each c In column.Cells
                    If TestRatioW(c.value) = False Then
                        'rowNumber = rowNumber & C.Address & ","
                        collAddress = colName & c.row
                        If rowNumber <> "" Then
                            rowNumber = rowNumber & "," & collAddress
                         Else
                            rowNumber = rowNumber & collAddress
                         End If
                    End If
                Next
            End If
        End With
    End If
    sheetName = value.Worksheet.name
    If rowNumber = "" Then
        rowNumber = "NE"
    End If
    validateRatioW = rowNumber
End Function

Public Function TestRatioW(ByVal value As String) As Boolean
    'Newly added Code...
    startRow = Worksheets("G_Partnership_Income").Range("ProfitShare.ListW").row
    lastRow = startRow + Worksheets("G_Partnership_Income").Range("ProfitShare.ListW").Rows.Count - 1
    cntRatioB = 0
    cntRatioI = 0
    cntRatioF = 0
    cntRatioR = 0
    cntRatioC = 0
    cntRatioOS = 0
    
       
    For Each c In Sheet33.Range("C" & (startRow) & ":C" & (lastRow)).Cells
        For Each r In Sheet33.Range("D" & (c.row) & ":D" & (c.row)).Cells
            If c.row > lastRow Then
                 Exit For
            End If

            If r.row > lastRow Then
                 Exit For
            End If
            'Business.
            If (Sheet33.Cells(c.row, c.column) = "Business" And c.column = 3 And value = "Business") Then
               cntRatioB = cntRatioB + r.value
            Else
               Exit For
            End If
        Next
    Next
        
    If (cntRatioB <> 100 And cntRatioB <> 0) Then
        TestRatioW = False
        Exit Function
    Else
        ' Exit For
        'End If
    'Interest.
        For Each c In Sheet33.Range("C" & (startRow) & ":C" & (lastRow)).Cells
            For Each r In Sheet33.Range("D" & (c.row) & ":D" & (c.row)).Cells
                If (Sheet33.Cells(c.row, c.column) = "Interest" And c.column = 3) Then
                    cntRatioI = cntRatioI + r.value
                Else
                    Exit For
                End If
            Next
        Next
        
        If (cntRatioI <> 100 And cntRatioI <> 0) Then
             TestRatioW = False
             Exit Function
        Else
            For Each c In Sheet33.Range("C" & (startRow) & ":C" & (lastRow)).Cells
                For Each r In Sheet33.Range("D" & (c.row) & ":D" & (c.row)).Cells
                'Farming.
                    If (Sheet33.Cells(c.row, c.column) = "Farming" And c.column = 3 And value = "Farming") Then
                        cntRatioF = cntRatioF + r.value
                    Else
                        Exit For
                    End If
                Next
            Next
            If (cntRatioF <> 100 And cntRatioF <> 0) Then
                TestRatioW = False
                Exit Function
            Else
                For Each c In Sheet33.Range("C" & (startRow) & ":C" & (lastRow)).Cells
                    For Each r In Sheet33.Range("D" & (c.row) & ":D" & (c.row)).Cells
                        'Rental.
                        If (Sheet33.Cells(c.row, c.column) = "Rental" And c.column = 3 And value = "Rental") Then
                            cntRatioR = cntRatioR + r.value
                        Else
                            Exit For
                        End If
                    Next
                Next
                If (cntRatioR <> 100 And cntRatioR <> 0) Then
                    TestRatioW = False
                    Exit Function
                Else
                    For Each c In Sheet33.Range("C" & (startRow) & ":C" & (lastRow)).Cells
                        For Each r In Sheet33.Range("D" & (c.row) & ":D" & (c.row)).Cells
                            'Commission.
                            If (Sheet33.Cells(c.row, c.column) = "Commission" And c.column = 3 And value = "Commission") Then
                                cntRatioC = cntRatioC + r.value
                            Else
                                Exit For
                            End If
                        Next
                    Next
                    If (cntRatioC <> 100 And cntRatioC <> 0) Then
                        TestRatioW = False
                        Exit Function
                    Else
                        For Each c In Sheet33.Range("C" & (startRow) & ":C" & (lastRow)).Cells
                            For Each r In Sheet33.Range("D" & (c.row) & ":D" & (c.row)).Cells
                                'Other Source.
                                If (Sheet33.Cells(c.row, c.column) = "Other Source" And c.column = 3 And value = "Other Source") Then
                                    cntRatioOS = cntRatioOS + r.value
                                Else
                                    Exit For
                                End If
                            Next
                        Next
                        If (cntRatioOS <> 100 And cntRatioOS <> 0) Then
                            TestRatioW = False
                            Exit Function
                        Else
                            TestRatioW = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function


Public Function validateSameIncomeRatioS(ByRef value As Range, ByVal colName As String) As String

    Dim rowNumber As String
    Dim sheetName As String
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    
    Set sName = value.name
    sheetName = value.Parent.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)
    
    
    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If checkPinIncomeTypeRatioS(c) = False Then
                collAddress = colName & c.row
                    If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
            End If
        Next
    End If
    End With

'End If
sheetName = value.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
validateSameIncomeRatioS = rowNumber
End Function



Public Function checkPinIncomeTypeRatioS(ByVal value As Range) As Boolean

Dim startRow As Long, endRow As Long
Dim testTypOfInc As String
Dim TypeOfIncome As String
Dim testRatioForInc As Double
Dim pinOfPartner As String

checkPinIncomeTypeRatioS = True

startRow = Worksheets("G_Partnership_Income").Range("ProfitShare.ListS").row
endRow = startRow + Worksheets("G_Partnership_Income").Range("ProfitShare.ListS").Rows.Count - 1
selectedRow = value.row
            pinOfPartnerJ = Sheet33.Cells(selectedRow, 1).value
            testTypOfIncJ = Sheet33.Cells(selectedRow, 3).value
            testRatioForIncJ = Sheet33.Cells(selectedRow, 4).value
For i = selectedRow + 1 To endRow
     pinOfPartnerI = Sheet33.Cells(i, 1).value
     testTypOfIncI = Sheet33.Cells(i, 3).value
     testRatioForIncI = Sheet33.Cells(i, 4).value
      
                If pinOfPartnerI = pinOfPartnerJ And pinOfPartnerJ <> "" Then
                    If testRatioForIncI <> testRatioForIncJ Then
                        checkPinIncomeTypeRatioS = False
                        Exit For
                    End If
                End If
            
Next

End Function

Public Function validateSameIncomeRatioW(ByRef value As Range, ByVal colName As String) As String

Dim rowNumber As String
Dim sheetName As String

'If Trim(colName) = "" Then
'    If checkPinIncomeTypeRatio(value.value) = False Then
'        rowNumber = value.Address
'     End If
'Else
 
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    sheetName = value.Parent.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)
    
    
    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If checkPinIncomeTypeRatioW(c) = False Then
                collAddress = colName & c.row
                    If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
            End If
        Next
    End If
    End With

'End If
sheetName = value.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
validateSameIncomeRatioW = rowNumber
End Function



Public Function checkPinIncomeTypeRatioW(ByVal value As Range) As Boolean

Dim startRow As Long, endRow As Long
Dim testTypOfInc As String
Dim TypeOfIncome As String
Dim testRatioForInc As Double
Dim pinOfPartner As String

checkPinIncomeTypeRatioW = True

startRow = Worksheets("G_Partnership_Income").Range("ProfitShare.ListW").row
endRow = startRow + Worksheets("G_Partnership_Income").Range("ProfitShare.ListW").Rows.Count - 1
selectedRow = value.row
            pinOfPartnerJ = Sheet33.Cells(selectedRow, 1).value
            testTypOfIncJ = Sheet33.Cells(selectedRow, 3).value
            testRatioForIncJ = Sheet33.Cells(selectedRow, 4).value
For i = selectedRow + 1 To endRow
     pinOfPartnerI = Sheet33.Cells(i, 1).value
     testTypOfIncI = Sheet33.Cells(i, 3).value
     testRatioForIncI = Sheet33.Cells(i, 4).value
      
                If pinOfPartnerI = pinOfPartnerJ And pinOfPartnerJ <> "" Then
                    If testRatioForIncI <> testRatioForIncJ Then
                        checkPinIncomeTypeRatioW = False
                        Exit For
                    End If
                End If
            
Next

End Function


'TestOtherPINNonId Function check entered PIN proper format or not and return Boolean Result Start
Public Function TestOtherPINNonId(lstr_check As String) As Boolean
'allowed characters A to Z, a to z And Blank
Dim i As Integer
Dim stlen As Integer
Dim alphabates As String
Dim numbers As String
TestOtherPINNonId = True

numbers = "0123456789"
alphabates = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"

stlen = Len(lstr_check)

For i = 1 To stlen
    
    If i = 1 Then
        If Mid(lstr_check, 1, 1) <> "P" And Mid(lstr_check, 1, 1) <> "p" Then
            TestOtherPINNonId = False
            Exit Function
        End If
    ElseIf i = 11 Then
        If InStr(alphabates, Mid(lstr_check, i, 1)) = 0 Then
            TestOtherPINNonId = False
            Exit Function
        End If
    Else
        If InStr(numbers, Mid(lstr_check, i, 1)) = 0 Then
            TestOtherPINNonId = False
            Exit Function
        End If
    End If
Next i
End Function
'TestOtherPINNonId Function check entered PIN proper format or not and return Boolean Result End

Function isyLeapYear(ByVal Date1 As Date) As Boolean
    yy = year(Date1)
    isyLeapYear = IIf(yy Mod 100 = 0, yy Mod 400 = 0, yy Mod 4 = 0)
End Function

Public Function validateAlphabetOnlyWithSpaceDot(ByRef value As Range, Optional ByVal colName As String) As String

Dim rowNumber As String
Dim sheetName As String
sheetName = value.Worksheet.name
If Trim(colName) = "" Then
    If TestAlphabetWithSpaceDot(value.value) = False Then
        rowNumber = value.Address
     End If
Else

    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)

    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If TestAlphabetWithSpaceDot(c.value) = False Then
                'rowNumber = rowNumber & C.Address & ","
                collAddress = colName & c.row
                If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
            End If
        Next
    End If
End With

End If
sheetName = value.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
validateAlphabetOnlyWithSpaceDot = rowNumber

End Function


Public Function TestAlphabetWithSpaceDot(lstr_check As String) As Boolean
'allowed characters A to Z, a to z And Blank
Dim i As Integer
Dim ia As Integer
Dim ina As Integer
Dim stlen As Integer

stlen = Len(lstr_check)
ia = 0
ina = 0
For i = 1 To stlen
    Select Case (Mid(lstr_check, i, 1))
        Case "A" To "Z"          'A to Z
            ia = ia + 1
        Case "a" To "z"          'a to z
            ia = ia + 1
        Case " ", "."               ' Blank
            ia = ia + 1
        Case Else
            ina = ina + 1
    End Select
Next i
If ina = 0 Then
    TestAlphabetWithSpaceDot = True
Else
    TestAlphabetWithSpaceDot = False
End If

End Function


Public Function TestDateLimit(ByVal value As String) As Boolean
Dim mm As Integer
Dim yr As Integer
Dim startDate As String
Dim endDate As String
Dim sheet As Worksheet

Dim tValueC As String
Dim tValueL As String
tValueL = "1/1/2001"
                    

If TestDate(value) = True Then
    tValueC = CDate(Format(value, "dd/mm/yyyy"))
    If (DateValue(tValueL) <= DateValue(value)) Then
          TestDateLimit = True
    Else
          TestDateLimit = False
    End If
Else
        TestDateLimit = False
End If
End Function

Public Function validateDateBtnRtnPeriodSysDate(ByRef value As Range, Optional ByVal colName As String) As String
Dim sheetName As String
Dim rowNumber As String
sheetName = value.Worksheet.name

If Trim(colName) = "" Then
    If TestDatebtnRtnPeriodSysDate(value.value) = False Then
        rowNumber = value.Address
    End If
Else
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)
    
    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If c.value <> "" And TestDatebtnRtnPeriodSysDate(c.value) = False Then
                    collAddress = colName & c.row
                    If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
            End If
        Next
    End If
End With

End If
sheetName = value.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
validateDateBtnRtnPeriodSysDate = rowNumber
End Function

Public Function TestDatebtnRtnPeriodSysDate(ByVal value As String) As Boolean
Dim mm As Integer
Dim yr As Integer
Dim startDate As String
Dim endDate As String
Dim tValue As String
Dim sheet As Worksheet

startDate = Range("RetInf.RetStartDate").value
endDate = Range("RetInf.RetEndDate").value
DsysDate = Date

If TestDate(value) = True And TestDate(startDate) = True And TestDate(endDate) = True Then
    startDate = CDate(Format(Sheet14.Range("RetInf.RetStartDate").value, "dd/mm/yyyy"))
    endDate = CDate(Format(Sheet14.Range("RetInf.RetEndDate").value, "dd/mm/yyyy"))
    tValue = CDate(Format(value, "dd/mm/yyyy"))
        If (DateValue(tValue) > DateValue(endDate) And DateValue(tValue) <= DateValue(DsysDate)) Then
            TestDatebtnRtnPeriodSysDate = True
        Else
            TestDatebtnRtnPeriodSysDate = False
        End If
Else
        TestDatebtnRtnPeriodSysDate = False
End If
End Function


'****************************Function added for checking Adv Tax Payable or not
Public Function validateAdvTax1(ByRef totAdvPmtVal As Range, ByRef advPmt As Range, ByRef selfAssPmt As Range) As String
Dim sheetName As String
Dim rowNumbers As String
Dim taxDue As Double
taxDue = Round(Sheet20.Range("TaxComp.TotTaxableIncomeLessRlfS").value - Sheet20.Range("TaxComp.PayeDedListSTO").value - Sheet20.Range("TaxComp.EstateTrustListTOS").value - Sheet20.Range("TaxComp.InstallmentTaxListSTO").value - Sheet20.Range("TaxComp.WithHoldingListSTO").value - Sheet20.Range("TaxComp.VehicleAdvTaxPaidListSTO").value - Sheet20.Range("TaxComp.IncomeDTACreditsS").value, 2)
If taxDue <= 0 And totAdvPmtVal.value > 0 Then
        rowNumbers = totAdvPmtVal.Address
End If
If rowNumbers = "" Then
    rowNumbers = "NE"
End If
validateAdvTax1 = rowNumbers
End Function


Public Function validateAdvTax2(ByRef totAdvPmtVal As Range, ByRef advPmt As Range, ByRef selfAssPmt As Range) As String
Dim sheetName As String
Dim rowNumbers As String
Dim taxDue As Double
taxDue = Round(Sheet20.Range("TaxComp.TotTaxableIncomeLessRlfS").value - Sheet20.Range("TaxComp.PayeDedListSTO").value - Sheet20.Range("TaxComp.EstateTrustListTOS").value - Sheet20.Range("TaxComp.InstallmentTaxListSTO").value - Sheet20.Range("TaxComp.WithHoldingListSTO").value - Sheet20.Range("TaxComp.VehicleAdvTaxPaidListSTO").value - Sheet20.Range("TaxComp.IncomeDTACreditsS").value, 2)
If (taxDue > 0 And (taxDue - totAdvPmtVal.value < 0) And advPmt.value > 0 And selfAssPmt.value > 0) Then
     rowNumbers = totAdvPmtVal.Address
End If
If rowNumbers = "" Then
    rowNumbers = "NE"
End If
validateAdvTax2 = rowNumbers
End Function
Public Function validateAdvTax4(ByRef totAdvPmtVal As Range, ByRef advPmt As Range, ByRef selfAssPmt As Range) As String
Dim sheetName As String
Dim rowNumbers As String
Dim taxDue As Double
taxDue = Round(Sheet20.Range("TaxComp.TotTaxableIncomeLessRlfS").value - Sheet20.Range("TaxComp.PayeDedListSTO").value - Sheet20.Range("TaxComp.EstateTrustListTOS").value - Sheet20.Range("TaxComp.InstallmentTaxListSTO").value - Sheet20.Range("TaxComp.WithHoldingListSTO").value - Sheet20.Range("TaxComp.VehicleAdvTaxPaidListSTO").value - Sheet20.Range("TaxComp.IncomeDTACreditsS").value, 2)
If taxDue > 0 And (taxDue - advPmt.value < 0 And selfAssPmt.value = 0) Then
      rowNumbers = advPmt.Address
 End If
If rowNumbers = "" Then
    rowNumbers = "NE"
End If
validateAdvTax4 = rowNumbers
End Function
Public Function validateAdvTax3(ByRef totAdvPmtVal As Range, ByRef advPmt As Range, ByRef selfAssPmt As Range) As String
Dim sheetName As String
Dim rowNumbers As String
Dim taxDue As Double
taxDue = Round(Sheet20.Range("TaxComp.TotTaxableIncomeLessRlfS").value - Sheet20.Range("TaxComp.PayeDedListSTO").value - Sheet20.Range("TaxComp.EstateTrustListTOS").value - Sheet20.Range("TaxComp.InstallmentTaxListSTO").value - Sheet20.Range("TaxComp.WithHoldingListSTO").value - Sheet20.Range("TaxComp.VehicleAdvTaxPaidListSTO").value - Sheet20.Range("TaxComp.IncomeDTACreditsS").value, 2)
If taxDue > 0 And (taxDue - selfAssPmt.value < 0 And advPmt.value = 0) Then
         rowNumbers = selfAssPmt.Address
End If
If rowNumbers = "" Then
    rowNumbers = "NE"
End If
validateAdvTax3 = rowNumbers
End Function

Public Function validateAdvTax1W(ByRef totAdvPmtVal As Range, ByRef advPmt As Range, ByRef selfAssPmt As Range) As String
Dim sheetName As String
Dim rowNumbers As String
Dim taxDue As Double
taxDue = Round(Sheet20.Range("TaxComp.TotTaxableIncomeLessRlfW").value - Sheet20.Range("TaxComp.PayeDedListWTO").value - Sheet20.Range("TaxComp.EstateTrustListTOW").value - Sheet20.Range("TaxComp.InstallmentTaxListWTO").value - Sheet20.Range("TaxComp.WithHoldingListWTO").value - Sheet20.Range("TaxComp.VehicleAdvTaxPaidListWTO").value - Sheet20.Range("TaxComp.IncomeDTACreditsW").value, 2)
If taxDue <= 0 And totAdvPmtVal.value > 0 Then
        rowNumbers = totAdvPmtVal.Address
End If
If rowNumbers = "" Then
    rowNumbers = "NE"
End If
validateAdvTax1W = rowNumbers
End Function

Public Function validateAdvTax2W(ByRef totAdvPmtVal As Range, ByRef advPmt As Range, ByRef selfAssPmt As Range) As String
Dim sheetName As String
Dim rowNumbers As String
Dim taxDue As Double
taxDue = Round(Sheet20.Range("TaxComp.TotTaxableIncomeLessRlfW").value - Sheet20.Range("TaxComp.PayeDedListWTO").value - Sheet20.Range("TaxComp.EstateTrustListTOW").value - Sheet20.Range("TaxComp.InstallmentTaxListWTO").value - Sheet20.Range("TaxComp.WithHoldingListWTO").value - Sheet20.Range("TaxComp.VehicleAdvTaxPaidListWTO").value - Sheet20.Range("TaxComp.IncomeDTACreditsW").value, 2)
If (taxDue > 0 And (taxDue - totAdvPmtVal.value < 0) And advPmt.value > 0 And selfAssPmt.value > 0) Then
     rowNumbers = totAdvPmtVal.Address
End If
If rowNumbers = "" Then
    rowNumbers = "NE"
End If
validateAdvTax2W = rowNumbers
End Function
Public Function validateAdvTax4W(ByRef totAdvPmtVal As Range, ByRef advPmt As Range, ByRef selfAssPmt As Range) As String
Dim sheetName As String
Dim rowNumbers As String
Dim taxDue As Double
taxDue = Round(Sheet20.Range("TaxComp.TotTaxableIncomeLessRlfW").value - Sheet20.Range("TaxComp.PayeDedListWTO").value - Sheet20.Range("TaxComp.EstateTrustListTOW").value - Sheet20.Range("TaxComp.InstallmentTaxListWTO").value - Sheet20.Range("TaxComp.WithHoldingListWTO").value - Sheet20.Range("TaxComp.VehicleAdvTaxPaidListWTO").value - Sheet20.Range("TaxComp.IncomeDTACreditsW").value, 2)
If taxDue > 0 And (taxDue - advPmt.value < 0 And selfAssPmt.value = 0) Then
      rowNumbers = advPmt.Address
 End If
If rowNumbers = "" Then
    rowNumbers = "NE"
End If
validateAdvTax4W = rowNumbers
End Function
Public Function validateAdvTax3W(ByRef totAdvPmtVal As Range, ByRef advPmt As Range, ByRef selfAssPmt As Range) As String
Dim sheetName As String
Dim rowNumbers As String
Dim taxDue As Double
taxDue = Round(Sheet20.Range("TaxComp.TotTaxableIncomeLessRlfW").value - Sheet20.Range("TaxComp.PayeDedListWTO").value - Sheet20.Range("TaxComp.EstateTrustListTOW").value - Sheet20.Range("TaxComp.InstallmentTaxListWTO").value - Sheet20.Range("TaxComp.WithHoldingListWTO").value - Sheet20.Range("TaxComp.VehicleAdvTaxPaidListWTO").value - Sheet20.Range("TaxComp.IncomeDTACreditsW").value, 2)
If taxDue > 0 And (taxDue - selfAssPmt.value < 0 And advPmt.value = 0) Then
         rowNumbers = selfAssPmt.Address
End If
If rowNumbers = "" Then
    rowNumbers = "NE"
End If
validateAdvTax3W = rowNumbers
End Function


'validateAlphaNumericOnly Function Check entered value AlphaNumericOnly or not Start
Public Function validateAlphaNumericSpaceOnly(ByRef value As Range, Optional ByVal colName As String) As String
Dim rowNumber As String
Dim sheetName As String

If Trim(colName) = "" Then
    If TestAlphanumericOnly(value.value) = False Then
       rowNumber = value.Address
     End If
Else

    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    sheetName = value.Parent.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)

    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If TestAlphanumericOnly(c.value) = False Then
                collAddress = colName & c.row
                If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
            End If
        Next
    End If
End With

End If
sheetName = value.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
validateAlphaNumericSpaceOnly = rowNumber
End Function
'validateAlphaNumericOnly Function Check entered value AlphaNumericOnly or not End

'Check alphaNumeric Field,Allow only AlphaNumeric Value Start
Public Function TestAlphanumericSpaceOnly(lstr_check As String) As Boolean
'allowed characters A to Z, a to z And Blank
Dim i As Integer
Dim ia As Integer
Dim ina As Integer
Dim stlen As Integer

stlen = Len(lstr_check)
ia = 0
ina = 0
For i = 1 To stlen
    Select Case (Mid(lstr_check, i, 1))
        Case "A" To "Z"          'A to Z
            ia = ia + 1
        Case "a" To "z"          'a to z
            ia = ia + 1
        Case "0" To "9"          '0 to 9
        ia = ia + 1
        Case " "          'Space
        ia = ia + 1
        Case Else
            ina = ina + 1
    End Select
Next i
If ina = 0 Then
    TestAlphanumericSpaceOnly = True
Else
    TestAlphanumericSpaceOnly = False
End If

End Function
'Check alphaNumeric Field,Allow only AlphaNumeric Value End

Public Function validatePRNNumeric(ByRef value As Range, Optional ByVal colName As String) As String
    Dim sheetName As String
    Dim rowNumber As String
    
    
    If Trim(colName) = "" Then
        If TestPRNNumber(value.value) = False Then
         rowNumber = value.Address
        End If
    Else
        Dim column As Range
        Dim columnRange As Range
        Dim blnFlag As Boolean
        Dim sName As name
        Set sName = value.name
        sheetName = Mid(sName, 2, InStrRev(sName, "!") - 2)
        Dim singleRow As Range
        Worksheets(sheetName).Activate
        Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
        Set column = Intersect(value, columnRange)
        
        With Worksheets(sheetName)
        If Not column Is Nothing Then
            For Each c In column.Cells
                If TestPRNNumber(c.value) = False Then
                    'rowNumber = rowNumber & C.Address & ","
                    collAddress = colName & c.row
                    If rowNumber <> "" Then
                            rowNumber = rowNumber & "," & collAddress
                         Else
                            rowNumber = rowNumber & collAddress
                         End If
                End If
            Next
        End If
    End With
    
    End If
    sheetName = value.Worksheet.name
    If rowNumber = "" Then
        rowNumber = "NE"
    End If
    validatePRNNumeric = rowNumber
End Function

Public Function TestPRNNumber(lstr_check As String) As Boolean
'allowed characters 0 to 9 and .
Dim i As Integer
Dim ia As Integer
Dim ina As Integer
Dim stlen As Integer

stlen = Len(lstr_check)
ia = 0
ina = 0
For i = 1 To stlen
    Select Case (Mid(lstr_check, i, 1))
        Case "0" To "9"          '0 to 9
        ia = ia + 1
        Case "-"
            ia = ia + 1
        Case Else
            ina = ina + 1
    End Select
Next i
If ina = 0 Then
    If InStr(1, lstr_check, ".") = 0 Then
        TestPRNNumber = True
    Else
        TestPRNNumber = True
    End If
Else
    TestPRNNumber = False
End If

End Function

Public Function TestAllZeroNumber(lstr_check As String) As Boolean

TestAllZeroNumber = False
Dim i As Integer
Dim stlen As Integer
Dim alphabates As String
Dim char As String
Dim numbers As String
Dim icnt As Integer
Dim iFlag As Boolean
icnt = 0
iFlag = True
numbers = "0123456789"
chars = "-"
If (lstr_check <> "") Then
    Dim new_lstr_check As String
    new_lstr_check = Replace(lstr_check, "-", "0", 1)
    stlen = Len(new_lstr_check)
    For i = 1 To stlen
        If (Mid(new_lstr_check, i, 1) <> "0") Then
            iFlag = True
            icnt = icnt + 1
        End If
    Next i
    
    If (icnt >= 1) Then
        TestAllZeroNumber = True
    End If
Else
    TestAllZeroNumber = True
End If

End Function


Public Function validatePRNAllzero(ByRef value As Range, Optional ByVal colName As String) As String

Dim sheetName As String
Dim rowNumber As String


If Trim(colName) = "" Then
    If TestAllZeroNumber(value.value) = False Then
     rowNumber = value.row
    End If
Else
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    sheetName = value.Parent.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)
    
    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If TestAllZeroNumber(c.value) = False Then
                collAddress = colName & c.row
                If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
            End If
        Next
    End If
End With

End If
sheetName = value.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
validatePRNAllzero = rowNumber
End Function


Public Function compareValuesZeroCheckProfitLossSelf(ByRef value1 As Range, ByRef value2 As Range, ByRef value3 As Range, ByRef value4 As Range, ByRef value5 As Range, ByRef value6 As Range, ByRef value7 As Range, ByRef value8 As Range, ByRef value9 As Range, ByRef value10 As Range, ByRef value11 As Range, ByRef value12 As Range) As String
Dim rowNumber As String
    If value1.value <> "" And value2.value <> "" And value3.value <> "" And value4.value <> "" And value5.value <> "" And value6.value <> "" And value7.value <> "" And value8.value <> "" And value9.value <> "" And value10.value <> "" And value11.value <> "" And value12.value <> "" Then
        If value1.value = 0 And value2.value = 0 And value3.value = 0 And value4.value = 0 And value5.value = 0 And value6.value = 0 And value7.value = 0 And value8.value = 0 And value9.value = 0 And value10.value = 0 And value11.value = 0 And value12.value = 0 Then
            rowNumber = value1.Address
        End If
    End If
sheetName = value1.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
compareValuesZeroCheckProfitLossSelf = rowNumber
End Function

Public Function compareValuesZeroCheckProfitLossWife(ByRef value1 As Range, ByRef value2 As Range, ByRef value3 As Range, ByRef value4 As Range, ByRef value5 As Range, ByRef value6 As Range, ByRef value7 As Range, ByRef value8 As Range, ByRef value9 As Range, ByRef value10 As Range, ByRef value11 As Range, ByRef value12 As Range) As String
Dim rowNumber As String
    If value1.value <> "" And value2.value <> "" And value3.value <> "" And value4.value <> "" And value5.value <> "" And value6.value <> "" And value7.value <> "" And value8.value <> "" And value9.value <> "" And value10.value <> "" And value11.value <> "" And value12.value <> "" Then
        If value1.value = 0 And value2.value = 0 And value3.value = 0 And value4.value = 0 And value5.value = 0 And value6.value = 0 And value7.value = 0 And value8.value = 0 And value9.value = 0 And value10.value = 0 And value11.value = 0 And value12.value = 0 Then
            rowNumber = value1.Address
        End If
    End If
sheetName = value1.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
compareValuesZeroCheckProfitLossWife = rowNumber
End Function



Public Function compareValuesZeroCheckSelf(ByRef value1 As Range, ByRef value2 As Range, ByRef value3 As Range) As String
Dim rowNumber As String
    If value1.value <> "" Or value2.value <> "" Or value3.value <> "" Then
        If value1.value = 0 And value2.value = 0 Then
            If (value3.value <= 0) Then
                rowNumber = value3.Address
            End If
        End If
    End If
sheetName = value1.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
compareValuesZeroCheckSelf = rowNumber
End Function

Public Function compareValuesZeroCheckWife(ByRef value1 As Range, ByRef value2 As Range, ByRef value3 As Range) As String
Dim rowNumber As String
    If value1.value <> "" Or value2.value <> "" Or value3.value <> "" Then
        If value1.value = 0 And value2.value = 0 Then
            If (value3.value <= 0) Then
                rowNumber = value3.Address
            End If
        End If
    End If
sheetName = value1.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
compareValuesZeroCheckWife = rowNumber
End Function




Public Sub lockUnlock_cell_rng_Hidden_without_clearing_contents(cursheet As String, rangeName As String, lockCellFlag As Boolean, Optional skipColumnFrom As String)

    Dim cellColor
    'On Error GoTo catch
    activeSheetName = ActiveSheet.name
    protectedStatus = ActiveSheet.ProtectContents
    
    Worksheets(cursheet).Activate
    ActiveSheet.Unprotect Password:=Pwd
    
    startRow = Range(rangeName).row
    lastRow = Range(rangeName).Rows.Count + startRow - 1
    
    startColumn = Range(rangeName).column
    lastColumn = startColumn + Range(rangeName).Columns.Count - 1 ' get actual last column
    
    If lockCellFlag = True Then
        cellColor = RGB(215, 215, 215) 'lock
    ElseIf lockCellFlag = False Then
        cellColor = RGB(255, 255, 255) 'unlock
    End If
    If skipColumnFrom = "" Then
    ' locked whole range
        Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).Locked = lockCellFlag
        With Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).Interior
            .Color = cellColor
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
        End With
    Else
        'skip columns and define new range
        lastColumn = skipColumnFrom - 1
        Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).Locked = lockCellFlag
        With Range(ActiveSheet.Cells(startRow, startColumn), ActiveSheet.Cells(lastRow, lastColumn)).Interior
           .Color = cellColor
           .Pattern = xlSolid
           .PatternColorIndex = xlAutomatic
        End With
        ActiveSheet.Cells(3, 1).Select
    End If
    
catch:
    If Err.Description <> "" Then
    '    MsgBox Err.Description
    End If
    
    ActiveSheet.Protect Password:=Pwd
    
    Worksheets(activeSheetName).Activate
    
    If protectedStatus Then
        ActiveSheet.Protect Password:=Pwd
    Else
        ActiveSheet.Unprotect Password:=Pwd
    End If
    
End Sub

Public Sub Zip_All_Files_in_Folder_Browse(FolderName As String, FileNameZip As String)
    
    Dim strDate As String, DefPath As String
    Dim oApp As Object

    FileNameZip = FolderName & ".zip"
    
    NewZip (FileNameZip)
    
    Set oApp = CreateObject("Shell.Application")

    If Right(FolderName, 1) <> "\" Then
        FolderName = FolderName & "\"
    End If

    'Copy the files to the compressed folder
    With CreateObject("Shell.Application") 'Crazy work around
    .Namespace("" & FileNameZip).CopyHere .Namespace("" & FolderName).items
    End With

    MsgBox "You find the zipfile here: " & FileNameZip

    
End Sub

Sub NewZip(sPath)
'Create empty Zip File
'Changed by keepITcool Dec-12-2005
    If Len(Dir(sPath)) > 0 Then
        Kill sPath
    End If
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub
Public Sub Delete_Folder(FolderName As String)
'Delete whole folder without removing the files first like in DeleteExample4
    Dim FSO As Object, parentFldName As String, MyPath As String
    Set FSO = CreateObject("scripting.filesystemobject")
    parentFldName = FSO.getFolder(FolderName).ParentFolder
    ChDir parentFldName
    MyPath = FolderName
    If Right(MyPath, 1) = "\" Then
        MyPath = Left(MyPath, Len(MyPath) - 1)
    End If
    FSO.deletefolder MyPath, True
    
End Sub

Public Sub resetSectionAHiddenDetails()
    ' Code to reset Return Period Month and Year
    
    Application.EnableEvents = False
    
    Dim rtnPrdFrom As String
    Dim year As String
    
    If Worksheets("A_Basic_Info").ProtectContents Then
        protectedStatus = True
    End If
    
    Worksheets("A_Basic_Info").Unprotect (Pwd)
    
    If (Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value <> "") Then
        If (TestDate(Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value) = True) Then
            rtnPrdTo = CDate(Format(Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value, "dd/mm/yyyy"))
        End If
    End If
    
    dd = Format(CDate(Trim(rtnPrdTo)), "DD")
    mm = Format(CDate(Trim(rtnPrdTo)), "MM")
    year = DatePart("yyyy", rtnPrdTo)
    Worksheets("A_Basic_Info").Range("RetInf.YearIncome").value = year
    Worksheets("A_Basic_Info").Range("RtnInf.MonthCode").value = mm
    
    ' Code to reset Return Type Code
    If (Worksheets("A_Basic_Info").Range("RetInf.RetType").value <> "") Then
        rowvalue = find_RetTypeCode(Worksheets("A_Basic_Info").Range("RetInf.RetType").value)
        RetTypeCode = Worksheets("Data").Cells(rowvalue, RetTypeCodeCol).value 'RetTypeCodeCol = 5th column of data sheet
         If (rowvalue <> 0) Then
            Worksheets("A_Basic_Info").Range("C4:C4").value = RetTypeCode
         End If
    End If
    
    
    If protectedStatus Then
        Worksheets("A_Basic_Info").Protect (Pwd)
    End If
    Application.EnableEvents = True
End Sub

Public Function validateDateBfrCurrentDate(ByRef value As Range, Optional ByVal colName As String) As String
Dim sheetName As String
Dim rowNumber As String
sheetName = value.Worksheet.name

If Trim(colName) = "" Then
    If value.value <> "" And checkFutureDate(value.value) = False Then
        rowNumber = value.Address
    End If
Else
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)
    
    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If c.value <> "" And c.value <> Empty And checkFutureDate(c.value) = False Then
                    collAddress = colName & c.row
                    If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
            End If
        Next
    End If
End With

End If
sheetName = value.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
validateDateBfrCurrentDate = rowNumber
End Function
'This function checks whether a list contains any records or not
Public Function isListNonEmpty(ByRef listName As Range) As String

    isRangeEmpty = False
    isListNonEmpty = "NE"
    isRangeEmpty = True
    startRow = listName.row
    lastRow = listName.Rows.Count + startRow - 1
    
    startColumn = listName.column
    lastColumn = startColumn + listName.Columns.Count - 1
    For i = startRow To lastRow
        If Worksheets("A_Basic_Info").Cells(i, startColumn).value <> "" Then
            isRangeEmpty = False
            Exit For
        End If
    Next
    
    If isRangeEmpty Then
        Dim vArr
        vArr = Split(Cells(1, startColumn).Address(True, False), "$")
        Col_Letter = vArr(0)
        isListNonEmpty = Col_Letter & listName.row
    End If
'    If WorksheetFunction.CountA(listName) = 0 Then  'List doesn't contain any data
'        isRangeEmpty = True
'        isListNonEmpty = listName.row & vbCrLf
'    End If

End Function

Public Function showHideWorksheet(flag As String, wsName As String)
    Dim showSection As Boolean
    showSection = False
    If UCase(flag) = "YES" Then showSection = True
    If UCase(flag) = "NO" Or flag = "" Then showSection = False
    
'    If showSection And isWifeSection _
'    And UCase(ThisWorkbook.Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value) = "NO" Then
'        showSection = False
'    End If
    
    If showSection Then
        Application.ThisWorkbook.Unprotect (Pwd)
        Worksheets(wsName).Visible = xlSheetVisible
        Application.ThisWorkbook.Protect (Pwd)
    Else
        Application.ThisWorkbook.Unprotect (Pwd)
        Dim unlockedCells As Range
        Set unlockedCells = GetUnlocked(wsName)
        If Not unlockedCells Is Nothing Then
            unlockedCells.ClearContents
        End If
        If wsName = "E1_IDA_CA" Then
            Call ClearSecE1Data
        ElseIf wsName = "G_Partnership_Income" Then
            Call ClearSecGData
        ElseIf wsName = "I_Computation_of_Car_Benefit" Then
            Call ClearSecIData
        ElseIf wsName = "P_Advance_Tax_Credits" Then
            Call clearSecPData
        End If
        Worksheets(wsName).Visible = xlHidden
        Application.ThisWorkbook.Protect (Pwd)
    End If
End Function

Public Function showHideWorksheetBI(flag As String, otherVal As String, wsName As String)
    Dim showSection As Boolean
    showSection = False
    If UCase(flag) = "YES" Or UCase(otherVal) = "YES" Then showSection = True
    If (UCase(flag) = "NO" Or flag = "") And (UCase(otherVal) = "NO" Or otherVal = "") Then showSection = False
    
'    If showSection And isWifeSection _
'    And UCase(ThisWorkbook.Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value) = "NO" Then
'        showSection = False
'    End If
    
    If showSection Then
        Application.ThisWorkbook.Unprotect (Pwd)
        Worksheets(wsName).Visible = xlSheetVisible
        Application.ThisWorkbook.Protect (Pwd)
    Else
        Application.ThisWorkbook.Unprotect (Pwd)
        Dim unlockedCells As Range
        Set unlockedCells = GetUnlocked(wsName)
        If Not unlockedCells Is Nothing Then
            unlockedCells.ClearContents
        End If
        If wsName = "E1_IDA_CA" Then
            Call ClearSecE1Data
        ElseIf wsName = "G_Partnership_Income" Then
            Call ClearSecGData
        ElseIf wsName = "I_Computation_of_Car_Benefit" Then
            Call ClearSecIData
        ElseIf wsName = "P_Advance_Tax_Credits" Then
            Call clearSecPData
        End If
        Worksheets(wsName).Visible = xlHidden
        Application.ThisWorkbook.Protect (Pwd)
    End If
End Function
Public Sub ClearSecE1Data()
    Dim wsName As String
    wsName = "E1_IDA_CA"
    Worksheets(wsName).Unprotect (Pwd)
    
    rangeName = Range("IniAllIBD.ListPart2S").Address
    startRow = Range(rangeName).row
    endRow = Range(rangeName).Rows.Count + startRow - 1
    For i = startRow To endRow
        Worksheets(wsName).Range("C" & i).value = ""
    Next
    
    rangeName = Range("IniAllIBD.ListPart2W").Address
    startRow = Range(rangeName).row
    endRow = Range(rangeName).Rows.Count + startRow - 1
    For i = startRow To endRow
        Worksheets(wsName).Range("C" & i).value = ""
    Next
    
    rangeName = Range("DeprIntengAst.ListS").Address
    startRow = Range(rangeName).row
    endRow = Range(rangeName).Rows.Count + startRow - 1
    For i = startRow To endRow
        Worksheets(wsName).Range("E" & i).value = ""
    Next
    
    rangeName = Range("DeprIntengAst.ListW").Address
    startRow = Range(rangeName).row
    endRow = Range(rangeName).Rows.Count + startRow - 1
    For i = startRow To endRow
        Worksheets(wsName).Range("E" & i).value = ""
    Next
    
    Worksheets(wsName).Protect (Pwd)
End Sub

Public Sub ClearSecGData()
    Dim wsName As String
    wsName = "G_Partnership_Income"
    Worksheets(wsName).Unprotect (Pwd)

    rangeName = Range("ProfitShare.ListS").Address
    startRow = Range(rangeName).row
    endRow = Range(rangeName).Rows.Count + startRow - 1
    For i = startRow To endRow
        Worksheets(wsName).Range("K" & i).value = ""
    Next
    
    rangeName = Range("ProfitShare.ListW").Address
    startRow = Range(rangeName).row
    endRow = Range(rangeName).Rows.Count + startRow - 1
    For i = startRow To endRow
        Worksheets(wsName).Range("K" & i).value = ""
    Next
    
    Worksheets(wsName).Protect (Pwd)
End Sub

Public Sub ClearSecIData()
    Dim wsName As String
    wsName = "I_Computation_of_Car_Benefit"
    Worksheets(wsName).Unprotect (Pwd)

    rangeName = Range("CarBenefit.ListS").Address
    startRow = Range(rangeName).row
    endRow = Range(rangeName).Rows.Count + startRow - 1
    For i = startRow To endRow
        Worksheets(wsName).Range("L" & i).value = ""
        Worksheets(wsName).Range("M" & i).value = ""
    Next
    
    rangeName = Range("CarBenefit.ListW").Address
    startRow = Range(rangeName).row
    endRow = Range(rangeName).Rows.Count + startRow - 1
    For i = startRow To endRow
        Worksheets(wsName).Range("L" & i).value = ""
        Worksheets(wsName).Range("M" & i).value = ""
    Next
    
    Worksheets(wsName).Protect (Pwd)
End Sub
Public Sub clearSecPData()
    Dim wsName As String
    wsName = "P_Advance_Tax_Credits"
    Worksheets(wsName).Unprotect (Pwd)

    rangeName = Range("VehicleAdvTaxPaid.ListS").Address
    startRow = Range(rangeName).row
    endRow = Range(rangeName).Rows.Count + startRow - 1
    For i = startRow To endRow
        Worksheets(wsName).Range("H" & i).value = ""
    Next
    
    rangeName = Range("VehicleAdvTaxPaid.ListW").Address
    startRow = Range(rangeName).row
    endRow = Range(rangeName).Rows.Count + startRow - 1
    For i = startRow To endRow
        Worksheets(wsName).Range("H" & i).value = ""
    Next
    
    Worksheets(wsName).Protect (Pwd)
End Sub



'This function returns all the unlocked cells as Range
Public Function GetUnlocked(myWSName As String) As Range
Dim myWS As Worksheet
Set myWS = Worksheets(myWSName)
Dim r As Range


Set GetUnlocked = Nothing
For Each r In myWS.UsedRange
    If Not r.Locked Then
        If GetUnlocked Is Nothing Then
            Set GetUnlocked = r
        Else
            Set GetUnlocked = Union(GetUnlocked, r)
        End If
    End If
Next r

End Function

Public Sub showModulerSectionsAmendment()
    Dim flag As String
    Dim otherVal As String
    
    flag = ThisWorkbook.Worksheets("A_Basic_Info").Range("B7").value
    otherVal = ThisWorkbook.Worksheets("A_Basic_Info").Range("B19").value
    Call showHideWorksheetAmd(flag, "B_Profit_Loss_Account_Self")
    Call showHideWorksheetAmdBI(flag, otherVal, "C_Balance_Sheet")
    Call showHideWorksheetAmdBI(flag, otherVal, "D_Stock_Analysis")
    Call showHideWorksheetAmdBI(flag, otherVal, "E1_IDA_CA")
    Call showHideWorksheetAmdBI(flag, otherVal, "E2_CA_WTA_WDV")
    Call showHideWorksheetAmdBI(flag, otherVal, "E2_CA_WTA_SLM")
    Call showHideWorksheetAmdBI(flag, otherVal, "E_Summary_of_Capital_Allowance")
    Call showHideWorksheetAmdBI(flag, otherVal, "N_Installment_Tax_Credits")
    Call showHideWorksheetAmdBI(flag, otherVal, "O_WHT_Credits")
    Call showHideWorksheetAmdBI(flag, otherVal, "S_Previous_Years_Losses")
    Call showHideWorksheetAmd(flag, "T_Income_Computation_Self")
    
    flag = ThisWorkbook.Worksheets("A_Basic_Info").Range("B8").value
    Call showHideWorksheetAmd(flag, "G_Partnership_Income")
    
    flag = ThisWorkbook.Worksheets("A_Basic_Info").Range("B9").value
    Call showHideWorksheetAmd(flag, "H_Estate_Trust_Income")
    
    flag = ThisWorkbook.Worksheets("A_Basic_Info").Range("B10").value
    Call showHideWorksheetAmd(flag, "I_Computation_of_Car_Benefit")
    
    flag = ThisWorkbook.Worksheets("A_Basic_Info").Range("B11").value
    Call showHideWorksheetAmd(flag, "J_Computation_of_Mortgage")
    
    flag = ThisWorkbook.Worksheets("A_Basic_Info").Range("B12").value
    Call showHideWorksheetAmd(flag, "K_Home_Ownership_Saving_Plan")
    
    flag = ThisWorkbook.Worksheets("A_Basic_Info").Range("B13").value
    Call showHideWorksheetAmd(flag, "L_Computation_of_Insu_Relief")
    
    flag = ThisWorkbook.Worksheets("A_Basic_Info").Range("B14").value
    Call showHideWorksheetAmd(flag, "P_Advance_Tax_Credits")
    
    flag = ThisWorkbook.Worksheets("A_Basic_Info").Range("B15").value
    Call showHideWorksheetAmd(flag, "R_DTAA_Credits")
    
    flag = ThisWorkbook.Worksheets("A_Basic_Info").Range("B19").value
    otherVal = ThisWorkbook.Worksheets("A_Basic_Info").Range("B7").value
    Call showHideWorksheet(flag, "B_Profit_Loss_Account_Wife")
    Call showHideWorksheetAmdBI(flag, otherVal, "C_Balance_Sheet")
    Call showHideWorksheetAmdBI(flag, otherVal, "D_Stock_Analysis")
    Call showHideWorksheetAmdBI(flag, otherVal, "E1_IDA_CA")
    Call showHideWorksheetAmdBI(flag, otherVal, "E2_CA_WTA_WDV")
    Call showHideWorksheetAmdBI(flag, otherVal, "E2_CA_WTA_SLM")
    Call showHideWorksheetAmdBI(flag, otherVal, "E_Summary_of_Capital_Allowance")
    Call showHideWorksheetAmdBI(flag, otherVal, "N_Installment_Tax_Credits")
    Call showHideWorksheetAmdBI(flag, otherVal, "O_WHT_Credits")
    Call showHideWorksheetAmdBI(flag, otherVal, "S_Previous_Years_Losses")
    Call showHideWorksheet(flag, "T_Income_Computation_Wife")
    
    'This is done because calls made above sets protection on workbook, which creates problem as still some hide/unhide are to be performed.
    Application.ThisWorkbook.Unprotect (Pwd)
End Sub

Public Sub showHideWorksheetAmd(flag As String, wsName As String)
    If UCase(flag) = "YES" Then
        Call showHideWorksheet(flag, wsName)
    End If
End Sub
Public Sub showHideWorksheetAmdBI(flag As String, otherVal As String, wsName As String)
    If UCase(flag) = "YES" Or UCase(otherVal) = "YES" Then
        flag = "YES"
        Call showHideWorksheet(flag, wsName)
    End If
End Sub

Public Function Initialize()
    str1 = Left(Pwd, 8)
    tmpStr = ""
    Initialize = ""


    For T = 1 To Len(str1)
        tmpStr = tmpStr & Chr(asc(Mid(str1, T, 1)) - 128)


        If Len(tmpStr) = 1000 Then


            DoEvents
                Initialize = Initialize & tmpStr
                tmpStr = ""
            End If
        Next T
        Initialize = Initialize & tmpStr
        Pwd = Initialize
End Function



Public Sub checkOfficeVersion()
    If Application.Version < "12.0" Or (Application.Version = "12.0" And Application.Build < 6425) Then
        MsgBox "Please upgrade your Microsoft Office to Office 2007 (Service Pack 2) or above to open this Excel Template."
        ThisWorkbook.Saved = True
        If Workbooks.Count < 2 Then
            Application.Quit
            ThisWorkbook.Close SaveChanges = False
        Else
            ThisWorkbook.Close SaveChanges = False
        End If
    End If
End Sub

'Code for Macro forceful enabling : start
Public Function HideAllSheets()
    Dim wasWBProtected As Boolean
    wasWBProtected = True
    If ThisWorkbook.ProtectStructure = False Then
        wasWBProtected = False
    End If
    If wasWBProtected Then
        ActiveWorkbook.Unprotect Password:=Pwd
    End If
    Application.ScreenUpdating = False

    Dim countSheet As Integer
    Dim i As Integer
    Dim visibleSheets As String
    countSheet = ThisWorkbook.Worksheets.Count
    ThisWorkbook.Worksheets("Macros_Disabled").Visible = True
    For i = 1 To countSheet
        If ThisWorkbook.Worksheets(i).Visible = True Then
            If visibleSheets = "" Then
                visibleSheets = ThisWorkbook.Worksheets(i).name
            Else
                visibleSheets = visibleSheets & ":" & ThisWorkbook.Worksheets(i).name
            End If
            If Not ThisWorkbook.Worksheets(i).name = "Macros_Disabled" Then
                ThisWorkbook.Worksheets(i).Visible = False
            End If
        End If
    Next
    ThisWorkbook.Worksheets("Data").Unprotect (Pwd)
    ThisWorkbook.Worksheets("Data").Range("visibleSheetsArr").value = visibleSheets
    ThisWorkbook.Worksheets("Data").Protect (Pwd)
    
    If wasWBProtected Then
        ActiveWorkbook.Protect Password:=Pwd
    End If
    
    isHideUnhidePerformed = True
End Function

Public Function MakeSheetsVisible()
    Dim visibleSheets As String
    visibleSheets = ThisWorkbook.Worksheets("Data").Range("visibleSheetsArr").value
    Dim visibleSheetsArr As Variant
    visibleSheetsArr = Split(visibleSheets, ":")
    
    Dim countSheet As Integer
    Dim i As Integer
    countSheet = ThisWorkbook.Worksheets.Count
    
    For i = 1 To countSheet
        If IsInArray(ThisWorkbook.Worksheets(i).name, visibleSheetsArr) Then
                ThisWorkbook.Worksheets(i).Visible = True
        End If
    Next
    
    ThisWorkbook.Worksheets("Macros_Disabled").Visible = False
    ThisWorkbook.Worksheets("Data").Unprotect (Pwd)
    ThisWorkbook.Worksheets("Data").Range("visibleSheetsArr").value = visibleSheets
    ThisWorkbook.Worksheets("Data").Protect (Pwd)
End Function
Public Function MacrosDisClose()
    If ThisWorkbook.Saved = False Then
        isCalledOnClose = True
    Else
'        isCalledOnClose = False
        ThisWorkbook.Saved = False
        Call HideAllSheets
        ThisWorkbook.Saved = True
        ThisWorkbook.Save
'        Application.EnableEvents = False
'        ThisWorkbook.Close SaveChanges:=True
    End If
End Function

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function
'Code for Macro forceful enabling : end

'This function enables/disables bank details fields based on value of liability : start
Function enableDisableBankDetails()
'    If Sheet20.Range("FinalTax.TaxRefundDueS").value <> "" Then
'       If Sheet20.Range("FinalTax.TaxRefundDueS").value < 0 Then
'            If Sheet14.Range("BankS").Locked = True Then
'                Call lockUnlock_cell_rng("A_Basic_Info", "BankS", False)
'                Call lockUnlock_cell_rng("A_Basic_Info", "BranchS", False)
'                Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.CityS", False)
'                Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNameS", False)
'                Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNumberS", False)
'            End If
'        Else
'            If Sheet14.Range("BankS").Locked = False Then
'                Call lockUnlock_cell_rng("A_Basic_Info", "BankS", True)
'                Call lockUnlock_cell_rng("A_Basic_Info", "BranchS", True)
'                Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.CityS", True)
'                Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNameS", True)
'                Call lockUnlock_cell_rng("A_Basic_Info", "BankDtl.AccNumberS", True)
'            End If
'        End If
'    End If
End Function
'This function enables/disables bank details fields based on value of liability : end



'Functions added for validating A > B+C type of validations which are missed in case of ImportCsv : Start

Public Function validateSalesQtyD1(ByRef value As Range) As String

    Dim rowNumber As String
    Dim sheetName As String
    Dim startRow As String
    Dim endRow As String
    Dim opStock As String
    Dim purchase As String
    Dim salesQty As String
    Dim diff As Double
    
    sheetName = value.Parent.name
    Worksheets(sheetName).Activate
    
    startRow = Worksheets(sheetName).Range(value.name).row
    endRow = Worksheets(sheetName).Range(value.name).Rows.Count + startRow - 1
    
    For i = startRow To endRow
        opStock = Worksheets(sheetName).Range("$C$" & i).value
        purchase = Worksheets(sheetName).Range("$D$" & i).value
        salesQty = Worksheets(sheetName).Range("$E$" & i).value
        If IsNumeric(opStock) And IsNumeric(purchase) And IsNumeric(salesQty) Then
            diff = (CDbl(opStock) + CDbl(purchase)) - CDbl(salesQty)
            If diff < 0 Then
                collAddress = "E" & i
                If rowNumber <> "" Then
                   rowNumber = rowNumber & "," & collAddress
                Else
                   rowNumber = rowNumber & collAddress
                End If
            End If
        End If
    Next

    If rowNumber = "" Then
        rowNumber = "NE"
    End If
    validateSalesQtyD1 = rowNumber
End Function

Public Function validateSalesQtyD2(ByRef value As Range) As String

    Dim rowNumber As String
    Dim sheetName As String
    Dim startRow As String
    Dim endRow As String
    Dim opStock As String
    Dim purchase As String
    Dim consumption As String
    Dim salesQty As String
    Dim diff As Double
    
    sheetName = value.Parent.name
    Worksheets(sheetName).Activate
    
    startRow = Worksheets(sheetName).Range(value.name).row
    endRow = Worksheets(sheetName).Range(value.name).Rows.Count + startRow - 1
    
    For i = startRow To endRow
        opStock = Worksheets(sheetName).Range("$C$" & i).value
        purchase = Worksheets(sheetName).Range("$D$" & i).value
        consumption = Worksheets(sheetName).Range("$E$" & i).value
        salesQty = Worksheets(sheetName).Range("$F$" & i).value
        If IsNumeric(opStock) And IsNumeric(purchase) And IsNumeric(salesQty) And IsNumeric(consumption) Then
            diff = CDbl(opStock) + CDbl(purchase) - CDbl(consumption) - CDbl(salesQty)
            If diff < 0 Then
                collAddress = "F" & i
                If rowNumber <> "" Then
                   rowNumber = rowNumber & "," & collAddress
                Else
                   rowNumber = rowNumber & collAddress
                End If
            End If
        End If
    Next

    If rowNumber = "" Then
        rowNumber = "NE"
    End If
    validateSalesQtyD2 = rowNumber
End Function

Public Function validateInvDedcE1_1(ByRef value As Range) As String

    Dim rowNumber As String
    Dim sheetName As String
    Dim startRow As String
    Dim endRow As String
    Dim cost As String
    Dim invDedc As String
    
    Dim diff As Double
    
    sheetName = value.Parent.name
    Worksheets(sheetName).Activate
    
    startRow = Worksheets(sheetName).Range(value.name).row
    endRow = Worksheets(sheetName).Range(value.name).Rows.Count + startRow - 1
    
    For i = startRow To endRow
        cost = Worksheets(sheetName).Range("$G$" & i).value
        invDedc = Worksheets(sheetName).Range("$H$" & i).value
        
        If IsNumeric(cost) And IsNumeric(invDedc) Then
            diff = CDbl(cost) - CDbl(invDedc)
            If diff < 0 Then
                collAddress = "H" & i
                If rowNumber <> "" Then
                   rowNumber = rowNumber & "," & collAddress
                Else
                   rowNumber = rowNumber & collAddress
                End If
            End If
        End If
    Next

    If rowNumber = "" Then
        rowNumber = "NE"
    End If
    validateInvDedcE1_1 = rowNumber
End Function

Public Function validateInvDedcE1_2(ByRef value As Range) As String

    Dim rowNumber As String
    Dim sheetName As String
    Dim startRow As String
    Dim endRow As String
    Dim cost As String
    Dim invDedc As String
    
    Dim diff As Double
    
    sheetName = value.Parent.name
    Worksheets(sheetName).Activate
    
    startRow = Worksheets(sheetName).Range(value.name).row
    endRow = Worksheets(sheetName).Range(value.name).Rows.Count + startRow - 1
    
    For i = startRow To endRow
        cost = Worksheets(sheetName).Range("$F$" & i).value
        invDedc = Worksheets(sheetName).Range("$G$" & i).value
        
        If IsNumeric(cost) And IsNumeric(invDedc) Then
            diff = CDbl(cost) - CDbl(invDedc)
            If diff < 0 Then
                collAddress = "G" & i
                If rowNumber <> "" Then
                   rowNumber = rowNumber & "," & collAddress
                Else
                   rowNumber = rowNumber & collAddress
                End If
            End If
        End If
    Next

    If rowNumber = "" Then
        rowNumber = "NE"
    End If
    validateInvDedcE1_2 = rowNumber
End Function

Public Function validateDedcE1_3(ByRef value As Range) As String

    Dim rowNumber As String
    Dim sheetName As String
    Dim startRow As String
    Dim endRow As String
    Dim exp As String
    Dim wdv As String
    Dim sale As String
    Dim dedc As String
    
    Dim diff As Double
    
    sheetName = value.Parent.name
    Worksheets(sheetName).Activate
    
    startRow = Worksheets(sheetName).Range(value.name).row
    endRow = Worksheets(sheetName).Range(value.name).Rows.Count + startRow - 1
    
    For i = startRow To endRow
        exp = Worksheets(sheetName).Range("$C$" & i).value
        wdv = Worksheets(sheetName).Range("$D$" & i).value
        sale = Worksheets(sheetName).Range("$E$" & i).value
        dedc = Worksheets(sheetName).Range("$F$" & i).value
        
        If IsNumeric(exp) And IsNumeric(wdv) And IsNumeric(sale) And IsNumeric(dedc) Then
            diff = CDbl(exp) + CDbl(wdv) - CDbl(sale) - CDbl(dedc)
            If diff < 0 Then
                collAddress = "F" & i
                If rowNumber <> "" Then
                   rowNumber = rowNumber & "," & collAddress
                Else
                   rowNumber = rowNumber & collAddress
                End If
            End If
        End If
    Next

    If rowNumber = "" Then
        rowNumber = "NE"
    End If
    validateDedcE1_3 = rowNumber
End Function

Public Function validateDispCostE1_4(ByRef value As Range) As String

    Dim rowNumber As String
    Dim sheetName As String
    Dim startRow As String
    Dim endRow As String
    Dim actCost As String
    Dim dispCost As String
    
    Dim diff As Double
    
    sheetName = value.Parent.name
    Worksheets(sheetName).Activate
    
    startRow = Worksheets(sheetName).Range(value.name).row
    endRow = Worksheets(sheetName).Range(value.name).Rows.Count + startRow - 1
    
    For i = startRow To endRow
        actCost = Worksheets(sheetName).Range("$G$" & i).value
        dispCost = Worksheets(sheetName).Range("$H$" & i).value
        
        If IsNumeric(actCost) And IsNumeric(dispCost) Then
            diff = CDbl(actCost) - CDbl(dispCost)
            If diff < 0 Then
                collAddress = "H" & i
                If rowNumber <> "" Then
                   rowNumber = rowNumber & "," & collAddress
                Else
                   rowNumber = rowNumber & collAddress
                End If
            End If
        End If
    Next

    If rowNumber = "" Then
        rowNumber = "NE"
    End If
    validateDispCostE1_4 = rowNumber
End Function

Public Function validateDispCostE2(ByRef value As Range) As String

    Dim rowNumber As String
    Dim sheetName As String
    Dim startRow As String
    Dim endRow As String
    Dim actCost As String
    Dim dispCost As String
    
    Dim diff As Double
    
    sheetName = value.Parent.name
    Worksheets(sheetName).Activate
    
    startRow = Worksheets(sheetName).Range(value.name).row
    endRow = Worksheets(sheetName).Range(value.name).Rows.Count + startRow - 1
    
    For i = startRow To endRow
        actCost = Worksheets(sheetName).Range("$F$" & i).value
        dispCost = Worksheets(sheetName).Range("$G$" & i).value
        
        If IsNumeric(actCost) And IsNumeric(dispCost) Then
            diff = CDbl(actCost) - CDbl(dispCost)
            If diff < 0 Then
                collAddress = "G" & i
                If rowNumber <> "" Then
                   rowNumber = rowNumber & "," & collAddress
                Else
                   rowNumber = rowNumber & collAddress
                End If
            End If
        End If
    Next

    If rowNumber = "" Then
        rowNumber = "NE"
    End If
    validateDispCostE2 = rowNumber
End Function

'Functions added for validating A > B+C type of validations which are missed in case of ImportCsv : End


'this is used to detect duplicate values of any single column (kind of primary key, BUT NOT COMPOSITE ONE :p)
Public Function checkCertificateDuplicacy(sheetName As String, rangeName As String, colName As String) As String
        
        Worksheets(sheetName).Activate
        Dim rowNumber As String
        checkDuplicacy = ""
        Dim k
        k = 0
        startRow = Worksheets(sheetName).Range(rangeName).row
        endRow = Worksheets(sheetName).Range(rangeName).Rows.Count + startRow - 1
        Dim strarray() As String
        ReDim strarray(endRow - startRow + 1)
        Dim str1 As String
        str1 = ""
        For i = startRow To endRow
            str1 = Worksheets(sheetName).Range(colName & i).value
            strarray(k) = str1
            k = k + 1
        Next

        If k > 1 Then
            For i = 0 To k - 1
                tmp = strarray(i)
                For l = 1 To k - 1
                    If i <> l And strarray(l) <> "" And tmp = strarray(l) Then
                        If rowNumber = "" Then
                        rowNumber = colName & l + startRow - 1
                        Else
                        rowNumber = rowNumber & "," & colName & l + startRow - 1
                        End If
                    End If
                Next
            Next
        End If
        If rowNumber = "" Then
            checkCertificateDuplicacy = "NE"
        Else
            checkCertificateDuplicacy = rowNumber
        End If
       
End Function
'Added by vaishali
'Added for validating Account name
Public Function validateAlphaNumericForAccName(ByRef value As Range, Optional ByVal colName As String) As String
'Worksheets("J_DTAA_Credits").Unprotect ("P@ssw0rd")
'Sheet13.Unprotect ("P@ssw0rd")
Dim rowNumber As String
Dim sheetName As String

If Trim(colName) = "" Then
    If TestAlphanumericForAccName(value.value) = False Then
        rowNumber = value.Address
     End If
Else
 
    Dim column As Range
    Dim columnRange As Range
    Dim blnFlag As Boolean
    Dim sName As name
    Set sName = value.name
    sheetName = Mid(sName, 2, InStrRev(sName, "!") - 2)
    Dim singleRow As Range
    Worksheets(sheetName).Activate
    Set columnRange = Worksheets(sheetName).Range(colName & "1", colName & Worksheets(sheetName).UsedRange.Rows.Count)
    Set column = Intersect(value, columnRange)
    
    With Worksheets(sheetName)
    If Not column Is Nothing Then
        For Each c In column.Cells
            If TestAlphanumericForAccName(c.value) = False Then
                'rowNumber = rowNumber & C.Address & ","
                collAddress = colName & c.row
                If rowNumber <> "" Then
                        rowNumber = rowNumber & "," & collAddress
                     Else
                        rowNumber = rowNumber & collAddress
                     End If
            End If
        Next
    End If
End With

End If
sheetName = value.Worksheet.name
If rowNumber = "" Then
    rowNumber = "NE"
End If
validateAlphaNumericForAccName = rowNumber
End Function
'Added by vaishali
'Added for validating Account name
Public Function TestAlphanumericForAccName(lstr_check As String) As Boolean
'allowed characters A to Z, a to z And Blank
Dim i As Integer
Dim ia As Integer
Dim ina As Integer
Dim stlen As Integer

stlen = Len(lstr_check)
ia = 0
ina = 0
For i = 1 To stlen
    Select Case (Mid(lstr_check, i, 1))
        Case "A" To "Z"          'A to Z
            ia = ia + 1
        Case "a" To "z"          'a to z
            ia = ia + 1
        Case " "                 ' Blank
           ia = ia + 1
        Case "0" To "9"          '0 to 9
            ia = ia + 1
        Case "/", ":", "-", "_", ",", ".", "'", "&", "(", ")"
            ia = ia + 1
        Case Else
            ina = ina + 1
    End Select
Next i
If ina = 0 Then
    TestAlphanumericForAccName = True
Else
    TestAlphanumericForAccName = False
End If

End Function

'Added by vaishali gohil
'validation for personal relief for year

Function validatePRS(ByVal PersonalRelief As Double, ByVal RtnYear As Integer) As Double
Application.Volatile

If PersonalRelief < 0 Then
    PersonalRelief = 0
Else
    If (RtnYear <= 2016) Then
        If PersonalRelief > 13944 Then
            Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
            Worksheets("T_Tax_Computation").Activate
            ActiveSheet.Range("TaxComp.PersonalReliefS").value = 0#
            Worksheets("T_Tax_Computation").Protect (Pwd)
            MsgBox "Please enter positive numeric value less than or equal to 13944."
        End If
    End If
    If (RtnYear = 2017) Then
        If PersonalRelief > 15360 Then
            Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
            Worksheets("T_Tax_Computation").Activate
            ActiveSheet.Range("TaxComp.PersonalReliefS").value = 0#
            Worksheets("T_Tax_Computation").Protect (Pwd)
            MsgBox "Please enter positive numeric value less than or equal to 15360."
        End If
    End If
    If (RtnYear = 2018 Or RtnYear = 2019) Then
        If PersonalRelief > 16896 Then
            Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
            Worksheets("T_Tax_Computation").Activate
            ActiveSheet.Range("TaxComp.PersonalReliefS").value = 0#
            Worksheets("T_Tax_Computation").Protect (Pwd)
            MsgBox "Please enter positive numeric value less than or equal to 16896."
        End If
    End If
    'Added by Ruth and Lawrence on 05/01/2021 to accomodate the 2020 tax relief
    If (RtnYear = 2020) Then
        If PersonalRelief > 25824 Or PersonalRelief < 0 Then
            Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
            Worksheets("T_Tax_Computation").Activate
            ActiveSheet.Range("TaxComp.PersonalReliefS").value = 0#
            Worksheets("T_Tax_Computation").Protect (Pwd)
            MsgBox "Please enter positive numeric value less than or equal to 25824."
        End If
    End If
    
    'Added 13/01/2022 to accomodate the 2021 tax relief
    If (RtnYear >= 2021) Then
        If ((PersonalRelief > 28800) Or (PersonalRelief < 0)) Then
            Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
            Worksheets("T_Tax_Computation").Activate
            ActiveSheet.Range("TaxComp.PersonalReliefS").value = 0#
            Worksheets("T_Tax_Computation").Protect (Pwd)
            MsgBox "Please enter positive numeric value less than or equal to 28800."
        End If
    End If
    'End 13/01/2022
End If
validatePRS = PersonalRelief
End Function
'validation for personal relief for year
'Added by vaishali gohil
Function validatePRW(ByVal PersonalRelief As Double, ByVal RtnYear As Integer) As Double
Application.Volatile

If PersonalRelief < 0 Then
    PersonalRelief = 0
Else
    If (RtnYear <= 2016) Then
        If PersonalRelief > 13944 Then
            Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
            Worksheets("T_Tax_Computation").Activate
            ActiveSheet.Range("TaxComp.PersonalReliefW").value = 0#
            Worksheets("T_Tax_Computation").Protect (Pwd)
            MsgBox "Please enter positive numeric value less than or equal to 13944."
        End If
    End If
    If (RtnYear = 2017) Then
        If PersonalRelief > 15360 Then
            Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
            Worksheets("T_Tax_Computation").Activate
            ActiveSheet.Range("TaxComp.PersonalReliefW").value = 0#
            Worksheets("T_Tax_Computation").Protect (Pwd)
            MsgBox "Please enter positive numeric value less than or equal to 15360."
        End If
    End If
    If (RtnYear = 2018 Or RtnYear = 2019) Then
        If PersonalRelief > 16896 Then
            Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
            Worksheets("T_Tax_Computation").Activate
            ActiveSheet.Range("TaxComp.PersonalReliefS").value = 0#
            Worksheets("T_Tax_Computation").Protect (Pwd)
            MsgBox "Please enter positive numeric value less than or equal to 16896."
        End If
    End If
    'Added by Ruth and Lawrence on 05/01/2021 to accomodate the 2020 tax relief
    If (RtnYear = 2020) Then
        If PersonalRelief > 25824 Or PersonalRelief < 0 Then
            Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
            Worksheets("T_Tax_Computation").Activate
            ActiveSheet.Range("TaxComp.PersonalReliefW").value = 0#
            Worksheets("T_Tax_Computation").Protect (Pwd)
            MsgBox "Please enter positive numeric value less than or equal to 25824."
        End If
     End If
     
     'Added 13/01/2022 to accomodate the 2020 tax relief
    If (RtnYear = 2020) Then
        If ((PersonalRelief > 28800) Or (PersonalRelief < 0)) Then
            Worksheets("T_Tax_Computation").Unprotect Password:=Pwd
            Worksheets("T_Tax_Computation").Activate
            ActiveSheet.Range("TaxComp.PersonalReliefW").value = 0#
            Worksheets("T_Tax_Computation").Protect (Pwd)
            MsgBox "Please enter positive numeric value less than or equal to 28800."
        End If
     End If
    'End 13/01/2022
End If
validatePRW = PersonalRelief
End Function

'validation for personal relief for year end
'Added by vaishali gohil end


Function toggleCellsInE1Sheet(ByVal year As String, ByVal mm As String) As Double
    Worksheets("E1_IDA_CA").Unprotect (Pwd)
    Worksheets("E2_CA_WTA_WDV").Unprotect (Pwd)
    rangeName = Worksheets("E1_IDA_CA").Range("IniAllPlanMach.ListPart1S").Address
    startRow = Worksheets("E1_IDA_CA").Range(rangeName).row
    endRow = Worksheets("E1_IDA_CA").Range(rangeName).Rows.Count + startRow - 1
    
                If year >= 2020 Then
                    If (year = 2020 And mm >= 4) Or year > 2020 Then
                        Call lockUnlock_cell_rng("E1_IDA_CA", "A" & startRow & ":A" & endRow, True)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "B" & startRow & ":B" & endRow, False)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "C" & startRow & ":C" & endRow, True)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "D" & startRow & ":D" & endRow, False)
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "D5" & ":D7", True)
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "E5" & ":E7", True)
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "B4" & ":C4", False)
                        Worksheets("E2_CA_WTA_WDV").Range("B4").value = "25"
                        Worksheets("E2_CA_WTA_WDV").Range("C4").value = "10"
                        Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_WDV", "B4:C4", True)
                        'Wife
                        If UCase(Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value) = "YES" Then
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "D17" & ":D19", True)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "E17" & ":E19", True)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "B17" & ":B19", False)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "C17" & ":C19", False)
                        End If
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "B16" & ":C16", False)
                        Worksheets("E2_CA_WTA_WDV").Range("B16").value = "25"
                        Worksheets("E2_CA_WTA_WDV").Range("C16").value = "10"
                        Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_WDV", "B16:C16", True)
                    Else
                        Call lockUnlock_cell_rng("E1_IDA_CA", "A" & startRow & ":A" & endRow, False)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "B" & startRow & ":B" & endRow, True)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "C" & startRow & ":C" & endRow, False)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "D" & startRow & ":D" & endRow, True)
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "D5" & ":D7", False)
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "E5" & ":E7", False)
                        Worksheets("E2_CA_WTA_WDV").Range("D5").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("D6").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("D7").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("E5").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("E6").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("E7").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("B5").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("B6").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("B7").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("C5").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("C6").value = "0"
                        Worksheets("E2_CA_WTA_WDV").Range("C7").value = "0"
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "B4" & ":C4", False)
                        Worksheets("E2_CA_WTA_WDV").Range("B4").value = "37.5"
                        Worksheets("E2_CA_WTA_WDV").Range("C4").value = "30"
                        Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_WDV", "B4:C4", True)
                        
                        'Wife
                        If UCase(Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value) = "YES" Then
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "D17" & ":D19", False)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "E17" & ":E19", False)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "B17" & ":B19", False)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "C17" & ":C19", False)
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
                        End If
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "B16" & ":C16", False)
                        Worksheets("E2_CA_WTA_WDV").Range("B16").value = "37.5"
                        Worksheets("E2_CA_WTA_WDV").Range("C16").value = "30"
                        Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_WDV", "B16:C16", True)
                        
                    End If
                Else
                    Call lockUnlock_cell_rng("E1_IDA_CA", "A" & startRow & ":A" & endRow, False)
                    Call lockUnlock_cell_rng("E1_IDA_CA", "B" & startRow & ":B" & endRow, True)
                    Call lockUnlock_cell_rng("E1_IDA_CA", "C" & startRow & ":C" & endRow, False)
                    Call lockUnlock_cell_rng("E1_IDA_CA", "D" & startRow & ":D" & endRow, True)
                    Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "D5" & ":D7", False)
                    Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "E5" & ":E7", False)
                    Worksheets("E2_CA_WTA_WDV").Range("D5").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("D6").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("D7").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("E5").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("E6").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("E7").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("B5").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("B6").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("B7").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("C5").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("C6").value = "0"
                    Worksheets("E2_CA_WTA_WDV").Range("C7").value = "0"
                    Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "B4" & ":C4", False)
                    Worksheets("E2_CA_WTA_WDV").Range("B4").value = "37.5"
                    Worksheets("E2_CA_WTA_WDV").Range("C4").value = "30"
                    Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_WDV", "B4:C4", True)
                    
                     'Wife
                        If UCase(Worksheets("A_Basic_Info").Range("RetInf.DeclareWifeIncome").value) = "YES" Then
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "D17" & ":D19", False)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "E17" & ":E19", False)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "B17" & ":B19", False)
                            Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "C17" & ":C19", False)
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
                        End If
                        Call lockUnlock_cell_rng("E2_CA_WTA_WDV", "B16" & ":C16", False)
                        Worksheets("E2_CA_WTA_WDV").Range("B16").value = "37.5"
                        Worksheets("E2_CA_WTA_WDV").Range("C16").value = "30"
                        Call lockUnlock_cell_rng_without_clearing_contents("E2_CA_WTA_WDV", "B16:C16", True)
                End If
                
                'For Section E1 Part 2
                rangeName = Worksheets("E1_IDA_CA").Range("IniAllIBD.ListPart2S").Address
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
                rangeName = Worksheets("E1_IDA_CA").Range("DeprIntengAst.ListS").Address
                startRow = Worksheets("E1_IDA_CA").Range(rangeName).row
                endRow = Worksheets("E1_IDA_CA").Range(rangeName).Rows.Count + startRow - 1
                
                If year >= 2020 Then
                    If (year = 2020 And mm >= 4) Or year > 2020 Then
                        Call lockUnlock_cell_rng("E1_IDA_CA", "A" & startRow & ":A" & endRow, True)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "B" & startRow & ":B" & endRow, False)
                        Call lockUnlock_cell_rng("E1_IDA_CA", "C" & startRow & ":C" & endRow, True)
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
                'by Palak for Enh-6 SR2 ends
End Function

Function checkIfReturnPeriodIsFilled() As Boolean
    endDate = Worksheets("A_Basic_Info").Range("RetInf.RetEndDate").value
    If endDate = "" Then
        MsgBox "Please first enter the Return Period From and Return Period To."
        Worksheets("A_Basic_Info").Activate
        Worksheets("A_Basic_Info").Range("RetInf.RetStartDate").Select
        checkIfReturnPeriodIsFilled = False
    Else
        checkIfReturnPeriodIsFilled = True
    End If
End Function

'Added by Lawrence for IT1 covid changes on 29/12/2020
Function validateJanDecIncome2020(ByRef janDecSum As Range, ByRef totalIncome As Range, ByRef cellToBeActivated As Range, ByVal RtnYear As Integer) As String

Dim rowNumber As String
Dim janDecSumLong As Long
Dim totalIncomeLong As Long

janDecSumLong = janDecSum.value
totalIncomeLong = totalIncome.value

    If (RtnYear = 2020) Then
        If totalIncomeLong - janDecSumLong = 0 Then
            rowNumber = "NE"
        Else
           rowNumber = cellToBeActivated.Address
        End If
    Else
            rowNumber = "NE"
    End If
validateJanDecIncome2020 = rowNumber
End Function

'Added by Ruth and Lawrence on 29/12/2020
'Validates if Input is Numeric
Function validateIsNumeric(ByVal targetCell As Range)
    If Not IsNumeric(targetCell.value) Then
        MsgBox "Only Numbers Allowed"
        targetCell.value = ""
        targetCell.Select
    End If
End Function

'Added by Ruth and Lawrence on 29/12/2020
'Validates if Input A is equal to input B
Function validateIsEqual(ByVal targetCell As Range, ByVal value2 As Double)
    If CStr(targetCell.value) <> "" Or CStr(value2) <> "" Then
        If Round(targetCell.value) <> Round(value2) Then
            MsgBox "Pension entered at T_Tax_Computation, should be equal to Total Pension at F_Employment_Income"
            
            targetCell.value = ""
            targetCell.Select
         End If
    End If
End Function




