Attribute VB_Name = "Module2"
#If Win64 Then
    Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
#Else
    Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
#End If


'Addded by Pranav fro Hex Code generation
'Declare Function OpenProcess Lib "kernel32" _
'                             (ByVal dwDesiredAccess As Long, _
'                              ByVal bInheritHandle As Long, _
'                              ByVal dwProcessId As Long) As Long
'
'Declare Function GetExitCodeProcess Lib "kernel32" _
'                                    (ByVal hProcess As Long, _
'                                     lpExitCode As Long) As Long
                                     
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STILL_ACTIVE = &H103
Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private State(4) As Long
Private ByteCounter As Long
Private ByteBuffer(63) As Byte
Private Const S11 = 7
Private Const S12 = 12
Private Const S13 = 17
Private Const S14 = 22
Private Const S21 = 5
Private Const S22 = 9
Private Const S23 = 14
Private Const S24 = 20
Private Const S31 = 4
Private Const S32 = 11
Private Const S33 = 16
Private Const S34 = 23
Private Const S41 = 6
Private Const S42 = 10
Private Const S43 = 15
Private Const S44 = 21

'--------------------------------------

Property Get RegisterA() As String
RegisterA = State(1)
End Property

Property Get RegisterB() As String
RegisterB = State(2)
End Property

Property Get RegisterC() As String
RegisterC = State(3)
End Property

Property Get RegisterD() As String
RegisterD = State(4)
End Property

'For Hex Code Generation

Public Function DigestStrToHexStr(SourceString As String) As String
MD5Init
MD5Update Len(SourceString), StringToArray(SourceString)
MD5Final
DigestStrToHexStr = GetValues
End Function

Public Sub MD5Update(ByVal InputLen As Long, InputBuffer() As Byte)
    Dim II As Integer, i As Long, j As Integer, k As Integer, lngBufferedBytes As Long, lngBufferRemaining As Long, lngRem As Long
    
    lngBufferedBytes = ByteCounter Mod 64
    lngBufferRemaining = 64 - lngBufferedBytes
    ByteCounter = ByteCounter + InputLen
    
    If InputLen >= lngBufferRemaining Then
    For II = 0 To lngBufferRemaining - 1
    ByteBuffer(lngBufferedBytes + II) = InputBuffer(II)
    Next II
    MD5Transform ByteBuffer
    lngRem = (InputLen) Mod 64
    For i = lngBufferRemaining To InputLen - II - lngRem Step 64
    For j = 0 To 63
    ByteBuffer(j) = InputBuffer(i + j)
    Next j
    MD5Transform ByteBuffer
    Next i
    lngBufferedBytes = 0
    Else
    i = 0
    End If
    
    If InputLen > 0 Then
    For k = 0 To InputLen - i - 1
    ByteBuffer(lngBufferedBytes + k) = InputBuffer(i + k)
    Next k
    End If
End Sub

Private Function StringToArray(InString As String) As Byte()
    Dim i As Long, bytBuffer() As Byte
    ReDim bytBuffer(Len(InString))
    
    For i = 0 To Len(InString) - 1
    bytBuffer(i) = asc(Mid$(InString, i + 1, 1))
    Next i
    StringToArray = bytBuffer
End Function
Public Function GetValues() As String
    GetValues = LongToString(State(1)) & LongToString(State(2)) & LongToString(State(3)) & LongToString(State(4))
End Function

Private Function LongToString(Num As Long) As String
    Dim a As Byte, b As Byte, c As Byte, d As Byte
    a = Num And &HFF&
    If a < 16 Then LongToString = "0" & Hex(a) Else LongToString = Hex(a)
    b = (Num And &HFF00&) \ 256
    If b < 16 Then LongToString = LongToString & "0" & Hex(b) Else LongToString = LongToString & Hex(b)
    c = (Num And &HFF0000) \ 65536
    If c < 16 Then LongToString = LongToString & "0" & Hex(c) Else LongToString = LongToString & Hex(c)
    If Num < 0 Then d = ((Num And &H7F000000) \ 16777216) Or &H80& Else d = (Num And &HFF000000) \ 16777216
    If d < 16 Then LongToString = LongToString & "0" & Hex(d) Else LongToString = LongToString & Hex(d)
End Function

Public Sub MD5Init()
    ByteCounter = 0
    State(1) = UnsignedToLong(1732584193#)
    State(2) = UnsignedToLong(4023233417#)
    State(3) = UnsignedToLong(2562383102#)
    State(4) = UnsignedToLong(271733878#)
End Sub
Private Function UnsignedToLong(value As Double) As Long
    If value < 0 Or value >= OFFSET_4 Then Error 6
    If value <= MAXINT_4 Then UnsignedToLong = value Else UnsignedToLong = value - OFFSET_4
End Function

Public Sub MD5Final()
    Dim dblBits As Double, padding(72) As Byte, lngBytesBuffered As Long
    padding(0) = &H80
    dblBits = ByteCounter * 8
    lngBytesBuffered = ByteCounter Mod 64
    If lngBytesBuffered <= 56 Then MD5Update 56 - lngBytesBuffered, padding Else MD5Update 120 - ByteCounter, padding
    padding(0) = UnsignedToLong(dblBits) And &HFF&
    padding(1) = UnsignedToLong(dblBits) \ 256 And &HFF&
    padding(2) = UnsignedToLong(dblBits) \ 65536 And &HFF&
    padding(3) = UnsignedToLong(dblBits) \ 16777216 And &HFF&
    padding(4) = 0
    padding(5) = 0
    padding(6) = 0
    padding(7) = 0
    MD5Update 8, padding
End Sub

'--------
Public Function DigestFileToHexStr(InFile As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
DigestFileToHexStr = ""
Exit Function

begin:
Dim FileO As Integer
FileO = FreeFile
Call FileLen(InFile)
Open InFile For Binary Access Read As #FileO
MD5Init
Do While Not EOF(FileO)
Get #FileO, , ByteBuffer
If Loc(FileO) < LOF(FileO) Then
ByteCounter = ByteCounter + 64
MD5Transform ByteBuffer
End If
Loop
ByteCounter = ByteCounter + (LOF(FileO) Mod 64)
Close #FileO
MD5Final
DigestFileToHexStr = GetValues
End Function

Private Sub Decode(Length As Integer, OutputBuffer() As Long, InputBuffer() As Byte)
Dim intDblIndex As Integer, intByteIndex As Integer, dblSum As Double
For intByteIndex = 0 To Length - 1 Step 4
dblSum = InputBuffer(intByteIndex) + InputBuffer(intByteIndex + 1) * 256# + InputBuffer(intByteIndex + 2) * 65536# + InputBuffer(intByteIndex + 3) * 16777216#
OutputBuffer(intDblIndex) = UnsignedToLong(dblSum)
intDblIndex = intDblIndex + 1
Next intByteIndex
End Sub


Private Sub MD5Transform(Buffer() As Byte)
Dim x(16) As Long, a As Long, b As Long, c As Long, d As Long

a = State(1)
b = State(2)
c = State(3)
d = State(4)
Decode 64, x, Buffer
FF a, b, c, d, x(0), S11, -680876936
FF d, a, b, c, x(1), S12, -389564586
FF c, d, a, b, x(2), S13, 606105819
FF b, c, d, a, x(3), S14, -1044525330
FF a, b, c, d, x(4), S11, -176418897
FF d, a, b, c, x(5), S12, 1200080426
FF c, d, a, b, x(6), S13, -1473231341
FF b, c, d, a, x(7), S14, -45705983
FF a, b, c, d, x(8), S11, 1770035416
FF d, a, b, c, x(9), S12, -1958414417
FF c, d, a, b, x(10), S13, -42063
FF b, c, d, a, x(11), S14, -1990404162
FF a, b, c, d, x(12), S11, 1804603682
FF d, a, b, c, x(13), S12, -40341101
FF c, d, a, b, x(14), S13, -1502002290
FF b, c, d, a, x(15), S14, 1236535329

GG a, b, c, d, x(1), S21, -165796510
GG d, a, b, c, x(6), S22, -1069501632
GG c, d, a, b, x(11), S23, 643717713
GG b, c, d, a, x(0), S24, -373897302
GG a, b, c, d, x(5), S21, -701558691
GG d, a, b, c, x(10), S22, 38016083
GG c, d, a, b, x(15), S23, -660478335
GG b, c, d, a, x(4), S24, -405537848
GG a, b, c, d, x(9), S21, 568446438
GG d, a, b, c, x(14), S22, -1019803690
GG c, d, a, b, x(3), S23, -187363961
GG b, c, d, a, x(8), S24, 1163531501
GG a, b, c, d, x(13), S21, -1444681467
GG d, a, b, c, x(2), S22, -51403784
GG c, d, a, b, x(7), S23, 1735328473
GG b, c, d, a, x(12), S24, -1926607734

HH a, b, c, d, x(5), S31, -378558
HH d, a, b, c, x(8), S32, -2022574463
HH c, d, a, b, x(11), S33, 1839030562
HH b, c, d, a, x(14), S34, -35309556
HH a, b, c, d, x(1), S31, -1530992060
HH d, a, b, c, x(4), S32, 1272893353
HH c, d, a, b, x(7), S33, -155497632
HH b, c, d, a, x(10), S34, -1094730640
HH a, b, c, d, x(13), S31, 681279174
HH d, a, b, c, x(0), S32, -358537222
HH c, d, a, b, x(3), S33, -722521979
HH b, c, d, a, x(6), S34, 76029189
HH a, b, c, d, x(9), S31, -640364487
HH d, a, b, c, x(12), S32, -421815835
HH c, d, a, b, x(15), S33, 530742520
HH b, c, d, a, x(2), S34, -995338651

II a, b, c, d, x(0), S41, -198630844
II d, a, b, c, x(7), S42, 1126891415
II c, d, a, b, x(14), S43, -1416354905
II b, c, d, a, x(5), S44, -57434055
II a, b, c, d, x(12), S41, 1700485571
II d, a, b, c, x(3), S42, -1894986606
II c, d, a, b, x(10), S43, -1051523
II b, c, d, a, x(1), S44, -2054922799
II a, b, c, d, x(8), S41, 1873313359
II d, a, b, c, x(15), S42, -30611744
II c, d, a, b, x(6), S43, -1560198380
II b, c, d, a, x(13), S44, 1309151649
II a, b, c, d, x(4), S41, -145523070
II d, a, b, c, x(11), S42, -1120210379
II c, d, a, b, x(2), S43, 718787259
II b, c, d, a, x(9), S44, -343485551

State(1) = LongOverflowAdd(State(1), a)
State(2) = LongOverflowAdd(State(2), b)
State(3) = LongOverflowAdd(State(3), c)
State(4) = LongOverflowAdd(State(4), d)
End Sub
Private Function FF(a As Long, b As Long, c As Long, d As Long, x As Long, S As Long, ac As Long) As Long
a = LongOverflowAdd4(a, (b And c) Or (Not (b) And d), x, ac)
a = LongLeftRotate(a, S)
a = LongOverflowAdd(a, b)
End Function

Private Function GG(a As Long, b As Long, c As Long, d As Long, x As Long, S As Long, ac As Long) As Long
a = LongOverflowAdd4(a, (b And d) Or (c And Not (d)), x, ac)
a = LongLeftRotate(a, S)
a = LongOverflowAdd(a, b)
End Function

Private Function HH(a As Long, b As Long, c As Long, d As Long, x As Long, S As Long, ac As Long) As Long
a = LongOverflowAdd4(a, b Xor c Xor d, x, ac)
a = LongLeftRotate(a, S)
a = LongOverflowAdd(a, b)
End Function

Private Function II(a As Long, b As Long, c As Long, d As Long, x As Long, S As Long, ac As Long) As Long
a = LongOverflowAdd4(a, c Xor (b Or Not (d)), x, ac)
a = LongLeftRotate(a, S)
a = LongOverflowAdd(a, b)
End Function
Function LongLeftRotate(value As Long, Bits As Long) As Long
Dim lngSign As Long, lngI As Long
Bits = Bits Mod 32
If Bits = 0 Then LongLeftRotate = value: Exit Function
For lngI = 1 To Bits
lngSign = value And &HC0000000
value = (value And &H3FFFFFFF) * 2
value = value Or ((lngSign < 0) And 1) Or (CBool(lngSign And &H40000000) And &H80000000)
Next
LongLeftRotate = value
End Function

Private Function LongOverflowAdd(Val1 As Long, Val2 As Long) As Long
Dim lngHighWord As Long, lngLowWord As Long, lngOverflow As Long
lngLowWord = (Val1 And &HFFFF&) + (Val2 And &HFFFF&)
lngOverflow = lngLowWord \ 65536
lngHighWord = (((Val1 And &HFFFF0000) \ 65536) + ((Val2 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
LongOverflowAdd = UnsignedToLong((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))
End Function

Private Function LongOverflowAdd4(Val1 As Long, Val2 As Long, val3 As Long, val4 As Long) As Long
Dim lngHighWord As Long, lngLowWord As Long, lngOverflow As Long
lngLowWord = (Val1 And &HFFFF&) + (Val2 And &HFFFF&) + (val3 And &HFFFF&) + (val4 And &HFFFF&)
lngOverflow = lngLowWord \ 65536
lngHighWord = (((Val1 And &HFFFF0000) \ 65536) + ((Val2 And &HFFFF0000) \ 65536) + ((val3 And &HFFFF0000) \ 65536) + ((val4 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
LongOverflowAdd4 = UnsignedToLong((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))
End Function



Public Sub ShellAndWait(ByVal PathName As String, Optional WindowState)
    Dim hProg As Long
    Dim hProcess As Long, ExitCode As Long
    'fill in the missing parameter and execute the program
    If IsMissing(WindowState) Then WindowState = 1
    hProg = Shell(PathName, WindowState)
    'hProg is a "process ID under Win32. To get the process handle:
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, hProg)
    Do
        'populate Exitcode variable
        GetExitCodeProcess hProcess, ExitCode
        DoEvents
    Loop While ExitCode = STILL_ACTIVE
End Sub

'With this example you browse to the folder you want to zip
'The zip file will be saved in: DefPath = Application.DefaultFilePath
'Normal if you have not change it this will be your Documents folder
'You can change this folder to this if you want to use another folder
'DefPath = "C:\Users\Ron\ZipFolder"
'There is no need to change the code before you test it

Function A_Zip_Folder_And_SubFolders_Browse(FolderName As String, fileName As String) As Boolean
    Dim PathZipProgram As String, NameZipFile As String
    Dim ShellStr As String, strDate As String, DefPath As String
    Dim Fld As Object

    'Path of the Zip program
    'PathZipProgram = "C:\program files\7-Zip\"
    Set myWS = CreateObject("WScript.Shell")
    PathZipProgram = myWS.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\7-Zip\Path")
    
    If Right(PathZipProgram, 1) <> "\" Then
        PathZipProgram = PathZipProgram & "\"
    End If

    'Check if this is the path where 7z is installed.
    If Dir(PathZipProgram & "7z.exe") = "" Then
        'MsgBox "Please find your copy of 7z.exe and try again"
        A_Zip_Folder_And_SubFolders_Browse = False
        Exit Function
    End If

    'Create Path and name of the new zip file
    'The zip file will be saved in: DefPath = Application.DefaultFilePath
    'Normal if you have not change it this will be your Documents folder
    'You can change the folder if you want to another folder like this
    'DefPath = "C:\Users\Ron\ZipFolder"
    DefPath = Application.DefaultFilePath
    If Right(DefPath, 1) <> "\" Then
        DefPath = DefPath & "\"
    End If

    'Create date/Time string, also the name of the Zip in this example
    'strDate = Format(Now, "yyyy-mm-dd h-mm-ss")

    'Set NameZipFile to the full path/name of the Zip file
    'If you want to add the word "MyZip" before the date/time use
    'NameZipFile = DefPath & "MyZip " & strDate & ".zip"
    'NameZipFile = DefPath & strDate & ".zip"
    
    NameZipFile = Application.GetSaveAsFilename(fileName, "Zip Files (*.zip), *.zip")
    
    If Dir(NameZipFile) <> "" Then
        MsgBox "The location you have selected already contains a file with same name. Please choose another name or location to save the file on validate."
        A_Zip_Folder_And_SubFolders_Browse = True
    Else
        If NameZipFile = "False" Then
            MsgBox ("Modifications Are Not Saved,Upload File Not Generated") '& Err.Description)
        Else
            'Browse to the folder with the files that you want to Zip
            'Set Fld = CreateObject("Shell.Application").BrowseForFolder(0, "Select folder to Zip", 512)
            'If Not Fld Is Nothing Then
    '        folderName = Fld.self.path
            If Right(FolderName, 1) <> "\" Then
                FolderName = FolderName & "\"
            End If
    
            'Zip all the files in the folder and subfolders, -r is Include subfolders
            ShellStr = PathZipProgram & "7z.exe a -r" _
                     & " " & Chr(34) & NameZipFile & Chr(34) _
                     & " " & Chr(34) & FolderName & fileName & ".xls" & Chr(34)
            ShellAndWait ShellStr, vbHide
            MsgBox "No Error Found,Zip file is saved: " & NameZipFile
        End If
        A_Zip_Folder_And_SubFolders_Browse = True
    End If
End Function


