' ============================================================================
'  modClashLite  -  minimal Revizto clash-test (.vimctst) generator
'
'  The SHAREABLE tool. No matrix engine, no codes/priorities/clearances.
'  Two sheets:
'     "Sets"   - column A: your search-set names (one per row, from row 2). Header in A1.
'     "Tests"  - written by the macro: the matrixed pairs (preview).
'
'  Run GenerateClashTests:
'     1. reads the Sets list,
'     2. builds every lower-triangular pair INCLUDING self-pairs (1v1, 2v1, 2v2, ...),
'     3. writes them to the Tests sheet,
'     4. saves "Clash Tests.vimctst" beside the workbook.
'  Revizto re-matches each side to a project search set BY NAME on import, so the names
'  in the Sets list must match your real Revizto set names. GUIDs are fabricated.
'
'  Paste this whole text into a standard module (Alt+F11: Insert > Module).
'  OneDrive note: ThisWorkbook.Path can be an https:// URL (run-time error 52) - run
'  from a LOCAL folder (e.g. C:\Temp) and copy the output back.
' ============================================================================
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (ByRef pGuid As Any) As Long
#Else
    Private Declare Function CoCreateGuid Lib "ole32.dll" (ByRef pGuid As Any) As Long
#End If

Private Const SHEET_SETS  As String = "Sets"
Private Const SHEET_TESTS As String = "Tests"
Private Const OUTPUT_FILE As String = "Clash Tests.vimctst"

' embedded clash-test byte template (generic defaults; GUIDs added per-test)
Private Const HEX_HEADER As String = "0A0B52436C617368546573747310011800"
Private Const HEX_TAIL4  As String = "20002801300038004800"
Private Const HEX_T5      As String = "0A1071347EF77093E54198BF66EFF01A861612107C5D60A1F2D9814AA19E1E939AFC95F71800200030003A109CBE1B0D8C76A244A613346CDC703A2A82012508001000180020002800300038004000480050005800600068007000800100880100900100880100980100A80100B80101C00103"
Private Const HEX_TAIL6   As String = "221C080010011A0F0DA0FA273E15A0FA273E1DA0FA273E200128881E58002A08080012020800180032071D0000000020003A280A0D08021209456C656D656E7420410A0808001204207673200A0D08041209456C656D656E7420424210080010001800200028003000380140004A270A1F0807120208001A050D48F9513F220A0D0000000015000000002A0208003200120408001200520808001000180020006D00000000750000000078008001008D0100000000950100000000"
Private Const OFF_TEST As Long = 2
Private Const OFF_MID  As Long = 20
Private Const OFF_RULE As Long = 44


' ===================== ENTRY POINT =====================
Public Sub GenerateClashTests()
    Dim t0 As Single: t0 = Timer

    ' ---- 1. read the Sets list (col A from row 2) ----
    Dim wsS As Worksheet
    On Error Resume Next
    Set wsS = ThisWorkbook.Worksheets(SHEET_SETS)
    On Error GoTo 0
    If wsS Is Nothing Then
        MsgBox "No sheet named '" & SHEET_SETS & "'. Put your set names in column A of a '" & SHEET_SETS & "' sheet.", vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = wsS.Cells(wsS.Rows.Count, "A").End(xlUp).Row
    Dim names() As String
    ReDim names(1 To Application.Max(1, lastRow))
    Dim n As Long, r As Long, v As String
    n = 0
    For r = 2 To lastRow
        v = Trim$(CStr(wsS.Cells(r, "A").Value))
        If Len(v) > 0 Then
            n = n + 1
            names(n) = v
        End If
    Next r
    If n < 1 Then
        MsgBox "No set names found in '" & SHEET_SETS & "' column A (from row 2).", vbExclamation
        Exit Sub
    End If

    ' ---- 2. templates ----
    Dim header() As Byte, tail4() As Byte, t5pl() As Byte, tail6() As Byte
    header = HexToBytes(HEX_HEADER)
    tail4 = HexToBytes(HEX_TAIL4)
    t5pl = HexToBytes(HEX_T5)
    tail6 = HexToBytes(HEX_TAIL6)
    Dim t5len As Long: t5len = UBound(t5pl) + 1

    ' ---- 3. buffers + Tests sheet ----
    Dim b4() As Byte, l4 As Long, b5() As Byte, l5 As Long, b6() As Byte, l6 As Long
    ReDim b4(0 To 65535): ReDim b5(0 To 65535): ReDim b6(0 To 65535)
    l4 = 0: l5 = 0: l6 = 0

    Dim wsT As Worksheet: Set wsT = EnsureSheet(SHEET_TESTS)
    wsT.Cells.Clear
    wsT.Range("A1").Value = "Element A"
    wsT.Range("B1").Value = "Element B"
    wsT.Range("C1").Value = "Test Name"
    Dim outRow As Long: outRow = 2

    Dim paths(0 To 3) As String
    paths(0) = "Clash Detection": paths(1) = "00_Categories": paths(2) = "Sets"

    ' ---- 4. matrix (lower-triangular incl. self-pairs) + emit ----
    Dim i As Long, j As Long, gen As Long
    Dim nameA As String, nameB As String, testName As String
    Dim gT() As Byte, gM() As Byte, gR() As Byte, nameB_() As Byte
    Dim nlen As Long, body4Len As Long, body6Len As Long, sideALen As Long, sideBLen As Long, k As Long

    gen = 0
    For i = 1 To n
        For j = 1 To i
            nameA = names(i)
            nameB = names(j)
            testName = nameA & " vs " & nameB

            wsT.Cells(outRow, 1).Value = nameA
            wsT.Cells(outRow, 2).Value = nameB
            wsT.Cells(outRow, 3).Value = testName
            outRow = outRow + 1

            gT = NewGuidBytes(): gM = NewGuidBytes(): gR = NewGuidBytes()
            nameB_ = AsciiBytes(testName)
            nlen = UBound(nameB_) + 1

            ' field4
            body4Len = 2 + 16 + 1 + VLen(nlen) + nlen + (UBound(tail4) + 1)
            PutB b4, l4, &H22
            PutVarint b4, l4, body4Len
            PutB b4, l4, &HA: PutB b4, l4, &H10: PutArr b4, l4, gT
            PutB b4, l4, &H12: PutVarint b4, l4, nlen: PutArr b4, l4, nameB_
            PutArr b4, l4, tail4

            ' field5 (clone, overwrite 3 guids)
            PutB b5, l5, &H2A: PutVarint b5, l5, t5len
            For k = 0 To t5len - 1
                If k >= OFF_TEST And k < OFF_TEST + 16 Then
                    PutB b5, l5, gT(k - OFF_TEST)
                ElseIf k >= OFF_MID And k < OFF_MID + 16 Then
                    PutB b5, l5, gM(k - OFF_MID)
                ElseIf k >= OFF_RULE And k < OFF_RULE + 16 Then
                    PutB b5, l5, gR(k - OFF_RULE)
                Else
                    PutB b5, l5, t5pl(k)
                End If
            Next k

            ' field6
            paths(3) = nameA: sideALen = SideLen(paths)
            paths(3) = nameB: sideBLen = SideLen(paths)
            body6Len = 2 + 16 + 1 + VLen(sideALen) + sideALen + 1 + VLen(sideBLen) + sideBLen + (UBound(tail6) + 1)
            PutB b6, l6, &H32: PutVarint b6, l6, body6Len
            PutB b6, l6, &HA: PutB b6, l6, &H10: PutArr b6, l6, gR
            paths(3) = nameA
            PutB b6, l6, &H12: PutVarint b6, l6, sideALen: PutSide b6, l6, NewGuidBytes(), paths
            paths(3) = nameB
            PutB b6, l6, &H1A: PutVarint b6, l6, sideBLen: PutSide b6, l6, NewGuidBytes(), paths
            PutArr b6, l6, tail6

            gen = gen + 1
        Next j
    Next i
    wsT.Columns("A:C").AutoFit

    ' ---- 5. write header + b4 + b5 + b6 ----
    Dim outPath As String: outPath = ThisWorkbook.Path & Application.PathSeparator & OUTPUT_FILE
    Dim f As Integer: f = FreeFile
    If Dir(outPath) <> "" Then Kill outPath
    Open outPath For Binary Access Write As #f
    PutRawArr f, header, UBound(header) + 1
    PutRawArr f, b4, l4
    PutRawArr f, b5, l5
    PutRawArr f, b6, l6
    Close #f

    MsgBox gen & " clash tests from " & n & " sets in " & Format(Timer - t0, "0.0") & "s" & vbCrLf & OUTPUT_FILE, vbInformation
End Sub


' ===================== helpers =====================
Private Function EnsureSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    End If
    Set EnsureSheet = ws
End Function

Private Function SideLen(ByRef paths() As String) As Long
    Dim m As Long, i As Long, L As Long
    m = 2 + 16
    For i = 0 To 3
        L = Len(paths(i))
        m = m + 1 + VLen(L) + L
    Next i
    SideLen = m
End Function

Private Sub PutSide(ByRef b() As Byte, ByRef pos As Long, ByRef guid() As Byte, ByRef paths() As String)
    Dim i As Long, pb() As Byte
    PutB b, pos, &HA: PutB b, pos, &H10: PutArr b, pos, guid
    For i = 0 To 3
        pb = AsciiBytes(paths(i))
        PutB b, pos, &H12: PutVarint b, pos, (UBound(pb) + 1): PutArr b, pos, pb
    Next i
End Sub

Private Sub PutB(ByRef b() As Byte, ByRef pos As Long, ByVal v As Byte)
    If pos > UBound(b) Then ReDim Preserve b(0 To (UBound(b) + 1) * 2 - 1)
    b(pos) = v: pos = pos + 1
End Sub

Private Sub PutArr(ByRef b() As Byte, ByRef pos As Long, ByRef src() As Byte)
    Dim i As Long
    For i = 0 To UBound(src): PutB b, pos, src(i): Next i
End Sub

Private Sub PutVarint(ByRef b() As Byte, ByRef pos As Long, ByVal n As Long)
    Dim x As Long
    Do
        x = n And &H7F
        n = n \ 128
        If n <> 0 Then x = x Or &H80
        PutB b, pos, CByte(x)
    Loop While n <> 0
End Sub

Private Sub PutRawArr(ByVal f As Integer, ByRef b() As Byte, ByVal n As Long)
    If n <= 0 Then Exit Sub
    Dim out() As Byte, i As Long
    ReDim out(0 To n - 1)
    For i = 0 To n - 1: out(i) = b(i): Next i
    Put #f, , out
End Sub

Private Function VLen(ByVal n As Long) As Long
    Dim c As Long: c = 1
    Do While n >= 128: n = n \ 128: c = c + 1: Loop
    VLen = c
End Function

Private Function HexToBytes(ByVal h As String) As Byte()
    Dim n As Long, i As Long, b() As Byte
    n = Len(h) \ 2
    ReDim b(0 To n - 1)
    For i = 0 To n - 1: b(i) = CByte(Val("&H" & Mid$(h, i * 2 + 1, 2))): Next i
    HexToBytes = b
End Function

Private Function AsciiBytes(ByVal s As String) As Byte()
    AsciiBytes = StrConv(s, vbFromUnicode)
End Function

Private Function NewGuidBytes() As Byte()
    Dim g(0 To 15) As Byte
    CoCreateGuid g(0)
    NewGuidBytes = g
End Function
