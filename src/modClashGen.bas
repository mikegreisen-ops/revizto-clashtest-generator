' ============================================================================
'  modClashGen  -  Revizto clash-test (.vimctst) generator  [EMBEDDED TEMPLATE]
'  Paste this ENTIRE text into a standard module in CLASH TESTS.xlsm
'  (VBA editor Alt+F11: Insert > Module, paste). Run: GenerateClashTests
'
'  SELF-CONTAINED: needs NO external files. Reads two sheets in this workbook:
'      SearchSets  - col C = 2-digit Code, col I = SearchSet name (free-form), col B = Discipline
'      MasterList  - col A = the |A|B| - ... test names (the pairs)
'  Set GUIDs are FABRICATED: Revizto re-matches search sets by NAME on import, so the set
'  GUID and folder path are cosmetic (confirmed by probe 2026-06-09). Only the set NAME in
'  col I must match a real set in the user's Revizto project. Writes:
'      DEVELOPMENT AREA clash tests_GENERATED.vimctst
'
'  v2 2026-06-08: template baked in (header/field4 tail/field5/field6 settings).
'  v3 2026-06-09: dropped the .vimsst dependency - set names now read from the SearchSets
'      sheet (col I), GUIDs fabricated. Fully self-contained / shareable.
'  If you ever want different default clash settings, make one test in Revizto,
'  export it, and re-harvest these 4 hex strings (ask Big C).
' ============================================================================
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (ByRef pGuid As Any) As Long
#Else
    Private Declare Function CoCreateGuid Lib "ole32.dll" (ByRef pGuid As Any) As Long
#End If

Private Const OUTPUT_FILE As String = "DEVELOPMENT AREA clash tests_GENERATED.vimctst"
Private Const SHEET_SETS  As String = "SearchSets"   ' source of code -> set name
Private Const COL_CODE    As String = "C"            ' 2-digit code (text)
Private Const COL_NAME    As String = "I"            ' SearchSet name (free-form, must match Revizto)
Private Const COL_DISC    As String = "B"            ' Discipline (cosmetic folder label)

' ---- embedded clash-test byte template (generic; GUIDs are added per-test) ----
Private Const HEX_HEADER As String = "0A0B52436C617368546573747310011800"
Private Const HEX_TAIL4  As String = "20002801300038004800"
Private Const HEX_T5      As String = "0A1071347EF77093E54198BF66EFF01A861612107C5D60A1F2D9814AA19E1E939AFC95F71800200030003A109CBE1B0D8C76A244A613346CDC703A2A82012508001000180020002800300038004000480050005800600068007000800100880100900100880100980100A80100B80101C00103"
Private Const HEX_TAIL6   As String = "221C080010011A0F0DA0FA273E15A0FA273E1DA0FA273E200128881E58002A08080012020800180032071D0000000020003A280A0D08021209456C656D656E7420410A0808001204207673200A0D08041209456C656D656E7420424210080010001800200028003000380140004A270A1F0807120208001A050D48F9513F220A0D0000000015000000002A0208003200120408001200520808001000180130005A6B0880E0C4B9FEFFFFFFFF011080DA80BF0318FF8597C405250000003F2D000070413000380042190A020801100018012000280030013800450000000048015001480050005800620808001000180020006D00000000750000000078008001008D0100000000950100000000"
' GUID offsets within the field5 template (f1=test, f2=mid, f7=rule)
Private Const OFF_TEST As Long = 2
Private Const OFF_MID  As Long = 20
Private Const OFF_RULE As Long = 44


' ===================== ENTRY POINT =====================
Public Sub GenerateClashTests()
    Dim t0 As Single
    t0 = Timer

    Dim folder As String
    folder = ThisWorkbook.Path & Application.PathSeparator

    ' ---------- 1. SearchSets sheet -> code -> {fabricated guid, folder, name} ----------
    ' Set names come straight from the SearchSets sheet (col I). GUIDs are fabricated and
    ' the folder is cosmetic - Revizto rebinds sets by NAME on import.
    Dim setGuid As Object, setFolder As Object, setName As Object
    Set setGuid = CreateObject("Scripting.Dictionary")
    Set setFolder = CreateObject("Scripting.Dictionary")
    Set setName = CreateObject("Scripting.Dictionary")

    Dim wsS As Worksheet
    Set wsS = ThisWorkbook.Worksheets(SHEET_SETS)
    Dim lastS As Long, rr As Long
    lastS = wsS.Cells(wsS.Rows.Count, COL_CODE).End(xlUp).Row

    Dim code As String, nm As String, disc As String
    For rr = 2 To lastS
        code = Trim$(CStr(wsS.Cells(rr, COL_CODE).Text))   ' .Text keeps the leading zero ("01")
        nm = Trim$(CStr(wsS.Cells(rr, COL_NAME).Value))    ' free-form Revizto set name
        disc = Trim$(CStr(wsS.Cells(rr, COL_DISC).Value))  ' discipline (folder label; cosmetic)
        If Len(disc) = 0 Then disc = "Categories"
        If Len(code) > 0 And Len(nm) > 0 Then
            If Not setName.Exists(code) Then
                setGuid(code) = NewGuidBytes()              ' one fabricated GUID per code
                setFolder(code) = disc
                setName(code) = nm
            End If
        End If
    Next rr

    ' ---------- 2. templates (embedded) ----------
    Dim header() As Byte, tail4() As Byte, t5pl() As Byte, tail6() As Byte
    Dim o51 As Long, o52 As Long, o57 As Long
    header = HexToBytes(HEX_HEADER)
    tail4 = HexToBytes(HEX_TAIL4)
    t5pl = HexToBytes(HEX_T5)
    tail6 = HexToBytes(HEX_TAIL6)
    o51 = OFF_TEST
    o52 = OFF_MID
    o57 = OFF_RULE

    ' ---------- 3. MasterList column A ----------
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("MasterList")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim arr As Variant
    arr = ws.Range("A1:A" & lastRow).Value

    ' ---------- 4/5. generate + write ----------
    Dim genCount As Long
    genCount = GenAll(arr, setGuid, setFolder, setName, header, tail4, t5pl, o51, o52, o57, tail6, folder & OUTPUT_FILE)

    MsgBox "Generated " & genCount & " clash tests in " & Format(Timer - t0, "0.0") & "s" & vbCrLf & OUTPUT_FILE, vbInformation
End Sub


' ===================== MAIN WORKER =====================
Private Function GenAll(ByVal arr As Variant, ByVal setGuid As Object, ByVal setFolder As Object, _
                        ByVal setName As Object, ByRef header() As Byte, ByRef tail4() As Byte, _
                        ByRef t5pl() As Byte, ByVal o51 As Long, ByVal o52 As Long, ByVal o57 As Long, _
                        ByRef tail6() As Byte, ByVal outPath As String) As Long

    Dim b4() As Byte, l4 As Long
    Dim b5() As Byte, l5 As Long
    Dim b6() As Byte, l6 As Long
    ReDim b4(0 To 65535)
    ReDim b5(0 To 65535)
    ReDim b6(0 To 65535)
    l4 = 0
    l5 = 0
    l6 = 0

    Dim t5len As Long, tail4Len As Long, tail6Len As Long
    t5len = UBound(t5pl) + 1
    tail4Len = UBound(tail4) + 1
    tail6Len = UBound(tail6) + 1

    Dim paths(0 To 3) As String
    paths(0) = "Clash Detection"
    paths(1) = "00_Categories"

    Dim gen As Long
    Dim r As Long
    Dim v As String, ca As String, cb As String
    Dim p2 As Long, p3 As Long, k As Long
    Dim gT() As Byte, gM() As Byte, gR() As Byte, nameB() As Byte
    Dim guidA() As Byte, guidB() As Byte
    Dim nlen As Long, body4Len As Long, body6Len As Long, sideALen As Long, sideBLen As Long

    gen = 0
    For r = 1 To UBound(arr, 1)
        v = CStr(arr(r, 1) & "")
        If Len(v) > 0 Then
            If Left$(v, 1) = "|" Then
                p2 = InStr(2, v, "|")
                p3 = InStr(p2 + 1, v, "|")
                ca = Mid$(v, 2, p2 - 2)
                cb = Mid$(v, p2 + 1, p3 - p2 - 1)
                If setGuid.Exists(ca) And setGuid.Exists(cb) Then
                    gT = NewGuidBytes()
                    gM = NewGuidBytes()
                    gR = NewGuidBytes()
                    nameB = AsciiBytes(v)
                    nlen = UBound(nameB) + 1

                    ' ----- field4 -----
                    body4Len = 2 + 16 + 1 + VLen(nlen) + nlen + tail4Len
                    PutB b4, l4, &H22
                    PutVarint b4, l4, body4Len
                    PutB b4, l4, &HA
                    PutB b4, l4, &H10
                    PutArr b4, l4, gT
                    PutB b4, l4, &H12
                    PutVarint b4, l4, nlen
                    PutArr b4, l4, nameB
                    PutArr b4, l4, tail4

                    ' ----- field5 (clone, overwrite 3 GUIDs) -----
                    PutB b5, l5, &H2A
                    PutVarint b5, l5, t5len
                    For k = 0 To t5len - 1
                        If k >= o51 And k < o51 + 16 Then
                            PutB b5, l5, gT(k - o51)
                        ElseIf k >= o52 And k < o52 + 16 Then
                            PutB b5, l5, gM(k - o52)
                        ElseIf k >= o57 And k < o57 + 16 Then
                            PutB b5, l5, gR(k - o57)
                        Else
                            PutB b5, l5, t5pl(k)
                        End If
                    Next k

                    ' ----- field6 -----
                    guidA = setGuid(ca)
                    guidB = setGuid(cb)

                    paths(2) = CStr(setFolder(ca))
                    paths(3) = CStr(setName(ca))
                    sideALen = SideLen(paths)

                    paths(2) = CStr(setFolder(cb))
                    paths(3) = CStr(setName(cb))
                    sideBLen = SideLen(paths)

                    body6Len = 2 + 16 + 1 + VLen(sideALen) + sideALen + 1 + VLen(sideBLen) + sideBLen + tail6Len
                    PutB b6, l6, &H32
                    PutVarint b6, l6, body6Len
                    PutB b6, l6, &HA
                    PutB b6, l6, &H10
                    PutArr b6, l6, gR

                    paths(2) = CStr(setFolder(ca))
                    paths(3) = CStr(setName(ca))
                    PutB b6, l6, &H12
                    PutVarint b6, l6, sideALen
                    PutSide b6, l6, guidA, paths

                    paths(2) = CStr(setFolder(cb))
                    paths(3) = CStr(setName(cb))
                    PutB b6, l6, &H1A
                    PutVarint b6, l6, sideBLen
                    PutSide b6, l6, guidB, paths

                    PutArr b6, l6, tail6

                    gen = gen + 1
                End If
            End If
        End If
    Next r

    ' ----- write header + b4 + b5 + b6 -----
    Dim f As Integer
    f = FreeFile
    If Dir(outPath) <> "" Then Kill outPath
    Open outPath For Binary Access Write As #f
    PutRawArr f, header, UBound(header) + 1
    PutRawArr f, b4, l4
    PutRawArr f, b5, l5
    PutRawArr f, b6, l6
    Close #f

    GenAll = gen
End Function


' length of one side message: 2 + 16 + sum(1 + VLen(len) + len) over 4 path strings (ASCII)
Private Function SideLen(ByRef paths() As String) As Long
    Dim n As Long, i As Long, L As Long
    n = 2 + 16
    For i = 0 To 3
        L = Len(paths(i))
        n = n + 1 + VLen(L) + L
    Next i
    SideLen = n
End Function

Private Sub PutSide(ByRef b() As Byte, ByRef pos As Long, ByRef guid() As Byte, ByRef paths() As String)
    Dim i As Long
    Dim pb() As Byte
    PutB b, pos, &HA
    PutB b, pos, &H10
    PutArr b, pos, guid
    For i = 0 To 3
        pb = AsciiBytes(paths(i))
        PutB b, pos, &H12
        PutVarint b, pos, (UBound(pb) + 1)
        PutArr b, pos, pb
    Next i
End Sub


' ===================== byte-buffer helpers =====================
Private Sub PutB(ByRef b() As Byte, ByRef pos As Long, ByVal v As Byte)
    If pos > UBound(b) Then ReDim Preserve b(0 To (UBound(b) + 1) * 2 - 1)
    b(pos) = v
    pos = pos + 1
End Sub

Private Sub PutArr(ByRef b() As Byte, ByRef pos As Long, ByRef src() As Byte)
    Dim i As Long
    For i = 0 To UBound(src)
        PutB b, pos, src(i)
    Next i
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
    Dim out() As Byte
    Dim i As Long
    ReDim out(0 To n - 1)
    For i = 0 To n - 1
        out(i) = b(i)
    Next i
    Put #f, , out
End Sub


' ===================== file + protobuf + guid helpers =====================
Private Function ReadFileBytes(ByVal path As String) As Byte()
    Dim f As Integer, n As Long
    Dim b() As Byte
    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    If n > 0 Then
        ReDim b(0 To n - 1)
        Get #f, , b
    End If
    Close #f
    ReadFileBytes = b
End Function

Private Function HexToBytes(ByVal h As String) As Byte()
    Dim n As Long, i As Long
    Dim b() As Byte
    n = Len(h) \ 2
    ReDim b(0 To n - 1)
    For i = 0 To n - 1
        b(i) = CByte(Val("&H" & Mid$(h, i * 2 + 1, 2)))
    Next i
    HexToBytes = b
End Function

' Returns Double so huge varints (e.g. .vimsst timestamps) never overflow.
Private Function ReadVarint(ByRef b() As Byte, ByRef pos As Long) As Double
    Dim result As Double, shift As Long, x As Long
    result = 0
    shift = 0
    Do
        x = b(pos)
        pos = pos + 1
        result = result + (x And &H7F) * (2 ^ shift)
        If (x And &H80) = 0 Then Exit Do
        shift = shift + 7
    Loop
    ReadVarint = result
End Function

Private Function BytesToStr(ByRef b() As Byte, ByVal s As Long, ByVal L As Long) As String
    Dim i As Long, r As String
    r = ""
    For i = 0 To L - 1
        r = r & Chr$(b(s + i))
    Next i
    BytesToStr = r
End Function

Private Function AsciiBytes(ByVal s As String) As Byte()
    AsciiBytes = StrConv(s, vbFromUnicode)
End Function

Private Function NewGuidBytes() As Byte()
    Dim g(0 To 15) As Byte
    CoCreateGuid g(0)
    NewGuidBytes = g
End Function

' 36-char GUID string -> 16 bytes in .NET Guid.ToByteArray() order
Private Function GuidStringToBytes(ByVal s As String) As Byte()
    Dim raw(0 To 15) As Byte, o(0 To 15) As Byte, i As Long
    s = Replace(s, "-", "")
    For i = 0 To 15
        raw(i) = CByte(Val("&H" & Mid$(s, i * 2 + 1, 2)))
    Next i
    o(0) = raw(3): o(1) = raw(2): o(2) = raw(1): o(3) = raw(0)
    o(4) = raw(5): o(5) = raw(4)
    o(6) = raw(7): o(7) = raw(6)
    For i = 8 To 15
        o(i) = raw(i)
    Next i
    GuidStringToBytes = o
End Function

Private Function VLen(ByVal n As Long) As Long
    Dim c As Long
    c = 1
    Do While n >= 128
        n = n \ 128
        c = c + 1
    Loop
    VLen = c
End Function

' Parse one protobuf message [s, s+length). Collection of 5-element Variant arrays:
'   (0)=field#  (1)=wire  (2)=payloadStart/valuePos  (3)=len(wire2)  (4)=varintValue(wire0)
Private Function ParseFields(ByRef b() As Byte, ByVal s As Long, ByVal length As Long) As Collection
    Dim c As Collection
    Dim p As Long, e As Long
    Dim tag As Long, fld As Long, wire As Long, L As Long
    Dim d() As Variant
    Set c = New Collection
    p = s
    e = s + length
    Do While p < e
        tag = ReadVarint(b, p)
        fld = tag \ 8
        wire = tag And 7
        ReDim d(0 To 4)
        d(0) = fld
        d(1) = wire
        d(2) = 0
        d(3) = 0
        d(4) = 0
        If wire = 2 Then
            L = ReadVarint(b, p)
            d(2) = p
            d(3) = L
            p = p + L
        ElseIf wire = 0 Then
            d(2) = p
            d(4) = ReadVarint(b, p)
        ElseIf wire = 5 Then
            d(2) = p
            d(3) = 4
            p = p + 4
        ElseIf wire = 1 Then
            d(2) = p
            d(3) = 8
            p = p + 8
        Else
            Exit Do
        End If
        c.Add d
    Loop
    Set ParseFields = c
End Function

Private Function Fld(ByVal c As Collection, ByVal n As Long) As Variant
    Dim it As Variant
    For Each it In c
        If it(0) = n Then
            Fld = it
            Exit Function
        End If
    Next it
    Fld = Empty
End Function
