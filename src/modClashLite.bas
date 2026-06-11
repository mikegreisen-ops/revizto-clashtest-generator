' ============================================================================
'  modClashLite  -  minimal Revizto clash-test (.vimctst) generator
'
'  The SHAREABLE tool. No matrix engine. Sheets:
'     "Sets"   - column A: your search-set names (one per row, from row 2). Header in A1.
'     "Tests"  - written by BuildTestList; read by ExportClashTests. Columns:
'                   A  Test Name        "<SetA> vs <SetB>"  (rename freely; it's the test's name)
'                   B  Tolerance (mm)   tolerance distance   (default 25)
'                   C  Clearance (mm)   clearance distance   (default blank)
'                   D  Priority         Trivial/Minor/Major/Critical/Blocker (default Minor)
'                   E  Stamp            Revizto stamp code, e.g. "0AR" - CASE-SENSITIVE (blank = none)
'                   F  Type             "Self" / "Cross" - filter to bulk-remove self-pairs; ignored on export
'                   (B-F centred, header row styled, set columns hidden far right - see LayOutTestsSheet)
'                   Y  Set A  ) hidden helper columns parked far right - the two search sets this
'                   Z  Set B  ) row clashes. Generate reads these; don't normally touch them.
'                Grouping is fixed at 15000mm on every test (Revizto manages it globally).
'
'  TWO STEPS:
'     1) BuildTestList          - Sets -> Tests sheet: every lower-triangular pair INCLUDING
'                              self-pairs (1v1, 2v1, 2v2, ...), with default Tol/Clr/Priority.
'                              Re-running PRESERVES any settings you'd already edited (matched
'                              by Test Name) and refreshes the pair list from Sets.
'     2) ExportClashTests  - Tests sheet -> "Clash Tests.vimctst" beside the workbook:
'                              one test per row, using your edited settings. Does NOT rebuild
'                              the matrix, so deleting rows / editing settings just works.
'  So: run BuildTestList, adjust Tol/Clr/Priority (or delete rows you don't want), then
'  run ExportClashTests.
'  ALREADY HAVE A LIST? The Tests sheet exists from the start - just paste your rows under the
'  headers (or run NewTestsSheet for a fresh blank one), then ExportClashTests. No BuildTestList needed.
'
'  Revizto re-matches each side to a project search set BY NAME on import, so the names
'  must match your real Revizto set names. GUIDs are fabricated.
'
'  Settings notes:
'   - Tolerance and Clearance are OPPOSITE functions; a test is one or the other, never
'     both. If any row has BOTH filled, ExportClashTests STOPS and lists those rows -
'     nothing is written until you clear one of the two on each.
'   - Distances are stored in the file in FEET (mm / 304.8); the macro converts for you.
'   - Blank Tol AND blank Clr => no proximity check (mode 0). Blank/None Priority => none.
'   - Per-test settings are written by overwriting fixed-length byte slots that the macro
'     LOCATES at run time by parsing the template (LocateSettings) - so re-harvesting the
'     HEX_TAIL6 template from a newer Revizto export won't break the offsets.
'
'  Paste this whole text into a standard module (Alt+F11: Insert > Module).
'  OneDrive note: ThisWorkbook.Path can be an https:// URL (run-time error 52) - run
'  from a LOCAL folder (e.g. C:\Temp) and copy the output back.
' ============================================================================
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (ByRef pGuid As Any) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dst As Any, ByRef src As Any, ByVal n As Long)
#Else
    Private Declare Function CoCreateGuid Lib "ole32.dll" (ByRef pGuid As Any) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dst As Any, ByRef src As Any, ByVal n As Long)
#End If

Private Const SHEET_SETS  As String = "Sets"
Private Const SHEET_TESTS As String = "Tests"
Private Const OUTPUT_FILE As String = "Clash Tests.vimctst"
Private Const MM_PER_FOOT As Double = 304.8     ' .vimctst stores clash distances in FEET

' ---- Tests-sheet columns (1-based) + defaults ----
Private Const COL_NAME  As Long = 1             ' A: test name
Private Const COL_TOL   As Long = 2             ' B: tolerance (mm)
Private Const COL_CLR   As Long = 3             ' C: clearance (mm)
Private Const COL_PRIO  As Long = 4             ' D: priority (text)
Private Const COL_STAMP As Long = 5             ' E: stamp code (text, e.g. "0AR"; blank = none)
Private Const COL_TYPE  As Long = 6             ' F: "Self" / "Cross" - for filtering self-pairs; ignored on export
' Set A / Set B are hidden helper columns parked FAR to the right (Y/Z) so they're well clear of
' the editable columns - resizing Priority etc. can't accidentally catch them. Cols F..X stay blank.
' Keep these indices in sync with the "Y"/"Z" column letters used in LayOutTestsSheet.
Private Const COL_SETA  As Long = 25            ' Y: set A name (hidden helper)
Private Const COL_SETB  As Long = 26            ' Z: set B name (hidden helper)
Private Const N_COLS    As Long = 26
Private Const DEF_TOL   As String = "25"
Private Const DEF_CLR   As String = ""
Private Const DEF_PRIO  As String = "Minor"
Private Const DEF_STAMP As String = ""
' Grouping is NOT a column - baked into every test at this default (Revizto manages it globally).
' NOTE: Revizto DISPLAYS twice the stored grouping distance, so the export stores GRP_MM / 2
' (a stored 7.5m-equivalent shows as 15m). GRP_MM is the value as shown in Revizto.
Private Const GRP_MM    As Double = 15000

' embedded clash-test byte template (current Revizto schema; GUIDs added per-test).
' header / field4-tail / field5 are unchanged from the older template; only field6 grew
' (it now carries f7 name-template, f8, f9 grouping, f10 priority/stamp, f11 display).
' HEX_TAIL6 below was re-harvested 2026-06-11 from a clean "tolerance 25mm, no priority"
' export so the per-test pokes have real slots to write into.
Private Const HEX_HEADER As String = "0A0B52436C617368546573747310011800"
Private Const HEX_TAIL4  As String = "20002801300038004800"
Private Const HEX_T5      As String = "0A1071347EF77093E54198BF66EFF01A861612107C5D60A1F2D9814AA19E1E939AFC95F71800200030003A109CBE1B0D8C76A244A613346CDC703A2A82012508001000180020002800300038004000480050005800600068007000800100880100900100880100980100A80100B80101C00103"
Private Const HEX_TAIL6   As String = "221C080010011A0F0DA0FAA73D15A0FAA73D1DA0FAA73D200228881E58002A08080012020800180032071D0000000020003A280A0D08021209456C656D656E7420410A0808001204207673200A0D08041209456C656D656E7420424210080010001800200028003000380140004A270A1F0807120208001A050DB3D9C441220A0D0000000015000000002A0208003200120408001200520808001000180130005A6B0880E0C4B9FEFFFFFFFF011080DA80BF0318FF8597C405250000003F2D000070413000380042190A020801100018012000280030013800450000000048015001480050005800620808001000180020006D00000000750000000078008001008D0100000000950100000000"
Private Const OFF_TEST As Long = 2
Private Const OFF_MID  As Long = 20
Private Const OFF_RULE As Long = 44


' ===================== STEP 1: BUILD THE TESTS SHEET =====================
Public Sub BuildTestList()
    Dim t0 As Single: t0 = Timer

    ' ---- read the Sets list (col A from row 2) ----
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

    ' ---- snapshot existing per-test settings (preserve edits across builds) ----
    Dim prev As Object: Set prev = CreateObject("Scripting.Dictionary")
    Dim wsT As Worksheet
    On Error Resume Next
    Set wsT = ThisWorkbook.Worksheets(SHEET_TESTS)
    On Error GoTo 0
    If Not wsT Is Nothing Then SnapshotSettings wsT, prev

    ' ---- build every lower-triangular pair (incl. self-pairs) into one array ----
    Dim total As Long: total = n * (n + 1) \ 2
    Dim outData() As Variant: ReDim outData(1 To total, 1 To N_COLS)
    Dim i As Long, j As Long, oi As Long
    Dim nameA As String, nameB As String, testName As String
    Dim sTol As String, sClr As String, sPrio As String, sStamp As String, pvv As Variant
    oi = 0
    For i = 1 To n
        For j = 1 To i
            nameA = names(i)
            nameB = names(j)
            testName = nameA & " vs " & nameB
            If prev.Exists(testName) Then
                pvv = prev(testName)
                sTol = pvv(0): sClr = pvv(1): sPrio = pvv(2): sStamp = pvv(3)
            Else
                sTol = DEF_TOL: sClr = DEF_CLR: sPrio = DEF_PRIO: sStamp = DEF_STAMP
            End If
            oi = oi + 1
            outData(oi, COL_NAME) = testName
            outData(oi, COL_TOL) = sTol
            outData(oi, COL_CLR) = sClr
            outData(oi, COL_PRIO) = sPrio
            outData(oi, COL_STAMP) = sStamp
            outData(oi, COL_TYPE) = IIf(i = j, "Self", "Cross")
            outData(oi, COL_SETA) = nameA
            outData(oi, COL_SETB) = nameB
        Next j
    Next i

    ' ---- write the sheet in one batch ----
    Dim su As Boolean: su = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set wsT = EnsureSheet(SHEET_TESTS)
    wsT.Cells.Clear
    If total > 0 Then wsT.Range("A2").Resize(total, N_COLS).Value = outData
    LayOutTestsSheet wsT
    Application.ScreenUpdating = su

    MsgBox "Built " & total & " tests from " & n & " sets in " & Format(Timer - t0, "0.0") & "s." & vbCrLf & vbCrLf & _
           "Adjust Tolerance / Clearance / Priority / Stamp on the '" & SHEET_TESTS & "' sheet (or delete rows you " & _
           "don't want), then run ExportClashTests.", vbInformation
End Sub


' ===================== START A BLANK TESTS SHEET (paste your own list) =====================
' For when you already have a test list/matrix to paste in: makes an empty, formatted Tests sheet.
' Paste your rows under the headers (Test Name as "SetA vs SetB", plus Tolerance/Clearance/Priority/
' Stamp), then run ExportClashTests. (BuildTestList builds the pairings for you instead.)
Public Sub NewTestsSheet()
    Dim wsT As Worksheet
    On Error Resume Next
    Set wsT = ThisWorkbook.Worksheets(SHEET_TESTS)
    On Error GoTo 0
    If Not wsT Is Nothing Then
        If wsT.Cells(wsT.Rows.Count, "A").End(xlUp).Row >= 2 Then
            If MsgBox("The '" & SHEET_TESTS & "' sheet already has rows - clear it and start blank?", _
                      vbExclamation Or vbOKCancel, "Reset Tests sheet?") <> vbOK Then Exit Sub
        End If
    End If
    Dim su As Boolean: su = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set wsT = EnsureSheet(SHEET_TESTS)
    wsT.Cells.Clear
    LayOutTestsSheet wsT
    Application.ScreenUpdating = su
    wsT.Activate
    MsgBox "Blank '" & SHEET_TESTS & "' sheet ready." & vbCrLf & vbCrLf & _
           "Paste your tests under the headers - Test Name as ""SetA vs SetB"", plus Tolerance / " & _
           "Clearance / Priority / Stamp - then run ExportClashTests.", vbInformation
End Sub


' ===================== STEP 2: GENERATE THE .vimctst =====================
Public Sub ExportClashTests()
    Dim t0 As Single: t0 = Timer

    ' ---- read the Tests sheet (the source of truth) ----
    Dim wsT As Worksheet
    On Error Resume Next
    Set wsT = ThisWorkbook.Worksheets(SHEET_TESTS)
    On Error GoTo 0
    If wsT Is Nothing Then
        MsgBox "No '" & SHEET_TESTS & "' sheet yet. Run BuildTestList first.", vbExclamation
        Exit Sub
    End If
    Dim last As Long: last = wsT.Cells(wsT.Rows.Count, "A").End(xlUp).Row
    If last < 2 Then
        MsgBox "The '" & SHEET_TESTS & "' sheet has no test rows. Run BuildTestList first.", vbExclamation
        Exit Sub
    End If
    Dim rowsArr As Variant: rowsArr = wsT.Range("A2:Z" & last).Value

    ' rows hidden by the AutoFilter (or hidden manually) count as excluded, so "filter out the
    ' Self-pairs" really leaves them out of the file. Collect the set of VISIBLE sheet rows.
    Dim visRows As Object: Set visRows = CreateObject("Scripting.Dictionary")
    Dim vArea As Range, vr0 As Long
    On Error Resume Next
    For Each vArea In wsT.Range("A2:A" & last).SpecialCells(xlCellTypeVisible).Areas
        For vr0 = vArea.Row To vArea.Row + vArea.Rows.Count - 1
            visRows(vr0) = True
        Next vr0
    Next vArea
    On Error GoTo 0

    ' ---- validate: a test may have Tol OR Clr, never both. Stop + list any that have both. ----
    Dim vr As Long, nConflict As Long, conflicts As String, dT As Double, dC As Double, vName As String
    For vr = 1 To UBound(rowsArr, 1)
        vName = Trim$(CStr(rowsArr(vr, COL_NAME) & ""))
        If Len(vName) > 0 And visRows.Exists(vr + 1) Then
            If ParseMm(CStr(rowsArr(vr, COL_TOL) & ""), dT) And ParseMm(CStr(rowsArr(vr, COL_CLR) & ""), dC) Then
                nConflict = nConflict + 1
                If nConflict <= 15 Then conflicts = conflicts & vbCrLf & "  - " & vName
            End If
        End If
    Next vr
    If nConflict > 0 Then
        Dim cmsg As String
        cmsg = nConflict & " test(s) have BOTH a Tolerance and a Clearance set - Revizto can't do both " & _
               "(they're opposite functions). Clear one of the two on each row, then run again:" & conflicts
        If nConflict > 15 Then cmsg = cmsg & vbCrLf & "  ... and " & (nConflict - 15) & " more."
        MsgBox cmsg, vbExclamation, "Tol/Clr conflict - nothing generated"
        Exit Sub
    End If

    ' ---- templates + writable-slot offsets ----
    Dim header() As Byte, tail4() As Byte, t5pl() As Byte, tail6() As Byte
    header = HexToBytes(HEX_HEADER)
    tail4 = HexToBytes(HEX_TAIL4)
    t5pl = HexToBytes(HEX_T5)
    tail6 = HexToBytes(HEX_TAIL6)
    Dim t5len As Long: t5len = UBound(t5pl) + 1
    Dim tail6Len As Long: tail6Len = UBound(tail6) + 1

    Dim clrA As Long, clrB As Long, clrC As Long, modeOff As Long, grpOff As Long
    LocateSettings tail6, tail6Len, clrA, clrB, clrC, modeOff, grpOff
    PokeSingle tail6, grpOff, CSng(GRP_MM / 2# / MM_PER_FOOT)   ' bake fixed grouping (Revizto shows 2x the stored value)

    ' ---- buffers ----
    Dim b4() As Byte, l4 As Long, b5() As Byte, l5 As Long, b6() As Byte, l6 As Long
    ReDim b4(0 To 65535): ReDim b5(0 To 65535): ReDim b6(0 To 65535)
    l4 = 0: l5 = 0: l6 = 0

    Dim paths(0 To 3) As String
    paths(0) = "Clash Detection": paths(1) = "00_Categories": paths(2) = "Sets"

    Dim r As Long, gen As Long, skipped As Long, badPrioCount As Long, hiddenCount As Long
    Dim testName As String, nameA As String, nameB As String, pvs As Long
    Dim sTol As String, sClr As String, sPrio As String, sStamp As String
    Dim gT() As Byte, gM() As Byte, gR() As Byte, nameBytes() As Byte, tailDyn() As Byte
    Dim nlen As Long, body4Len As Long, body6Len As Long, sideALen As Long, sideBLen As Long, k As Long
    Dim prioVal As Long, known As Boolean, dynLen As Long

    For r = 1 To UBound(rowsArr, 1)
        testName = Trim$(CStr(rowsArr(r, COL_NAME) & ""))
        If Len(testName) > 0 Then
          If visRows.Exists(r + 1) Then
            nameA = Trim$(CStr(rowsArr(r, COL_SETA) & ""))
            nameB = Trim$(CStr(rowsArr(r, COL_SETB) & ""))
            ' fallback for a hand-added row: split the test name on " vs "
            If Len(nameA) = 0 Or Len(nameB) = 0 Then
                pvs = InStr(testName, " vs ")
                If pvs > 0 Then
                    nameA = Trim$(Left$(testName, pvs - 1))
                    nameB = Trim$(Mid$(testName, pvs + 4))
                End If
            End If

            If Len(nameA) > 0 And Len(nameB) > 0 Then
                sTol = CStr(rowsArr(r, COL_TOL) & "")
                sClr = CStr(rowsArr(r, COL_CLR) & "")
                sPrio = CStr(rowsArr(r, COL_PRIO) & "")
                sStamp = Trim$(CStr(rowsArr(r, COL_STAMP) & ""))

                gT = NewGuidBytes(): gM = NewGuidBytes(): gR = NewGuidBytes()
                nameBytes = AsciiBytes(testName)
                nlen = UBound(nameBytes) + 1

                ' field4
                body4Len = 2 + 16 + 1 + VLen(nlen) + nlen + (UBound(tail4) + 1)
                PutB b4, l4, &H22
                PutVarint b4, l4, body4Len
                PutB b4, l4, &HA: PutB b4, l4, &H10: PutArr b4, l4, gT
                PutB b4, l4, &H12: PutVarint b4, l4, nlen: PutArr b4, l4, nameBytes
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

                ' build the per-test tail FIRST: proximity is a fixed poke, but priority+stamp
                ' rebuild f10 and change the tail length - so compute body6Len from the result.
                tailDyn = CloneBytes(tail6, tail6Len)
                ApplyProximity tailDyn, sTol, sClr, clrA, clrB, clrC, modeOff
                prioVal = PriorityValue(sPrio, known)
                If Not known Then badPrioCount = badPrioCount + 1
                tailDyn = RebuildF10(tailDyn, tail6Len, prioVal, sStamp, dynLen)

                ' field6 header + sides + tail
                paths(3) = nameA: sideALen = SideLen(paths)
                paths(3) = nameB: sideBLen = SideLen(paths)
                body6Len = 2 + 16 + 1 + VLen(sideALen) + sideALen + 1 + VLen(sideBLen) + sideBLen + dynLen
                PutB b6, l6, &H32: PutVarint b6, l6, body6Len
                PutB b6, l6, &HA: PutB b6, l6, &H10: PutArr b6, l6, gR
                paths(3) = nameA
                PutB b6, l6, &H12: PutVarint b6, l6, sideALen: PutSide b6, l6, NewGuidBytes(), paths
                paths(3) = nameB
                PutB b6, l6, &H1A: PutVarint b6, l6, sideBLen: PutSide b6, l6, NewGuidBytes(), paths
                PutArr b6, l6, tailDyn

                gen = gen + 1
            Else
                skipped = skipped + 1
            End If
          Else
            hiddenCount = hiddenCount + 1
          End If
        End If
    Next r

    If gen = 0 Then
        MsgBox "No usable test rows found on the '" & SHEET_TESTS & "' sheet.", vbExclamation
        Exit Sub
    End If

    Dim genSecs As Single: genSecs = Timer - t0

    ' ---- ask where to save (default name + the workbook's folder); never silently overwrite ----
    Dim defPath As String: defPath = ThisWorkbook.Path
    If InStr(1, defPath, "://") > 0 Then defPath = ""          ' OneDrive/SharePoint URL - let the dialog choose the folder
    If Len(defPath) > 0 Then defPath = defPath & Application.PathSeparator
    defPath = defPath & OUTPUT_FILE
    Dim chosen As Variant
    chosen = Application.GetSaveAsFilename(InitialFileName:=defPath, _
                 FileFilter:="Revizto clash tests (*.vimctst), *.vimctst", Title:="Save clash tests as")
    If VarType(chosen) = vbBoolean Then Exit Sub               ' user cancelled the Save dialog
    Dim outPath As String: outPath = CStr(chosen)

    ' ---- write header + b4 + b5 + b6 ----
    Dim f As Integer: f = FreeFile
    If Dir(outPath) <> "" Then Kill outPath
    Open outPath For Binary Access Write As #f
    PutRawArr f, header, UBound(header) + 1
    PutRawArr f, b4, l4
    PutRawArr f, b5, l5
    PutRawArr f, b6, l6
    Close #f

    Dim msg As String
    msg = gen & " clash tests written in " & Format(genSecs, "0.0") & "s" & vbCrLf & outPath
    If hiddenCount > 0 Then msg = msg & vbCrLf & vbCrLf & hiddenCount & " filtered/hidden row(s) were NOT exported."
    If skipped > 0 Then msg = msg & vbCrLf & vbCrLf & skipped & " row(s) skipped (no Set A / Set B - rebuild with BuildTestList)."
    If badPrioCount > 0 Then msg = msg & vbCrLf & badPrioCount & " row(s) had an unrecognised Priority - treated as none (use Trivial/Minor/Major/Critical/Blocker)."
    MsgBox msg, vbInformation
End Sub


' ===================== IMPORT: .vimctst -> Tests sheet =====================
' The inverse of ExportClashTests: read an existing Revizto clash-test export and
' populate the Tests sheet (name, set A/B, tol/clr, priority) so you can edit the
' settings and re-export. Grouping/stamp/display settings are not surfaced (Lite scope).
Public Sub ImportClashTests()
    Dim path As String: path = PickVimctst()
    If Len(path) = 0 Then Exit Sub

    Dim b() As Byte: b = ReadFileBytes(path)
    If (Not Not b) = 0 Then
        MsgBox "That file is empty (0 bytes). If it came from a OneDrive folder, re-export to a local folder.", vbExclamation
        Exit Sub
    End If

    ' collect field4 (identity) and field6 (sides+settings) records in document order;
    ' the file is grouped [field4 x N][field5 x N][field6 x N], so they pair by index.
    Dim top As Collection: Set top = ParseFields(b, 0, UBound(b) + 1)
    Dim f4recs As Collection: Set f4recs = New Collection
    Dim f6recs As Collection: Set f6recs = New Collection
    Dim it As Variant
    For Each it In top
        If it(1) = 2 Then
            If it(0) = 4 Then f4recs.Add it
            If it(0) = 6 Then f6recs.Add it
        End If
    Next it

    Dim nTests As Long: nTests = Application.Min(f4recs.Count, f6recs.Count)
    If nTests < 1 Then
        MsgBox "No clash tests found in that file." & vbCrLf & path, vbExclamation
        Exit Sub
    End If

    ' confirm before clobbering an existing Tests sheet
    Dim wsT As Worksheet
    On Error Resume Next
    Set wsT = ThisWorkbook.Worksheets(SHEET_TESTS)
    On Error GoTo 0
    If Not wsT Is Nothing Then
        If wsT.Cells(wsT.Rows.Count, "A").End(xlUp).Row >= 2 Then
            If MsgBox("The '" & SHEET_TESTS & "' sheet already has tests. Replace them with the " & _
                      nTests & " test(s) from this file?", vbExclamation Or vbOKCancel, "Replace tests?") <> vbOK Then
                Exit Sub
            End If
        End If
    End If

    ' decode each test into a row
    Dim outData() As Variant: ReDim outData(1 To nTests, 1 To N_COLS)
    Dim i As Long
    Dim nm As String, setA As String, setB As String, tolS As String, clrS As String
    Dim dirCount As Long
    For i = 1 To nTests
        nm = ReadTestName(b, f4recs(i))
        DecodeSides b, f6recs(i), setA, setB
        DecodeProximity b, f6recs(i), tolS, clrS, dirCount
        outData(i, COL_NAME) = nm
        outData(i, COL_TOL) = tolS
        outData(i, COL_CLR) = clrS
        outData(i, COL_PRIO) = DecodePriority(b, f6recs(i))
        outData(i, COL_STAMP) = DecodeStamp(b, f6recs(i))
        outData(i, COL_TYPE) = IIf(setA = setB, "Self", "Cross")
        outData(i, COL_SETA) = setA
        outData(i, COL_SETB) = setB
    Next i

    ' write the sheet in one batch
    Dim su As Boolean: su = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set wsT = EnsureSheet(SHEET_TESTS)
    wsT.Cells.Clear
    wsT.Range("A2").Resize(nTests, N_COLS).Value = outData
    LayOutTestsSheet wsT
    Application.ScreenUpdating = su

    Dim msg As String
    msg = nTests & " test(s) imported into '" & SHEET_TESTS & "'." & vbCrLf & vbCrLf & _
          "Edit Tolerance / Clearance / Priority / Stamp, then run ExportClashTests to write a new .vimctst." & vbCrLf & vbCrLf & _
          "NOTE: re-importing adds NEW tests (fresh GUIDs) - delete the originals in Revizto first, " & _
          "or you'll get duplicates."
    If dirCount > 0 Then msg = msg & vbCrLf & vbCrLf & dirCount & " test(s) used DIRECTIONAL clearance - " & _
          "only the horizontal value was imported (Lite uses a single clearance)."
    MsgBox msg, vbInformation
End Sub

' --- decode helpers (read one test out of the parsed records) ---
Private Function ReadTestName(ByRef b() As Byte, ByVal f4 As Variant) As String
    Dim rf As Collection: Set rf = ParseFields(b, f4(2), f4(3))
    Dim f2 As Variant: f2 = Fld(rf, 2)
    If Not IsEmpty(f2) Then ReadTestName = BytesToStr(b, f2(2), f2(3))
End Function

Private Sub DecodeSides(ByRef b() As Byte, ByVal f6 As Variant, ByRef setA As String, ByRef setB As String)
    Dim rf As Collection: Set rf = ParseFields(b, f6(2), f6(3))
    setA = SideSetName(b, Fld(rf, 2))     ' f6.f2 = side A
    setB = SideSetName(b, Fld(rf, 3))     ' f6.f3 = side B
End Sub

' a side = guid + path strings; the set name is the LAST string (works for 2- or 4-string sides)
Private Function SideSetName(ByRef b() As Byte, ByVal side As Variant) As String
    If IsEmpty(side) Then Exit Function
    Dim sf As Collection: Set sf = ParseFields(b, side(2), side(3))
    Dim it As Variant, lastStr As String
    For Each it In sf
        If it(0) = 2 And it(1) = 2 Then lastStr = BytesToStr(b, it(2), it(3))
    Next it
    SideSetName = lastStr
End Function

' f6.f4: mode (f4.f4) + distance vector (f4.f3). mode 1=clearance, 2=tolerance,
' 3=directional clearance (import the horizontal axis only), 0=none.
Private Sub DecodeProximity(ByRef b() As Byte, ByVal f6 As Variant, _
                            ByRef tolS As String, ByRef clrS As String, ByRef dirCount As Long)
    tolS = "": clrS = ""
    Dim rf As Collection: Set rf = ParseFields(b, f6(2), f6(3))
    Dim f4 As Variant: f4 = Fld(rf, 4)
    If IsEmpty(f4) Then Exit Sub
    Dim f4f As Collection: Set f4f = ParseFields(b, f4(2), f4(3))
    Dim modeV As Variant: modeV = Fld(f4f, 4)
    Dim mode As Long
    If Not IsEmpty(modeV) Then mode = CLng(modeV(4))
    If mode = 0 Then Exit Sub
    Dim f3 As Variant: f3 = Fld(f4f, 3)
    If IsEmpty(f3) Then Exit Sub
    Dim f3f As Collection: Set f3f = ParseFields(b, f3(2), f3(3))
    Dim a1 As Variant: a1 = Fld(f3f, 1)
    If IsEmpty(a1) Then Exit Sub
    Dim mmStr As String: mmStr = CStr(RoundMm(ReadSingle(b, a1(2)) * MM_PER_FOOT))
    If mode = 2 Then
        tolS = mmStr
    Else
        clrS = mmStr                       ' mode 1 (clearance) or 3 (directional)
        If mode = 3 Then dirCount = dirCount + 1
    End If
End Sub

Private Function DecodePriority(ByRef b() As Byte, ByVal f6 As Variant) As String
    Dim rf As Collection: Set rf = ParseFields(b, f6(2), f6(3))
    Dim f10 As Variant: f10 = Fld(rf, 10)
    If IsEmpty(f10) Then Exit Function
    Dim f10f As Collection: Set f10f = ParseFields(b, f10(2), f10(3))
    Dim pv As Variant: pv = Fld(f10f, 6)
    If IsEmpty(pv) Then Exit Function
    Select Case CLng(pv(4))
        Case 1: DecodePriority = "Trivial"
        Case 2: DecodePriority = "Minor"
        Case 3: DecodePriority = "Major"
        Case 4: DecodePriority = "Critical"
        Case 5: DecodePriority = "Blocker"
        Case Else: DecodePriority = ""     ' 0 / none
    End Select
End Function

' stamp code = f6.f10.f4 (string); "" if the test has no stamp
Private Function DecodeStamp(ByRef b() As Byte, ByVal f6 As Variant) As String
    Dim rf As Collection: Set rf = ParseFields(b, f6(2), f6(3))
    Dim f10 As Variant: f10 = Fld(rf, 10)
    If IsEmpty(f10) Then Exit Function
    Dim f10f As Collection: Set f10f = ParseFields(b, f10(2), f10(3))
    Dim cv As Variant: cv = Fld(f10f, 4)
    If IsEmpty(cv) Then Exit Function
    If cv(1) = 2 Then DecodeStamp = BytesToStr(b, cv(2), cv(3))
End Function

' round to 2 dp, half-up (clash distances are clean numbers; float32 introduces tiny noise)
Private Function RoundMm(ByVal mm As Double) As Double
    RoundMm = Int(mm * 100# + 0.5) / 100#
End Function

Private Function ReadSingle(ByRef b() As Byte, ByVal pos As Long) As Single
    Dim v As Single
    CopyMemory v, b(pos), 4
    ReadSingle = v
End Function

Private Function ReadFileBytes(ByVal path As String) As Byte()
    Dim f As Integer, n As Long, b() As Byte
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

Private Function BytesToStr(ByRef b() As Byte, ByVal s As Long, ByVal L As Long) As String
    Dim i As Long, r As String
    For i = 0 To L - 1: r = r & Chr$(b(s + i)): Next i
    BytesToStr = r
End Function

Private Function PickVimctst() As String
    Dim fd As Object: Set fd = Application.FileDialog(3)   ' msoFileDialogFilePicker
    fd.Title = "Pick a Revizto clash-test export (.vimctst)"
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "Revizto clash tests", "*.vimctst"
    fd.Filters.Add "All files", "*.*"
    If fd.Show = -1 Then PickVimctst = fd.SelectedItems(1) Else PickVimctst = ""
End Function


' ===================== per-test settings =====================
' Apply one row's Tol/Clr onto a cloned field6 tail (fixed-length overwrite of f4.f3 + f4.f4).
' Grouping is baked into the template before the loop; priority + stamp are done by RebuildF10.
Private Sub ApplyProximity(ByRef tail() As Byte, ByVal sTol As String, ByVal sClr As String, _
                           ByVal clrA As Long, ByVal clrB As Long, ByVal clrC As Long, ByVal modeOff As Long)
    Dim tolMm As Double, clrMm As Double
    Dim hasTol As Boolean, hasClr As Boolean
    hasTol = ParseMm(sTol, tolMm)
    hasClr = ParseMm(sClr, clrMm)

    ' clearance / tolerance are mutually exclusive - callers reject rows with both set first
    Dim mode As Byte, dist As Double
    If hasClr Then
        mode = 1: dist = clrMm
    ElseIf hasTol Then
        mode = 2: dist = tolMm
    Else
        mode = 0: dist = 0
    End If
    Dim fv As Single: fv = CSng(dist / MM_PER_FOOT)   ' uniform: same distance on all 3 axes
    PokeSingle tail, clrA, fv
    PokeSingle tail, clrB, fv
    PokeSingle tail, clrC, fv
    tail(modeOff) = mode
End Sub

' Rebuild field6.f10 (priority + stamp) and splice it into the tail. Unlike the fixed-length
' pokes, a stamp INSERTS f4 (code) + f11 (GUID), so f10 - and the tail - grow; the new tail
' length comes back in outLen. f10 sub-fields: f1 flag (stamp=1|priority=4), f2=0, f3=1,
' [f4=code], f6=priority, [f11=fabricated GUID, cosmetic - stamps bind by the code string].
Private Function RebuildF10(ByRef tail() As Byte, ByVal tailLen As Long, ByVal priorityVal As Long, _
                            ByVal stampCode As String, ByRef outLen As Long) As Byte()
    Dim tc As Collection: Set tc = ParseFields(tail, 0, tailLen)
    Dim f10 As Variant: f10 = Fld(tc, 10)
    Dim payStart As Long: payStart = f10(2)
    Dim payLen As Long: payLen = f10(3)
    Dim recStart As Long: recStart = payStart - 1 - VLen(payLen)   ' back over the len varint + 0x52 tag
    Dim recEnd As Long: recEnd = payStart + payLen

    Dim hasStamp As Boolean: hasStamp = (Len(stampCode) > 0)
    Dim f1v As Long: f1v = 0
    If hasStamp Then f1v = f1v Or 1
    If priorityVal > 0 Then f1v = f1v Or 4

    Dim pay() As Byte, pl As Long: ReDim pay(0 To 255): pl = 0
    PutB pay, pl, &H8: PutB pay, pl, CByte(f1v)              ' f1 = flag bits
    PutB pay, pl, &H10: PutB pay, pl, 0                      ' f2 = 0
    PutB pay, pl, &H18: PutB pay, pl, 1                      ' f3 = 1
    If hasStamp Then
        Dim cb() As Byte: cb = AsciiBytes(stampCode)
        PutB pay, pl, &H22: PutVarint pay, pl, (UBound(cb) + 1): PutArr pay, pl, cb   ' f4 = code
    End If
    PutB pay, pl, &H30: PutVarint pay, pl, priorityVal      ' f6 = priority
    If hasStamp Then
        PutB pay, pl, &H5A: PutB pay, pl, &H10: PutArr pay, pl, NewGuidBytes()        ' f11 = GUID
    End If

    ' out = tail[0..recStart) + (0x52 + len(pl) + payload) + tail[recEnd..tailLen)
    Dim total As Long: total = recStart + (1 + VLen(pl) + pl) + (tailLen - recEnd)
    Dim out() As Byte: ReDim out(0 To total - 1)
    Dim ol As Long, i As Long: ol = 0
    For i = 0 To recStart - 1: out(ol) = tail(i): ol = ol + 1: Next i
    PutB out, ol, &H52: PutVarint out, ol, pl
    For i = 0 To pl - 1: out(ol) = pay(i): ol = ol + 1: Next i
    For i = recEnd To tailLen - 1: out(ol) = tail(i): ol = ol + 1: Next i

    outLen = total
    RebuildF10 = out
End Function

' "25" / "25 mm" / 25 -> True + mm; blank/non-numeric -> False
Private Function ParseMm(ByVal s As String, ByRef mm As Double) As Boolean
    s = Trim$(s)
    If Len(s) = 0 Then ParseMm = False: Exit Function
    If IsNumeric(s) Then
        mm = CDbl(s): ParseMm = True
    Else
        ParseMm = False
    End If
End Function

Private Function PriorityValue(ByVal s As String, ByRef known As Boolean) As Long
    known = True
    Select Case LCase$(Trim$(s))
        Case "", "none", "0": PriorityValue = 0
        Case "trivial":       PriorityValue = 1
        Case "minor":         PriorityValue = 2
        Case "major":         PriorityValue = 3
        Case "critical":      PriorityValue = 4
        Case "blocker":       PriorityValue = 5
        Case Else:            known = False: PriorityValue = 0
    End Select
End Function

' Find the writable setting slots inside the field6 tail by parsing it (survives re-harvest):
'   clrA/B/C = the 3 proximity floats f4.f3.f1/f2/f3   modeOff = f4.f4 mode varint
'   grpOff   = grouping float f9.f1.f3.f1   (priority + stamp live in f10, rebuilt by RebuildF10)
Private Sub LocateSettings(ByRef tail6() As Byte, ByVal tail6Len As Long, _
                           ByRef clrA As Long, ByRef clrB As Long, ByRef clrC As Long, ByRef modeOff As Long, _
                           ByRef grpOff As Long)
    Dim tc As Collection, f4f As Collection, f3f As Collection
    Dim f4rec As Variant, f3rec As Variant
    Set tc = ParseFields(tail6, 0, tail6Len)

    ' proximity (f4.f3 floats + f4.f4 mode)
    f4rec = Fld(tc, 4)
    Set f4f = ParseFields(tail6, f4rec(2), f4rec(3))
    f3rec = Fld(f4f, 3)
    Set f3f = ParseFields(tail6, f3rec(2), f3rec(3))
    clrA = Fld(f3f, 1)(2)
    clrB = Fld(f3f, 2)(2)
    clrC = Fld(f3f, 3)(2)
    modeOff = Fld(f4f, 4)(2)

    ' grouping (f9.f1.f3.f1)
    Dim f9rec As Variant, f9f As Collection, f91rec As Variant, f91f As Collection
    Dim g3rec As Variant, g3f As Collection
    f9rec = Fld(tc, 9)
    Set f9f = ParseFields(tail6, f9rec(2), f9rec(3))
    f91rec = Fld(f9f, 1)
    Set f91f = ParseFields(tail6, f91rec(2), f91rec(3))
    g3rec = Fld(f91f, 3)
    Set g3f = ParseFields(tail6, g3rec(2), g3rec(3))
    grpOff = Fld(g3f, 1)(2)
End Sub


' ===================== Tests-sheet snapshot =====================
' Read existing Tests rows in one shot; key by test name (col A) -> Array(tol, clr, prio).
Private Sub SnapshotSettings(ByVal ws As Worksheet, ByVal dict As Object)
    Dim last As Long
    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If last < 2 Then Exit Sub
    Dim a As Variant: a = ws.Range("A2:E" & last).Value
    Dim r As Long, key As String
    For r = 1 To UBound(a, 1)
        key = Trim$(CStr(a(r, 1)))
        If Len(key) > 0 And Not dict.Exists(key) Then
            dict(key) = Array(CStr(a(r, COL_TOL)), CStr(a(r, COL_CLR)), CStr(a(r, COL_PRIO)), CStr(a(r, COL_STAMP)))
        End If
    Next r
End Sub


' ===================== byte / protobuf / guid helpers =====================
' Write the Tests-sheet headers and apply all formatting (BuildTestList, ImportClashTests and
' NewTestsSheet all share this). Set A/Set B live hidden far right; everything visible is A:F.
Private Sub LayOutTestsSheet(ByVal wsT As Worksheet)
    wsT.Range("A1:F1").Value = Array("Test Name", "Tolerance (mm)", "Clearance (mm)", "Priority", "Stamp", "Type")
    wsT.Cells(1, COL_SETA).Value = "Set A"
    wsT.Cells(1, COL_SETB).Value = "Set B"
    wsT.Columns("A").AutoFit                                  ' Test Name fits its content
    wsT.Columns("B:E").ColumnWidth = 16                       ' the four setting columns: equal + wider
    wsT.Columns("B:F").HorizontalAlignment = xlCenter         ' settings + Type centred
    wsT.Columns("F").AutoFit                                  ' Type fits "Cross"/"Self"
    With wsT.Range("A1:F1")                                   ' header row: dark-grey box, white bold text
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(64, 64, 64)
    End With
    wsT.Range(wsT.Cells(1, COL_SETA), wsT.Cells(1, COL_SETB)).EntireColumn.Hidden = True
    ' filter dropdowns on the header row (e.g. open the Type dropdown and untick "Self")
    On Error Resume Next
    wsT.AutoFilterMode = False                                ' clear any stale filter first
    On Error GoTo 0
    wsT.Range("A1").AutoFilter                                ' applies to A1's region (A:F, header + data)
End Sub

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

Private Function CloneBytes(ByRef src() As Byte, ByVal n As Long) As Byte()
    Dim b() As Byte, i As Long
    ReDim b(0 To n - 1)
    For i = 0 To n - 1: b(i) = src(i): Next i
    CloneBytes = b
End Function

' overwrite 4 bytes at pos with a little-endian IEEE-754 single (matches the file order)
Private Sub PokeSingle(ByRef b() As Byte, ByVal pos As Long, ByVal v As Single)
    Dim tmp(0 To 3) As Byte
    CopyMemory tmp(0), v, 4
    b(pos) = tmp(0): b(pos + 1) = tmp(1): b(pos + 2) = tmp(2): b(pos + 3) = tmp(3)
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

' Returns Double so large varints never overflow a Long.
Private Function ReadVarint(ByRef b() As Byte, ByRef pos As Long) As Double
    Dim result As Double, shift As Long, x As Long
    result = 0: shift = 0
    Do
        x = b(pos): pos = pos + 1
        result = result + (x And &H7F) * (2 ^ shift)
        If (x And &H80) = 0 Then Exit Do
        shift = shift + 7
    Loop
    ReadVarint = result
End Function

' Parse one protobuf message [s, s+length). Collection of 5-element Variant arrays:
'   (0)=field#  (1)=wire  (2)=payloadStart/valuePos  (3)=len(wire2)  (4)=varintValue(wire0)
Private Function ParseFields(ByRef b() As Byte, ByVal s As Long, ByVal length As Long) As Collection
    Dim c As Collection, p As Long, e As Long
    Dim tag As Long, fld As Long, wire As Long, L As Long
    Dim d() As Variant
    Set c = New Collection
    p = s: e = s + length
    Do While p < e
        tag = ReadVarint(b, p)
        fld = tag \ 8
        wire = tag And 7
        ReDim d(0 To 4)
        d(0) = fld: d(1) = wire: d(2) = 0: d(3) = 0: d(4) = 0
        If wire = 2 Then
            L = ReadVarint(b, p): d(2) = p: d(3) = L: p = p + L
        ElseIf wire = 0 Then
            d(2) = p: d(4) = ReadVarint(b, p)
        ElseIf wire = 5 Then
            d(2) = p: d(3) = 4: p = p + 4
        ElseIf wire = 1 Then
            d(2) = p: d(3) = 8: p = p + 8
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
        If it(0) = n Then Fld = it: Exit Function
    Next it
    Fld = Empty
End Function
