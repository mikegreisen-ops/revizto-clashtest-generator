' ============================================================================
'  modImportSets  -  Workflow B: import search-set NAMES from a Revizto
'  .vimsst export into this workbook, so generated tests reference the user's
'  real Revizto sets (Revizto rebinds clash-test sets by NAME on import).
'
'  Paste this ENTIRE text into a NEW standard module (Alt+F11: Insert > Module).
'  Run: ImportSetsFromVimsst  (it shows a file picker).
'
'  Writes a sheet "ImportedSets":
'     A = Set name (exactly as stored in the .vimsst)
'     B = Folder it lives under
'     D = SearchSets col I names that were NOT found in the import (misalignments)
'  Use col A to fill SearchSets col I; fix anything listed in col D.
'
'  Built 2026-06-09 by Big C. Parse logic mirrors the proven vsst_names.ps1 /
'  the old modClashGen .vimsst reader. Self-contained (own helpers).
' ============================================================================
Option Explicit

Private Const SHEET_IMPORT As String = "ImportedSets"
Private Const SHEET_SETS   As String = "SearchSets"
Private Const COL_SETNAME  As String = "I"     ' SearchSets set-name column


' ===================== ENTRY POINT =====================
Public Sub ImportSetsFromVimsst()
    Dim path As String
    path = PickVimsst()
    If Len(path) = 0 Then Exit Sub

    Dim names As Collection, folders As Collection
    Set names = New Collection
    Set folders = New Collection
    ParseVimsst path, names, folders

    If names.Count = 0 Then
        MsgBox "No search sets found in that file." & vbCrLf & path, vbExclamation
        Exit Sub
    End If

    ' ---- write the ImportedSets sheet ----
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_IMPORT)
    ws.Cells.Clear
    ws.Range("A1").Value = "Set name (from .vimsst)"
    ws.Range("B1").Value = "Folder"
    ws.Range("D1").Value = "SearchSets col I names NOT in this import"

    Dim i As Long
    For i = 1 To names.Count
        ws.Cells(i + 1, 1).Value = names(i)
        ws.Cells(i + 1, 2).Value = folders(i)
    Next i

    ' ---- build a lookup of imported names ----
    Dim have As Object
    Set have = CreateObject("Scripting.Dictionary")
    have.CompareMode = 1                     ' text/case-insensitive
    For i = 1 To names.Count
        have(CStr(names(i))) = True
    Next i

    ' ---- flag SearchSets col I names that don't match any imported set ----
    Dim missing As Long
    missing = 0
    On Error Resume Next
    Dim wsS As Worksheet
    Set wsS = ThisWorkbook.Worksheets(SHEET_SETS)
    On Error GoTo 0
    If Not wsS Is Nothing Then
        Dim lastS As Long, rr As Long, nm As String
        lastS = wsS.Cells(wsS.Rows.Count, COL_SETNAME).End(xlUp).Row
        For rr = 2 To lastS
            nm = Trim$(CStr(wsS.Cells(rr, COL_SETNAME).Value))
            If Len(nm) > 0 Then
                If Not have.Exists(nm) Then
                    missing = missing + 1
                    ws.Cells(missing + 1, 4).Value = nm
                End If
            End If
        Next rr
    End If

    ws.Columns("A:D").AutoFit

    MsgBox names.Count & " sets imported to '" & SHEET_IMPORT & "'." & vbCrLf & vbCrLf & _
           missing & " SearchSets col I name(s) had no match in the import" & _
           IIf(missing > 0, " (listed in column D - fix these)", " - all aligned."), vbInformation
End Sub


' ===================== .vimsst parse =====================
' Fills parallel collections: names(i) = set name, folders(i) = its folder name.
' Walks top-level field-4 records; f6 = kind (2=folder, 1=set), f7 = name.
Private Sub ParseVimsst(ByVal path As String, ByRef names As Collection, ByRef folders As Collection)
    Dim b() As Byte
    b = ReadFileBytes(path)
    If (Not Not b) = 0 Then Exit Sub        ' empty

    Dim p As Long, e As Long
    Dim tag As Double, fld As Long, wire As Long, L As Long
    Dim recEnd As Long
    Dim curFolder As String
    curFolder = ""
    p = 0
    e = UBound(b) + 1

    Do While p < e
        tag = ReadVarint(b, p)
        fld = CLng(tag) \ 8
        wire = CLng(tag) And 7
        If wire = 2 Then
            L = CLng(ReadVarint(b, p))
            If fld = 4 Then
                ' parse this record's inner fields
                recEnd = p + L
                Dim kind As Long, nm As String
                kind = 0
                nm = ""
                ParseRecord b, p, recEnd, kind, nm
                If kind = 2 Then
                    curFolder = nm
                ElseIf kind = 1 Then
                    names.Add nm
                    folders.Add curFolder
                End If
                p = recEnd
            Else
                p = p + L
            End If
        ElseIf wire = 0 Then
            ReadVarint b, p
        ElseIf wire = 5 Then
            p = p + 4
        ElseIf wire = 1 Then
            p = p + 8
        Else
            Exit Do
        End If
    Loop
End Sub

' Reads one record [p, recEnd): sets kind (from f6) and nm (from f7).
Private Sub ParseRecord(ByRef b() As Byte, ByVal p As Long, ByVal recEnd As Long, _
                        ByRef kind As Long, ByRef nm As String)
    Dim tag As Double, fld As Long, wire As Long, L As Long
    Do While p < recEnd
        tag = ReadVarint(b, p)
        fld = CLng(tag) \ 8
        wire = CLng(tag) And 7
        If wire = 2 Then
            L = CLng(ReadVarint(b, p))
            If fld = 7 Then nm = BytesToStr(b, p, L)
            p = p + L
        ElseIf wire = 0 Then
            Dim v As Double
            v = ReadVarint(b, p)
            If fld = 6 Then kind = CLng(v)
        ElseIf wire = 5 Then
            p = p + 4
        ElseIf wire = 1 Then
            p = p + 8
        Else
            Exit Do
        End If
    Loop
End Sub


' ===================== helpers =====================
Private Function PickVimsst() As String
    Dim fd As Object
    Set fd = Application.FileDialog(3)       ' msoFileDialogFilePicker
    fd.Title = "Pick a Revizto search-set export (.vimsst)"
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "Revizto search sets", "*.vimsst"
    fd.Filters.Add "All files", "*.*"
    If fd.Show = -1 Then PickVimsst = fd.SelectedItems(1) Else PickVimsst = ""
End Function

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

' Returns Double so huge .vimsst timestamp varints (~6e17) never overflow a Long.
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
