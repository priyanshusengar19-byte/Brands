Attribute VB_Name = "PincodeMapper"
Option Explicit

' ╔══════════════════════════════════════════════════════════════════╗
' ║           PINCODE → METRO / TIER  MAPPER  (Pure Excel VBA)      ║
' ║                                                                  ║
' ║  Runs entirely inside Excel — no Python, no add-ins, no         ║
' ║  internet. All processing is in-memory using arrays.            ║
' ║                                                                  ║
' ║  BEFORE RUNNING:                                                 ║
' ║  1. Copy your 10k reference sheet into THIS workbook            ║
' ║     and name it exactly:  Ref_10k                               ║
' ║  2. Confirm the column numbers in CONFIG below match your data   ║
' ║  3. Alt+F8 → Run  MapPincodesToMetroTier                        ║
' ╚══════════════════════════════════════════════════════════════════╝

' ┌──────────────────────────────────────────────────────────────────┐
' │  CONFIG — Update these to match YOUR sheet and column layout     │
' └──────────────────────────────────────────────────────────────────┘

' Sheet names
Private Const SHEET_19K    As String = "Population by Pincode"
Private Const SHEET_10K    As String = "Ref_10k"

' ── 19k sheet column numbers (A=1, B=2, C=3 …)
Private Const C19_PINCODE  As Integer = 1   ' A  Pincode
Private Const C19_LAT      As Integer = 3   ' C  Latitude    ← adjust if different
Private Const C19_LON      As Integer = 4   ' D  Longitude   ← adjust if different

' ── 10k sheet column numbers
Private Const C10_PINCODE  As Integer = 2   ' B  Pincode
Private Const C10_LAT      As Integer = 11  ' K  Latitude    ← adjust if different
Private Const C10_LON      As Integer = 12  ' L  Longitude   ← adjust if different
Private Const C10_METRO    As Integer = 7   ' G  Metro/City Area
Private Const C10_TIER     As Integer = 9   ' I  T3 (Tier)
Private Const C10_STATE    As Integer = 8   ' H  State
Private Const C10_DISTRICT As Integer = 6   ' F  District

' Header row numbers (usually 1, but 10k file has title in row 1 so data header is row 2)
Private Const HDR_19K As Integer = 1
Private Const HDR_10K As Integer = 2   ' ← set to 1 if your 10k header is on row 1

' ─────────────────────────────────────────────────────────────────────
Public Sub MapPincodesToMetroTier()
' ─────────────────────────────────────────────────────────────────────

    ' ── Validate both sheets exist ─────────────────────────────────
    If Not SheetExists(SHEET_19K) Then
        MsgBox "Cannot find sheet: """ & SHEET_19K & """" & vbNewLine & _
               "Please check CONFIG at the top of the macro.", vbCritical
        Exit Sub
    End If
    If Not SheetExists(SHEET_10K) Then
        MsgBox "Cannot find sheet: """ & SHEET_10K & """" & vbNewLine & _
               "Please copy your 10k reference data into this workbook" & vbNewLine & _
               "and rename that sheet to:  Ref_10k", vbCritical
        Exit Sub
    End If

    Dim ws19 As Worksheet, ws10 As Worksheet
    Set ws19 = ThisWorkbook.Sheets(SHEET_19K)
    Set ws10 = ThisWorkbook.Sheets(SHEET_10K)

    ' ── Speed optimisations ─────────────────────────────────────────
    Application.ScreenUpdating   = False
    Application.Calculation      = xlCalculationManual
    Application.EnableEvents     = False
    Dim t0 As Double: t0 = Timer

    ' ── Load EVERYTHING into arrays (fastest possible approach) ─────
    Dim lastR19 As Long, lastR10 As Long
    lastR19 = ws19.Cells(ws19.Rows.Count, C19_PINCODE).End(xlUp).Row
    lastR10 = ws10.Cells(ws10.Rows.Count, C10_PINCODE).End(xlUp).Row

    Dim maxCol19 As Long, maxCol10 As Long
    maxCol19 = ws19.Cells(HDR_19K, ws19.Columns.Count).End(xlToLeft).Column
    maxCol10 = ws10.Cells(HDR_10K, ws10.Columns.Count).End(xlToLeft).Column
    If maxCol19 < 4 Then maxCol19 = 4
    If maxCol10 < 12 Then maxCol10 = 12

    Dim arr19 As Variant, arr10 As Variant
    arr19 = ws19.Range(ws19.Cells(HDR_19K + 1, 1), ws19.Cells(lastR19, maxCol19)).Value
    arr10 = ws10.Range(ws10.Cells(HDR_10K + 1, 1), ws10.Cells(lastR10, maxCol10)).Value

    Dim n19 As Long, n10 As Long
    n19 = UBound(arr19, 1)
    n10 = UBound(arr10, 1)

    ' ── Prepare output array (7 new columns) ────────────────────────
    Dim out() As Variant
    ReDim out(1 To n19, 1 To 7)

    Dim i As Long, j As Long, bj As Long
    Dim p19 As Long, p10 As Long
    Dim la1 As Double, lo1 As Double, la2 As Double, lo2 As Double
    Dim d As Double, dMin As Double
    Dim hit As Boolean
    Dim nExact As Long, nNN As Long, nFail As Long

    ' ── Main loop ───────────────────────────────────────────────────
    For i = 1 To n19

        ' Progress bar every 200 rows
        If i Mod 200 = 0 Then
            Application.StatusBar = "Row " & i & "/" & n19 & _
                "   Exact=" & nExact & "  NN=" & nNN & _
                "  Elapsed=" & Format(Timer - t0, "0") & "s"
            DoEvents
        End If

        ' Skip empty/error pincodes
        If IsError(arr19(i, C19_PINCODE)) Then GoTo Skip19
        If IsEmpty(arr19(i, C19_PINCODE)) Then GoTo Skip19
        On Error Resume Next
        p19 = CLng(arr19(i, C19_PINCODE))
        If Err.Number <> 0 Then Err.Clear: GoTo Skip19
        On Error GoTo 0

        hit = False

        ' ── PASS 1: Exact pincode match ─────────────────────────────
        For j = 1 To n10
            If Not IsError(arr10(j, C10_PINCODE)) Then
                If Not IsEmpty(arr10(j, C10_PINCODE)) Then
                    On Error Resume Next
                    p10 = CLng(arr10(j, C10_PINCODE))
                    On Error GoTo 0
                    If p10 = p19 Then
                        out(i, 1) = arr10(j, C10_METRO)
                        out(i, 2) = arr10(j, C10_TIER)
                        out(i, 3) = arr10(j, C10_STATE)
                        out(i, 4) = arr10(j, C10_DISTRICT)
                        out(i, 5) = "Exact"
                        out(i, 6) = p10
                        out(i, 7) = 0
                        nExact = nExact + 1
                        hit = True
                        Exit For
                    End If
                End If
            End If
        Next j
        If hit Then GoTo Skip19

        ' ── PASS 2: Nearest neighbour via Haversine ─────────────────
        If IsError(arr19(i, C19_LAT)) Or IsEmpty(arr19(i, C19_LAT)) Then
            out(i, 5) = "No Lat/Lon"
            nFail = nFail + 1
            GoTo Skip19
        End If
        On Error Resume Next
        la1 = CDbl(arr19(i, C19_LAT))
        lo1 = CDbl(arr19(i, C19_LON))
        If Err.Number <> 0 Then
            Err.Clear: out(i, 5) = "Bad Coords": nFail = nFail + 1: GoTo Skip19
        End If
        On Error GoTo 0

        dMin = 1E+18
        bj = 0

        For j = 1 To n10
            If Not IsError(arr10(j, C10_LAT)) And Not IsEmpty(arr10(j, C10_LAT)) Then
                If Not IsEmpty(arr10(j, C10_METRO)) And Not IsError(arr10(j, C10_METRO)) Then
                    On Error Resume Next
                    la2 = CDbl(arr10(j, C10_LAT))
                    lo2 = CDbl(arr10(j, C10_LON))
                    If Err.Number = 0 Then
                        d = Haversine(la1, lo1, la2, lo2)
                        If d < dMin Then dMin = d: bj = j
                    Else
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
            End If
        Next j

        If bj > 0 Then
            out(i, 1) = arr10(bj, C10_METRO)
            out(i, 2) = arr10(bj, C10_TIER)
            out(i, 3) = arr10(bj, C10_STATE)
            out(i, 4) = arr10(bj, C10_DISTRICT)
            out(i, 5) = "Nearest Neighbour"
            out(i, 6) = arr10(bj, C10_PINCODE)
            out(i, 7) = Format(dMin, "0.00")
            nNN = nNN + 1
        Else
            out(i, 5) = "Not Found"
            nFail = nFail + 1
        End If

Skip19:
    Next i

    ' ── Write output headers ────────────────────────────────────────
    Dim cStart As Long
    cStart = ws19.Cells(HDR_19K, ws19.Columns.Count).End(xlToLeft).Column + 1

    Dim hdrRange As Range
    Set hdrRange = ws19.Range(ws19.Cells(HDR_19K, cStart), ws19.Cells(HDR_19K, cStart + 6))
    hdrRange.Value = Array("Metro/City Area", "Tier (T3)", "State", "District", _
                           "Match Type", "Matched Pincode", "Distance (km)")
    With hdrRange
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 111, 235)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With

    ' ── Bulk-write all output in one call (fastest) ─────────────────
    ws19.Range(ws19.Cells(HDR_19K + 1, cStart), _
               ws19.Cells(HDR_19K + n19, cStart + 6)).Value = out

    ' ── Colour-code Match Type column ───────────────────────────────
    ColorCodeMatchType ws19, HDR_19K + 1, HDR_19K + n19, cStart + 4

    ' ── Auto-fit new columns ────────────────────────────────────────
    ws19.Range(ws19.Cells(1, cStart), ws19.Cells(1, cStart + 6)).EntireColumn.AutoFit

    ' ── Restore Excel ───────────────────────────────────────────────
    Application.ScreenUpdating   = True
    Application.Calculation      = xlCalculationAutomatic
    Application.EnableEvents     = True
    Application.StatusBar        = False

    Dim elapsed As Double: elapsed = Timer - t0
    MsgBox "  Done in " & Format(elapsed, "0.0") & " seconds!" & vbNewLine & vbNewLine & _
           "Exact Matches     : " & Format(nExact, "#,##0") & vbNewLine & _
           "Nearest Neighbour : " & Format(nNN, "#,##0") & vbNewLine & _
           "Could Not Match   : " & Format(nFail, "#,##0") & vbNewLine & vbNewLine & _
           "New columns added to: " & SHEET_19K, vbInformation, "Pincode Mapper — Complete"
End Sub

' ─────────────────────────────────────────────────────────────────────
Private Sub ColorCodeMatchType(ws As Worksheet, rStart As Long, rEnd As Long, col As Long)
' ─────────────────────────────────────────────────────────────────────
    Dim r As Long
    For r = rStart To rEnd
        Select Case ws.Cells(r, col).Value
            Case "Exact"
                ws.Cells(r, col).Interior.Color = RGB(209, 250, 223) ' green
                ws.Cells(r, col).Font.Color      = RGB(14, 100, 55)
            Case "Nearest Neighbour"
                ws.Cells(r, col).Interior.Color = RGB(255, 243, 205) ' amber
                ws.Cells(r, col).Font.Color      = RGB(133, 77, 14)
            Case Else
                ws.Cells(r, col).Interior.Color = RGB(254, 226, 226) ' red
                ws.Cells(r, col).Font.Color      = RGB(153, 27, 27)
        End Select
    Next r
End Sub

' ─────────────────────────────────────────────────────────────────────
Private Function Haversine(lat1 As Double, lon1 As Double, _
                            lat2 As Double, lon2 As Double) As Double
' Returns distance in km between two lat/lon points
' ─────────────────────────────────────────────────────────────────────
    Const PI As Double = 3.14159265358979323846
    Dim dLat As Double, dLon As Double, a As Double
    dLat = (lat2 - lat1) * PI / 180
    dLon = (lon2 - lon1) * PI / 180
    a = Sin(dLat / 2) ^ 2 + _
        Cos(lat1 * PI / 180) * Cos(lat2 * PI / 180) * Sin(dLon / 2) ^ 2
    Haversine = 6371 * 2 * Atn(Sqr(a) / Sqr(1 - a))
End Function

' ─────────────────────────────────────────────────────────────────────
Private Function SheetExists(name As String) As Boolean
' ─────────────────────────────────────────────────────────────────────
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(name)
    SheetExists = (Not ws Is Nothing)
    On Error GoTo 0
End Function
