Attribute VB_Name = "NearestCityMapper"
Option Explicit

' ==========================================================================
'  NEAREST CITY MAPPER  (Haversine, nearest-neighbour)
' --------------------------------------------------------------------------
'  Maps every pincode to its nearest city using lat/long + Haversine distance.
'  Built for: ~19,000 pincodes  x  4,232 cities  (~80M distance comparisons)
'
'  SETUP REQUIRED before running:
'   1) Sheet "Pincodes"  -> Col A: Pincode | Col B: Lat | Col C: Lon
'      (row 1 = headers, data starts row 2)
'   2) Sheet "Cities"    -> Col A: City    | Col B: Lat | Col C: Lon
'      (row 1 = headers, data starts row 2)
'   3) Rename sheets / columns in the "SETUP" section below if yours differ.
'
'  RECOMMENDED ORDER TO RUN:
'   1) ValidateSourceData        <- run this FIRST, always
'   2) MapPincodesToNearestCity
'   3) FlagDistantPincodes
'
'  OUTPUT: writes to Pincodes sheet, Col D = Nearest City, Col E = Distance(km)
'
'  Speed notes:
'   - All data is loaded into arrays once (no cell-by-cell reads in the loop)
'   - cos(lat) and radians are pre-computed once per pincode/city, not per pair
'   - Latitude-only lower bound is used to skip the full Haversine calc for
'     any city that can't possibly beat the current best match (big speedup,
'     since |lat1-lat2| in km is always <= true great-circle distance)
'   - Expect roughly 1-4 minutes total depending on machine
'
'  ROBUSTNESS: blank cells, text, "N/A", stray spaces etc. in the lat/lon
'  columns will NOT crash the macro anymore. Bad pincode rows are marked
'  "INVALID LAT/LON" in the output and skipped from distance calculation.
'  Bad city rows are excluded from the candidate list entirely. Both are
'  reported in the final summary and logged to a "DataQuality_Issues" sheet.
' ==========================================================================


' ==========================================================================
'  STEP 1 - VALIDATE SOURCE DATA  (run this before anything else)
' --------------------------------------------------------------------------
'  Scans Pincodes!B:C and Cities!B:C for blanks / non-numeric / out-of-range
'  lat-long values and lists every offending row on a new sheet called
'  "DataQuality_Issues", with the sheet name, row number, and the raw values
'  found - so you can jump straight to the cell and fix it before running
'  the (slow) nearest-city mapper on possibly-bad data.
' ==========================================================================

Sub ValidateSourceData()

    Const PIN_SHEET As String = "Pincodes"
    Const CITY_SHEET As String = "Cities"
    Const PIN_HEADER_ROW As Long = 1
    Const CITY_HEADER_ROW As Long = 1

    Dim wsPin As Worksheet, wsCity As Worksheet, wsIssues As Worksheet
    Dim lastPinRow As Long, lastCityRow As Long
    Dim pinData As Variant, cityData As Variant
    Dim i As Long, issueRow As Long
    Dim badLat As Boolean, badLon As Boolean
    Dim reasonTxt As String

    On Error GoTo ErrHandler

    Set wsPin = ThisWorkbook.Sheets(PIN_SHEET)
    Set wsCity = ThisWorkbook.Sheets(CITY_SHEET)

    lastPinRow = wsPin.Cells(wsPin.Rows.Count, "A").End(xlUp).Row
    lastCityRow = wsCity.Cells(wsCity.Rows.Count, "A").End(xlUp).Row

    pinData = wsPin.Range("A" & (PIN_HEADER_ROW + 1) & ":C" & lastPinRow).Value
    cityData = wsCity.Range("A" & (CITY_HEADER_ROW + 1) & ":C" & lastCityRow).Value

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("DataQuality_Issues").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler

    Set wsIssues = ThisWorkbook.Sheets.Add(After:=wsPin)
    wsIssues.Name = "DataQuality_Issues"
    wsIssues.Range("A1").Value = "Source Sheet"
    wsIssues.Range("B1").Value = "Row #"
    wsIssues.Range("C1").Value = "ID (Pincode/City)"
    wsIssues.Range("D1").Value = "Lat Value Found"
    wsIssues.Range("E1").Value = "Lon Value Found"
    wsIssues.Range("F1").Value = "Problem"
    wsIssues.Range("A1:F1").Font.Bold = True

    issueRow = 2

    ' ---- check Pincodes sheet ----
    For i = 1 To UBound(pinData, 1)
        badLat = Not IsLatValid(pinData(i, 2))
        badLon = Not IsLonValid(pinData(i, 3))
        If badLat Or badLon Then
            reasonTxt = ""
            If badLat Then reasonTxt = "Bad/blank latitude"
            If badLon Then reasonTxt = reasonTxt & IIf(reasonTxt <> "", " + ", "") & "Bad/blank longitude"
            wsIssues.Cells(issueRow, 1).Value = PIN_SHEET
            wsIssues.Cells(issueRow, 2).Value = i + PIN_HEADER_ROW
            wsIssues.Cells(issueRow, 3).Value = CStr(pinData(i, 1))
            wsIssues.Cells(issueRow, 4).Value = CStr(pinData(i, 2))
            wsIssues.Cells(issueRow, 5).Value = CStr(pinData(i, 3))
            wsIssues.Cells(issueRow, 6).Value = reasonTxt
            issueRow = issueRow + 1
        End If
    Next i

    ' ---- check Cities sheet ----
    For i = 1 To UBound(cityData, 1)
        badLat = Not IsLatValid(cityData(i, 2))
        badLon = Not IsLonValid(cityData(i, 3))
        If badLat Or badLon Then
            reasonTxt = ""
            If badLat Then reasonTxt = "Bad/blank latitude"
            If badLon Then reasonTxt = reasonTxt & IIf(reasonTxt <> "", " + ", "") & "Bad/blank longitude"
            wsIssues.Cells(issueRow, 1).Value = CITY_SHEET
            wsIssues.Cells(issueRow, 2).Value = i + CITY_HEADER_ROW
            wsIssues.Cells(issueRow, 3).Value = CStr(cityData(i, 1))
            wsIssues.Cells(issueRow, 4).Value = CStr(cityData(i, 2))
            wsIssues.Cells(issueRow, 5).Value = CStr(cityData(i, 3))
            wsIssues.Cells(issueRow, 6).Value = reasonTxt
            issueRow = issueRow + 1
        End If
    Next i

    wsIssues.Columns("A:F").AutoFit

    If issueRow = 2 Then
        MsgBox "Validation passed - no bad lat/long values found in either sheet." & vbCrLf & _
               "You're clear to run MapPincodesToNearestCity.", vbInformation
        Application.DisplayAlerts = False
        wsIssues.Delete
        Application.DisplayAlerts = True
    Else
        MsgBox (issueRow - 2) & " row(s) with bad/blank/invalid lat-long found." & vbCrLf & _
               "See 'DataQuality_Issues' sheet for the exact rows." & vbCrLf & vbCrLf & _
               "You can fix these first, OR just proceed - MapPincodesToNearestCity will " & _
               "automatically skip these rows instead of crashing.", vbExclamation
    End If
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Check sheet names 'Pincodes' / 'Cities' match your workbook.", vbCritical
End Sub

' ---- helper: is this a usable latitude value? (-90 to 90, numeric, non-blank) ----
Private Function IsLatValid(v As Variant) As Boolean
    If Trim(v & "") = "" Then IsLatValid = False: Exit Function
    If Not IsNumeric(v) Then IsLatValid = False: Exit Function
    If CDbl(v) < -90 Or CDbl(v) > 90 Then IsLatValid = False: Exit Function
    IsLatValid = True
End Function

' ---- helper: is this a usable longitude value? (-180 to 180, numeric, non-blank) ----
Private Function IsLonValid(v As Variant) As Boolean
    If Trim(v & "") = "" Then IsLonValid = False: Exit Function
    If Not IsNumeric(v) Then IsLonValid = False: Exit Function
    If CDbl(v) < -180 Or CDbl(v) > 180 Then IsLonValid = False: Exit Function
    IsLonValid = True
End Function


' ==========================================================================
'  STEP 2 - MAP EACH PINCODE TO ITS NEAREST CITY
' ==========================================================================

Sub MapPincodesToNearestCity()

    ' ---------------- SETUP: edit these if your layout differs ----------------
    Const PIN_SHEET As String = "Pincodes"
    Const CITY_SHEET As String = "Cities"
    Const PIN_HEADER_ROW As Long = 1
    Const CITY_HEADER_ROW As Long = 1
    ' ----------------------------------------------------------------------------

    Dim wsPin As Worksheet, wsCity As Worksheet
    Dim lastPinRow As Long, lastCityRow As Long
    Dim nPin As Long, nCity As Long, nCityValid As Long
    Dim i As Long, j As Long, k As Long

    Dim pinData As Variant, cityData As Variant

    Dim pinLatRad() As Double, pinLonRad() As Double, pinValid() As Boolean
    Dim cityLatRad() As Double, cityLonRad() As Double, cityCosLat() As Double
    Dim cityName() As String

    Dim resultCity() As String, resultDist() As Variant

    Dim degToRad As Double, R As Double
    Dim cosPinLat As Double
    Dim dLat As Double, dLon As Double, a As Double, dist As Double
    Dim minDist As Double, minIdx As Long
    Dim latBoundKm As Double
    Dim startTime As Double
    Dim badPinCount As Long, badCityCount As Long

    On Error GoTo ErrHandler

    Set wsPin = ThisWorkbook.Sheets(PIN_SHEET)
    Set wsCity = ThisWorkbook.Sheets(CITY_SHEET)

    lastPinRow = wsPin.Cells(wsPin.Rows.Count, "A").End(xlUp).Row
    lastCityRow = wsCity.Cells(wsCity.Rows.Count, "A").End(xlUp).Row

    nPin = lastPinRow - PIN_HEADER_ROW
    nCity = lastCityRow - CITY_HEADER_ROW

    If nPin < 1 Or nCity < 1 Then
        MsgBox "No data found. Check sheet names / header row settings.", vbExclamation
        Exit Sub
    End If

    R = 6371.0088 ' mean Earth radius in km
    degToRad = 3.14159265358979 / 180

    ' ---------------- Bulk read both tables in one shot ----------------
    pinData = wsPin.Range("A" & (PIN_HEADER_ROW + 1) & ":C" & lastPinRow).Value
    cityData = wsCity.Range("A" & (CITY_HEADER_ROW + 1) & ":C" & lastCityRow).Value

    ReDim pinLatRad(1 To nPin): ReDim pinLonRad(1 To nPin): ReDim pinValid(1 To nPin)
    ReDim resultCity(1 To nPin): ReDim resultDist(1 To nPin)

    ' ---------------- Validate + convert pincodes (skip bad rows, don't crash) ----------------
    badPinCount = 0
    For i = 1 To nPin
        If IsLatValid(pinData(i, 2)) And IsLonValid(pinData(i, 3)) Then
            pinValid(i) = True
            pinLatRad(i) = CDbl(pinData(i, 2)) * degToRad
            pinLonRad(i) = CDbl(pinData(i, 3)) * degToRad
        Else
            pinValid(i) = False
            badPinCount = badPinCount + 1
            resultCity(i) = "INVALID LAT/LON"
            resultDist(i) = ""
        End If
    Next i

    ' ---------------- Validate + convert cities (drop bad rows from candidate list) ----------------
    ReDim cityLatRad(1 To nCity): ReDim cityLonRad(1 To nCity): ReDim cityCosLat(1 To nCity)
    ReDim cityName(1 To nCity)
    nCityValid = 0
    badCityCount = 0
    For j = 1 To nCity
        If IsLatValid(cityData(j, 2)) And IsLonValid(cityData(j, 3)) Then
            nCityValid = nCityValid + 1
            cityName(nCityValid) = CStr(cityData(j, 1))
            cityLatRad(nCityValid) = CDbl(cityData(j, 2)) * degToRad
            cityLonRad(nCityValid) = CDbl(cityData(j, 3)) * degToRad
            cityCosLat(nCityValid) = Cos(cityLatRad(nCityValid))
        Else
            badCityCount = badCityCount + 1
        End If
    Next j

    If nCityValid < 1 Then
        MsgBox "No valid city coordinates found - cannot proceed. Run ValidateSourceData to see why.", vbCritical
        Exit Sub
    End If

    Erase pinData
    Erase cityData

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    startTime = Timer

    ' ================== MAIN NEAREST-NEIGHBOUR LOOP ==================
    For i = 1 To nPin
        If pinValid(i) Then
            cosPinLat = Cos(pinLatRad(i))
            minDist = 1E+18
            minIdx = 0

            For k = 1 To nCityValid
                latBoundKm = R * Abs(cityLatRad(k) - pinLatRad(i))
                If latBoundKm < minDist Then
                    dLat = cityLatRad(k) - pinLatRad(i)
                    dLon = cityLonRad(k) - pinLonRad(i)
                    a = (Sin(dLat / 2)) ^ 2 + cosPinLat * cityCosLat(k) * (Sin(dLon / 2)) ^ 2
                    If a > 1 Then a = 1
                    dist = 2 * R * Application.WorksheetFunction.Asin(Sqr(a))
                    If dist < minDist Then
                        minDist = dist
                        minIdx = k
                    End If
                End If
            Next k

            resultCity(i) = cityName(minIdx)
            resultDist(i) = Round(minDist, 2)
        End If

        If i Mod 250 = 0 Or i = nPin Then
            Application.StatusBar = "Mapping nearest city... " & Format(i / nPin, "0%") & _
                " (" & i & " of " & nPin & " pincodes, " & Format(Timer - startTime, "0") & "s elapsed)"
            DoEvents
        End If
    Next i
    ' ===================================================================

    ' ---------------- Bulk write results back in one shot ----------------
    Dim outArr() As Variant
    ReDim outArr(1 To nPin, 1 To 2)
    For i = 1 To nPin
        outArr(i, 1) = resultCity(i)
        outArr(i, 2) = resultDist(i)
    Next i
    wsPin.Range("D" & (PIN_HEADER_ROW + 1) & ":E" & lastPinRow).Value = outArr

    wsPin.Cells(PIN_HEADER_ROW, 4).Value = "Nearest City"
    wsPin.Cells(PIN_HEADER_ROW, 5).Value = "Distance (km)"
    wsPin.Cells(PIN_HEADER_ROW, 4).Font.Bold = True
    wsPin.Cells(PIN_HEADER_ROW, 5).Font.Bold = True

    For i = 1 To nPin
        If Not pinValid(i) Then
            wsPin.Cells(PIN_HEADER_ROW + i, 4).Interior.Color = RGB(255, 235, 156)
        End If
    Next i

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False

    If Err.Number = 0 Then
        Dim msg As String
        msg = "Done." & vbCrLf & _
              (nPin - badPinCount) & " of " & nPin & " pincodes mapped successfully" & vbCrLf & _
              nCityValid & " of " & nCity & " cities used as valid candidates" & vbCrLf & _
              "Time taken: " & Format(Timer - startTime, "0.0") & " seconds."
        If badPinCount > 0 Or badCityCount > 0 Then
            msg = msg & vbCrLf & vbCrLf & "NOTE: " & badPinCount & " pincode row(s) and " & badCityCount & _
                  " city row(s) had invalid/blank lat-long and were skipped." & vbCrLf & _
                  "Skipped pincodes are marked 'INVALID LAT/LON' and highlighted in Col D." & vbCrLf & _
                  "Run ValidateSourceData for the full list with row numbers."
        End If
        MsgBox msg, vbInformation
    End If
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Check that sheet names 'Pincodes' / 'Cities' and column layout (A=Name/Pincode, B=Lat, C=Lon) match your workbook.", vbCritical
    Resume CleanExit

End Sub


' ==========================================================================
'  STEP 3 - FLAG DISTANT PINCODES  (data-quality check on the OUTPUT)
' ==========================================================================

Sub FlagDistantPincodes()

    Const PIN_SHEET As String = "Pincodes"
    Const PIN_HEADER_ROW As Long = 1
    Const DISTANCE_THRESHOLD_KM As Double = 150   ' <-- adjust as needed

    Dim wsPin As Worksheet, wsFlag As Worksheet
    Dim lastPinRow As Long, nPin As Long
    Dim i As Long, flagCount As Long
    Dim pinData As Variant
    Dim flaggedRows() As Variant
    Dim r As Long

    On Error GoTo ErrHandler

    Set wsPin = ThisWorkbook.Sheets(PIN_SHEET)
    lastPinRow = wsPin.Cells(wsPin.Rows.Count, "A").End(xlUp).Row
    nPin = lastPinRow - PIN_HEADER_ROW

    If wsPin.Cells(PIN_HEADER_ROW, 5).Value <> "Distance (km)" Then
        MsgBox "Column E doesn't look like it has 'Distance (km)' yet." & vbCrLf & _
               "Run MapPincodesToNearestCity first.", vbExclamation
        Exit Sub
    End If

    pinData = wsPin.Range("A" & (PIN_HEADER_ROW + 1) & ":E" & lastPinRow).Value

    wsPin.Cells(PIN_HEADER_ROW, 6).Value = "Flag"
    wsPin.Cells(PIN_HEADER_ROW, 6).Font.Bold = True

    ReDim flaggedRows(1 To nPin, 1 To 3)
    flagCount = 0

    Application.ScreenUpdating = False

    For i = 1 To nPin
        Dim dist As Double
        dist = 0
        If IsNumeric(pinData(i, 5)) Then dist = CDbl(pinData(i, 5))

        If dist > DISTANCE_THRESHOLD_KM Then
            wsPin.Cells(PIN_HEADER_ROW + i, 6).Value = "FLAG"
            wsPin.Cells(PIN_HEADER_ROW + i, 5).Interior.Color = RGB(255, 199, 206)
            wsPin.Cells(PIN_HEADER_ROW + i, 5).Font.Color = RGB(156, 0, 6)

            flagCount = flagCount + 1
            flaggedRows(flagCount, 1) = pinData(i, 1)
            flaggedRows(flagCount, 2) = pinData(i, 4)
            flaggedRows(flagCount, 3) = dist
        ElseIf pinData(i, 4) <> "INVALID LAT/LON" Then
            wsPin.Cells(PIN_HEADER_ROW + i, 6).Value = ""
            wsPin.Cells(PIN_HEADER_ROW + i, 5).Interior.ColorIndex = xlNone
        End If
    Next i

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Flagged_Review").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler

    Set wsFlag = ThisWorkbook.Sheets.Add(After:=wsPin)
    wsFlag.Name = "Flagged_Review"

    wsFlag.Range("A1").Value = "Pincode"
    wsFlag.Range("B1").Value = "Nearest City"
    wsFlag.Range("C1").Value = "Distance (km)"
    wsFlag.Range("A1:C1").Font.Bold = True

    If flagCount > 0 Then
        Dim tmpP As Variant, tmpC As Variant, tmpD As Double
        Dim a As Long, b As Long
        For a = 1 To flagCount - 1
            For b = 1 To flagCount - a
                If flaggedRows(b, 3) < flaggedRows(b + 1, 3) Then
                    tmpP = flaggedRows(b, 1): tmpC = flaggedRows(b, 2): tmpD = flaggedRows(b, 3)
                    flaggedRows(b, 1) = flaggedRows(b + 1, 1)
                    flaggedRows(b, 2) = flaggedRows(b + 1, 2)
                    flaggedRows(b, 3) = flaggedRows(b + 1, 3)
                    flaggedRows(b + 1, 1) = tmpP
                    flaggedRows(b + 1, 2) = tmpC
                    flaggedRows(b + 1, 3) = tmpD
                End If
            Next b
        Next a

        Dim outArr() As Variant
        ReDim outArr(1 To flagCount, 1 To 3)
        For r = 1 To flagCount
            outArr(r, 1) = flaggedRows(r, 1)
            outArr(r, 2) = flaggedRows(r, 2)
            outArr(r, 3) = Round(flaggedRows(r, 3), 2)
        Next r
        wsFlag.Range("A2:C" & (flagCount + 1)).Value = outArr
    End If

    wsFlag.Columns("A:C").AutoFit
    wsFlag.Range("A1").Select

    Application.ScreenUpdating = True

    MsgBox flagCount & " of " & nPin & " pincodes (" & Format(flagCount / nPin, "0.0%") & _
           ") are more than " & DISTANCE_THRESHOLD_KM & " km from their nearest city." & vbCrLf & _
           "See 'Flagged_Review' sheet, sorted worst-first.", vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical

End Sub
