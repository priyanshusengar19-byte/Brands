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
'  OUTPUT: writes to Pincodes sheet, Col D = Nearest City, Col E = Distance(km)
'
'  Speed notes:
'   - All data is loaded into arrays once (no cell-by-cell reads in the loop)
'   - cos(lat) and radians are pre-computed once per pincode/city, not per pair
'   - Latitude-only lower bound is used to skip the full Haversine calc for
'     any city that can't possibly beat the current best match (big speedup,
'     since |lat1-lat2| in km is always <= true great-circle distance)
'   - Expect roughly 1-4 minutes total depending on machine (was previously
'     ~10-30 sec for the smaller 10k pin run; this run is ~4x the pair count)
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
    Dim nPin As Long, nCity As Long
    Dim i As Long, j As Long

    Dim pinData As Variant, cityData As Variant

    Dim pinLatRad() As Double, pinLonRad() As Double
    Dim cityLatRad() As Double, cityLonRad() As Double, cityCosLat() As Double
    Dim cityName() As String

    Dim resultCity() As String, resultDist() As Double

    Dim degToRad As Double, R As Double
    Dim cosPinLat As Double
    Dim dLat As Double, dLon As Double, a As Double, dist As Double
    Dim minDist As Double, minIdx As Long
    Dim latBoundKm As Double
    Dim startTime As Double

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

    ReDim pinLatRad(1 To nPin): ReDim pinLonRad(1 To nPin)
    ReDim cityLatRad(1 To nCity): ReDim cityLonRad(1 To nCity): ReDim cityCosLat(1 To nCity)
    ReDim cityName(1 To nCity)
    ReDim resultCity(1 To nPin): ReDim resultDist(1 To nPin)

    ' ---------------- Pre-convert everything to radians once ----------------
    For i = 1 To nPin
        pinLatRad(i) = CDbl(pinData(i, 2)) * degToRad
        pinLonRad(i) = CDbl(pinData(i, 3)) * degToRad
    Next i

    For j = 1 To nCity
        cityName(j) = CStr(cityData(j, 1))
        cityLatRad(j) = CDbl(cityData(j, 2)) * degToRad
        cityLonRad(j) = CDbl(cityData(j, 3)) * degToRad
        cityCosLat(j) = Cos(cityLatRad(j))
    Next j

    Erase pinData
    Erase cityData

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    startTime = Timer

    ' ================== MAIN NEAREST-NEIGHBOUR LOOP ==================
    For i = 1 To nPin
        cosPinLat = Cos(pinLatRad(i))
        minDist = 1E+18
        minIdx = 0

        For j = 1 To nCity
            ' ---- cheap lower-bound check first: skip full Haversine if
            '      even the latitude-only distance can't beat current best
            latBoundKm = R * Abs(cityLatRad(j) - pinLatRad(i))
            If latBoundKm < minDist Then
                dLat = cityLatRad(j) - pinLatRad(i)
                dLon = cityLonRad(j) - pinLonRad(i)
                a = (Sin(dLat / 2)) ^ 2 + cosPinLat * cityCosLat(j) * (Sin(dLon / 2)) ^ 2
                If a > 1 Then a = 1  ' guard against float rounding
                dist = 2 * R * Application.WorksheetFunction.Asin(Sqr(a))
                If dist < minDist Then
                    minDist = dist
                    minIdx = j
                End If
            End If
        Next j

        resultCity(i) = cityName(minIdx)
        resultDist(i) = minDist

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
        outArr(i, 2) = Round(resultDist(i), 2)
    Next i
    wsPin.Range("D" & (PIN_HEADER_ROW + 1) & ":E" & lastPinRow).Value = outArr

    wsPin.Cells(PIN_HEADER_ROW, 4).Value = "Nearest City"
    wsPin.Cells(PIN_HEADER_ROW, 5).Value = "Distance (km)"
    wsPin.Cells(PIN_HEADER_ROW, 4).Font.Bold = True
    wsPin.Cells(PIN_HEADER_ROW, 5).Font.Bold = True

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False

    If Err.Number = 0 Then
        MsgBox "Done." & vbCrLf & _
               nPin & " pincodes mapped against " & nCity & " cities" & vbCrLf & _
               "Time taken: " & Format(Timer - startTime, "0.0") & " seconds.", vbInformation
    End If
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Check that sheet names 'Pincodes' / 'Cities' and column layout (A=Name/Pincode, B=Lat, C=Lon) match your workbook.", vbCritical
    Resume CleanExit

End Sub


' ==========================================================================
'  FLAG DISTANT PINCODES  (data-quality check)
' --------------------------------------------------------------------------
'  Run this AFTER MapPincodesToNearestCity has populated Col D/E on the
'  Pincodes sheet. It flags any pincode whose "nearest city" distance is
'  suspiciously large (default 150 km) - usually a sign of a bad lat/long
'  in either the pincode table or the city table, rather than a genuinely
'  remote location. Produces:
'   - Col F on Pincodes sheet: "FLAG" for anything over the threshold
'   - Red fill on the flagged rows' distance cell for quick visual scan
'   - A new sheet "Flagged_Review" listing all flagged rows, sorted by
'     distance descending, so you can work top-down through the worst ones
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

    ' Read Pincode, City-name(col A), Nearest City(D), Distance(E)
    pinData = wsPin.Range("A" & (PIN_HEADER_ROW + 1) & ":E" & lastPinRow).Value

    wsPin.Cells(PIN_HEADER_ROW, 6).Value = "Flag"
    wsPin.Cells(PIN_HEADER_ROW, 6).Font.Bold = True

    ReDim flaggedRows(1 To nPin, 1 To 3)   ' Pincode | Nearest City | Distance
    flagCount = 0

    Application.ScreenUpdating = False

    For i = 1 To nPin
        Dim dist As Double
        dist = 0
        If IsNumeric(pinData(i, 5)) Then dist = CDbl(pinData(i, 5))

        If dist > DISTANCE_THRESHOLD_KM Then
            wsPin.Cells(PIN_HEADER_ROW + i, 6).Value = "FLAG"
            wsPin.Cells(PIN_HEADER_ROW + i, 5).Interior.Color = RGB(255, 199, 206) ' light red
            wsPin.Cells(PIN_HEADER_ROW + i, 5).Font.Color = RGB(156, 0, 6)

            flagCount = flagCount + 1
            flaggedRows(flagCount, 1) = pinData(i, 1) ' Pincode
            flaggedRows(flagCount, 2) = pinData(i, 4) ' Nearest City
            flaggedRows(flagCount, 3) = dist
        Else
            wsPin.Cells(PIN_HEADER_ROW + i, 6).Value = ""
            wsPin.Cells(PIN_HEADER_ROW + i, 5).Interior.ColorIndex = xlNone
        End If
    Next i

    ' ---------------- Build / refresh Flagged_Review sheet ----------------
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
        ' Simple bubble sort by distance descending (flagCount is small, so this is fine)
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
