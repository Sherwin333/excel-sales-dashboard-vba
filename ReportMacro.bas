
Attribute VB_Name = "ReportMacro"
Option Explicit

Sub GenerateMonthlyReport()
    Dim wsRaw As Worksheet, wsSum As Worksheet
    Dim lastRow As Long, i As Long
    Dim month As String, region As String, category As String
    Dim sales As Double, units As Double, target As Double

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wsRaw = ThisWorkbook.Sheets("Raw Data")
    Set wsSum = ThisWorkbook.Sheets("Monthly Summary")

    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row

    ' --- Clear old values in summary (keep formatting) ---
    Dim clearRanges As Variant
    clearRanges = Array("B6:G11", "B15:G17", "B21:F21")
    Dim cr As Variant
    For Each cr In clearRanges
        wsSum.Range(cr).ClearContents
    Next cr

    ' --- Month rows: A6:A11 already have month labels ---
    Dim mths(1 To 6) As String
    mths(1) = "Jan" : mths(2) = "Feb" : mths(3) = "Mar"
    mths(4) = "Apr" : mths(5) = "May" : mths(6) = "Jun"

    Dim mthSales(1 To 6) As Double
    Dim mthTarget(1 To 6) As Double
    Dim mthUnits(1 To 6) As Double
    Dim mthRegionSales(1 To 6, 1 To 4) As Double  ' regions: N,S,E,W

    Dim regions(1 To 4) As String
    regions(1) = "North" : regions(2) = "South"
    regions(3) = "East"  : regions(4) = "West"

    Dim catSales(1 To 3) As Double
    Dim catUnits(1 To 3) As Double
    Dim catMonthSales(1 To 3, 1 To 6) As Double
    Dim cats(1 To 3) As String
    cats(1) = "Electronics" : cats(2) = "Clothing" : cats(3) = "Furniture"

    ' --- Loop raw data ---
    For i = 2 To lastRow
        month    = Trim(wsRaw.Cells(i, 1).Value)
        region   = Trim(wsRaw.Cells(i, 2).Value)
        category = Trim(wsRaw.Cells(i, 3).Value)
        sales    = wsRaw.Cells(i, 4).Value
        units    = wsRaw.Cells(i, 5).Value
        target   = wsRaw.Cells(i, 6).Value

        Dim mi As Integer, ri As Integer, ci As Integer
        For mi = 1 To 6
            If month = mths(mi) Then
                mthSales(mi)  = mthSales(mi) + sales
                mthTarget(mi) = mthTarget(mi) + target
                mthUnits(mi)  = mthUnits(mi) + units
                For ri = 1 To 4
                    If region = regions(ri) Then
                        mthRegionSales(mi, ri) = mthRegionSales(mi, ri) + sales
                    End If
                Next ri
                Exit For
            End If
        Next mi

        For ci = 1 To 3
            If category = cats(ci) Then
                catSales(ci) = catSales(ci) + sales
                catUnits(ci) = catUnits(ci) + units
                For mi = 1 To 6
                    If month = mths(mi) Then
                        catMonthSales(ci, mi) = catMonthSales(ci, mi) + sales
                        Exit For
                    End If
                Next mi
                Exit For
            End If
        Next ci
    Next i

    ' --- Write Section A: Monthly ---
    Dim totalRevenue As Double, bestSales As Double, bestMonth As String
    bestSales = 0

    For mi = 1 To 6
        Dim r As Integer : r = mi + 5
        Dim variance As Double, pct As Double
        variance = mthSales(mi) - mthTarget(mi)
        If mthTarget(mi) > 0 Then pct = mthSales(mi) / mthTarget(mi) Else pct = 0

        wsSum.Cells(r, 2).Value = mthSales(mi)
        wsSum.Cells(r, 2).NumberFormat = Chr(8377) & "#,##0"

        wsSum.Cells(r, 3).Value = mthTarget(mi)
        wsSum.Cells(r, 3).NumberFormat = Chr(8377) & "#,##0"

        wsSum.Cells(r, 4).Value = variance
        wsSum.Cells(r, 4).NumberFormat = Chr(8377) & "#,##0"
        If variance >= 0 Then
            wsSum.Cells(r, 4).Font.Color = RGB(30, 126, 52)
        Else
            wsSum.Cells(r, 4).Font.Color = RGB(192, 57, 43)
        End If

        wsSum.Cells(r, 5).Value = pct
        wsSum.Cells(r, 5).NumberFormat = "0.0%"
        If pct >= 1 Then
            wsSum.Cells(r, 5).Font.Color = RGB(30, 126, 52)
        Else
            wsSum.Cells(r, 5).Font.Color = RGB(192, 57, 43)
        End If

        If pct >= 1 Then
            wsSum.Cells(r, 6).Value = "HIT"
            wsSum.Cells(r, 6).Font.Color = RGB(30, 126, 52)
            wsSum.Cells(r, 6).Font.Bold = True
        Else
            wsSum.Cells(r, 6).Value = "MISS"
            wsSum.Cells(r, 6).Font.Color = RGB(192, 57, 43)
            wsSum.Cells(r, 6).Font.Bold = True
        End If

        ' Top region this month
        Dim topReg As String : topReg = ""
        Dim topVal As Double : topVal = 0
        For ri = 1 To 4
            If mthRegionSales(mi, ri) > topVal Then
                topVal = mthRegionSales(mi, ri)
                topReg = regions(ri)
            End If
        Next ri
        wsSum.Cells(r, 7).Value = topReg

        totalRevenue = totalRevenue + mthSales(mi)
        If mthSales(mi) > bestSales Then
            bestSales = mthSales(mi)
            bestMonth = mths(mi)
        End If
    Next mi

    ' --- Write Section B: Category ---
    Dim totalSales As Double : totalSales = totalRevenue
    For ci = 1 To 3
        Dim cr2 As Integer : cr2 = ci + 14
        wsSum.Cells(cr2, 2).Value = catSales(ci)
        wsSum.Cells(cr2, 2).NumberFormat = Chr(8377) & "#,##0"

        wsSum.Cells(cr2, 3).Value = catUnits(ci)
        wsSum.Cells(cr2, 3).NumberFormat = "#,##0"

        If catUnits(ci) > 0 Then
            wsSum.Cells(cr2, 4).Value = catSales(ci) / catUnits(ci)
        End If
        wsSum.Cells(cr2, 4).NumberFormat = Chr(8377) & "#,##0"

        If totalSales > 0 Then
            wsSum.Cells(cr2, 5).Value = catSales(ci) / totalSales
        End If
        wsSum.Cells(cr2, 5).NumberFormat = "0.0%"

        ' Best month for category
        Dim bestCatMonth As String : bestCatMonth = ""
        Dim bestCatVal As Double   : bestCatVal = 0
        For mi = 1 To 6
            If catMonthSales(ci, mi) > bestCatVal Then
                bestCatVal = catMonthSales(ci, mi)
                bestCatMonth = mths(mi)
            End If
        Next mi
        wsSum.Cells(cr2, 6).Value = bestCatMonth
    Next ci

    ' --- Write Section C: KPIs ---
    wsSum.Cells(21, 1).Value = totalRevenue
    wsSum.Cells(21, 1).NumberFormat = Chr(8377) & "#,##0"

    Dim totalUnits As Double
    For mi = 1 To 6 : totalUnits = totalUnits + mthUnits(mi) : Next mi
    wsSum.Cells(21, 2).Value = totalUnits
    wsSum.Cells(21, 2).NumberFormat = "#,##0"

    wsSum.Cells(21, 3).Value = totalRevenue / 6
    wsSum.Cells(21, 3).NumberFormat = Chr(8377) & "#,##0"

    wsSum.Cells(21, 4).Value = bestMonth

    Dim totalTarget As Double
    For mi = 1 To 6 : totalTarget = totalTarget + mthTarget(mi) : Next mi
    If totalTarget > 0 Then wsSum.Cells(21, 5).Value = totalRevenue / totalTarget
    wsSum.Cells(21, 5).NumberFormat = "0.0%"

    Dim topCat As String : topCat = ""
    Dim topCatVal As Double : topCatVal = 0
    For ci = 1 To 3
        If catSales(ci) > topCatVal Then topCatVal = catSales(ci) : topCat = cats(ci)
    Next ci
    wsSum.Cells(21, 6).Value = topCat

    ' --- Timestamp ---
    wsSum.Cells(2, 7).Value = "Last run: " & Format(Now, "dd-mmm-yyyy hh:mm")
    wsSum.Cells(2, 7).Font.Italic = True
    wsSum.Cells(2, 7).Font.Size = 9
    wsSum.Cells(2, 7).Font.Color = RGB(119, 119, 119)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Report generated successfully!" & Chr(10) & _
           "Total Revenue: " & Chr(8377) & Format(totalRevenue, "#,##0") & Chr(10) & _
           "Best Month: " & bestMonth & Chr(10) & _
           "Top Category: " & topCat, vbInformation, "Deloitte Sales Report"
End Sub
