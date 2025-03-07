Dim ticker As String
Dim currentYear As Integer
    
Sub FillDCF()
    
    ' Set the DCF sheet as active
    Sheets("DCF").Activate
    
    ' Variables
    Dim i As Integer
    Dim cell As Range
    Dim key As Variant
    
    ' Scale by millions
    Dim scaling As Double
    scaling = 1 / 1000000
    
    ' Read ticker and year from Excel cells
    ticker = Range("D3").Value
    currentYear = Range("I8").Value
    
    ' Fill company name
    Range("B2") = "=TR(""" & ticker & """, ""TR.CompanyName"")"
    
    ' Calculate year range
    Dim firstPart As String
    Dim secondPart As String
    firstPart = "'" & Right(currentYear + 1, 2)
    secondPart = "'" & Right(currentYear + 5, 2)
    Range("O8") = "(" & firstPart & " - " & secondPart & ")"
    
    ' Cell locations as a dictionary
    Dim historicals As Object
    Set historicals = CreateObject("Scripting.Dictionary")
    ' EDIT AS NECESSARY TO MATCH FINANCIALS
    historicals.Add "I9", "TotRevenue"
    historicals.Add "I11", "COGSTot"
    historicals.Add "I14", "SGATot"
    historicals.Add "I17", "DeprDeplAmortTot"
    historicals.Add "I24", "CAPEXTot"
    
    ' Loop to fetch data for current year and previous 3 years
    For i = 0 To 3
        For Each key In historicals.keys
            ' Define the target cell (starting from C3)
            Set cell = Range(key).Offset(0, -i)
            ' Fetch data from Refinitiv and place the TR function directly into the cell
        cell.Formula = "=IFERROR(TR(""" & ticker & """, ""TR.F." & (historicals(key)) & """, ""Period=" & (currentYear - i) & """) * " & scaling & ", 0)"
        Next key
    Next i
    
    ' Loop to fill the tax rates
    For i = 0 To 3
        ' Define the target cell
        Set cell = Range("I57").Offset(0, -i)
        ' Fetch data from Refinitiv and place the TR function directly into the cell
        cell.Formula = "=IFERROR(TR(""" & ticker & """, ""TR.TaxRateActValue"", ""Period=" & (currentYear - i) & """) / 100, 0)"
    Next i
    
    ' Additional data
    Range("K36").Formula = "=IFERROR(TR(""" & ticker & """, ""TR.F.DebtTot"") * " & scaling & ", 0)"
    Range("K39").Formula = "=IFERROR(TR(""" & ticker & """, ""TR.F.CashCashEquivTot"") * " & scaling & ", 0)"
    Range("P43").Formula = "=IFERROR(TR(""" & ticker & """, ""TR.F.EBITDA"", ""Period=LTM"") * " & scaling & ", 0)"
    Range("K43").Formula = "=IFERROR(TR(""" & ticker & """, ""TR.SharesOutstanding"") * " & scaling & ", 0)"
    Range("K37").Formula = "=IFERROR(TR(""" & ticker & """, ""TR.F.PrefShHoldEq"") * " & scaling & ", 0)"
    Range("K38").Formula = "=IFERROR(TR(""" & ticker & """, ""TR.F.MinIntrTot"") * " & scaling & ", 0)"
    
    ' Call FillWACC
    FillWACC ticker, currentYear
    FillNWC ticker, currentYear, scaling
    
    ' Set the Assumptions sheet as active
    Sheets("Assumptions").Activate
    
    MsgBox "Macro Ran!"
    
End Sub

Sub FillWACC(ticker As String, currentYear As Integer)
    
    ' Set the WACC sheet as active
    Sheets("WACC").Activate
    
    Range("E9").Formula = "=TR(""" & ticker & """, ""TR.WACCDebtWeight"") / 100"
    Range("E14").Formula = "=TR(""" & ticker & """, ""TR.WACCCostofDebt"") / 100"
    Range("E15").Formula = "=TR(""" & ticker & """, ""TR.WACCTaxRate"") / 100"
    Range("E20").Formula = "=TR(""US10YT=RR"", ""TR.BidYield"") / 100"
    Range("E21").Value = 4.33 / 100
    Range("E22").Formula = "=TR(""" & ticker & """, ""TR.WACCBeta"")"

End Sub

Sub FillNWC(ticker As String, currentYear As Integer, scaling As Double)

    ' Set the NWC sheet as active
    Sheets("NWC").Activate

    ' Cell locations as a dictionary
    Dim historicals As Object
    Set historicals = CreateObject("Scripting.Dictionary")
    ' EDIT AS NECESSARY TO MATCH FINANCIALS
    historicals.Add "G13", "LoansRcvblNetST"
    historicals.Add "G14", "InvntTot"
    historicals.Add "G15", "OthCurrAssetsTot"
    historicals.Add "G19", "TradeAcctTradeNotesPbleSt"
    historicals.Add "G20", "AccrExpnSt"
    historicals.Add "G21", "OthCurrLiabTot"
    
    ' Loop to fetch data for current year and previous 3 years
    For i = 0 To 3
        For Each key In historicals.keys
            ' Define the target cell (starting from C3)
            Set cell = Range(key).Offset(0, -i)
            ' Fetch data from Refinitiv and place the TR function directly into the cell
            cell.Formula = "=IFERROR(TR(""" & ticker & """, ""TR.F." & (historicals(key)) & """, ""Period=" & (currentYear - i) & """) * " & scaling & ", 0)"
        Next key
    Next i
    
End Sub

Sub FillAssumptions()
    Dim ws As Worksheet
    Dim rowNumbers As Variant
    Dim avgChanges() As Double
    Dim rng As Range
    Dim i As Integer
    
    ' Set the active sheet to "DCF"
    Sheets("DCF").Activate
    Set ws = Sheets("DCF")

    ' List of row numbers to calculate average percentage change for
    rowNumbers = Array(9, 11, 14, 17, 24)

    ' Redimension the array to match the number of rows
    ReDim avgChanges(LBound(rowNumbers) To UBound(rowNumbers))

    ' Loop through each row in the list
    For i = LBound(rowNumbers) To UBound(rowNumbers)
        ' Set the range for F:G:H:I in the current row
        Set rng = ws.Range(ws.Cells(rowNumbers(i), 6), ws.Cells(rowNumbers(i), 9)) ' Columns F to I

        ' Store the calculated average percentage change in the array
        avgChanges(i) = AveragePercentageChange(rng)
    Next i
    
    ' The Sales row
    Set rng = ws.Range(ws.Cells(rowNumbers(0), 6), ws.Cells(rowNumbers(0), 9))
    ' Store the calculated average percentage change in the array
    avgChanges(0) = RealAveragePercentageChange(rng)
    
    ' Set the active sheet to "Assumptions"
    Sheets("Assumptions").Activate
    Set ws = Sheets("Assumptions")

    ' Rows to enter data in
    assumRows = Array(11, 18, 25, 32, 40)

    ' Write stored averages directly into columns F:J
    For i = LBound(assumRows) To UBound(assumRows)
        ' Set original values in F:J
        ws.Range(ws.Cells(assumRows(i), 6), ws.Cells(assumRows(i), 10)).Value = avgChanges(i)
    
        ' Set +2 row (Increase by 1)
        ws.Range(ws.Cells(assumRows(i) + 2, 6), ws.Cells(assumRows(i) + 2, 10)).Value = avgChanges(i) + 0.01
    
        ' Set +3 row (Decrease by 1)
        ws.Range(ws.Cells(assumRows(i) + 3, 6), ws.Cells(assumRows(i) + 3, 10)).Value = avgChanges(i) - 0.01
    Next i
    
    ' Set the active sheet to "NWC"
    Sheets("NWC").Activate
    Set ws = Sheets("NWC")

    ' List of row numbers to calculate average percentage change for
    rowNumbers = Array(13, 14, 15, 19, 20, 21)

    ' Redimension the array to match the number of rows
    ReDim avgChanges(LBound(rowNumbers) To UBound(rowNumbers))

    ' Loop through each row in the list
    For i = LBound(rowNumbers) To UBound(rowNumbers)
        ' Set the range for F:G:H:I in the current row
        Set rng = ws.Range(ws.Cells(rowNumbers(i), 4), ws.Cells(rowNumbers(i), 7)) ' Columns D to G

        ' Store the calculated average percentage change in the array
        avgChanges(i) = AveragePercentageChange(rng)
    Next i
    
    ' Set the active sheet to "Assumptions"
    Sheets("Assumptions").Activate
    Set ws = Sheets("Assumptions")

    ' Rows to enter data in
    assumRows = Array(48, 55, 62, 69, 76, 83)

    ' Write stored averages directly into columns F:J
    For i = LBound(assumRows) To UBound(assumRows)
        ' Set original values in F:J
        ws.Range(ws.Cells(assumRows(i), 6), ws.Cells(assumRows(i), 10)).Value = avgChanges(i)
    
        ' Set +2 row (Increase by 1)
        ws.Range(ws.Cells(assumRows(i) + 2, 6), ws.Cells(assumRows(i) + 2, 10)).Value = avgChanges(i) + 0.01
    
        ' Set +3 row (Decrease by 1)
        ws.Range(ws.Cells(assumRows(i) + 3, 6), ws.Cells(assumRows(i) + 3, 10)).Value = avgChanges(i) - 0.01
    Next i
    
    ' Cleanup
    Set ws = Nothing
    Set rng = Nothing
    
    MsgBox "Macro Ran!"

End Sub

' avg % of sales
Function AveragePercentageChange(rng As Range) As Double
    Dim i As Integer
    Dim totalPct As Double
    Dim count As Integer
    Dim salesRow As Range

    ' Reference to sales values in DCF!F9:I9
    Set salesRow = Sheets("DCF").Range("F9:I9")

    totalPct = 0
    count = 0

    ' Loop through each cell in the provided range
    For i = 1 To rng.Columns.count
        If salesRow.Cells(1, i).Value <> 0 Then
            ' Calculate percentage of sales
            totalPct = totalPct + (rng.Cells(1, i).Value / salesRow.Cells(1, i).Value)
            count = count + 1
        End If
    Next i

    ' Calculate average percentage
    If count > 0 Then
        AveragePercentageChange = totalPct / count
    Else
        AveragePercentageChange = 0
    End If
End Function

' Function to calculate average percentage change
Function RealAveragePercentageChange(rng As Range) As Double
    Dim i As Integer
    Dim totalChange As Double
    Dim count As Integer
    
    totalChange = 0
    count = 0

    ' Loop through each column (F to I) and calculate percentage change
    For i = 2 To rng.Columns.count
        If rng.Cells(1, i - 1).Value <> 0 Then
            totalChange = totalChange + ((rng.Cells(1, i).Value - rng.Cells(1, i - 1).Value) / rng.Cells(1, i - 1).Value)
            count = count + 1
        End If
    Next i

    ' Calculate the average percentage change
    If count > 0 Then
        RealAveragePercentageChange = totalChange / count
    Else
        RealAveragePercentageChange = 0
    End If
End Function