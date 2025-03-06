
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
    ticker = Range("D3").Value & ".O"
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
    historicals.Add "I57", "WACCTaxRate"
    
    ' Loop to fetch data for current year and previous 3 years
    For i = 0 To 3
        For Each key In historicals.keys
            ' Define the target cell (starting from C3)
            Set cell = Range(key).Offset(0, -i)
            ' Fetch data from Refinitiv and place the TR function directly into the cell
            cell.Formula = "=TR(""" & ticker & """, ""TR.F." & (historicals(key)) & """, ""Period=" & (currentYear - i) & """) * " & scaling
        Next key
    Next i
    ' Loop to fill the tax rates
    For i = 0 To 3
        ' Define the target cell
        Set cell = Range("I57").Offset(0, -i)
        ' Fetch data from Refinitiv and place the TR function directly into the cell
        cell.Formula = "=TR(""" & ticker & """, ""TR.WACCTaxRate"", ""Period=" & (currentYear - i) & """)"
    Next i
    
    ' Additional data
    Range("K36").Formula = "=TR(""" & ticker & """, ""TR.F.DebtTot"") * " & scaling
    Range("K39").Formula = "=TR(""" & ticker & """, ""TR.F.CashCashEquivTot"") * " & scaling
    Range("P43").Formula = "=TR(""" & ticker & """, ""TR.F.EBITDA"", ""Period=LTM"") * " & scaling
    Range("K43").Formula = "=TR(""" & ticker & """, ""TR.SharesOutstanding"") * " & scaling
    Range("K37").Formula = "=TR(""" & ticker & """, ""TR.F.PrefShHoldEq"") * " & scaling
    Range("K38").Formula = "=TR(""" & ticker & """, ""TR.F.MinIntrEq"") * " & scaling
    
    Debug.Print "Macro is running!"
    MsgBox "Button Clicked!"
    
End Sub

Sub FillWACC()
    
    ' Set the WACC sheet as active
    Sheets("WACC").Activate
    Range("").Formula = "=TR(""" & ticker & """, ""TR.F.MinIntrEq"") * " & scaling

End Sub

Sub FillAssumptions()
    Dim wsDCF As Worksheet, wsAssumptions As Worksheet
    Dim lastCol As Integer
    Dim firstProjectionCol As Integer
    Dim scaleFactor As Double
    Dim grossMarginAvg As Double
    Dim i As Integer
    
    ' Set worksheets
    Set wsDCF = ThisWorkbook.Sheets("DCF")
    Set wsAssumptions = ThisWorkbook.Sheets("Assumptions")

    ' Define scaling factor (convert to millions)
    scaleFactor = 1000000

    ' Identify the last column for projection data
    lastCol = wsAssumptions.Cells(9, wsAssumptions.Columns.Count).End(xlToLeft).Column

    ' First projection column (skip C to E, start at F)
    firstProjectionCol = 6

    ' Calculate average Gross Margin % from historical data (Row 13 in DCF)
    grossMarginAvg = Application.WorksheetFunction.Average(wsDCF.Range("F13:J13"))

    ' Fill Base Case % Growth (Sales) - Row 11 (only light yellow regions)
    For i = 0 To (lastCol - firstProjectionCol)
        If wsAssumptions.Cells(11, firstProjectionCol + i).Interior.Color = RGB(255, 255, 153) Then ' Light Yellow
            wsAssumptions.Cells(11, firstProjectionCol + i).Formula = "=" & wsDCF.Range("I8").Address(True, True, xlA1, True) & " / " & scaleFactor
        End If
    Next i

    ' Fill Base Case COGS % (Cost of Goods Sold) - Row 18
    For i = 0 To (lastCol - firstProjectionCol)
        If wsAssumptions.Cells(18, firstProjectionCol + i).Interior.Color = RGB(255, 255, 153) Then
            wsAssumptions.Cells(18, firstProjectionCol + i).Formula = "=" & wsDCF.Range("I9").Address(True, True, xlA1, True) & " / " & scaleFactor
        End If
    Next i

    ' Fill Base Case SG&A % - Row 25
    For i = 0 To (lastCol - firstProjectionCol)
        If wsAssumptions.Cells(25, firstProjectionCol + i).Interior.Color = RGB(255, 255, 153) Then
            wsAssumptions.Cells(25, firstProjectionCol + i).Formula = "=" & wsDCF.Range("I12").Address(True, True, xlA1, True) & " / " & scaleFactor
        End If
    Next i

    ' Fill Base Case Depreciation & Amortization % Sales - Row 32
    For i = 0 To (lastCol - firstProjectionCol)
        If wsAssumptions.Cells(32, firstProjectionCol + i).Interior.Color = RGB(255, 255, 153) Then
            wsAssumptions.Cells(32, firstProjectionCol + i).Formula = "=" & wsDCF.Range("I15").Address(True, True, xlA1, True) & " / " & scaleFactor
        End If
    Next i

    ' Fill Base Case Gross Margin % - Row 13 (Light Yellow Cells)
    For i = 0 To (lastCol - firstProjectionCol)
        If wsAssumptions.Cells(13, firstProjectionCol + i).Interior.Color = RGB(255, 255, 153) Then
            wsAssumptions.Cells(13, firstProjectionCol + i).Value = grossMarginAvg
        End If
    Next i

    MsgBox "Base Case assumptions filled successfully, including calculated Gross Margin!", vbInformation
End Sub


