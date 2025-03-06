
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
    Sub FillAssumptions()
    Dim wsDCF As Worksheet, wsAssumptions As Worksheet, wsNWC As Worksheet
    Dim firstCol As Integer, lastCol As Integer
    Dim i As Integer
    Dim grossMarginAvg As Double, cogsAvg As Double, capexAvg As Double, wcAvg As Double
    Dim arAvg As Double, invAvg As Double, prepaidAvg As Double, apAvg As Double, accruedLiabAvg As Double, otherLiabAvg As Double
    Dim downsideCase As Double, upsideCase As Double
    
    ' Set worksheets
    Set wsDCF = ThisWorkbook.Sheets("DCF")
    Set wsNWC = ThisWorkbook.Sheets("NWC")
    Set wsAssumptions = ThisWorkbook.Sheets("Assumptions")

    ' Define first and last columns (2021-2024 range)
    firstCol = 6  ' Column F (assuming first projection starts at F)
    lastCol = 9   ' Column I (corresponding to 2024)

    ' Calculate Gross Margin % Average from 2021-2024 (Row 13 in DCF)
    grossMarginAvg = Application.WorksheetFunction.Average(wsDCF.Range(wsDCF.Cells(13, firstCol), wsDCF.Cells(13, lastCol)))

    ' Calculate Downside and Upside Cases
    downsideCase = grossMarginAvg - 0.01 ' Downside Case (Base - 1%)
    upsideCase = grossMarginAvg + 0.01   ' Upside Case (Base + 1%)

    ' 💾 Balance Sheet Assumptions (From NWC Tab - Calculated as % of Gross Margin)
    ' Accounts Receivable % Sales
    arAvg = Application.WorksheetFunction.Average(wsNWC.Range(wsNWC.Cells(48, firstCol), wsNWC.Cells(48, lastCol))) / grossMarginAvg

    ' Inventories % Sales
    invAvg = Application.WorksheetFunction.Average(wsNWC.Range(wsNWC.Cells(55, firstCol), wsNWC.Cells(55, lastCol))) / grossMarginAvg

    ' Prepaid Expenses % Sales
    prepaidAvg = Application.WorksheetFunction.Average(wsNWC.Range(wsNWC.Cells(62, firstCol), wsNWC.Cells(62, lastCol))) / grossMarginAvg

    ' Accounts Payable % Sales
    apAvg = Application.WorksheetFunction.Average(wsNWC.Range(wsNWC.Cells(69, firstCol), wsNWC.Cells(69, lastCol))) / grossMarginAvg

    ' Accrued Liabilities % Sales
    accruedLiabAvg = Application.WorksheetFunction.Average(wsNWC.Range(wsNWC.Cells(76, firstCol), wsNWC.Cells(76, lastCol))) / grossMarginAvg

    ' Other Current Liabilities % Sales
    otherLiabAvg = Application.WorksheetFunction.Average(wsNWC.Range(wsNWC.Cells(83, firstCol), wsNWC.Cells(83, lastCol))) / grossMarginAvg

    ' Fill Base Case (2021-2024)
    wsAssumptions.Range(wsAssumptions.Cells(48, firstCol), wsAssumptions.Cells(48, lastCol)).Value = arAvg ' Accounts Receivable
    wsAssumptions.Range(wsAssumptions.Cells(55, firstCol), wsAssumptions.Cells(55, lastCol)).Value = invAvg ' Inventories
    wsAssumptions.Range(wsAssumptions.Cells(62, firstCol), wsAssumptions.Cells(62, lastCol)).Value = prepaidAvg ' Prepaid Expenses
    wsAssumptions.Range(wsAssumptions.Cells(69, firstCol), wsAssumptions.Cells(69, lastCol)).Value = apAvg ' Accounts Payable
    wsAssumptions.Range(wsAssumptions.Cells(76, firstCol), wsAssumptions.Cells(76, lastCol)).Value = accruedLiabAvg ' Accrued Liabilities
    wsAssumptions.Range(wsAssumptions.Cells(83, firstCol), wsAssumptions.Cells(83, lastCol)).Value = otherLiabAvg ' Other Liabilities

    ' Fill Downside Case (Base - 1%)
    wsAssumptions.Range(wsAssumptions.Cells(49, firstCol), wsAssumptions.Cells(49, lastCol)).Value = arAvg - 0.01
    wsAssumptions.Range(wsAssumptions.Cells(56, firstCol), wsAssumptions.Cells(56, lastCol)).Value = invAvg - 0.01
    wsAssumptions.Range(wsAssumptions.Cells(63, firstCol), wsAssumptions.Cells(63, lastCol)).Value = prepaidAvg - 0.01
    wsAssumptions.Range(wsAssumptions.Cells(70, firstCol), wsAssumptions.Cells(70, lastCol)).Value = apAvg - 0.01
    wsAssumptions.Range(wsAssumptions.Cells(77, firstCol), wsAssumptions.Cells(77, lastCol)).Value = accruedLiabAvg - 0.01
    wsAssumptions.Range(wsAssumptions.Cells(84, firstCol), wsAssumptions.Cells(84, lastCol)).Value = otherLiabAvg - 0.01

    ' Fill Upside Case (Base + 1%)
    wsAssumptions.Range(wsAssumptions.Cells(50, firstCol), wsAssumptions.Cells(50, lastCol)).Value = arAvg + 0.01
    wsAssumptions.Range(wsAssumptions.Cells(57, firstCol), wsAssumptions.Cells(57, lastCol)).Value = invAvg + 0.01
    wsAssumptions.Range(wsAssumptions.Cells(64, firstCol), wsAssumptions.Cells(64, lastCol)).Value = prepaidAvg + 0.01
    wsAssumptions.Range(wsAssumptions.Cells(71, firstCol), wsAssumptions.Cells(71, lastCol)).Value = apAvg + 0.01
    wsAssumptions.Range(wsAssumptions.Cells(78, firstCol), wsAssumptions.Cells(78, lastCol)).Value = accruedLiabAvg + 0.01
    wsAssumptions.Range(wsAssumptions.Cells(85, firstCol), wsAssumptions.Cells(85, lastCol)).Value = otherLiabAvg + 0.01

    MsgBox "Assumptions filled successfully for Base, Downside (-1%), and Upside (+1%) Cases (2021-2024)!", vbInformation
End Sub




