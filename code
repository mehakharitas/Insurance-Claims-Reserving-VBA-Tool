Option Explicit
Dim i As Integer, j As Integer
Dim rng As Range

Sub SetupClaimsTriangle()
    Dim n As Integer, m As Integer, startYear As Integer
    Dim ws As Worksheet
    Dim wsmacros As Worksheet
    
    
    ' Set up worksheets
    Set wsmacros = ThisWorkbook.Sheets("macros")
    
    ' Check if output sheet exists; if not, create it.
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Input")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=wsmacros)
        ws.Name = "Input"
    Else
        ws.Cells.Clear
    End If
    
    ' Inputs
    n = InputBox("Enter number of Accident Years (n):")
    m = InputBox("Enter number of Development Years (m):")
    startYear = InputBox("Enter Start Year (e.g., 2010):")
    
    ' Headers
    ws.Cells(1, 1).Value = "Accident Year"
    For j = 0 To m - 1
        ws.Cells(1, j + 2).Value = "Dev " & j
    Next j
    
    ' Accident Years
    For i = 1 To n
        ws.Cells(i + 1, 1).Value = startYear + (i - 1)
    Next i
    
    ' Input triangle (incremental)
    Set rng = ws.Range(ws.Cells(2, 2), ws.Cells(n + 1, m + 1))
    rng.Interior.Color = RGB(255, 255, 153) ' yellow
    
    ' Add cumulative title
    ws.Cells(n + 3, 1).Value = "Cumulative Claims"
    
    ' Accident years again for cumulative
    For i = 1 To n
        ws.Cells(n + 3 + i, 1).Value = startYear + (i - 1)
    Next i
    
    ' Dev year headers for cumulative
    For j = 0 To m - 1
        ws.Cells(n + 3, j + 2).Value = "Dev " & j
    Next j
    
    ' Cumulative formulas (upper triangle only)
    For i = 1 To n
        For j = 1 To m
            If j <= m - i + 1 Then
                ws.Cells(n + 3 + i, j + 1).FormulaR1C1 = _
                   "=SUM(R" & (i + 1) & "C2:R" & (i + 1) & "C" & (j + 1) & ")"
            Else
                ws.Cells(n + 3 + i, j + 1).Value = ""
            End If
        Next j
    Next i
    
    ws.Range(ws.Cells(n + 4, 2), ws.Cells(n + 3 + n, m + 1)).Interior.Color = RGB(204, 229, 255) ' light blue
    
    MsgBox "Triangle setup complete! Enter non - incremental claims in the upper triangle."
End Sub

Sub ChainLadder()
    Dim wsInput As Worksheet, wsOutput As Worksheet
    Dim n As Long, m As Long
    Dim cumulativeStartRowInput As Long
    
    ' Set up input and output worksheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    
    ' Check if output sheet exists; if not, create it.
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets("ChainLadder")
    On Error GoTo 0
    If wsOutput Is Nothing Then
        Set wsOutput = ThisWorkbook.Sheets.Add(After:=wsInput)
        wsOutput.Name = "ChainLadder"
    Else
        wsOutput.Cells.Clear
    End If
    
    ' Determine n and m based on the incremental table in your SetupClaimsTriangle macro.
    ' n is the number of rows in the incremental data, which ends at row n+1.
    n = wsInput.Range("A1").End(xlDown).Row - 1
    ' m is the number of columns in the incremental data, which ends at column m+1.
    m = wsInput.Range("A1").End(xlToRight).Column - 1
    
    ' The cumulative table starts exactly 2 rows below the last row of incremental data.
    ' Last row of incremental data is n + 1. So, the cumulative table starts at (n + 1) + 2 = n + 3.
    cumulativeStartRowInput = n + 3
    
    ' Define the range of the cumulative claims table on the Input sheet.
    ' It includes the headers and the data below it.
    Dim cumulativeTableRange As Range
    Set cumulativeTableRange = wsInput.Range(wsInput.Cells(cumulativeStartRowInput, 1), wsInput.Cells(cumulativeStartRowInput + n, m + 1))
    
    ' Copy the entire cumulative claims table and paste it to the new sheet, starting at cell A1.
    cumulativeTableRange.Copy
    wsOutput.Cells(1, 1).PasteSpecial Paste:=xlPasteAll
    
    ' Auto-fit columns for better visibility.
    wsOutput.Columns.AutoFit

    
    ' === Step 2: Calculate and Display Development Factors ===
    Dim startRowLDF As Long
    startRowLDF = n + 3 ' This places the LDF table below the copied claims data
    
    wsOutput.Cells(startRowLDF, 1).Value = "Development Factors (LDFs)"
    wsOutput.Cells(startRowLDF, 1).Font.Bold = True
    wsOutput.Cells(startRowLDF + 1, 1).Value = "Dev Period"
    wsOutput.Cells(startRowLDF + 1, 2).Value = "LDF"
    
    Dim ldfArray() As Double
    ReDim ldfArray(1 To m - 1)
    
    ' Loop through each development period to calculate the LDF
    For j = 1 To m - 1
        Dim numerator As Double, denominator As Double
        numerator = 0
        denominator = 0
        
        ' Iterate through the copied cumulative claims triangle for the ratio calculation
        For i = 1 To n - j
            ' The copied data starts at row 2 on the ChainLadder sheet
            numerator = numerator + wsOutput.Cells(i + 1, j + 2).Value
            denominator = denominator + wsOutput.Cells(i + 1, j + 1).Value
        Next i
        
        If denominator > 0 Then
            ldfArray(j) = numerator / denominator
        Else
            ldfArray(j) = 1 ' Avoid division by zero
        End If
        
        ' Output LDFs to the ChainLadder sheet
        wsOutput.Cells(startRowLDF + j + 1, 1).Value = "Dev " & j & " to " & "Dev " & j + 1
        wsOutput.Cells(startRowLDF + j + 1, 2).Value = ldfArray(j)
        wsOutput.Cells(startRowLDF + j + 1, 2).Interior.Color = RGB(255, 255, 153) ' Yellow color
    Next j
    
    ' Auto-fit columns for better visibility
    wsOutput.Columns.AutoFit
    
    ' === Step 3: Create Detailed Projected Claims Table ===
    Dim startRowProjection As Long
    startRowProjection = startRowLDF + m + 2
    
    wsOutput.Cells(startRowProjection, 1).Value = "Projected Claims"
    wsOutput.Cells(startRowProjection, 1).Font.Bold = True
    
    ' Header for the new table
    wsOutput.Cells(startRowProjection + 1, 1).Value = "Accident Year"
    Dim headerCol As Long
    For headerCol = 2 To m + 1
        wsOutput.Cells(startRowProjection + 1, headerCol).Value = "Dev " & headerCol - 2
    Next headerCol
    wsOutput.Cells(startRowProjection + 1, m + 2).Value = "Ultimate Claim"
    
    For i = 1 To n
        ' Get the latest known cumulative claim for this accident year
        Dim latestCumulative As Double
        latestCumulative = wsOutput.Cells(i + 1, m - i + 2).Value
        
        ' Start a new row for the accident year
        wsOutput.Cells(startRowProjection + i + 1, 1).Value = wsInput.Cells(i + 1, 1).Value
        wsOutput.Cells(startRowProjection + i + 1, m - i + 2).Value = latestCumulative
        
        ' Loop to apply development factors and fill in the projected values
        Dim projectedClaim As Double
        projectedClaim = latestCumulative
        
        For j = m - i + 1 To m - 1
            ' Apply the development factor
            If j >= 1 And j <= m - 1 Then
                projectedClaim = projectedClaim * ldfArray(j)
                ' Place the intermediate projected value in the correct column
                wsOutput.Cells(startRowProjection + i + 1, j + 2).Value = projectedClaim
                wsOutput.Cells(startRowProjection + i + 1, j + 2).Font.Color = RGB(255, 0, 0) ' Red color
            End If
        Next j
        
        ' Place the final ultimate claim in the last column
        wsOutput.Cells(startRowProjection + i + 1, m + 2).Value = projectedClaim
        wsOutput.Cells(startRowProjection + i + 1, m + 2).Font.Color = RGB(255, 0, 0) ' Red color
        Next i
' Auto-fit columns for better visibility
    wsOutput.Columns.AutoFit
' === Step 4: Add the Reserves Table ===
    Dim startRowReserves As Long
    startRowReserves = startRowProjection + n + 3
    
    wsOutput.Cells(startRowReserves, 1).Value = "Reserves"
    wsOutput.Cells(startRowReserves, 1).Font.Bold = True
    wsOutput.Cells(startRowReserves + 1, 1).Value = "Accident Year"
    wsOutput.Cells(startRowReserves + 1, 2).Value = "Reserve"
    
    Dim totalReserve As Double
    totalReserve = 0
    
    For i = 1 To n
        Dim latestCumulativeFromTable As Double
        latestCumulativeFromTable = wsOutput.Cells(i + 1, m - i + 2).Value
        
        Dim ultimateClaimFromTable As Double
        ultimateClaimFromTable = wsOutput.Cells(startRowProjection + i + 1, m + 2).Value
        
        Dim reserve As Double
        reserve = ultimateClaimFromTable - latestCumulativeFromTable
        totalReserve = totalReserve + reserve
        
        wsOutput.Cells(startRowReserves + i + 1, 1).Value = wsInput.Cells(i + 1, 1).Value
        wsOutput.Cells(startRowReserves + i + 1, 2).Value = reserve
    Next i
    
    wsOutput.Cells(startRowReserves + n + 2, 1).Value = "Total Reserve:"
    wsOutput.Cells(startRowReserves + n + 2, 1).Font.Bold = True
    wsOutput.Cells(startRowReserves + n + 2, 2).Value = totalReserve
    wsOutput.Cells(startRowReserves + n + 2, 2).Interior.Color = RGB(204, 229, 255)
    
    wsOutput.Columns.AutoFit
    MsgBox "Chain Ladder calculations complete! Results are on the 'ChainLadder' sheet."
End Sub

Sub BornhuetterFerguson()
    Dim wsInput As Worksheet, wsOutput As Worksheet
    Dim n As Long, m As Long
    Dim cumulativeStartRowInput As Long
    Dim wsChainLadder As Worksheet
    
    
    ' Set up input and output worksheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsChainLadder = ThisWorkbook.Sheets("ChainLadder")
    
    ' Check if output sheet exists; if not, create it.
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets("BornhuetterFerguson")
    On Error GoTo 0
    If wsOutput Is Nothing Then
        Set wsOutput = ThisWorkbook.Sheets.Add(After:=wsChainLadder)
        wsOutput.Name = "BornhuetterFerguson"
    Else
        wsOutput.Cells.Clear
    End If
    
    ' Determine n and m based on the incremental table in your SetupClaimsTriangle macro.
    ' n is the number of rows in the incremental data, which ends at row n+1.
    n = wsInput.Range("A1").End(xlDown).Row - 1
    ' m is the number of columns in the incremental data, which ends at column m+1.
    m = wsInput.Range("A1").End(xlToRight).Column - 1
    
    ' The cumulative table starts exactly 2 rows below the last row of incremental data.
    ' Last row of incremental data is n + 1. So, the cumulative table starts at (n + 1) + 2 = n + 3.
    cumulativeStartRowInput = n + 3
    
    ' Define the range of the cumulative claims table on the Input sheet.
    ' It includes the headers and the data below it.
    Dim cumulativeTableRange As Range
    Set cumulativeTableRange = wsInput.Range(wsInput.Cells(cumulativeStartRowInput, 1), wsInput.Cells(cumulativeStartRowInput + n, m + 1))
    
    ' Copy the entire cumulative claims table and paste it to the new sheet, starting at cell A1.
    cumulativeTableRange.Copy
    wsOutput.Cells(1, 1).PasteSpecial Paste:=xlPasteAll
    
    ' Auto-fit columns for better visibility.
    wsOutput.Columns.AutoFit
    
    ' === Step 2: Get Inputs and Set Up Table Headers ===
    Dim startRowTable As Long
    startRowTable = n + 3
    
    ' Prompt for Expected Loss Ratio
    Dim expectedLossRatio As Double
    expectedLossRatio = InputBox("Enter the Expected Loss Ratio (e.g., 0.65 for 65%):")
    Dim lossRatioCell As Range
    Set lossRatioCell = wsOutput.Cells(startRowTable + 1, 10)
    lossRatioCell.Offset(-1, 0).Value = "Expected Loss Ratio:"
    lossRatioCell.Value = expectedLossRatio
    ' Get Premium for each Accident Year
    Dim premiums() As Double
    ReDim premiums(1 To n)
    For i = 1 To n
        premiums(i) = InputBox("Enter premium for Accident Year " & wsInput.Cells(i + 1, 1).Value & ":")
    Next i
    
    ' Set up table headers
    wsOutput.Cells(startRowTable, 1).Value = "Bornhuetter-Ferguson Calculation"
    wsOutput.Cells(startRowTable, 1).Font.Bold = True
    
    wsOutput.Cells(startRowTable + 1, 1).Value = "Accident Year"
    wsOutput.Cells(startRowTable + 1, 2).Value = "Premium"
    wsOutput.Cells(startRowTable + 1, 3).Value = "Latest Cumulative"
    wsOutput.Cells(startRowTable + 1, 4).Value = "development Ratio"
    wsOutput.Cells(startRowTable + 1, 5).Value = "Development Factor"
    wsOutput.Cells(startRowTable + 1, 6).Value = "1-1/F"
    wsOutput.Cells(startRowTable + 1, 7).Value = "IUL"
    wsOutput.Cells(startRowTable + 1, 8).Value = "Emerging Liabilities"
    
    ' Calculate LDFs for all development periods
    Dim ldfArray() As Double
    ReDim ldfArray(1 To m - 1)
    For j = 1 To m - 1
        Dim numerator As Double, denominator As Double
        numerator = 0
        denominator = 0
        For i = 1 To n - j
            numerator = numerator + wsOutput.Cells(i + 1, j + 2).Value
            denominator = denominator + wsOutput.Cells(i + 1, j + 1).Value
        Next i
        If denominator > 0 Then
            ldfArray(j) = numerator / denominator
        Else
            ldfArray(j) = 1
        End If
         wsOutput.Cells(startRowTable + 1 + i, 4).Value = ldfArray(j)
    Next j
    
    ' === Step 3: Populate Table with Data and Calculations ===
    Dim totalReserve As Double
    totalReserve = 0
    
    For i = 1 To n
        Dim latestCumulative As Double
        latestCumulative = wsOutput.Cells(i + 1, wsOutput.Columns.Count).End(xlToLeft).Value
        
        Dim devFactor As Double
        devFactor = 1
        For j = m - i + 1 To m - 1
            If j >= 1 And j <= m - 1 Then
                devFactor = devFactor * ldfArray(j)
            End If
        Next j
        
        Dim oneMinusOneOverF As Double
        If devFactor > 0 Then
            oneMinusOneOverF = 1 - (1 / devFactor)
        Else
            oneMinusOneOverF = 0
        End If
        
        Dim iul As Double
        iul = premiums(i) * expectedLossRatio
        
        Dim emergingLiabilities As Double
        emergingLiabilities = iul * oneMinusOneOverF
        
        Dim reserve As Double
        ' Corrected Reserve Calculation based on your request
        reserve = emergingLiabilities
        totalReserve = totalReserve + reserve
        
        wsOutput.Cells(startRowTable + 1 + i, 1).Value = wsInput.Cells(i + 1, 1).Value
        wsOutput.Cells(startRowTable + 1 + i, 2).Value = premiums(i)
        wsOutput.Cells(startRowTable + 1 + i, 3).Value = latestCumulative
        wsOutput.Cells(startRowTable + 1 + i, 5).Value = devFactor
        wsOutput.Cells(startRowTable + 1 + i, 6).Value = oneMinusOneOverF
        wsOutput.Cells(startRowTable + 1 + i, 7).Value = iul
        wsOutput.Cells(startRowTable + 1 + i, 8).Value = emergingLiabilities
        wsOutput.Cells(startRowTable + 1 + i, 8).Interior.Color = RGB(204, 229, 255)
        
    Next i
 
    
    ' Display the total reserve
    wsOutput.Cells(startRowTable + n + 2, 7).Value = "Total Reserve:"
    wsOutput.Cells(startRowTable + n + 2, 8).Value = totalReserve
    wsOutput.Cells(startRowTable + n + 2, 8).Interior.Color = RGB(255, 255, 153)
    
    wsOutput.Columns.AutoFit
    MsgBox "Bornhuetter-Ferguson calculations are complete and displayed on the sheet."
End Sub

Sub ScenarioAnalysis()
    Dim wsInput As Worksheet, wsOutput As Worksheet, wsBF As Worksheet
    Dim n As Long, m As Long
    Dim cumulativeStartRowInput As Long
    
    ' Set up input and output worksheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsBF = ThisWorkbook.Sheets("BornhuetterFerguson")
    
    ' Check if output sheet exists; if not, create it.
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets("ScenarioAnalysis")
    On Error GoTo 0
    If wsOutput Is Nothing Then
        Set wsOutput = ThisWorkbook.Sheets.Add(After:=wsBF)
        wsOutput.Name = "ScenarioAnalysis"
    Else
        wsOutput.Cells.Clear
    End If
    
    ' Determine n and m from Input sheet
    n = wsInput.Range("A1").End(xlDown).Row - 1
    m = wsInput.Range("A1").End(xlToRight).Column - 1
    cumulativeStartRowInput = n + 3
    
    ' Copy the cumulative claims table from Input sheet
    Dim cumulativeTableRange As Range
    Set cumulativeTableRange = wsInput.Range(wsInput.Cells(cumulativeStartRowInput, 1), wsInput.Cells(cumulativeStartRowInput + n, m + 1))
    cumulativeTableRange.Copy
    wsOutput.Cells(1, 1).PasteSpecial Paste:=xlPasteAll
    wsOutput.Columns.AutoFit
    
    ' Get Expected Loss Ratio and Premiums from BF sheet
    Dim expectedLossRatio As Double
    Dim premiums() As Double
    ReDim premiums(1 To n)
    
    ' Find the Expected Loss Ratio from BF sheet (assuming it's in a specific location)
    expectedLossRatio = wsBF.Cells(n + 4, 10).Value ' Adjust if needed
    
    ' Get premiums from BF sheet
    For i = 1 To n
        premiums(i) = wsBF.Cells(n + 4 + i, 2).Value
    Next i
    
    ' Calculate base LDFs
    Dim ldfArray() As Double
    ReDim ldfArray(1 To m - 1)
    For j = 1 To m - 1
        Dim numerator As Double, denominator As Double
        numerator = 0
        denominator = 0
        For i = 1 To n - j
            numerator = numerator + wsOutput.Cells(i + 1, j + 2).Value
            denominator = denominator + wsOutput.Cells(i + 1, j + 1).Value
        Next i
        If denominator > 0 Then
            ldfArray(j) = numerator / denominator
        Else
            ldfArray(j) = 1
        End If
    Next j
    
    ' Create stressed LDF arrays
    Dim ldfArrayPlus10() As Double, ldfArrayMinus10() As Double
    ReDim ldfArrayPlus10(1 To m - 1)
    ReDim ldfArrayMinus10(1 To m - 1)
    
    For j = 1 To m - 1
        ldfArrayPlus10(j) = ldfArray(j) * 1.1    ' +10% stress
        ldfArrayMinus10(j) = ldfArray(j) * 0.9   ' -10% stress
    Next j
    
    ' === Scenario 1: Development Factors +10% ===
    Dim startRowScenario1 As Long
    startRowScenario1 = n + 3
    
    wsOutput.Cells(startRowScenario1, 1).Value = "Scenario 1: Development Factors +10%"
    wsOutput.Cells(startRowScenario1, 1).Font.Bold = True
    wsOutput.Cells(startRowScenario1, 1).Interior.Color = RGB(255, 204, 204) ' Light red
    
    ' Display Expected Loss Ratio for Scenario 1
    wsOutput.Cells(startRowScenario1 + 1, 10).Offset(-1, 0).Value = "Expected Loss Ratio:"
    wsOutput.Cells(startRowScenario1 + 1, 10).Value = expectedLossRatio
    
    ' Set up headers for Scenario 1
    wsOutput.Cells(startRowScenario1 + 1, 1).Value = "Accident Year"
    wsOutput.Cells(startRowScenario1 + 1, 2).Value = "Premium"
    wsOutput.Cells(startRowScenario1 + 1, 3).Value = "Latest Cumulative"
    wsOutput.Cells(startRowScenario1 + 1, 4).Value = "Development Ratio"
    wsOutput.Cells(startRowScenario1 + 1, 5).Value = "Development Factor (+10%)"
    wsOutput.Cells(startRowScenario1 + 1, 6).Value = "1-1/F"
    wsOutput.Cells(startRowScenario1 + 1, 7).Value = "IUL"
    wsOutput.Cells(startRowScenario1 + 1, 8).Value = "Emerging Liabilities"
    
    Dim totalReserveScenario1 As Double
    totalReserveScenario1 = 0
    
    ' Calculate Scenario 1 values (+10% LDFs)
    For i = 1 To n
        Dim latestCumulative As Double
        latestCumulative = wsOutput.Cells(i + 1, wsOutput.Columns.Count).End(xlToLeft).Value
        
        ' Get the next development ratio for this accident year (using +10% stressed LDFs)
        Dim devRatio1 As Double
        If m - i + 1 <= m - 1 And m - i + 1 >= 1 Then
            devRatio1 = ldfArrayPlus10(m - i + 1) ' +10% stressed LDF for next step
        Else
            devRatio1 = 1 ' No further development needed
        End If
        
        ' Calculate development factor using +10% stressed LDFs
        Dim devFactor1 As Double
        devFactor1 = 1
        For j = m - i + 1 To m - 1
            If j >= 1 And j <= m - 1 Then
                devFactor1 = devFactor1 * ldfArrayPlus10(j) ' Use +10% stressed LDFs
            End If
        Next j
        
        Dim oneMinusOneOverF1 As Double
        If devFactor1 > 0 Then
            oneMinusOneOverF1 = 1 - (1 / devFactor1)
        Else
            oneMinusOneOverF1 = 0
        End If
        
        Dim iul1 As Double
        iul1 = premiums(i) * expectedLossRatio
        
        Dim emergingLiabilities1 As Double
        emergingLiabilities1 = iul1 * oneMinusOneOverF1
        
        totalReserveScenario1 = totalReserveScenario1 + emergingLiabilities1
        
        ' Populate Scenario 1 table
        wsOutput.Cells(startRowScenario1 + 1 + i, 1).Value = wsInput.Cells(i + 1, 1).Value
        wsOutput.Cells(startRowScenario1 + 1 + i, 2).Value = premiums(i)
        wsOutput.Cells(startRowScenario1 + 1 + i, 3).Value = latestCumulative
        wsOutput.Cells(startRowScenario1 + 1 + i, 4).Value = devRatio1
        wsOutput.Cells(startRowScenario1 + 1 + i, 5).Value = devFactor1
        wsOutput.Cells(startRowScenario1 + 1 + i, 6).Value = oneMinusOneOverF1
        wsOutput.Cells(startRowScenario1 + 1 + i, 7).Value = iul1
        wsOutput.Cells(startRowScenario1 + 1 + i, 8).Value = emergingLiabilities1
        wsOutput.Cells(startRowScenario1 + 1 + i, 8).Interior.Color = RGB(255, 230, 230) ' Light red
    Next i
    
    ' Total Reserve for Scenario 1
    wsOutput.Cells(startRowScenario1 + n + 2, 7).Value = "Total Reserve (+10%):"
    wsOutput.Cells(startRowScenario1 + n + 2, 7).Font.Bold = True
    wsOutput.Cells(startRowScenario1 + n + 2, 8).Value = totalReserveScenario1
    wsOutput.Cells(startRowScenario1 + n + 2, 8).Interior.Color = RGB(255, 153, 153) ' Red
    wsOutput.Cells(startRowScenario1 + n + 2, 8).Font.Bold = True
    
    ' === Scenario 2: Development Factors -10% ===
    Dim startRowScenario2 As Long
    startRowScenario2 = startRowScenario1 + n + 5
    
    wsOutput.Cells(startRowScenario2, 1).Value = "Scenario 2: Development Factors -10%"
    wsOutput.Cells(startRowScenario2, 1).Font.Bold = True
    wsOutput.Cells(startRowScenario2, 1).Interior.Color = RGB(204, 255, 204) ' Light green
    
    ' Display Expected Loss Ratio for Scenario 2
    wsOutput.Cells(startRowScenario2 + 1, 10).Offset(-1, 0).Value = "Expected Loss Ratio:"
    wsOutput.Cells(startRowScenario2 + 1, 10).Value = expectedLossRatio
    
    ' Set up headers for Scenario 2
    wsOutput.Cells(startRowScenario2 + 1, 1).Value = "Accident Year"
    wsOutput.Cells(startRowScenario2 + 1, 2).Value = "Premium"
    wsOutput.Cells(startRowScenario2 + 1, 3).Value = "Latest Cumulative"
    wsOutput.Cells(startRowScenario2 + 1, 4).Value = "Development Ratio"
    wsOutput.Cells(startRowScenario2 + 1, 5).Value = "Development Factor (-10%)"
    wsOutput.Cells(startRowScenario2 + 1, 6).Value = "1-1/F"
    wsOutput.Cells(startRowScenario2 + 1, 7).Value = "IUL"
    wsOutput.Cells(startRowScenario2 + 1, 8).Value = "Emerging Liabilities"
    
    Dim totalReserveScenario2 As Double
    totalReserveScenario2 = 0
    
    ' Calculate Scenario 2 values (-10% LDFs)
    For i = 1 To n
        latestCumulative = wsOutput.Cells(i + 1, wsOutput.Columns.Count).End(xlToLeft).Value
        
        ' Get the next development ratio for this accident year (using -10% stressed LDFs)
        Dim devRatio2 As Double
        If m - i + 1 <= m - 1 And m - i + 1 >= 1 Then
            devRatio2 = ldfArrayMinus10(m - i + 1) ' -10% stressed LDF for next step
        Else
            devRatio2 = 1 ' No further development needed
        End If
        
        ' Calculate development factor using -10% stressed LDFs
        Dim devFactor2 As Double
        devFactor2 = 1
        For j = m - i + 1 To m - 1
            If j >= 1 And j <= m - 1 Then
                devFactor2 = devFactor2 * ldfArrayMinus10(j) ' Use -10% stressed LDFs
            End If
        Next j
        
        Dim oneMinusOneOverF2 As Double
        If devFactor2 > 0 Then
            oneMinusOneOverF2 = 1 - (1 / devFactor2)
        Else
            oneMinusOneOverF2 = 0
        End If
        
        Dim iul2 As Double
        iul2 = premiums(i) * expectedLossRatio
        
        Dim emergingLiabilities2 As Double
        emergingLiabilities2 = iul2 * oneMinusOneOverF2
        
        totalReserveScenario2 = totalReserveScenario2 + emergingLiabilities2
        
        ' Populate Scenario 2 table
        wsOutput.Cells(startRowScenario2 + 1 + i, 1).Value = wsInput.Cells(i + 1, 1).Value
        wsOutput.Cells(startRowScenario2 + 1 + i, 2).Value = premiums(i)
        wsOutput.Cells(startRowScenario2 + 1 + i, 3).Value = latestCumulative
        wsOutput.Cells(startRowScenario2 + 1 + i, 4).Value = devRatio2
        wsOutput.Cells(startRowScenario2 + 1 + i, 5).Value = devFactor2
        wsOutput.Cells(startRowScenario2 + 1 + i, 6).Value = oneMinusOneOverF2
        wsOutput.Cells(startRowScenario2 + 1 + i, 7).Value = iul2
        wsOutput.Cells(startRowScenario2 + 1 + i, 8).Value = emergingLiabilities2
        wsOutput.Cells(startRowScenario2 + 1 + i, 8).Interior.Color = RGB(230, 255, 230) ' Light green
    Next i
    
    ' Total Reserve for Scenario 2
    wsOutput.Cells(startRowScenario2 + n + 2, 7).Value = "Total Reserve (-10%):"
    wsOutput.Cells(startRowScenario2 + n + 2, 7).Font.Bold = True
    wsOutput.Cells(startRowScenario2 + n + 2, 8).Value = totalReserveScenario2
    wsOutput.Cells(startRowScenario2 + n + 2, 8).Interior.Color = RGB(153, 255, 153) ' Green
    wsOutput.Cells(startRowScenario2 + n + 2, 8).Font.Bold = True
    
    ' === Summary Comparison Table ===
    Dim startRowSummary As Long
    startRowSummary = startRowScenario2 + n + 5
    
    wsOutput.Cells(startRowSummary, 1).Value = "Scenario Summary"
    wsOutput.Cells(startRowSummary, 1).Font.Bold = True
    wsOutput.Cells(startRowSummary, 1).Interior.Color = RGB(255, 255, 153) ' Yellow
    
    wsOutput.Cells(startRowSummary + 1, 1).Value = "Scenario"
    wsOutput.Cells(startRowSummary + 1, 2).Value = "Total Reserve"
    wsOutput.Cells(startRowSummary + 1, 3).Value = "Difference from Base"
    
    ' Get base reserve from BF sheet for comparison
    Dim baseReserve As Double
    baseReserve = wsBF.Cells(wsBF.Cells.Find("Total Reserve:").Row, wsBF.Cells.Find("Total Reserve:").Column + 1).Value
    
    wsOutput.Cells(startRowSummary + 2, 1).Value = "Base Case"
    wsOutput.Cells(startRowSummary + 2, 2).Value = baseReserve
    wsOutput.Cells(startRowSummary + 2, 3).Value = 0
    
    wsOutput.Cells(startRowSummary + 3, 1).Value = "Dev Factors +10%"
    wsOutput.Cells(startRowSummary + 3, 2).Value = totalReserveScenario1
    wsOutput.Cells(startRowSummary + 3, 3).Value = totalReserveScenario1 - baseReserve
    wsOutput.Cells(startRowSummary + 3, 2).Interior.Color = RGB(255, 230, 230)
    wsOutput.Cells(startRowSummary + 3, 3).Interior.Color = RGB(255, 230, 230)
    
    wsOutput.Cells(startRowSummary + 4, 1).Value = "Dev Factors -10%"
    wsOutput.Cells(startRowSummary + 4, 2).Value = totalReserveScenario2
    wsOutput.Cells(startRowSummary + 4, 3).Value = totalReserveScenario2 - baseReserve
    wsOutput.Cells(startRowSummary + 4, 2).Interior.Color = RGB(230, 255, 230)
    wsOutput.Cells(startRowSummary + 4, 3).Interior.Color = RGB(230, 255, 230)
    
    wsOutput.Columns.AutoFit
    MsgBox "Scenario Analysis complete! Results show stress testing with ±10% development factors."
End Sub

