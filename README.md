<p align="center">
  <img src="https://github.com/theidari/VBA-challenge/blob/main/Result%20and%20File%20IMGs/Stock%20Header.jpg" width="400" title="Multiple Year Stock Analysis">
<h1 align="center">
<b>Multiple Year Stock Analysis</b>
</h1>
</p>
<p align="center">
<sup><i> VBA challenge - UofT Data Analytics BootCamp</i></sup>
</P>


<details open><summary>Table of Contents</summary>

1. [Overview of Project](https://github.com/theidari/VBA-challenge#1-overview-of-project)  
   1. [Objective](https://github.com/theidari/VBA-challenge#i-objective)
   2. [Methods and Software](https://github.com/theidari/VBA-challenge#ii-methods-and-software)
2. [Code](https://github.com/theidari/VBA-challenge#2-code)
3. [Result](https://github.com/theidari/VBA-challenge#3-results)
4. [Explore The Docs](https://github.com/theidari/VBA-challenge#4-Explore-The-Docs)
5. [References](https://github.com/theidari/VBA-challenge#5-References)
</details>

## 1. Overview of Project
  This project used Visual Basic for Applications (VBA) programming language to analyze generated more than <b>750K</b> stock market data.
The results of the analyses provide insights into about <b>3K</b> unique stock's ➊ Ticker Symbol, ➋ Yearly Change,➌ Percent Change, and ➍ Total Stock Volume. The analyses will also provide information on the <b>"Greatest % increase", "Greatest % decrease", and "Greatest total volume"</b>.

### i. Objective
Create a script that loops through all the stocks for one year and outputs the following information:
| ${\color{red}Main \space \color{red}Part}$ | ${\color{red}Bonus}$ |
| ------------- | ------------- |
| The ticker symbol  | Greatest % increase  |
| Yearly change  | Greatest % decrease  |
| The percentage change  | Greatest total volume  |
| The total stock volume  |  |



### ii. Methods and Software
The analyses were performed using the [Multiple Year Stock Data](https://github.com/theidari/VBA-challenge/blob/main/Multiple_year_stock_data.xlsm) dataset.
Following Software were used in this project:

<img src="https://user-images.githubusercontent.com/17062794/200467306-1b06a964-0384-4a87-a0e5-6d4ba32fc9de.png" width="150" title="VBA"><img src="https://user-images.githubusercontent.com/17062794/200467777-f2df83a5-5964-4de3-9389-bcd7190cdde3.png" width="200" title="Excel"><img src="https://user-images.githubusercontent.com/17062794/200468097-278f79e9-9eb8-44e9-a31c-58b92e7efcca.png" width="150" title="VS">




<p align="center">
  
[<sup>⬆ BACK TO TOP ⬆</sup>](#multiple-year-stock-analysis)
<a name="multiple-year-stock-analysis"></a> 
  
</p>

## 2. Code
<details><summary>VBA Code (Click Me)</summary>


    For Each ws In Worksheets
        
        Dim Ticker_name As String
        Dim GrInTicker As String
        Dim GrDeTicker As String
        Dim TotVoTicker As String
        Dim Ticker_Summary As Integer
        Dim TotalStockVolume As Double
        Dim Openvalue As Double
        Dim Closevalue As Double
        Dim YearlyChange As Double
        Dim PrecentChange As Double
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestTotalVolume As Double
        Dim PercentChangeRange As Range
        Dim YearlyChangeRange As Range
        Set YearlyChangeRange = ws.Range("J:J")
        
 
        Ticker_Summary = 2
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        TotalStockVolume = ws.Cells(2, 7).Value
        Openvalue = ws.Cells(2, 3).Value
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestTotalVolume = 0

        ' CORRECT CELLS FORMAT
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("J:J").NumberFormat = "0.00"
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker_name = ws.Cells(i, 1).Value
                YearlyChange = Closevalue - Openvalue
                PrecentChange = YearlyChange / Openvalue
                ws.Range("I" & Ticker_Summary).Value = Ticker_name
                ws.Range("J" & Ticker_Summary).Value = YearlyChange
                ws.Range("K" & Ticker_Summary).Value = PrecentChange

                    'Greatest % Increase & Greatest % Decrease
                    If PrecentChange > GreatestIncrease Then
                        GreatestIncrease = PrecentChange
                        GrInTicker = Ticker_name
                    ElseIf PrecentChange < GreatestDecrease Then
                        GreatestDecrease = PrecentChange
                        GrDeTicker = Ticker_name
                    End If
                    
                ws.Range("L" & Ticker_Summary).Value = TotalStockVolume

                    'Greatest Total Volume
                    If TotalStockVolume > GreatestTotalVolume Then
                        GreatestTotalVolume = TotalStockVolume
                        TotVoTicker = Ticker_name
                    End If
                
                TotalStockVolume = ws.Cells(i + 1, 7).Value
                Ticker_Summary = Ticker_Summary + 1
                Openvalue = ws.Cells(i + 1, 3).Value

            ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
                TotalStockVolume = TotalStockVolume + ws.Cells(i + 1, 7).Value
                Closevalue = ws.Cells(i + 1, 6).Value
            End If

        Next i
        'Yearly Change Column Color (Column J)
        For Each Cell In YearlyChangeRange
            If Cell.Value > 0 Then
                Cell.Interior.ColorIndex = 4
            ElseIf Cell.Value < 0 Then
                Cell.Interior.ColorIndex = 3
            Else
                Cell.Interior.ColorIndex = xlNone
            End If
        Next
        ws.Cells(1, 10).Interior.ColorIndex = xlNone

        ws.Cells(2, 16).Value = GrInTicker
        ws.Cells(3, 16).Value = GrDeTicker
        ws.Cells(4, 16).Value = TotVoTicker
        ws.Cells(2, 17).Value = GreatestIncrease
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).Value = GreatestDecrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 17).Value = GreatestTotalVolume
        ws.Cells(4, 17).NumberFormat = "0.00E+00"
        'Columns width Autofit
        ws.Columns("A:R").AutoFit
    Next
    
</xyz>

</details>


[<sup>⬆ BACK TO TOP ⬆</sup>](#multiple-year-stock-analysis)
<a name="multiple-year-stock-analysis"></a>

## 3. Results
<b>Stock Performance Comparison:</b>
<br>
<b>2018:</b> <b>THB</b> had the greatest yearly percent change <i>increase</i> of all stocks at 141.42%. and <b>RKS</b> had the greatest yearly percent change <i>decrease</i> with falling down about 90%.
in addition, <b>QKN</b> had highest total volume in 3K unique stock trickers with a amount of 1.69x10<sup>+12</sup>. 
<br></br>

> <sub>2018 Calculation and Result</sub>
<p align="center">
<img src="https://github.com/theidari/VBA-challenge/blob/main/Result%20and%20File%20IMGs/2018-Result.png" width="800" title="Multiple Year Stock Analysis">
</p>
<b>2019:</b> <b>RYU</b> had the greatest yearly percent change <i>increase</i> of all stocks at 190.03%. and <b>RKS</b> had the greatest yearly percent change <i>decrease</i> with falling down about 91.6%.
in addition, <b>ZQD</b> had highest total volume in 3K unique stock trickers with a amount of 4.37x10<sup>+12</sup>.
<br></br>

> <sub>2019 Calculation and Result</sub>
<p align="center">
<img src="https://github.com/theidari/VBA-challenge/blob/main/Result%20and%20File%20IMGs/2019-Result.png" width="800" title="Multiple Year Stock Analysis">
</p>
<b>2020:</b> <b>YDI</b> had the greatest yearly percent change <i>increase</i> of all stocks at 188.76%. and <b>VNG</b> had the greatest yearly percent change <i>decrease</i> with falling down about 89%.
in addition, <b>QKN</b> had highest total volume in 3K unique stock trickers with a amount of 3.45x10<sup>+12</sup>.
<br></br>

> <sub>2020 Calculation and Result</sub>
<p align="center">
<img src="https://github.com/theidari/VBA-challenge/blob/main/Result%20and%20File%20IMGs/2020-Result.png" width="800" title="Multiple Year Stock Analysis">
</p>

[<sup>⬆ BACK TO TOP ⬆</sup>](#multiple-year-stock-analysis)
<a name="multiple-year-stock-analysis"></a> 
## 4. Explore The Docs
[Multiple Year Stock Data](https://github.com/theidari/VBA-challenge/blob/main/Multiple_year_stock_data.xlsm)

[Results Images](https://github.com/theidari/VBA-challenge/tree/main/Result%20and%20File%20IMGs)
## 5. References
Dataset created by Trilogy Education Services, a 2U, Inc. brand.

[<sup>⬆ BACK TO TOP ⬆</sup>](#multiple-year-stock-analysis)
<a name="multiple-year-stock-analysis"></a> 
