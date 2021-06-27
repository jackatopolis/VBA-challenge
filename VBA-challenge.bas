Attribute VB_Name = "Module2"
Public CalcState As Long
Public EventState As Boolean
Public PageBreakState As Boolean

' Subs to make code run faster
' https://www.thespreadsheetguru.com/blog/2015/2/25/best-way-to-improve-vba-macro-performance-and-prevent-slow-code-execution

Sub OptimizeCode_Begin()

Application.ScreenUpdating = False

EventState = Application.EnableEvents
Application.EnableEvents = False

CalcState = Application.Calculation
Application.Calculation = xlCalculationManual

PageBreakState = ActiveSheet.DisplayPageBreaks
ActiveSheet.DisplayPageBreaks = False

End Sub

Sub OptimizeCode_End()

ActiveSheet.DisplayPageBreaks = PageBreakState
Application.Calculation = CalcState
Application.EnableEvents = EventState
Application.ScreenUpdating = True

End Sub

Sub stock1()

    Call OptimizeCode_Begin
    
    ' Initialize Variables
    Dim sht As Worksheet
    Dim i As Long
    Dim j As Long
    Dim new_open_price As Double
    Dim running_vol As Double
    Dim lastrow As Double
    Dim MyRange As Range
    Dim lrow As Double
    Dim volume As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim current_ticker As String
    Dim previous_ticker As String
    
    ' Loop through all sheets in workbook
    For Each sht In Worksheets
    
        ' Establish current worksheet, find last row, add headers for new table
        lastrow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row
        sht.Range("I1").Value2 = "Ticker"
        sht.Range("J1").Value2 = "Yearly Change"
        sht.Range("K1").Value2 = "Percent Change"
        sht.Range("L1").Value2 = "Total Volume"
        current_close_price = 0
        current_open_price = 0
        
        ' Iterate over all rows in active sheet
        For i = 2 To lastrow
            lrow = sht.Cells(sht.Rows.Count, "I").End(xlUp).Row ' last row of new table
            
            current_ticker = sht.Cells(i, 1).Value2 ' save current ticker
            previous_ticker = sht.Cells(i - 1, 1).Value2 ' save previous ticker
            open_price = sht.Cells(i, 3).Value2 ' save current open price
            close_price = sht.Cells(i, 6).Value2 ' save current close price
            volume = sht.Cells(i, 7).Value2 ' save current volume
            
            If current_ticker <> previous_ticker Then ' if found new ticker
                
                ' Calculate and save values before working with new ticker
                sht.Cells(lrow + 1, 9).Value2 = current_ticker ' add new ticker
                sht.Cells(lrow, 12).Value2 = running_vol ' populate total volume cell with running volume before resetting
                yearly_change = current_close_price - current_open_price ' calculate yearly price change
                sht.Cells(lrow, 10).Value2 = yearly_change ' save yearly change before new ticker
                
                If i <> 2 Then ' if not the first iteration
                    If current_open_price = 0 Or yearly_change = 0 Then ' fixes divide by zero if all data is zero
                        percent_change = 0
                    Else
                        percent_change = (yearly_change / current_open_price) ' calculate yearly percent change
                    End If
                    sht.Cells(lrow, 11).Value2 = percent_change ' save yearly percent change before new ticker
                End If
                
                ' Reset variables for new ticker
                running_vol = volume ' reset volume for new ticker
                current_open_price = open_price ' save open price for new ticker
        
            Else ' if current ticker same as previous ticker
                running_vol = running_vol + volume ' update running volume
                current_close_price = close_price
                
                ' Checking to see if ticker is last ticker, so the last row's cell are populated
                If IsEmpty(sht.Cells(i + 1, 1).Value2) Then
                    sht.Cells(lrow, 12).Value2 = running_vol
                    yearly_change = current_close_price - current_open_price ' calculate yearly price change for last ticker
                    sht.Cells(lrow, 10).Value2 = yearly_change ' save yearly change for last ticker
                    percent_change = yearly_change / current_open_price ' calculate yearly percent change
                    sht.Cells(lrow, 11).Value2 = percent_change ' save yearly percent change before new ticker
                End If
            End If
        Next i
        
        ' Override Changed Header Values
        sht.Range("I1").Value2 = "Ticker"
        sht.Range("J1").Value2 = "Yearly Change"
        sht.Range("K1").Value2 = "Percent Change"
        sht.Range("L1").Value2 = "Total Volume"
        
        ' Conditional Formatting
        Set MyRange = sht.Range("J2:J" & (lrow))
        MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=0"
        MyRange.FormatConditions(1).Interior.Color = vbGreen ' zero or positive change is green
        MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        MyRange.FormatConditions(2).Interior.Color = vbRed ' negative yearly change is red
        sht.Range("K2:K" & (lrow)).NumberFormat = "0.00%" ' percent change formatted to percent
        
        ' Find greatest percent increase, greatest percent decrease, greatest total volume; store in new table
        sht.Range("O1").Value2 = "Ticker"
        sht.Range("P1").Value2 = "Value"
        sht.Range("N2").Value2 = "Greatest % Increase"
        sht.Range("N3").Value2 = "Greatest % Decrease"
        sht.Range("N4").Value2 = "Greatest Total Volume"
        Dim MyRange1 As Range
        Dim MyRange2 As Range
        
        ' Find greatest percent increase and corresponding ticker
        Set MyRange1 = sht.Range("K2:K" & lrow)
        sht.Range("P2").Value2 = WorksheetFunction.Max(MyRange1)
        pos1 = WorksheetFunction.Match(g_increase, MyRange1, 0) + MyRange1.Row - 1
        sht.Range("O2").Value2 = sht.Cells(pos1, 9).Value
        
        ' Find greatest percent decrease and corresponding ticker
        sht.Range("P3").Value2 = WorksheetFunction.Min(MyRange1)
        pos2 = WorksheetFunction.Match(g_decrease, MyRange1, 0) + MyRange1.Row - 1
        sht.Range("O3").Value2 = sht.Cells(pos2, 9).Value
        
        ' Find greatest total volume and corresponding ticker
        Set MyRange2 = sht.Range("L2:L" & lrow)
        sht.Range("P4").Value2 = WorksheetFunction.Max(MyRange2)
        pos3 = WorksheetFunction.Match(g_volume, MyRange2, 0) + MyRange2.Row - 1
        sht.Range("O4").Value2 = sht.Cells(pos3, 9).Value2
        
        sht.Range("P2:P3").NumberFormat = "0.00%" ' percent formatting
    
    Next sht
    
    Call OptimizeCode_End

End Sub
