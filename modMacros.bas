Attribute VB_Name = "modMacros"
'note: requires reference to Microsoft Scripting Runtime
'summary variables
    Dim greatest_increase_percent As Double
    Dim greatest_increase_ticker As String
    Dim greatest_decrease_percent As Double
    Dim greatest_decrease_ticker As String
    Dim greatest_total_volume As LongLong
    Dim greatest_total_volume_ticker As String
    
Sub multiple_sheets()
    For Each ws In ActiveWorkbook.Sheets
        ws.Activate
        one_sheet
    Next ws
End Sub

Sub one_sheet()
    greatest_increase_percent = greatest_decrease_percent = 0
    greatest_total_volume = 0
    write_headers
    process_data
    write_summary
    Columns("l").EntireColumn.AutoFit
End Sub

Private Sub process_data()
    Dim rng As Range
    Set rng = Cells(1).CurrentRegion.Columns(1)
    If Not rng.Cells(1).Formula = "<ticker>" Then
        MsgBox "Please select a tab with <ticker> in the upper left corner!", vbCritical, "Warning"
        Exit Sub
    End If
    
    Dim cctype As String
    Dim cctypeold As String
    Dim totalvolume As Double
    Dim rowcounter As Integer
    Dim openprice As Double
    Dim closeprice As Double
    Dim yearlychange As Double
    Dim percentchange As Double

    'cctypeold = ""
    rowcounter = 2
    For Each cell In rng.Cells
        cctype = cell.Formula
        If cell.Row = 1 Then  'this is the header row
            cctypeold = cell.Offset(1, 0).Formula
            openprice = cell.Offset(1, 3 - 1).Value
        ElseIf cctype = cctypeold Then 'this means we're in a group continuing
            closeprice = cell.Offset(0, 6 - 1).Value
            totalvolume = totalvolume + cell.Offset(0, 7 - 1).Value
        Else 'this means we're starting a new group
            'write out
            Cells(rowcounter, 9).Formula = cctypeold
            'yearly change
            yearly_change = Round(closeprice - openprice, 2)
            Cells(rowcounter, 10).Formula = yearly_change
            If yearly_change >= 0 Then
                Cells(rowcounter, 10).Interior.Color = RGB(0, 255, 0)
            Else
                Cells(rowcounter, 10).Interior.Color = RGB(255, 0, 0)
            End If
            'percent change
            percent_change = Round(yearly_change / openprice, 4)
            Cells(rowcounter, 11).Formula = percent_change
            Cells(rowcounter, 11).NumberFormat = "0.00%"
            Cells(rowcounter, 12).Formula = totalvolume
            update_summary_variables cctypeold, percent_change, totalvolume
            'start things new
            totalvolume = 0
            rowcounter = rowcounter + 1
            cctypeold = cctype
            openprice = cell.Offset(0, 3 - 1).Value
            closeprice = cell.Offset(0, 6 - 1).Value
            totalvolume = totalvolume + cell.Offset(0, 7 - 1).Value
            
        End If
    Next cell
    
    'write out the values of the last group
    Cells(rowcounter, 9).Formula = cctypeold
    'yearly_change
    yearly_change = Round(closeprice - openprice, 2)
    Cells(rowcounter, 10).Formula = yearly_change
    If yearly_change >= 0 Then
        Cells(rowcounter, 10).Interior.Color = RGB(0, 255, 0)
    Else
        Cells(rowcounter, 10).Interior.Color = RGB(255, 0, 0)
    End If
    'percent change
    percent_change = Round(yearly_change / openprice, 4)
    Cells(rowcounter, 11).Formula = percent_change
    Cells(rowcounter, 11).NumberFormat = "0.00%"
    Cells(rowcounter, 12).Formula = totalvolume
   
    'update summary
    update_summary_variables cctypeold, percent_change, totalvolume
End Sub

Private Sub write_headers()
    Cells(9).Formula = "Ticker"
    Cells(10).Formula = "Yearly Change"
    Cells(11).Formula = "Percent Change"
    Cells(12).Formula = "Total Stock Volume"
    Cells(16).Formula = "Ticker"
    Cells(17).Formula = "Value"
    Cells(2, 15).Formula = "Greatest % Increase"
    Cells(3, 15).Formula = "Greatest % Decrease"
    Cells(4, 15).Formula = "Greatest Total Volume"
End Sub

Private Sub write_summary()
    Cells(2, 16).Formula = greatest_increase_ticker
    Cells(2, 17).Formula = greatest_increase_percent
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 16).Formula = greatest_decrease_ticker
    Cells(3, 17).Formula = greatest_decrease_percent
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(4, 16).Formula = greatest_total_volume_ticker
    Cells(4, 17).Formula = greatest_total_volume
End Sub

Private Sub update_summary_variables(ByVal ticker As String, ByVal percent_change As Double, ByVal volume As Double)
    If percent_change >= 0 And percent_change > greatest_increase_percent Then
        greatest_increase_percent = percent_change
        greatest_increase_ticker = ticker
    End If
    
    If percent_change < 0 And percent_change < greatest_decrease_percent Then
        greatest_decrease_percent = percent_change
        greatest_decrease_ticker = ticker
    End If
    
    If volume > greatest_total_volume Then
        greatest_total_volume = volume
        greatest_total_volume_ticker = ticker
    End If
End Sub
