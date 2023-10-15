# VBA-challenge
This Assignment created by: Hanen Zeddini on Friday, October 13, 2023. 
The Location of the source code: VBA_Scripts.rtf
***************** The source code Description **************
The source code is composed with the following two Subs: 
1- Analyse: Calculates the analysis in the entire worksheet.
2- Reset: Deletes the additional column's analysis.
***************** The source code **************************

Sub Analyse()
    Dim ws As Worksheet
    Dim cellA, PlageHeader As Range
    Dim lastRow, k As Long
    
    '*****************************Looping Across the entire Worksheet***************************
   For Each ws In ActiveWorkbook.Worksheets
        '**************************Column Creation************************************
        'Column's Hearder
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "YearlyChange"
        ws.Cells(1, 11) = "PercentChange"
        ws.Cells(1, 12) = "TotalStockVolume"
        ws.Cells(1, 15) = "Ticker"
        ws.Cells(1, 16) = "Value"
        ws.Cells(2, 14) = "Greatest % increase"
        ws.Cells(3, 14) = "Greatest % decrease"
        ws.Cells(4, 14) = "Greatest Total volume"
                            
        Set PlageHeader = ws.Range("A1:" & ws.Cells(1, Columns.Count).End(xlToLeft).Address)
        PlageHeader.Font.Bold = True ' Set Bold to the header's line
        
        ws.UsedRange.Columns.AutoFit ' Column's Automatic Adjustment
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ws.Range("K2:K" & lastRow).NumberFormat = "0.00%" ' Set % (2 decimals) Format to K column ( 11)
        

        i = 2
        Count = 252
        '**************************Retrieval of Data**********************************************
        ticker = ws.Cells(i, 1) 'Memorize the first appearance of ticker in a variable named ticker
        opens = ws.Cells(i, 3) 'Memorize the opens of the first day of the year
        ws.Cells(i, 9) = ticker 'Write the ticker in the analysis column
        j = i
        stocks = 0
        
        Top_perc_Increase = -100
        Top_perc_decrease = 100
        Top_Total_Volume = 0

        k = 2
        '**************************************Looping Across sheet*******************************
        Do While ((i < (i + Count)) And (k <= lastRow))
            If (ws.Cells(i, 1) = ticker) Then 'Continue to add up the stocks as long as it is the same ticker
                        closes = ws.Cells(i, 6)
                        stocks = stocks + ws.Cells(i, 7)
            Else
                        'MsgBox ("Ticker: " & ticker & "Opens: " & opens & " Close: " & closes)
                        ws.Cells(j, 9) = ticker 'Write the ticker in the analysis column
                        ws.Cells(j, 10) = (closes - opens) 'Yearly changes
                        If ws.Cells(j, 10).Value < 0 Then 'Conditionnal Formating of Yearly changes
                            ws.Cells(j, 10).Interior.Color = RGB(255, 0, 0) ' Red
                        Else
                            ws.Cells(j, 10).Interior.Color = RGB(0, 255, 0) ' Green
                        End If
                        ws.Cells(j, 11) = ((closes - opens) / opens) ' Percentage change
                        '********************Conditional Formatting************************************
                        If ws.Cells(j, 11).Value < 0 Then 'Conditionnal Formating of Percentage  change
                            ws.Cells(j, 11).Interior.Color = RGB(255, 0, 0) ' Red
                        Else
                            ws.Cells(j, 11).Interior.Color = RGB(0, 255, 0) ' Green
                        End If

                        ws.Cells(j, 12) = stocks
                        '******************Calculated Values********************************************
                        If Top_perc_Increase <= ws.Cells(j, 11) Then
                            Top_perc_Increase = ws.Cells(j, 11)
                            ws.Cells(2, 15) = ticker
                            ws.Cells(2, 16) = Top_perc_Increase
                            ws.Cells(2, 16).NumberFormat = "0.00%" ' Set % (2 decimals)
                        End If
                        If Top_perc_decrease >= ws.Cells(j, 11) Then
                            Top_perc_decrease = ws.Cells(j, 11)
                            ws.Cells(3, 15) = ticker
                            ws.Cells(3, 16) = Top_perc_decrease
                            ws.Cells(3, 16).NumberFormat = "0.00%" ' Set % (2 decimals)
                        End If
                        
                        If Top_Total_Volume <= stocks Then
                            Top_Total_Volume = stocks
                            ws.Cells(4, 15) = ticker
                            ws.Cells(4, 16) = Top_Total_Volume
                        End If
                        
                        '******************************************************
                        
                        stocks = 0
                        opens = ws.Cells(i, 3) 'Memorize the opens of the first day of the year
                        ticker = ws.Cells(i, 1) 'Memorize the first appearance of ticker in a variable named ticker
                        stocks = stocks + ws.Cells(i, 7)
                        j = j + 1
            End If
              i = i + 1
              k = k + 1
    Loop
            
    Next ws
    MsgBox ("All data analysis in all sheets are completed successfully")
End Sub

Sub Reset()
    '*****************Reset The Analysis*************************
    Dim ws As Worksheet
    
    answer = MsgBox("Are you sure you want to delete Analyse's columns ?", vbQuestion + vbYesNo + vbDefaultButton2, "Attention !")

   If answer = vbYes Then
    
            For Each ws In ActiveWorkbook.Worksheets
                ws.Columns("I:P").Delete
            Next ws
            MsgBox ("The deletion of column I to L was successfully done")
   End If
End Sub


