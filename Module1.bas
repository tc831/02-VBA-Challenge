Attribute VB_Name = "Module1"
Sub VBAchallenge():
    Dim MaxRow As Long
    Dim MinRow As Long
    Dim resulttable As Integer
    Dim tickers As String
    Dim totalvol As Double
    Dim Summary_Table_Row As Integer
    Dim openvaule As Double
    Dim closevalue As Double
    Dim yearlychange As Double
    Dim changepercentage As Double
    Dim greatdecrease As String
    Dim greatincrease As String

    ' Looping through worksheets
    For Each WS In Worksheets
    
        If WS.Name <> "Summary" Then
           WS.Activate
           totalvol = 0
           resulttable = 9
    
           'setting the format for each column
            Columns(resulttable).NumberFormat = "@"
            Columns(resulttable + 2).NumberFormat = "0.00%"
            Columns(resulttable + 3).NumberFormat = "#,##0 "
        
            'Assign Headers
            Range("I1").Value = "Ticker"
            Range("J1").Value = "Yearly Change"
            Range("K1").Value = "Percentage"
            Range("L1").Value = "Total Volume"
    
            'Defining the first row of data
            Summary_Table_Row = 2
            lastrow = Cells(Rows.Count, 1).End(xlUp).Row
 
            'Setting initial values
            closevalue = Cells(2, 6).Value
            openvalue = Cells(2, 3).Value
      
            For i = 2 To lastrow
            
                'Finding last day of the year for each ticker
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                    'Extracting values for each variable
                    tickers = Cells(i, 1).Value
                    totalvol = totalvol + Cells(i, 7).Value
                    yearlychange = closevalue - openvalue
   
                    'If function as to prevent calculation error for
                    If yearlychange <> 0 Then
    
                        changepercentage = yearlychange / openvalue
    
                    Else
    
                        'Resetting each value for tickers
                        changepercentage = 0
    
                    End If
    
                    'Assigning data into destinated cells
                    Cells(Summary_Table_Row, resulttable) = tickers
                    Cells(Summary_Table_Row, resulttable + 1) = yearlychange
                    Cells(Summary_Table_Row, resulttable + 2) = changepercentage
                    Cells(Summary_Table_Row, resulttable + 3) = totalvol
                    Summary_Table_Row = Summary_Table_Row + 1
          
                   'resetting ticker
                    totalvol = 0
                    closevalue = Cells(i + 1, 6).Value
                    openvalue = Cells(i + 1, 3).Value
            
                Else
    
                    'Extracting last day of the year
                    If Cells(i + 1, 2).Value > Cells(i, 2).Value Then
       
                        closevalue = Cells(i + 1, 6).Value
       
                    'Extracting first day of the year
                    ElseIf Cells(i + 1, 2).Value < Cells(i, 2).Value Then
       
                        openvalue = Cells(i + 1, 3).Value
                       
                    End If
                
                    'Set the value for open value if first day of the year open value is 0
                    If openvalue = 0 Then
                    
                        openvalue = Cells(i + 1, 3).Value
                    
                    End If
       
                    'Sum total value of each ticker
                    totalvol = totalvol + Cells(i, 7).Value
    
                End If
         
            Next i
            
            'Setting the format for each column
            Columns("K:K").Select
            Selection.Style = "Percent"
            Selection.NumberFormat = "0.00%"
            Columns("L:L").Select
            Selection.NumberFormat = "#,##0"
            Columns("I:L").Select
            Columns("I:L").EntireColumn.AutoFit
            Range("J2:J" & Summary_Table_Row).Select
                
            'Setting the conditional formatting for positives and negatives
            Selection.FormatConditions.Delete
            Set condition1 = Range("J2:J" & Summary_Table_Row).FormatConditions.Add(xlCellValue, xlGreater, "=0")
            Set condition2 = Range("J2:J" & Summary_Table_Row).FormatConditions.Add(xlCellValue, xlLess, "=0")
            With condition1

                 .Interior.ColorIndex = 4

            End With

            With condition2

                 .Interior.ColorIndex = 3

            End With
            
            'Assign values to cells
            Range("O2").Value = "Greatest % Increase"
            Range("O3").Value = "Greatest % Decrease"
            Range("O4").Value = "Greatest Total Volume"
            Range("P1").Value = Range("I1").Value
            Range("Q1").Value = "Value"
            
            'Extract Max Value of the column to cell
            Range("Q2").Value = Application.WorksheetFunction.Max(Range("K:K"))
            
            'Find the Row Number for Max value cell in column
            MaxRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(Range("K:K")), Range("K:K"), 0)
            
            'Find the Row Number of Min value cell in column
            MinRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(Range("K:K")), Range("K:K"), 0)
            Range("P2").Value = Cells(MaxRow, 9)
            Range("Q2").Style = "Percent"
            Range("Q2").NumberFormat = "0.00%"
            
            'Extracting Minimum Value of the column to cell
            Range("Q3").Value = Application.WorksheetFunction.Min(Range("K:K"))
            Range("P3").Value = Cells(MinRow, 9)
            Range("Q3").Style = "Percent"
            Range("Q3").NumberFormat = "0.00%"
            Range("Q4").Value = Application.WorksheetFunction.Max(Range("L:L"))
            
            'Re-assigne value to variable
            MaxRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(Range("L:L")), Range("L:L"), 0)
            Range("P4").Value = Cells(MaxRow, 9)
            Range("Q4").NumberFormat = "#,##0"
            Range("P1:Q1,O2:O4").Select
            Selection.Font.Bold = True
            Range("O1:Q4").Select
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                
                .Weight = xlThin
            
            End With
          
            
        End If
            
    Next WS
     
End Sub

