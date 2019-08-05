Attribute VB_Name = "Module1"

'Create a script that will loop through all the stocks for one year for each run and take the following information.
' The ticker symbol.
' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
' The total stock volume of the stock.
'You should also have conditional formatting that will highlight positive change in green and negative change in red.

Sub Wall_Street_Stock_Data_Moderate()
    
    Dim Ticker As String
    Dim Year As Long
    Dim Total_Volume As LongLong
    Dim LastRow As Long
    Dim ws As Worksheet
    Dim starting_ws As Worksheet
    Dim Yearly_change As Single
    Dim percent_change As Single
    Dim Summary_Table_Row As Integer
    Dim LastRowTicker As Integer
    Dim ColMPer As String
    Dim rg As Range
    Dim cond1 As FormatCondition, cond2 As FormatCondition
   
    'remember which worksheet is active in the beginning
    Set starting_ws = ActiveSheet
    
    'Looping through the worksheet
    For Each ws In ThisWorkbook.Worksheets
    ws.Activate

    'Initializing the veriable
    Total_Volume = 0
    Summary_Table_Row = 2
    
    'Setting the Headders
    Range("K" & 1).Value = "Ticker"
    Range("L" & 1).Value = "Yearly change"
    Range("M" & 1).Value = "percent change"
    Range("N" & 1).Value = "Total_Volume"
   
    'To Count a Last Row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    On Error Resume Next
    
    'Looping the Rows
    For I = 2 To LastRow
    
    'Assigning the Valus of the Rows
    Ticker = Cells(I, 1).Value
    Year = Cells(I, 2).Value
   
    'Comparing the Tickers in each Row
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
    
    'Finding a Year Closing Values for the percitular Ticker
    Ticker = Cells(I, 1).Value
    Yearly_change_close = Cells(I, 6)
    
    'Formulas for finding Percentage Change
    Yearly_change = Yearly_change_close - Yearly_change_open
    percent_change = Yearly_change / Yearly_change_open
    Total_Volume = Total_Volume + Cells(I, 7).Value
    ColMPer = Format(percent_change, "Percent")
    
    'Arrangeing the Summary Table row
    Range("K" & Summary_Table_Row).Value = Ticker
    Range("L" & Summary_Table_Row).Value = Yearly_change
    Range("M" & Summary_Table_Row).Value = ColMPer
    Range("N" & Summary_Table_Row).Value = Total_Volume
     
    'Adding to the Next Row
    Summary_Table_Row = Summary_Table_Row + 1
      
    'Resetting the Total Volume
    Total_Volume = 0
    Else
      
    'Adding the Total Volume
    Total_Volume = Total_Volume + Cells(I, 7).Value
    
        'Getting the opening value for the Ticker
        If Year = Range("B2").Value Then
           Yearly_change_open = Cells(I, 3)
            
        End If
    
    End If
    
    Next I
    
    'Getting the Range of the Summary Table
    Set rg = Range("L2", Range("L2").End(xlDown))
    
    'clear any existing conditional formatting
    rg.FormatConditions.Delete
    
    'Creating the Conditions fot the Format
    Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, 0)
    Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, 0)
    
    'define the format applied for each conditional format
        With cond1
            .Interior.Color = vbGreen
            .Font.Color = vbBlackl
        End With
 
        With cond2
            .Interior.Color = vbRed
            .Font.Color = vbBlack
        End With
    
    Next
    
    'activate the worksheet that was originally active
    starting_ws.Activate
    
End Sub


