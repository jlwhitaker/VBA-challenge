Attribute VB_Name = "Module1"
Sub Stocks()

'Set CurrentWS as the active worksheet
'Dim CurrentWs As Worksheet
Dim Ws As Worksheet


'Dim Need_Summary_Table_Header As Boolean
'Need_Summary_Table_Header = False

'Set variables for each CurrentWs
'For Each CurrentWs In ThisWorkbook.Worksheets
For Each Ws In ThisWorkbook.Worksheets

    'MsgBox CurrentWs.Range("A5").Value
    

    'Set an initial variable to hold the ticker name
    Dim Ticker_name As String
    Ticker_name = ""

    'Set an initial variable for holding the opening price
    Dim Opening_Price As Double
    Opening_Price = 0

    'Set an initial variable for holding the closing price
    Dim Closing_Price As Double
    Closing_Price = 0

    'Set an initial variable for holding the delta price
    Dim Delta_Price As Double
    Delta_Price = 0

    'Set an initial variable for holing the delta percentage
    Dim Delta_Percent As Double
    Delta_Percent = 0

    'Set an initial variable for holding the total stock volume
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0

    'Set location for output
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2

    'Locate the last row of each CurrentWs
    Dim Last_Row As Long
    Dim i As Long
    'Last_Row = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row
    Last_Row = Ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'Set titles for the summary table
    'CurrentWs.Range("I1").Value = "Ticker"
    'CurrentWs.Range("J1").Value = "Yearly Change"
    'CurrentWs.Range("K1").Value = "Percent Change"
    'CurrentWs.Range("L1").Value = "Total Stock Volume"

    Ws.Range("I1").Value = "Ticker"
    Ws.Range("J1").Value = "Yearly Change"
    Ws.Range("K1").Value = "Percent Change"
    Ws.Range("L1").Value = "Total Stock Volume"
          

    'Set value for first ticker
    'Opening_Price = CurrentWs.Cells(2, 3).Value
    Opening_Price = Ws.Cells(2, 3).Value

    
    'Create loop to locate Ticker_Name
    For i = 2 To Last_Row
    'For i = 2 To 5

        'Check to see if we are still in the same Ticker
        'If CurrentWs.Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
            'MsgBox CurrentWs.Cells(i, 1).Value
        
        
            'Set Ticker Name
            'Ticker_name = CurrentWs.Cells(i, 1).Value
            Ticker_name = Ws.Cells(i, 1).Value

            'Print the Ticker Name
            'CurrentWs.Range("I" & Summary_Table_Row).Value = Ticker_name
            
            'MsgBox
            'MsgBox Ticker_name
                        
            
            'Calculate Yearly Change
            'Closing_Price = CurrentWs.Cells(i, 6).Value
            'Closing_Price = CurrentWs.Cells(2, 6).Value
            Closing_Price = Ws.Cells(i, 6).Value
            Delta_Price = Closing_Price - Opening_Price

            'Calculate Percent Change
            If Opening_Price <> 0 Then

                Delta_Percent = (Delta_Price / Opening_Price)
    
                Else
                Delta_Percent = 0
            
            End If
        
        
            'Add to the Stock Volume
            'Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        
                    
        
            'Print the Delta Price
            'CurrentWs.Range("J" & Summary_Table_Row).Value = Delta_Price
            Ws.Range("J" & Summary_Table_Row).Value = Delta_Price
            
        'Else
        'Add to the Stock Volume
        'Total_Stock_Volume = Total_Stock_Volume + CurrentWs.Cells(i, 7).Value
        'Total_Stock_Volume = Total_Stock_Volume + Ws.Cells(i, 7).Value


        'End If
        
        
            'Print the Delta Percent
            'CurrentWs.Range("K" & Summary_Table_Row).Value = Delta_Percent
            Ws.Range("K" & Summary_Table_Row).Value = Delta_Percent
            
                If (Delta_Percent > 0) Then
                    'Fill column with GREEN color - good
                    Ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (Delta_Percent <= 0) Then
                    'Fill column with RED color - bad
                    Ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                End If

   
            'Print the Stock Volume
            
            
            'Print the Ticker Name
            'CurrentWs.Range("I" & Summary_Table_Row).Value = Ticker_name
            Ws.Range("I" & Summary_Table_Row).Value = Ticker_name


        
    'Add another line to summary table to reset before moving to the next ticker.
    'Summary_Table_Row = Summary_Table_Row + 1

    'Reset ticker
    Delta_Price = 0
    Closing_Price = 0
    'Opening_Price = CurrentWs.Cells(i + 1, 3).Value
    Opening_Price = Ws.Cells(i + 1, 3).Value
    
        Else
        'Add to the Stock Volume
        'Total_Stock_Volume = Total_Stock_Volume + CurrentWs.Cells(i, 7).Value
        Total_Stock_Volume = Total_Stock_Volume + Ws.Cells(i, 7).Value

        'CurrentWs.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        Ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        
        'Add another line to summary table to reset before moving to the next ticker.
        Summary_Table_Row = Summary_Table_Row + 1


        End If
    
    Next i

'Next CurrentWs
Next Ws



End Sub
