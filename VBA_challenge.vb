Option Explicit
Sub Stocksummarize()

Application.ScreenUpdating = False
Dim ws As Worksheet
Dim rowread As Long
Dim row_limit As Long
Dim i As Long
Dim ticker_no As Integer

Dim ticker As String
Dim volume As LongLong
Dim price_openT As Single
Dim price_closeT As Single
Dim price_open As Single
Dim price_close As Single
Dim date_ As Long
Dim date_NY As Long
Dim date_YE As Long
Dim date_open As Long
Dim date_close As Long

Dim greatest_ins As Single
Dim greatest_dec As Single
Dim greatest_vol As Single
Dim giTicker As String
Dim gdTicker As String
Dim gvTicker As String

Dim booVerify As String
'"default" for assuming all tickers have price records for opening and closing date
'"Zero" for set first or last day opening price to zero
'"Avail" for using first or last available date for price
Dim answer As String
Dim msg As String

ticker_no = 0
booVerify = "default"

'Loop through each worksheet
For Each ws In Worksheets
    ws.Activate
        row_limit = Range("A1").End(xlDown).Row
        
        'headings
        Columns("I:P").Clear
        Range("P2:P3").NumberFormat = "0.00%"
        Range("I1") = "Ticker"
        Range("J1") = "Yearly Change"
        Range("K1") = "Percent Change"
        Range("L1") = "Total Stock Volume"
        Range("O1") = "Ticker"
        Range("P1") = "Value"
        Range("N2") = "Greatest % Increase"
        Range("N3") = "Greatest % Decrease"
        Range("N4") = "Greatest Total Volume"
        
        'reset
        price_open = 0
        price_close = 0
        volume = 0
        ticker_no = 0
        greatest_ins = 0
        greatest_dec = 0
        greatest_vol = 0
        date_NY = WorksheetFunction.Max(Columns("B:B"))
        date_YE = WorksheetFunction.Min(Columns("B:B"))
        date_open = date_NY
        date_close = date_YE
        Cells(ticker_no + 2, "I") = Range("A2")
        
        For rowread = 2 To row_limit + 1
            ticker = Cells(rowread, "A")
            date_ = Cells(rowread, "B")
            price_openT = Cells(rowread, "C")
            price_closeT = Cells(rowread, "F")
            

            'check if ticker is new
            If Cells(ticker_no + 2, "I") = ticker Then
                If date_ < date_open Then
                    date_open = date_
                    price_open = price_openT
                End If
                If date_ > date_close Then
                    date_close = date_
                    price_close = price_closeT
                End If
                volume = volume + Cells(rowread, "G")
                
            Else
                'verify 1st and last day of year match with data
                    If date_open <> date_YE Or date_close <> date_NY Then
                        Select Case booVerify
                            Case "default"
                                msg = "Ticker " & ticker & " does not have price available on the begining/end of the year" & vbNewLine & vbNewLine
                                msg = msg & "Would you like to using earliest/latest available price for begining/end of year price ?" & vbNewLine
                                msg = msg & "(choose ""no"" for using 0 as opening price or closing price)"
                                answer = MsgBox(msg, vbYesNo, "Warning")
                                If answer = vbYes Then
                                    booVerify = "Avail"
                                Else
                                    booVerify = "Zero"
                                    If date_open <> date_YE Then price_open = 0
                                    If date_close <> date_NY Then price_close = 0
                                End If
                                
                            Case "Zero"
                                If date_open <> date_YE Then price_open = 0
                                If date_close <> date_NY Then price_close = 0
                                    
                        End Select
                    End If
                             
                
                'write ticker to summary
                    Cells(ticker_no + 2, "J") = price_close - price_open
                        If price_close >= price_open Then
                            Cells(ticker_no + 2, "J").Interior.ColorIndex = 4
                        Else
                            Cells(ticker_no + 2, "J").Interior.ColorIndex = 3
                        End If
                        
                    If price_open = 0 Then
                        Cells(ticker_no + 2, "K") = 0
                    Else
                        Cells(ticker_no + 2, "K") = (price_close - price_open) / price_open
                    End If
                    
                    Cells(ticker_no + 2, "L") = volume
                                    
                    
                'update greatest ticker
                    If greatest_ins < Cells(ticker_no + 2, "K") Then
                        greatest_ins = (price_close - price_open) / price_open
                        giTicker = Cells(ticker_no + 2, "I")
                    End If
                    
                    If greatest_dec > Cells(ticker_no + 2, "K") Then
                        greatest_dec = (price_close - price_open) / price_open
                        gdTicker = Cells(ticker_no + 2, "I")
                    End If
                    
                    If greatest_vol < volume Then
                        greatest_vol = volume
                        gvTicker = Cells(ticker_no + 2, "I")
                    End If
                    
                'reset tickers variables
                    date_open = date_NY
                    date_close = date_YE
                    price_open = 0
                    price_close = 0
                    volume = 0
                    ticker_no = ticker_no + 1
                    Cells(ticker_no + 2, "I") = ticker
                    
                'process new ticket
                    If date_ < date_open Then
                        date_open = date_
                        price_open = price_openT
                    End If
                    If date_ > date_close Then
                        date_close = date_
                        price_close = price_closeT
                    End If
                    volume = volume + Cells(rowread, "G")
                    
            End If
                
        Next rowread
        
        'Write greatest
         Range("o2") = giTicker
         Range("o3") = gdTicker
         Range("o4") = gvTicker
         
         Range("p2") = greatest_ins
         Range("p3") = greatest_dec
         Range("p4") = greatest_vol
         
        'set format
         Columns("J:J").NumberFormat = "#,##0.00"
         Columns("K:K").NumberFormat = "0.00%"
         Columns("A:P").EntireColumn.AutoFit
         Rows("1:1").Font.Bold = True
Next ws
       
Application.ScreenUpdating = True

End Sub



