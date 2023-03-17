VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub MagicButton_Click()

MsgBox ("It's time for some magic!")

For Each ws In Worksheets

    'Step 0: declare and create variables and assign values accordingly
    Dim header_1, header_2, header_3, header_4, header_5 As String
    Dim non_header_1, non_header_2, non_header_3 As String
    Dim index As Integer
    Dim opening As Double
    Dim closing As Double
    Dim yearlychange As Double
    Dim tickervolume As Double
    
    WorksheetName = ws.Name

    header_1 = "Ticker"
    header_2 = "Yearly Change"
    header_3 = "Percent Change"
    header_4 = "Total Stock Volume"
    header_5 = "Ticker"
    header_6 = "Value"
    
    non_header_1 = "Greatest % Increase"
    non_header_2 = "Greatest % Decrease"
    non_header_3 = "Greatest Total Volume"


    totalrows = ws.Cells(Rows.Count, 1).End(xlUp).Row
    totaltickerrows = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    index = 2
    firstyear = 2
    tickervolume = 0

    'Step 2: Run Non Iteration commands
    
    ws.Range("I1").Value = header_1
    ws.Range("J1").Value = header_2
    ws.Range("K1").Value = header_3
    ws.Range("L1").Value = header_4
    
    ws.Cells(index, 9).Value = ws.Cells(2, 1).Value

    'Step 3: Run iterations and their respective computations

    For i = 2 To totalrows
        'increment only if the next one does not equal the current row being iterated upon
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        index = index + 1
        ws.Cells(index, 9).Value = ws.Cells(i + 1, 1)
        
        End If
        
    Next i
    
    index = 2
    'produce the yearly change, percentage change, by stock.

    For i = 2 To totalrows
    
    'loop through rows and pull the first opening and last year value if the ticker has changed
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            opening = ws.Cells(i, 3).Value
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            closing = ws.Cells(i, 6).Value
        End If
    
    'compute both the yearly and percentage change based on the outputs of the above conditions
        If closing > 0 And opening > 0 Then
            yearlychange = closing - opening
            percentchange = yearlychange / opening
            
            ws.Cells(index, 10).Value = yearlychange
            ws.Cells(index, 11).Value = FormatPercent(percentchange)
            closing = 0
            opening = 0
            index = index + 1
        End If
    Next i

    'loop through each yearly change value and assign a color accordingly
    For i = 2 To totalrows
        If ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 4
        End If
    Next i
    
    index = 2
    'produce the total stock volume, by stock. Only apply if the ticker symbol hasn't changed, else reset.
    
    For i = 2 To totalrows
        presentrow = ws.Cells(i, 1).Value
        previousrow = ws.Cells(i - 1, 1).Value
   
   'compute the total stock volume based on the outputs of the above conditions
        If presentrow <> previousrow And i > 2 Then
            ws.Cells(index, 12).Value = tickervolume
            index = index + 1
            tickervolume = 0
        Else
            tickervolume = tickervolume + ws.Cells(i, 7).Value
        End If
    
    Next i
          
ws.Range("T1").Value = header_5
ws.Range("U1").Value = header_6
ws.Range("S2").Value = non_header_1
ws.Range("S3").Value = non_header_2
ws.Range("S4").Value = non_header_3

'Set baseline values for greatest increase, decrease and total volume

percentageincrease = 0
percentagedecrease = 0
greatestvolume = 0

For i = 2 To totalrows

    If ws.Cells(i, 11).Value > percentageincrease Then
        percentageincrease = ws.Cells(i, 11)
        percentincreaseticker = ws.Cells(i, 9)
    ElseIf ws.Cells(i, 11).Value < percentagedecrease Then
        percentagedecrease = ws.Cells(i, 11)
        percentdecreaseticker = ws.Cells(i, 9)
    ElseIf greatestvolume < ws.Cells(i, 12).Value Then
        greatestvolume = ws.Cells(i, 12)
        greatestvolumeticker = ws.Cells(i, 9)
    End If
    
Next i

    ws.Range("T2").Value = percentincreaseticker
    ws.Range("T3").Value = percentdecreaseticker
    ws.Range("T4").Value = greatestvolumeticker
    ws.Range("U2").Value = FormatPercent(percentageincrease)
    ws.Range("U3").Value = FormatPercent(percentagedecrease)
    ws.Range("U4").Value = greatestvolume
    
Next ws

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub nuclearbutton_Click()

'clear all contents in the specified range for all worksheets

Range("I:U").Select
Selection.ClearContents
Selection.ClearFormats

End Sub
