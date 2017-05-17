Sub Future_Values()


ActiveSheet.Range("B4:M30").Clear
ActiveWindow.DisplayGridlines = False


Initial_Value = InputBox("What is the initial value?")
Cells(5, 2) = "Initial Value:"
Cells(5, 3) = Format(Initial_Value, "$0.00")



Interest_Rate = InputBox("What is the interest/compounding rate?")
Cells(6, 2) = "Interest Rate:"
Cells(6, 3) = Format(Interest_Rate, "%0.00")


Compound_Freq = InputBox("What is the max frequency of compounding?")

Cells(13, 3) = "Compounding Frequency" 'Label column heading
For i = 1 To Compound_Freq
    Cells(14, i + 2) = i
    Next i
    Range(Cells(13, 3), Cells(14, Compound_Freq + 2)).Select
    Selection.Font.Bold = True
    
    

Num_Periods = InputBox("How many periods for compounding would you like to include?")

Cells(14, 2) = "Number of Periods" 'Label column heading
For j = 1 To Num_Periods
    If j = 1 Then Cells(j + 14, 2) = j
    If j > 1 Then Cells(j + 14, 2) = j
    Next j
    Range(Cells(14, 2), Cells(Num_Periods + 14, 2)).Select
    Selection.Font.Bold = True



For h = 1 To Num_Periods
    For k = 1 To Compound_Freq
        Compound_Rate = ((1 + ((Interest_Rate / 100))) ^ k)
        Total = Initial_Value * Compound_Rate
        Per_Total = Total * h
        Cells(14 + h, k + 2) = Format(Per_Total, "$0.00")
        Next k
        Next h
        
Range("A:AA").EntireColumn.AutoFit
With ActiveSheet.PageSetup.Orientation = xlLandscape
    End With


End Sub
