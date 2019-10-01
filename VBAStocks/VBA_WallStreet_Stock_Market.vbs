Sub WallStreet()


Dim Ws As Worksheet

Dim tiker As String
Dim lastRow As Long
Dim j, k, n, c, max, mim As Integer
Dim sum, tsum As Long
Dim YChange, tmax As Double
Dim PChange As Double
Dim PercentLastRow As Long
Dim TotalLastRow As Long
Dim OpenP As Double
Dim CloseP As Double

' ----Loop through each Worksheet

For Each Ws In Worksheets

    '-----Insert the header for Unique Tiker,Yearly Change,Percent Change and Total stock volume
    Ws.Cells(1, 10).Value = "Tiker"
    Ws.Cells(1, 11).Value = "Yearly Change"
    Ws.Cells(1, 12).Value = "Percent Change"
    Ws.Cells(1, 13).Value = "Total Stock volume"
   
    '-----Fetch lastRow of the records
    lastRow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    '-----Initialising all increment veriables
    tiker = Ws.Cells(2, 1).Value
    Ws.Cells(2, 10).Value = tiker
    O = 2
    t = 2
    c = 0
    sum = 0
    Total = 0

    '-----Loop through each row in active sheet
    '-----Calculate Total Stock volume,Yearly change,Tiker and Percent change
    For I = 2 To lastRow
        If tiker = Ws.Cells(I, 1).Value Then
            
            sum = Ws.Cells(I, 7).Value
            Total = Total + sum
            c = c + 1
        Else
           
            t = t + 1
            tiker = Ws.Cells(I, 1).Value
            Ws.Cells(t, 10).Value = tiker
            Ws.Cells(t - 1, 13).Value = Total
        
            OpenP = Ws.Cells(O, 3).Value
            CloseP = Ws.Cells(I - 1, 6).Value
            YChange = CloseP - OpenP
            Ws.Cells(t - 1, 11).Value = YChange
            '------Conditional formatting if value is -ve then cell color is red and value is +ve then cell color is green
            If YChange < 0 Then
                Ws.Cells(t - 1, 11).Interior.Color = RGB(255, 0, 0)
            Else
                Ws.Cells(t - 1, 11).Interior.Color = RGB(0, 255, 0)
            End If
       
             If OpenP = 0 Then
                PChange = 0
             Else
                PChange = YChange / OpenP
             End If
             Ws.Cells(t - 1, 12).Value = PChange
             Ws.Cells(t - 1, 12).Style = "percent"
      
             O = O + c
             sum = 0
             Total = 0
             sum = Ws.Cells(I, 7).Value
             Total = Total + sum
             c = 1
        End If
    Next I
    '----Insert and Clculate last value
      t = t + 1
      OpenP = Ws.Cells(O, 3).Value
      CloseP = Ws.Cells(I - 1, 6).Value
      YChange = CloseP - OpenP
      Ws.Cells(t - 1, 11).Value = YChange
      If YChange < 0 Then
            Ws.Cells(t - 1, 11).Interior.Color = RGB(255, 0, 0)
      Else
            Ws.Cells(t - 1, 11).Interior.Color = RGB(0, 255, 0)
      End If
       
     If OpenP = 0 Then
        PChange = 0
     Else
       PChange = YChange / OpenP
     End If
     Ws.Cells(t - 1, 12).Value = PChange
     Ws.Cells(t - 1, 12).Style = "percent"
      
     Total = Total + Cells(I, 7).Value
     Ws.Cells(t - 1, 13).Value = Total
     
     '------------Calculate Gretest% change, Lowest% change and Gretest Total Volume with tiker
     Ws.Cells(2, 15).Value = "Gretest % Increase"
     Ws.Cells(3, 15).Value = "Gretest % Decrease"
     Ws.Cells(4, 15).Value = "Gretest Total Volume"
     Ws.Cells(1, 16).Value = "Tiker"
     Ws.Cells(1, 17).Value = "Value"
     
     max = 0
     Min = 0
     
     PercentLastRow = Ws.Cells(Rows.Count, "L").End(xlUp).Row
     TotalLastRow = Ws.Cells(Rows.Count, "M").End(xlUp).Row
     For I = 2 To PercentLastRow
        If Ws.Cells(I, 12).Value > max Then
            max = Ws.Cells(I, 12).Value
            Ws.Cells(2, 16).Value = Ws.Cells(I, 10).Value
        ElseIf Ws.Cells(I, 12).Value < Min Then
            Min = Ws.Cells(I, 12).Value
            Ws.Cells(3, 16).Value = Ws.Cells(I, 10).Value
        End If
     Next I
     
     Ws.Cells(2, 17).Value = max
     Ws.Cells(3, 17).Value = Min
     tmax = 0
     
    For I = 2 To TotalLastRow
       If Ws.Cells(I, 13).Value > tmax Then
            tmax = Ws.Cells(I, 13).Value
          Ws.Cells(4, 16).Value = Ws.Cells(I, 10).Value
        End If
     Next I
     Ws.Cells(4, 17).Value = tmax
        
        
Next Ws

End Sub




