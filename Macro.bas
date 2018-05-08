Sub Fact()

S_Count = ActiveWorkbook.Worksheets.Count

PassCounter = 0
FailCounter = 0

'For First Sheet
For i = 1 To S_Count
       SheetName = ActiveWorkbook.Worksheets(i).Name
       
    For Row = 2 To Excel.Sheets(SheetName).UsedRange.Rows.Count
        If SheetName = "SearchManagement" Then
           col = 9
        Else
           col = 2
        End If
            Sysval = Trim(Excel.Sheets(SheetName).Cells(Row, col).Value)
            If Sysval = "Passed" Then
                PassCounter = PassCounter + 1
            Else
                FailCounter = FailCounter + 1
            End If
    Next
Next

    Excel.Sheets.Add.Name = "Charts"
    Excel.Sheets("Charts").Cells(1, 1).Value = "Total"
    Excel.Sheets("Charts").Cells(1, 2).Value = "Passed"
    Excel.Sheets("Charts").Cells(1, 3).Value = "Failed"
    Excel.Sheets("Charts").Cells(2, 1).Value = PassCounter + FailCounter
    Excel.Sheets("Charts").Cells(2, 2).Value = PassCounter
    Excel.Sheets("Charts").Cells(2, 3).Value = FailCounter
    
'create a bar chart in excel with this macro

Excel.ActiveSheet.Shapes.AddChart.Select
Excel.ActiveChart.SetSourceData Source:=Range("'Charts'!$B$1:$C$3")
Excel.ActiveChart.ChartType = xl3DPie



    Set rPatterns = Excel.ActiveSheet.Range("B1:C2")
    vPatterns = rPatterns.Value
    
    With Excel.ActiveChart.SeriesCollection(1)
    vValues = .Values
    
    For IPoint = 1 To UBound(vValues)
     
       ' If vPatterns(1, IPoint) = "Total" Then
        '    If vValues(IPoint) > 0 Then
         '      .Points(IPoint).Interior.Color = RGB(0, 0, 255)
               
          '  End If
        If vPatterns(1, IPoint) = "Passed" Then
            If vValues(IPoint) > 0 Then
               .Points(IPoint).Interior.Color = RGB(0, 150, 0)
               
            End If
        ElseIf vPatterns(1, IPoint) = "Failed" Then
            If vValues(IPoint) > 0 Then
               .Points(IPoint).Interior.Color = RGB(255, 0, 0)
               
            End If
        End If
     
    Next
  End With

End Sub



