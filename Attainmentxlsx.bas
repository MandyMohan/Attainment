Attribute VB_Name = "Attainmentxlsx"
Sub Pivot_Loopx()
  Dim sheetsArray As Sheets
  Dim xArray As Sheets
  Dim xWs As Worksheet
  Dim last As Long
  Dim pt As PivotTable
  Dim pi As PivotItem
  Dim pf As PivotField
  Dim pi2 As PivotItem
  Dim j As Integer
  Set pt = Sheets("Graph").PivotTables("PivotTable1")
  Set pf = pt.PivotFields("School Code")
  
  'make an array of pages to be filtered
  
  Set sheetsArray = ActiveWorkbook.Sheets(Array("ATTAIN (atleast 1)", "Performance Report 2013", "Performance Report 2014", "Performance Report 2015", "Performance Report 2016", "Performance Report 2017", "Performance Report 2018", "Performance Report 2019", "Performance Report 2020", "Performance Report 2021", "Performance Report 2022"))
  
  'make an array of pages to be copied and pasted once filtered
  
  Set xArray = ActiveWorkbook.Sheets(Array("Graph", "ATTAIN (atleast 1)", "Performance Report 2013", "Performance Report 2014", "Performance Report 2015", "Performance Report 2016", "Performance Report 2017", "Performance Report 2018", "Performance Report 2019", "Performance Report 2020", "Performance Report 2021", "Performance Report 2022"))
  
  'go through pivot filter and filter array accordingly
  
  For Each pi In pf.PivotItems
     pf.PivotItems(pf.PivotItems.Count - 1).Visible = True
  For Each pi2 In pf.PivotItems
      If pi2 = pi Then
        pi2.Visible = True
        For Each xWs In sheetsArray
              xWs.Range("B4").AutoFilter 2, pi2
              j = WorksheetFunction.Count(xWs.Range("B4:B5000").Cells.SpecialCells(xlCellTypeVisible))
              If j = 0 Then
                xWs.Visible = False
              End If
        Next
      Else: pi2.Visible = False
   End If
   Next
   
   'Create a chart trendline
   
   Sheets("Graph").ChartObjects("Chart 1").Chart.SeriesCollection(1).Trendlines.Add
   With Sheets("Graph").ChartObjects("Chart 1").Chart.SeriesCollection(1).Trendlines(1)
    .Type = xlLinear
    .DisplayEquation = True
    .Format.Line.DashStyle = msoLineSysDot
    .Format.Line.Weight = 2.5
    .DataLabel.Font.Size = 18
    .DataLabel.Font.Color = vbBlack
   End With
   
   'add new workbook
  
  Set newBook = Workbooks.Add(xlWBATWorksheet)
  
  'go through xArray and copy and paste each filtered sheet to new workbook

  With xArray(1)
    .Range("A1:F50").SpecialCells(xlCellTypeVisible).Copy
    newBook.Activate
    ActiveSheet.Name = xArray(1).Name
    ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
    ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteColumnWidths
    With ActiveSheet
        .Range("E1:F1,A5").Font.Size = 18
        .Range("A4").Font.Size = 24
        .Range("A8:A11").EntireRow.Delete
        .Range("A8:C21").Font.Size = 14
        .Range("A1").ColumnWidth = 18.86
        .Range("C1").ColumnWidth = 62
        .Range("D1").ColumnWidth = 11.86
        .Range("E1").ColumnWidth = 16.43
        .Range("F1").ColumnWidth = 27.86
    End With
    xArray(1).ChartObjects("Chart 1").Activate
    ActiveChart.ChartArea.Select
    ActiveChart.ChartArea.Copy
    newBook.Sheets("Graph").Activate
    newBook.Sheets("Graph").Range("A25").Select
    ActiveSheet.Paste
    With ActiveChart
        .ChartTitle.Font.Size = 18
        .Axes(xlValue).TickLabels.Font.Size = 12
        .Axes(xlCategory).TickLabels.Font.Size = 12
        .Parent.Left = Sheets("Graph").Range("A25").Left
        .Parent.Top = Sheets("Graph").Range("A25").Top
        If Sheets("Graph").Range("F1") = "Victoria" Or Sheets("Graph").Range("F1") = "Caroni" Or Sheets("Graph").Range("F1") = "Tobago" Then
            .Parent.Width = Sheets("Graph").Range("A25:F25").Width - 125
            .Shapes("Rectangle 1").Width = Sheets("Graph").Range("A25:F25").Width - 135
        Else
            .Parent.Width = Sheets("Graph").Range("A25:G25").Width + 15
            .Shapes("Rectangle 1").Width = Sheets("Graph").Range("A25:G25").Width - 10
        End If
        .Parent.Height = Sheets("Graph").Range("A25:A55").Height
        .Shapes("Rectangle 1").Top = 425
        .Shapes("Rectangle 1").Left = 5
        
    End With
            
 End With
 
 With xArray(2)
    .Range("A1:G3000").SpecialCells(xlCellTypeVisible).Copy
    newBook.Worksheets.Add(After:=Sheets(Sheets.Count)).Name = xArray(2).Name
    ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    last = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    With ActiveSheet
        .Range("A1").Font.Size = 24
        .Range("A4:G" & last).Borders.LineStyle = xlContinuous
        .Range("A4:G" & last).VerticalAlignment = xlVAlignCenter
        .Range("A4:G" & last).HorizontalAlignment = xlHAlignCenter
        .Range("C5:C" & last).HorizontalAlignment = xlHAlignLeft
        .Range("A4:G" & last).Font.Size = 14
        .Range("A1").ColumnWidth = 10.71
        .Range("B1").ColumnWidth = 18
        .Range("C1").ColumnWidth = 53.71
        .Range("D1:G1").ColumnWidth = 20
        .Range("A2:A3").RowHeight = 15
        .Range("A4:G4").Interior.Color = RGB(208, 206, 206)
    End With
End With

'add chart to second sheet in array
    
  Dim Ws As Worksheet
  Dim Rang As Range
  Dim MyChart As Object
  
  Set Ws = ActiveSheet
  Set Rang = Ws.Range("A4:A" & last & "," & "G4:G" & last)
  Set MyChart = Ws.Shapes.AddChart2
  
  With MyChart.Chart
        .SetSourceData Source:=Rang, PlotBy:=xlRows
        .ChartType = xlXYScatterSmooth
        .ChartTitle.Text = "% Attained Atleast 1 Subject"   'Title
        .ChartTitle.Font.Size = 18
        .ChartTitle.Font.Color = vbBlack
        .HasLegend = False
        .PlotBy = xlColumns
        .Axes(xlCategory).MinimumScale = 2013    'Adjust scale
        .Axes(xlCategory).MaximumScale = 2022
        .Axes(xlValue).TickLabels.NumberFormat = "0.0%"     'Remove decimals from scale
        .Axes(xlValue).HasMajorGridlines = True  'Remove Gridlines
        .Axes(xlCategory).HasMajorGridlines = True
        .Axes(xlValue).TickLabels.Font.Color = vbBlack
        .Axes(xlValue).TickLabels.Font.Size = 12
        .Axes(xlCategory).TickLabels.Font.Color = vbBlack
        .Axes(xlCategory).TickLabels.Font.Size = 12
        .SeriesCollection(1).Trendlines.Add
         With .SeriesCollection(1).Trendlines(1)
            .Type = xlLinear
            .DisplayEquation = True
            .Format.Line.DashStyle = msoLineSysDot
            .Format.Line.Weight = 2.5
            .DataLabel.Font.Size = 18
            .DataLabel.Font.Color = vbBlack
        End With
        With .Parent
           .Left = Ws.Range("A" & last + 2).Left
           .Top = Ws.Range("A" & last + 2).Top
           .Width = Ws.Range("A16:G16").Width
           .Height = Ws.Range("A" & last + 2, "A41").Height
        End With
  End With
 
 For i = 3 To 12
    xArray(i).Range("A1:O5000").SpecialCells(xlCellTypeVisible).Copy
    newBook.Worksheets.Add(After:=Sheets(Sheets.Count)).Name = xArray(i).Name
    ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteColumnWidths
    ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
    last = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    ActiveSheet.Range("B" & last).Value = ActiveSheet.Range("B" & last - 1).Value
    ActiveSheet.Range("C" & last).Value = ActiveSheet.Range("C" & last - 1).Value
    With ActiveSheet
        .Range("A1").Font.Size = 36
        .Range("A4:O" & last).Borders.LineStyle = xlContinuous
        .Range("A4:O" & last).Font.Size = 14
        .Range("A1").ColumnWidth = 6.57
        .Range("B1").ColumnWidth = 10.14
        .Range("C1").ColumnWidth = 33.14
        .Range("D1").ColumnWidth = 29.57
        .Range("E1").ColumnWidth = 12.29
        .Range("F1").ColumnWidth = 7.71
        .Range("K1").ColumnWidth = 9.14
        .Range("G1:J1,L1:O1").ColumnWidth = 6
        .Range("A2").RowHeight = 15
        .Range("A5:A" & last - 1).RowHeight = 60
    End With
Next
    
  
  'export sheets as an excel wbk and place in a folder based on district
  
   If Sheets("Graph").Range("F1") = "Victoria" Then
    newBook.SaveAs _
            Filename:="C:\Users\" & Environ("username") & "\Documents\Victoria\" & Sheets("Graph").Range("A4") & " Performance Report " & Sheets("Graph").Range("B12") & "-" & Sheets("Graph").Range("B21") & ".xlsx"
    newBook.Close SaveChanges:=False
   

   ElseIf Sheets("Graph").Range("F1") = "Caroni" Then
    newBook.SaveAs _
             Filename:="C:\Users\" & Environ("username") & "\Documents\Caroni\" & Sheets("Graph").Range("A4") & " Performance Report " & Sheets("Graph").Range("B12") & "-" & Sheets("Graph").Range("B21") & ".xlsx"
     newBook.Close SaveChanges:=False
   

    ElseIf Sheets("Graph").Range("F1") = "North Eastern" Then
        newBook.SaveAs _
            Filename:="C:\Users\" & Environ("username") & "\Documents\North Eastern\" & Sheets("Graph").Range("A4") & " Performance Report " & Sheets("Graph").Range("B12") & "-" & Sheets("Graph").Range("B21") & ".xlsx"
        newBook.Close SaveChanges:=False
       
    ElseIf Sheets("Graph").Range("F1") = "South Eastern" Then
    newBook.SaveAs _
            Filename:="C:\Users\" & Environ("username") & "\Documents\South Eastern\" & Sheets("Graph").Range("A4") & " Performance Report " & Sheets("Graph").Range("B12") & "-" & Sheets("Graph").Range("B21") & ".xlsx"
    newBook.Close SaveChanges:=False
   
   ElseIf Sheets("Graph").Range("F1") = "St George East" Then
    newBook.SaveAs _
            Filename:="C:\Users\" & Environ("username") & "\Documents\St. George East\" & Sheets("Graph").Range("A4") & " Performance Report " & Sheets("Graph").Range("B12") & "-" & Sheets("Graph").Range("B21") & ".xlsx"
    newBook.Close SaveChanges:=False

   ElseIf Sheets("Graph").Range("F1") = "Port Of Spain" Then
    newBook.SaveAs _
            Filename:="C:\Users\" & Environ("username") & "\Documents\Port of Spain\" & Sheets("Graph").Range("A4") & " Performance Report " & Sheets("Graph").Range("B12") & "-" & Sheets("Graph").Range("B21") & ".xlsx"
    newBook.Close SaveChanges:=False

   ElseIf Sheets("Graph").Range("F1") = "Tobago" Then
    newBook.SaveAs _
            Filename:="C:\Users\" & Environ("username") & "\Documents\Tobago\" & Sheets("Graph").Range("A4") & " Performance Report " & Sheets("Graph").Range("B12") & "-" & Sheets("Graph").Range("B21") & ".xlsx"
    newBook.Close SaveChanges:=False

   Else
   newBook.SaveAs _
            Filename:="C:\Users\" & Environ("username") & "\Documents\St. Patrick\" & Sheets("Graph").Range("A4") & " Performance Report " & Sheets("Graph").Range("B12") & "-" & Sheets("Graph").Range("B21") & ".xlsx"
    newBook.Close SaveChanges:=False
   
   End If
   
   'restore all sheets in the array to visible
   
    For Each xWs In sheetsArray
        xWs.Visible = True
    Next

    'set a next school in the pivot filter to visible and repeat
    
   Next
   
   'delete trendline
   
   Sheets("Graph").ChartObjects("Chart 1").Chart.SeriesCollection(1).Trendlines(1).Delete
End Sub

