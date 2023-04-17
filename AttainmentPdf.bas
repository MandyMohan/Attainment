Attribute VB_Name = "AttainmentPdf"
Sub Pivot_Loop()
  Dim sheetsArray As Sheets
  Dim xArray As Sheets
  Dim xWs As Worksheet
  Dim pt As PivotTable
  Dim pi As PivotItem
  Dim pf As PivotField
  Dim pi2 As PivotItem
  Dim j As Integer
  Set pt = Sheets("Graph").PivotTables("PivotTable1")
  Set pf = pt.PivotFields("School Code")
  
  'make an array of pages to be filtered
  
  Set sheetsArray = ActiveWorkbook.Sheets(Array("ATTAIN (atleast 1)", "Performance Report 2013", "Performance Report 2014", "Performance Report 2015", "Performance Report 2016", "Performance Report 2017", "Performance Report 2018", "Performance Report 2019", "Performance Report 2020", "Performance Report 2021", "Performance Report 2022"))
  
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
    .Format.Line.Weight = 3
    .DataLabel.Font.Size = 24
    .DataLabel.Font.Color = vbBlack
   End With
   
   'Create a chart trendline
   
   Sheets("ATTAIN (atleast 1)").ChartObjects("Chart 2").Chart.SeriesCollection(1).Trendlines.Add
    With Sheets("ATTAIN (atleast 1)").ChartObjects("Chart 2").Chart.SeriesCollection(1).Trendlines(1)
        .Type = xlLinear
        .DisplayEquation = True
        .Format.Line.DashStyle = msoLineSysDot
        .Format.Line.Weight = 3
        .DataLabel.Font.Size = 24
        .DataLabel.Font.Color = vbBlack
    End With
   
   'Format sheets to be exported
   
   For Each xWs In ActiveWorkbook.Worksheets
    With ActiveSheet.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
      End With
  Next
  
  'export to pdf and place in a folder based on district
  
   If Sheets("Graph").Range("F1") = "Victoria" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "C:\Users\" & Environ("username") & "\Documents\Victoria\" & Sheets("Graph").Range("A4") & " Performance Report " & Sheets("Graph").Range("B16") & "-" & Sheets("Graph").Range("B25") & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   ElseIf Sheets("Graph").Range("F1") = "Caroni" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "C:\Users\" & Environ("username") & "\Documents\Caroni\" & Sheets("Graph").Range("A4") & " Performance Report " & Sheets("Graph").Range("B16") & "-" & Sheets("Graph").Range("B25") & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
    ElseIf Sheets("Graph").Range("F1") = "North Eastern" Then
       ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
       "C:\Users\" & Environ("username") & "\Documents\North Eastern\" & Sheets("Graph").Range("A4") & " Performance Report " & Sheets("Graph").Range("B16") & "-" & Sheets("Graph").Range("B25") & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
    :=False, OpenAfterPublish:=False
    ElseIf Sheets("Graph").Range("F1") = "South Eastern" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "C:\Users\" & Environ("username") & "\Documents\South Eastern\" & Sheets("Graph").Range("A4") & " Performance Report " & Sheets("Graph").Range("B16") & "-" & Sheets("Graph").Range("B25") & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   ElseIf Sheets("Graph").Range("F1") = "St George East" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "C:\Users\" & Environ("username") & "\Documents\St. George East\" & Sheets("Graph").Range("A4") & " Performance Report " & Sheets("Graph").Range("B16") & "-" & Sheets("Graph").Range("B25") & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   ElseIf Sheets("Graph").Range("F1") = "Port Of Spain" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "C:\Users\" & Environ("username") & "\Documents\Port of Spain\" & Sheets("Graph").Range("A4") & " Performance Report " & Sheets("Graph").Range("B16") & "-" & Sheets("Graph").Range("B25") & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   ElseIf Sheets("Graph").Range("F1") = "Tobago" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "C:\Users\" & Environ("username") & "\Documents\Tobago" & Sheets("Graph").Range("A4") & " Performance Report " & Sheets("Graph").Range("B16") & "-" & Sheets("Graph").Range("B25") & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   Else
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "C:\Users\" & Environ("username") & "\Documents\St. Patrick\" & Sheets("Graph").Range("A4") & " Performance Report " & Sheets("Graph").Range("B16") & "-" & Sheets("Graph").Range("B25") & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   End If
   
   'restore all sheets in the array to visible
   
    For Each xWs In sheetsArray
        xWs.Visible = True
    Next
    
    'delete trendline
    
    Sheets("ATTAIN (atleast 1)").ChartObjects("Chart 2").Chart.SeriesCollection(1).Trendlines(1).Delete
    
    'set a next school in the pivot filter to visible and repeat
    
   Next
   
   'delete trendline
   
   Sheets("Graph").ChartObjects("Chart 1").Chart.SeriesCollection(1).Trendlines(1).Delete
End Sub

