Attribute VB_Name = "Attainment"
Sub apply_autofilter_across_worksheets()
    Dim sheetsArray As Sheets
    Dim xWs As Worksheet
    Dim j As Integer
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim pi As PivotItem
    Set pt = Sheets("Graph").PivotTables("PivotTable1")
    Set pf = pt.PivotFields("School Code")
    Set sheetsArray = ActiveWorkbook.Sheets(Array("Performance Report 2012", "Performance Report 2013", "Performance Report 2014", "Performance Report 2015", "Performance Report 2016", "Performance Report 2017", "Performance Report 2018", "Performance Report 2019", "Performance Report 2020", "Performance Report 2021", "Performance Report 2022"))
    On Error Resume Next
    For Each pi In pf.PivotItems
        If pi.Visible = True Then
           For Each xWs In sheetsArray
              xWs.Range("B4").AutoFilter 2, pi
              j = WorksheetFunction.Count(xWs.Range("B4:B5000").Cells.SpecialCells(xlCellTypeVisible))
              If j = 0 Then
                xWs.Visible = False
              End If
           Next
        End If
    Next
End Sub
Sub Anomaly()
   Sheets("Graph").ChartObjects("Chart 1").Chart.SeriesCollection(1).Trendlines.Add
   With Sheets("Graph").ChartObjects("Chart 1").Chart.SeriesCollection(1).Trendlines(1)
    .Type = xlLinear
    .DisplayEquation = True
    .Format.Line.DashStyle = msoLineSysDot
    .Format.Line.Weight = 4
    .DataLabel.Font.Size = 32
    .DataLabel.Font.Color = vbBlack
   End With
   If Sheets("Graph").Range("F1") = "Victoria" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "Z:\MANDY\CSEC Performance Report Attainment Data 2012-2022\CSEC Performance Reports for Schools 2013-2022(1)\Victoria\" & Sheets("Graph").Range("A4") & " Performance Report 2013-2022.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   ElseIf Sheets("Graph").Range("F1") = "Caroni" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "Z:\MANDY\CSEC Performance Report Attainment Data 2012-2022\CSEC Performance Reports for Schools 2013-2022(1)\Caroni\" & Sheets("Graph").Range("A4") & " Performance Report 2013-2022.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
    ElseIf Sheets("Graph").Range("F1") = "North Eastern" Then
       ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
       "Z:\MANDY\CSEC Performance Report Attainment Data 2012-2022\CSEC Performance Reports for Schools 2013-2022(1)\North Eastern\" & Sheets("Graph").Range("A4") & " Performance Report 2013-2022.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
    :=False, OpenAfterPublish:=False
    ElseIf Sheets("Graph").Range("F1") = "South Eastern" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "Z:\MANDY\CSEC Performance Report Attainment Data 2012-2022\CSEC Performance Reports for Schools 2013-2022(1)\South Eastern\" & Sheets("Graph").Range("A4") & " Performance Report 2013-2022.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   ElseIf Sheets("Graph").Range("F1") = "St George East" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "Z:\MANDY\CSEC Performance Report Attainment Data 2012-2022\CSEC Performance Reports for Schools 2013-2022(1)\St. George East\" & Sheets("Graph").Range("A4") & " Performance Report 2013-2022.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   ElseIf Sheets("Graph").Range("F1") = "Port Of Spain" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "Z:\MANDY\CSEC Performance Report Attainment Data 2012-2022\CSEC Performance Reports for Schools 2013-2022(1)\Port of Spain\" & Sheets("Graph").Range("A4") & " Performance Report 2013-2022.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   ElseIf Sheets("Graph").Range("F1") = "Tobago" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "Z:\MANDY\CSEC Performance Report Attainment Data 2012-2022\CSEC Performance Reports for Schools 2013-2022(1)\Tobago\" & Sheets("Graph").Range("A4") & " Performance Report 2013-2022.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   Else
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "Z:\MANDY\CSEC Performance Report Attainment Data 2012-2022\CSEC Performance Reports for Schools 2013-2022(1)\St. Patrick\" & Sheets("Graph").Range("A4") & " Performance Report 2013-2022.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   End If
End Sub
Sub Pivot_Loop()
  Dim sheetsArray As Sheets
  Dim xWs As Worksheet
  Dim pt As PivotTable
  Dim pi As PivotItem
  Dim pf As PivotField
  Dim pi2 As PivotItem
  Dim j As Integer
  Set pt = Sheets("Graph").PivotTables("PivotTable1")
  Set pf = pt.PivotFields("School Code")
  
  'make an array of performance reports for each year, 2013-2022
  'performance reports include performance data for every secondary school in Trinidad and Tobago
  
  Set sheetsArray = ActiveWorkbook.Sheets(Array("Performance Report 2013", "Performance Report 2014", "Performance Report 2015", "Performance Report 2016", "Performance Report 2017", "Performance Report 2018", "Performance Report 2019", "Performance Report 2020", "Performance Report 2021", "Performance Report 2022"))
  
  'go through each school in the pivot filter
  'since a school must always be visible in the pivot filter, we set
  
  For Each pi In pf.PivotItems
     pf.PivotItems(pf.PivotItems.Count - 1).Visible = True
     
  'if a school is visible then filter the array of performance reports to show only the visible school's data
  'if a specific performance report does not contain that school's data, then hide that worksheet
  
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
   
   'Create a chart trendline for the visible school's chart
   
   Sheets("Graph").ChartObjects("Chart 1").Chart.SeriesCollection(1).Trendlines.Add
   With Sheets("Graph").ChartObjects("Chart 1").Chart.SeriesCollection(1).Trendlines(1)
    .Type = xlLinear
    .DisplayEquation = True
    .Format.Line.DashStyle = msoLineSysDot
    .Format.Line.Weight = 4
    .DataLabel.Font.Size = 32
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
  On Error Resume Next
  
  'export sheets as a pdf and place in a folder based on district
  
   If Sheets("Graph").Range("F1") = "Victoria" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "Z:\MANDY\CSEC Performance Report Attainment Data 2012-2022\CSEC Performance Reports for Schools 2013-2022(1)\Victoria\" & Sheets("Graph").Range("A4") & " Performance Report 2013-2022.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   ElseIf Sheets("Graph").Range("F1") = "Caroni" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "Z:\MANDY\CSEC Performance Report Attainment Data 2012-2022\CSEC Performance Reports for Schools 2013-2022(1)\Caroni\" & Sheets("Graph").Range("A4") & " Performance Report 2013-2022.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
    ElseIf Sheets("Graph").Range("F1") = "North Eastern" Then
       ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
       "Z:\MANDY\CSEC Performance Report Attainment Data 2012-2022\CSEC Performance Reports for Schools 2013-2022(1)\North Eastern\" & Sheets("Graph").Range("A4") & " Performance Report 2013-2022.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
    :=False, OpenAfterPublish:=False
    ElseIf Sheets("Graph").Range("F1") = "South Eastern" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "Z:\MANDY\CSEC Performance Report Attainment Data 2012-2022\CSEC Performance Reports for Schools 2013-2022(1)\South Eastern\" & Sheets("Graph").Range("A4") & " Performance Report 2013-2022.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   ElseIf Sheets("Graph").Range("F1") = "St George East" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "Z:\MANDY\CSEC Performance Report Attainment Data 2012-2022\CSEC Performance Reports for Schools 2013-2022(1)\St. George East\" & Sheets("Graph").Range("A4") & " Performance Report 2013-2022.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   ElseIf Sheets("Graph").Range("F1") = "Port Of Spain" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "Z:\MANDY\CSEC Performance Report Attainment Data 2012-2022\CSEC Performance Reports for Schools 2013-2022(1)\Port of Spain\" & Sheets("Graph").Range("A4") & " Performance Report 2013-2022.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   ElseIf Sheets("Graph").Range("F1") = "Tobago" Then
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "Z:\MANDY\CSEC Performance Report Attainment Data 2012-2022\CSEC Performance Reports for Schools 2013-2022(1)\Tobago\" & Sheets("Graph").Range("A4") & " Performance Report 2013-2022.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   Else
   ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
   "Z:\MANDY\CSEC Performance Report Attainment Data 2012-2022\CSEC Performance Reports for Schools 2013-2022(1)\St. Patrick\" & Sheets("Graph").Range("A4") & " Performance Report 2013-2022.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False
   End If
   
   'restore all performance reports in the array to visible
   
    For Each xWs In sheetsArray
        xWs.Visible = True
    Next
    
    'set a next school in the pivot filter to visible and repeat
    
   Next
End Sub

