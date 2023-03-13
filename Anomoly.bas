Attribute VB_Name = "Anomoly"
Sub Anomaly()
    Dim sheetsArray As Sheets
    Dim xWs As Worksheet
    Dim j As Integer
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim pi As PivotItem
    Set pt = Sheets("Graph").PivotTables("PivotTable1")
    Set pf = pt.PivotFields("School Code")
    Set sheetsArray = ActiveWorkbook.Sheets(Array("Performance Report 2013", "Performance Report 2014", "Performance Report 2015", "Performance Report 2016", "Performance Report 2017", "Performance Report 2018", "Performance Report 2019", "Performance Report 2020", "Performance Report 2021", "Performance Report 2022"))
    Sheets("Graph").ChartObjects("Chart 1").Chart.SeriesCollection(1).Trendlines.Add
    With Sheets("Graph").ChartObjects("Chart 1").Chart.SeriesCollection(1).Trendlines(1)
        .Type = xlLinear
        .DisplayEquation = True
        .Format.Line.DashStyle = msoLineSysDot
        .Format.Line.Weight = 4
        .DataLabel.Font.Size = 32
        .DataLabel.Font.Color = vbBlack
    End With
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

