Attribute VB_Name = "AddPerformanceSheets"
Sub Reports()
Dim wb As Workbook
Dim i As Integer
Dim yearArr() As Variant

'make an array of years

yearArr = Array("2013", "2014", "2015", "2016", "2017", "2018", "2019", "2020", "2021", "2022")
Sheet_Name = "PERFORMANCE REPORT"
Set New_Wbk = ThisWorkbook

'open folder of performance reports and copy performance report sheet
'paste to open workbook and rename according to year

    Dim FolderPath As String
    Dim FilePath As String
        FolderPath = "C:\Users\mandy.mohan\Documents\Analysis\Performance Reports (Updated)\"
        FilePath = Dir(FolderPath & "*.xls*")
        Do While FilePath <> ""
            Set wb = Workbooks.Open(FolderPath & FilePath)
            wb.Worksheets(Sheet_Name).UsedRange.Copy
            ActiveColumn = ActiveColumn + 1
            New_Wbk.Sheets.Add(After:=Sheets(Sheets.Count)).Name = Sheet_Name & yearArr(i)
            New_Wbk.Worksheets(Sheet_Name & yearArr(i)).Cells(1, 1).PasteSpecial Paste:=xlPasteAll
            i = i + 1
            FilePath = Dir
            Loop
End Sub
