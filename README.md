# Attainment
# Produce Attainment Performance Reports for 156 Secondary Schools in Trinidad and Tobago

Attainment reports for secondary schools include the percentage of students that pass five or more subjects at CSEC (ie. obtain a full certificate). The necessary data for these reports are first extracted from a CSEC Combined Raw Database and compiled into an excel workbook using Power Query. This data is then used to create a pivot table and a subsequent pivot chart. The pivot table and subsequent chart shows the percent of students attaining a full certificate for a ten year period for each school. The pivot filter includes each school code which is unique for every individual school. The second page in the excel workbook shows the Number of Students Registering, Attempting and Attaining atleast 1 CSEC Subject. The data for this sheet is also extracted from a CSEC Combined Raw Database using Power Query and a subsequent scatter plot is also added. Performance sheets for each consecutive year within a ten year period is then added to the workbook using the AddPerformanceSheet code. The workbook now contains the necessary attainment data to produce attainment reports.

Now individual attainment reports for each school are produced using the codes Attainmentpdf and Attainmentxlsx. The code is generally the same for both Attainmentpdf and Attainmentxlsx except that Attainmentpdf produces pdf reports and Attainmentxlsx produces excel reports. Similarly there is the AnomolyPDF and AnomalyXLSX codes. However,these 
Now individual attainment reports for each school are produced using the codes Attainmentpdf and Attainmentxlsx. The code is generally the same for both Attainmentpdf and Attainmentxlsx except that Attainmentpdf produces pdf reports and Attainmentxlsx produces excel reports. Similarly there is the AnomolyPDF and AnomalyXLSX codes. However,these codes can only produce a single report at a time and hence are only used in the event of an anomaly ie. the data has been updated for a school and a subsequent new report has to be made.
 
# Codes Summary

The codes are based on a loop which goes through each pivot item (ie. school code) in the pivot table on the first page of the workbook. Once a particular school code is visible, the remaining sheets in the workbook are subsequently filtered to show data pertaining to this school code only. The active sheets are then formatted to be exported and then subsequently exported into the relevant district folders.

# Attainmentpdf

First an array of sheets is made based on the subsequent sheets that will be filtered. Going through each school code (pi) in the pivot filter, the second to last school code is set to visible since an item must always be visible in a pivot filter. Now going through each school code (pi2) in the pivot filter again, if the item from the first loop is equal to the item from the second loop (ie. pi = pi2), then school code (pi2) is made visible in the pivot filter and the subsequent sheets in the array is filtered to show data pertaining to that school code (pi2). If a sheet in the array does not have data pertaining to that particular code, then the sheet is hidden in the workbook. Now if the item from the first loop is not equal to the item from the second loop (ie. pi <> pi2), then the school code (pi2) is not visible. This includes the second to last school code which was originally set to visible. This second loop is repeated for all the school codes (pi2) within the pivot filter till only one school code (pi2) is visible. Now that only one school code is visible, chart trendlines can be added to charts. Note that these trendlines weren't added directly when the charts were created because during the code, more that 1 school code is visible, which will cause the trendline to drop off. The worksheets within the workbook are then formatted to fit to 1 page and are exported to the school's located district folder as a pdf named accordingly. All sheets within the array are then reset to visible and the trendline for the chart on the second worksheet is deleted. This is because based on the loop a new trendline will be added once the next school code is set to visible. The trendline from the pivot chart does not need to deleted since it will automatically drop off when more that one school code is visible. A next school code (pi) is then set to visible and the loop is repeated till all school codes have been set to visible and exported. The trendline from the pivot chart of the workbook is then removed so the code can be rerun directly if need be.

# Attainmentxlsx

First an array of sheets (sheetsArray) is made based on the subsequent sheets that will be filtered. A subsequent array (xArray) containing each active sheet of the workbook is also made. A double pivot loop is made identical to the pivot loop in Attainmentpdf and sheetsArray is subsequently filtered like in Attainmentpdf. Once only one school code is visible, a trendline is added to the pivot chart on the first page. A new workbook is then added. Then each filtered worksheet from xArray is copied and pasted within the new workbook. For the first page which include the pivot table and chart (xArray(1)), both the page and chart are copied directly to the new workbook. However for the second worksheet, only the page was copied directly. A new chart was created within the new workbook using the copied, filtered data. This is because when copied and pasted directly, the chart is still linked to the original and changes when the original changes. Once each sheet is pasted, the new workbook is saved, named appropriately, closed and then exported to the school's located district folder. Once the loop is finished, the trendline from the pivot chart of the workbook is then removed so the code can be rerun directly if need be.

# AnomalyPDF and AnomalyXLSX

These codes follow Attainmentpdf and Attainmentxlsx closely. The only difference is that AnomalyPDF and AnomalyXLSX only have 1 pivot loop and one needs to select the relevant school code in the pivot filter before the code is run. These codes go through each school code in the pivot filter and if one of them is visible, then all sheets in sheetsArray are filtered accordingly.
