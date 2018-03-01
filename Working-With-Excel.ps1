﻿function Remove-ComObject {
 # Requires -Version 2.0
 [CmdletBinding()]
 param()
 end {
  Start-Sleep -Milliseconds 500
  [Management.Automation.ScopedItemOptions]$scopedOpt = 'ReadOnly, Constant'
  Get-Variable -Scope 1 | Where-Object {
   $_.Value.pstypenames -contains 'System.__ComObject' -and -not ($scopedOpt -band $_.Options)
  } | Remove-Variable -Scope 1 -Verbose:([Bool]$PSBoundParameters['Verbose'].IsPresent)
  [gc]::Collect()
 }
}
 
Function Using-Culture (
[System.Globalization.CultureInfo]$culture,
[ScriptBlock]$script)
{
    $OldCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
    trap
    {
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $OldCulture
    }
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $culture
    $ExecutionContext.InvokeCommand.InvokeScript($script)
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $OldCulture
} # End Function
 
Using-Culture en-us {
$xl= New-Object -COM Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $False
$xl.Workbooks.Open("C:\Scripts\PowerShell\test.xls")
 
$xl.quit() 
Remove-ComObject -Verbose
Start-Sleep -Milliseconds 250
Get-Process Excel
#[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
# ======================================================
#_______________________________________________________________________
# How do I create an Excel object? 
  
$xl = new-object -comobject excel.application 
 
# ________________________________________________________________________ 
#How do I make Excel visible?
 
$xl.Visible = $true
 
# ________________________________________________________________________
#
# How do I add a workbook? 
  
$wb = $xl.Workbooks.Add()
  
# By default this adds three empty worksheets
# ________________________________________________________________________
#
#
# How do I open an existing Workbook ?
#
$xl = New-Object -comobject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $False
$wb = $xl.Workbooks.Open("C:\Scripts\powershell\test.xls")
 
# ________________________________________________________________________
#
# 
#How do I add a worksheet to an existing workbook? 
  
$xl = new-object -comobject excel.application
$xl.Visible = $true
$xl.DisplayAlerts = $False
$wb = $xl.Workbooks.Open("C:\Scripts\test.xls")
$ws = $xl.Sheets.Add()
 
#
# ________________________________________________________________________
#
# How do I activate a a worksheet?
# Create Excel.Application object
$xl = New-Object -comobject Excel.Application
# Show Excel
$xl.visible = $true
$xl.DisplayAlerts = $False
# Create a workbook
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
# Get sheets
$ws1 = $wb.worksheets | where {$_.name -eq "sheet1"} #<------- Selects sheet 1
$ws2 = $wb.worksheets | where {$_.name -eq "sheet2"} #<------- Selects sheet 2
$ws3 = $wb.worksheets | where {$_.name -eq "sheet3"} #<------- Selects sheet 3
# Activate sheet 1
$ws1.activate()
Start-Sleep 1 
# Activate sheet 2
$ws2.activate()
Start-Sleep 1 
# Activate sheet 3
$ws3.activate()
# _________________________________________________________________
# How do I change the value of a cell? 
  
$ws1.cells.Item(1,1).Value() = "x"
or
$ws1.cells.Item($row,$col).Value() = "x"  # where $row = the row your on and col = the column
# ________________________________________________________________________
#
# How do I select an entire row?
$range = $ws1.Cells.Item(1,1).EntireRow
$range.font.bold = $true # sets the top row to bold
#
#
# How do I Autofit the entire worksheet?
[void]$ws1.cells.entireColumn.Autofit()
#
#
# How do I name a worksheet?
$Date = (Get-Date -format "MM-dd-yyyy")
$xl.Worksheets.item(1).name = $Date
#
#
# How do I Find a specified cell? 
#
 
$xlCellTypeLastCell = 11
 
$used = $ws.usedRange 
$lastCell = $used.SpecialCells($xlCellTypeLastCell) 
$row = $lastCell.row # goes to the last used row in the worksheet
 
for ($i = 1; $i -le $row; $i++) {
If ($ws1.cells.Item(1,2).Value() = "y") {
# "do something"
          }
}
 
# And another way:
#
$mainRng = $ws.usedRange
$mainRng.Select()
$objSearch = $mainRng.Find("Grand Total")
$objSearch.Select()
#
#
# Find a very specific item:
 
$Rng = $ws1.range("AS:AS")
$oSearch = "IBSS-Corporate-OSE-PrdSupp-Win"
#search
$ySearch = $xl.WorksheetFunction.Match($oSearch,$Rng,0) # gives you the ROW # of the found cell
$range3 = $ws1.range("AS$ySearch")
[void]$range3.select()
$z = $ySearch - 1
Write-Host $z
# [void]$range3.Delete()
$range3 = $ws1.range("A2:CI$z")
[void]$range3.select()
[void]$range3.Delete()
#
#
# How do I match an Item?
#
$used = $ws1.usedRange 
$Rng = $ws1.range("A:A")
$oSearch = $Null
$ySearch = $Rng.Find($oSearch)
$zSearch = $ySearch.Row
$R = $zSearch
 
For ($i = $R; $i -ge 2; $i--)  {
  Write-Progress -Activity "Searching $R Old's..." `
     -PercentComplete ($i/$R*100) -CurrentOperation `
     "$i of $R to go" -Status "Please wait."
      $Rng = $ws1.Cells.Item($i, 1).value()
If ($Rng -match '\bold\b|-new|zz\.|-INACTIVE|-DELETE|\.\.+') {
# Match at Begining and end of word '\b   \b' <-- looking for old? '\bold\b' 
# must put a \ before a period \. othewise will catch every charactor.
   Write-Host $Rng
   $Range = $ws1.Cells.Item($i, 1).EntireRow
     [void]$Range.Select()
     [void]$Range.Delete()
     $R = $R - 1
      } 
    }
     Write-Progress -Activity "Searching -Old's..." `
  -Completed -Status "All done."
 
# ________________________________________________________________________
#
# How do I delete a row in Excel?
#
 
Function DelRow($y) {   
    $Range = $ws1.Cells.Item($y, 1).EntireRow
    [void]$Range.Select()
    [void]$Range.Delete()
}
DelRow($y)
 
# How do I find and delete empty rows in Excel?
 
Function DeleteEmptyRows {
$used = $ws.usedRange 
$lastCell = $used.SpecialCells($xlCellTypeLastCell) 
$row = $lastCell.row 
 
for ($i = 1; $i -le $row; $i++) {
    If ($ws.Cells.Item($i, 1).Value() -eq $Null) {
        $Range = $ws.Cells.Item($i, 1).EntireRow
        $Range.Delete()
        }
    }
} 
 
$xlCellTypeLastCell = 11 
 
$xl = New-Object -comobject excel.application 
$xl.Visible = $true
$xl.DisplayAlerts = $False
$wb = $xl.Workbooks.Open("C:\Scripts\Test.xls") # <-- Change as required!
$ws = $wb.worksheets | where {$_.name -eq "Servers" } # <-- Or Sheet1 or Whatever 
 
DeleteEmptyRows   # <—Call Function 
 
  
# ________________________________________________________________________
 
#
#
# How do I set a range with variables? 
  
$ws1.range("a${Sumrow}:b$Sumrow").font.bold = "true"
 
# separate the : from the $ with {} on the left hand side
# ________________________________________________________________________
#
# How do I Set range to a value?
  
$range4=$ws.range("3:3")
$range4.cells="Row 3"
 
# ________________________________________________________________________
#
# How do I list the workbook's name? 
  
$wb.Name
 
# ______________________________________________________________
#
# How do I find the last used row number?
 
$xlCellTypeLastCell = 11
$used = $ws.usedRange
$lastCell = $used.SpecialCells($xlCellTypeLastCell)
$lastrow = $lastCell.row 
 
# Or
 
$mainRng = $ws1.UsedRange.Cells 
$RowCount = $mainRng.Rows.Count  
$R = $RowCount
 
# ________________________________________________________________________
#
# How do I find the last used column number? 
#This works for columns and rows 
  
$mainRng = $ws1.UsedRange.Cells 
$ColCount = $mainRng.Columns.Count 
$RowCount = $mainRng.Rows.Count  
$xRow = $RowCount
$xCol = $ColCount
 
# ________________________________________________________________________
#
# How do I loop through a range of cells by row number? 
  
$xl = New-Object -comobject excel.application
$xl.Visible = $true
$xl.DisplayAlerts = $False
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
for ($row = 1; $row -lt 11; $row++)
{
    $ws.Cells.Item($row,1) = $row
}
# ________________________________________________________________________
#
# How do I get today's date and format it as a string? 
  
$m = (get-date).month
$d = (get-date).day
$y = [string] (get-date).year
$y = $y.substring($y.length - 2, 2)
$f = "C:\Scripts\" + $m + "-" + $d + "-" + $y + ".xlsx"
$wb.SaveAs($F) # C:\Scripts\6-18-10.xlsx 
 
# OR
 
$Date = (Get-Date -format "MM-dd-yyyy")
 
#
# ________________________________________________________________________
#
# How do I write a list of files to Excel? 
  
$xl = new-object -comobject excel.application
$xl.Visible = $true
$xl.DisplayAlerts = $False
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$row = 1
$s = dir Z:\MBSA_Report\ScanData\*.mbsa
$s | foreach -process {
    $ws.Cells.Item($row,1) = $_;
    $row++
}
 
#
# this takes a long file name with spaces and splits it up,
# It then picks out the 3rd element and writes it out.
# The 3rd element is [2] because the first one is [0].
# sample file name: AD-Dom - 3IDCT001 (5-21-2010 4-28 PM).mbsa 
  
$xl = new-object -comobject excel.application
$xl.Visible = $true
$xl.DisplayAlerts = $False
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$row = 1
dir Z:\MBSA_Report\ScanData\*.mbsa |
ForEach-Object {
$FileName = $_.name
$N = $FileName.tostring()
$E = $N.split()
$F = $E[2]
$ws.Cells.Item($row,1) = $f;
    $row++
}
#Or for the whole file name
$ws.Cells.Item($row,1) = $_;
 
# ________________________________________________________________________
#
# How do I write a list of processes to Excel? 
  
function Release-Ref ($ref) {
#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ref)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
}
# -----------------------------------------------------
$xl = New-Object -comobject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $False
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$range = $ws.Cells.Item(1,1)
$row = 1
$s = Get-Process | Select-Object name
$s | foreach -process {
    $range = $ws.Cells.Item($row,1);
    $range.Value = $_.Name;
    $row++ } 
$wb.SaveAs("C:\Scripts\Get_Process.xls")
Release-Ref $range
Release-Ref $ws
Release-Ref $wb
$xl.Quit()
Release-Ref $xl
#***For a remote machine try
$strComputer = (remote machine name)
$P = gwmi win32_process -comp $strComputer
 
# ________________________________________________________________________
#
# How do I write the command history to Excel? 
  
function Release-Ref ($ref) {
#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ref)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
}
 
$xl = New-Object -comobject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $False
$wb = $excel.Workbooks.Add()
$ws = $workbook.Worksheets.Item(1)
$range = $worksheet.Cells.Item(1,1)
$row = 1
$s = Get-History | Select-Object CommandLine $s | foreach -process { `
$range = $worksheet.Cells.Item($row,1); `
$range.Value = $_.CommandLine; `
$row++ }
$xl.DisplayAlerts = $False
$wb.SaveAs("C:\Scripts\Get_CommandLine.xls")
Release-Ref $range
Release-Ref $ws
Release-Ref $wb
$xl.Quit()
Release-Ref $xl
 
# ________________________________________________________________________
#
# How Can I Convert a Tilde-Delimited File to Microsoft Excel Format?
# Script name: ConvertTilde.ps1
# Created on: 2007-01-06
# Author: Kent Finkle
# Purpose: How Can I Convert a Tilde-Delimited File to Microsoft Excel Format? 
  
$s = gc C:\Scripts\Test.txt
$s = $s -replace("~","`t")
$s | sc C:\Scripts\Test.txt
$xl = new-object -comobject excel.application
$xl.Visible = $true
$xl.DisplayAlerts = $False
$wb = $xl.Workbooks.Open("C:\Scripts\Test.txt")
 
# ________________________________________________________________________
#
# How can I add Validation to an Excel Worksheet?
#
# $comments = @'
# Script name: Add-Validation.ps1
# Created on: Wednesday, September 19, 2007
# Author: Kent Finkle
# Purpose: How can I use Windows Powershell to Add Validation to an
# Excel Worksheet?
# '@
#-----------------------------------------------------
 
function Release-Ref ($ref) {
([System.Runtime.InteropServices.Marshal]::ReleaseComObject(
[System.__ComObject]$ref) -gt 0)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
} # End Function 
$xlValidateWholeNumber = 1
$xlValidAlertStop = 1
$xlBetween = 1
$xl = new-object -comobject excel.application
$xl.Visible = $true
$xl.DisplayAlerts = $False
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$r = $ws.Range("e5")
$r.Validation.Add($xlValidateWholeNumber,$xlValidAlertStop, $xlBetween, "5", "10")
$r.Validation.InputTitle = "Integers"
$r.Validation.ErrorTitle = "Integers"
$r.Validation.InputMessage = "Enter an integer from five to ten"
$r.Validation.ErrorMessage = "You must enter a number from five to ten"
$a = Release-Ref $r
$a = Release-Ref $ws
$a = Release-Ref $wb
$a = Release-Ref $xl
 
# ________________________________________________________________________
#
# How do I add a Chart to an Excel Worksheet? 
  
$xRow = 1
$yrow = 8
$xl = New-Object -c excel.application
$xl.visible = $true
$xl.DisplayAlerts = $False
$wb = $xl.workbooks.add()
$ws = $wb.sheets.item(1)
1..8 | % { $ws.Cells.Item(1,$_) = $_ }
1..8 | % { $ws.Cells.Item(2,$_) = 9-$_ }
$range = $ws.range("a${xrow}:h$yrow")
$range.activate
# create and assign the chart to a variable
$ch = $xl.charts.add()   # This will open a new sheet
$ch = $ws.shapes.addChart().chart # This will put the Chart in the selected WorkSheet
$ch.chartType = 58
$ch.setSourceData($range)
$ch.HasTitle = $true
$ch.ChartTitle.Text = "Count of KB's"
$ch.export("C:\test.jpg")
$xl.quit() 
# excel has 48 chart styles, you can cycle through all
1..48 | % {$ch.chartStyle = $_; $xl.speech.speak("Style $_"); sleep 1}
$ch.chartStyle = 27      # <-- use the one you like 
 
# And another Chart sample:
 
Function XLcharts {
$xlColumnClustered = 51
$xlColumns = 2
$xlLocationAsObject = 2
$xlCategory = 1
$xlPrimary = 1
$xlValue = 2
$xlRows = 1
$xlLocationAsNewSheet = 1
$xlRight = -4152
$xlBuiltIn =21
$xlCategory = 1 
$xl = New-Object -comobject excel.application
$xl.Visible = $true
$xl.DisplayAlerts = $False
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws = $wb.Sheets.Add()
$ws.Cells.Item(1, 2) =  "Jan"
$ws.Cells.Item(1, 3) =  "Feb"
$ws.Cells.Item(1, 4) =  "Mar"
$ws.Cells.Item(2, 1) =  "John"
$ws.Cells.Item(3, 1) =  "Mae"
$ws.Cells.Item(4, 1) =  "Al"
$ws.Cells.Item(2, 2) =  100
$ws.Cells.Item(2, 3) =  200
$ws.Cells.Item(2, 4) =  300
$ws.Cells.Item(3, 2) =  400
$ws.Cells.Item(3, 3) =  500
$ws.Cells.Item(3, 4) =  600
$ws.Cells.Item(4, 2) =  900
$ws.Cells.Item(4, 3) =  800
$ws.Cells.Item(4, 4) =  700 
$Range = $ws.range("A1:D4")
$ch = $xl.charts.add()
$ch.chartType = 58
$ch.name ="Bar Chart"
$ch.Tab.ColorIndex = 3
$ch.setSourceData($Range)
[void]$ch.Location, $xlLocationAsObject, "Bar Chart"
$ch.HasTitle = $False
$ch.Axes($xlCategory, $xlPrimary).HasTitle = $False
$ch.Axes($xlValue, $xlPrimary).HasTitle = $False
$ch2 = $xl.Charts.Add() | Out-Null
$ch2.HasTitle = $true
$ch2.ChartTitle.Text = "Sales"
$ch2.Axes($xlCategory).HasTitle = $true
$ch2.Axes($xlCategory).AxisTitle.Text = "1st Quarter"
$ch2.Axes($xlValue).HasTitle = $True
$ch2.Axes($xlValue).AxisTitle.Text = "Dollars"
[void]$ch2.Axes($xlValue).Select
$ch2.name ="Columns with Depth"
$ch2.Tab.ColorIndex = 5
[void]$ws.cells.entireColumn.Autofit()
} # End Function
# ________________________________________________________________________
#
# How do I Move and resize a chart in an Excel Worksheet? 
 
$xl = New-Object -comobject Excel.Application   # Opens Excel and 3 empty Worksheets
# Show Excel
$xl.visible = $true
$xl.DisplayAlerts = $False
# Open a workbook
$wb = $xl.workbooks.add() 
#Create Worksheets
$ws = $wb.Worksheets.Item(1)
1..8 | % { $ws.Cells.Item(1,$_) = $_ }          # adds some data
1..8 | % { $ws.Cells.Item(2,$_) = 9-$_ }        # adds some data
# < ----- This is the good part ------------------------------------------------------->
$range = $ws.range("a${xrow}:h$yrow")           # sets the Data range we want to chart
# create and assign the chart to a variable
#$ch = $xl.charts.add()                         # This will open a new sheet
$ch = $ws.shapes.addChart().chart               # This will put the Chart in the selected WorkSheet
$ch.chartType = 58                              # Select Chart Type
$ch.setSourceData($range)                       # Create the Chart
$ch.HasTitle = $true
$ch.ChartTitle.Text = "Sales"
$RngToCover = $ws.Range("D5:J19")               # This is where we want the chart
$ChtOb = $ch.Parent                             # This selects the curent Chart
$ChtOb.Top = $RngToCover.Top                    # This moves it up to row 5
$ChtOb.Left = $RngToCover.Left                  # and to column D 
$ChtOb.Height = $RngToCover.Height              # resize This sets the height of your chart to Rows 5 - 19
$ChtOb.Width = $RngToCover.Width                # resize This sets the width to Columns D - J
# ________________________________________________________________________
#
# How do I sort a column in an Excel Worksheet? 
  
$xlSummaryAbove = 0
$xlSortValues = $xlPinYin = 1
$xlAscending = 1
$xlDescending = 2
$xlYes = 1 
$xl = New-Object -comobject excel.application
$xl.Visible = $true
$xl.DisplayAlerts = $False
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1) 
# one-column sort
 
$R = $ws.UsedRange
$r2 = $ws.Range("B2") # Sorts on Column B and leaves Header alone 
$a = $r.sort($r2, $xlAscending) 
 
# two-column sort
 
[void]$ws1.Activate()
$last = $ws1.UsedRange.SpecialCells(11).Address($False,$false) 
$range1 = $ws1.range("A2:$last" )
$range2 = $ws1.range("D2") 
$range3 = $ws1.range("A2")
#two-column sort --> works 
$ws1.sort.sortFields.clear()
[void]$ws1.sort.sortFields.add($range2, $xlSortOnValues, $xlAscending,`
$xlSortNormal)
[void]$ws1.sort.sortFields.add($range3, $xlSortOnValues, $xlAscending,`
$xlSortNormal)
$ws1.sort.setRange($range1)
$ws1.sort.header = $xlNo
$ws1.sort.orientation = $xlTopToBottom
$ws1.sort.apply()
#-----------------------------------------------------
# How do I use xlConstants? 
$xlOpenXMLWorkbook = 51
$xlAutomatic=-4105
$xlBottom = -4107
$xlCenter = -4108
$xlContext = -5002
$xlContinuous=1
$xlDiagonalDown=5
$xlDiagonalUp=6
$xlEdgeBottom=9
$xlEdgeLeft=7
$xlEdgeRight=10
$xlEdgeTop=8
$xlInsideHorizontal=12
$xlInsideVertical=11
$xlNone=-4142
$xlThin=2 
$xl = new-object -com excel.application
$xl.visible=$true
$wb = $xl.workbooks.open("d:\book1.xls")
$ws = $wb.worksheets | where {$_.name -eq "sheet1"}
$selection = $ws.range("A1:D1")
$selection.select() 
$selection.HorizontalAlignment = $xlCenter
$selection.VerticalAlignment = $xlBottom
$selection.WrapText = $false
$selection.Orientation = 0
$selection.AddIndent = $false
$selection.IndentLevel = 0
$selection.ShrinkToFit = $false
$selection.ReadingOrder = $xlContext
$selection.MergeCells = $false
$selection.Borders.Item($xlInsideHorizontal).Weight = $xlThin
 
## -----------------------------------------------------------
# How do I set the column width ? 
 
$ws.columns.item(1).columnWidth = 50
$ws.columns.item('a').columnWidth = 50
# Or
$ws.cells.item(1).columnWidth = 50
$ws.range('a:a').columnwidth = 50
 
## -----------------------------------------------------------
# How do I center a column?  
## You can try this to center a column: 
  
[reflection.assembly]::loadWithPartialname("Microsoft.Office.Interop.Excel") |
Out-Null
$xlConstants = "microsoft.office.interop.excel.Constants" -as [type] 
  
$ws.columns.item("F").HorizontalAlignment = $xlConstants::xlCenter
$ws.columns.item("K").HorizontalAlignment = $xlConstants::xlCenter
 
  
# The next four lines of code create four enumeration types.
# Enumeration types are used to tell Excel which values are allowed
# for specific types of options. As an example, xlLineStyle enumeration
# is used to determine the kind of line to draw: double, dashed, and so on.
# These enumeration values are documented on MSDN.
# To make the code easier to read, we create shortcut aliases for each
# of the four enumeration types we will be using.
# Essentially, we're casting a string that represents the name of the
# enumeration to a [type]. This technique is actually a pretty cool trick:
# http://technet.microsoft.com/en-us/magazine/2009.01.heyscriptingguy.aspx 
  
$lineStyle = "microsoft.office.interop.excel.xlLineStyle" -as [type]
$colorIndex = "microsoft.office.interop.excel.xlColorIndex" -as [type]
$borderWeight = "microsoft.office.interop.excel.xlBorderWeight" -as [type]
$chartType = "microsoft.office.interop.excel.xlChartType" -as [type]
 
For($b = 1 ; $b -le 2 ; $b++)
{
$ws.cells.item(1,$b).font.bold = $true
$ws.cells.item(1,$b).borders.LineStyle = $lineStyle::xlDashDot
$ws.cells.item(1,$b).borders.ColorIndex = $colorIndex::xlColorIndexAutomatic
$ws.cells.item(1,$b).borders.weight = $borderWeight::xlMedium
}
$workbook.ActiveChart.chartType = $chartType::xl3DPieExploded
$workbook.ActiveChart.SetSourceData($range) 
 
# ________________________________________________________________________
#
# How do I use autofill in excel?
  
$xlFillWeekdays = 6 
$xl = New-Object -com excel.application
$xl.visible=$true
$wb = $xl.workbooks.add()
$ws = $wb.worksheets | where {$_.name -eq "sheet1"}
$range1= $ws.range("A1")
$range1.value() = (get-date).toString("d")
$range2 = $ws.range("A1:A25")
$range1.AutoFill($range2,$xlFillWeekdays)
$range1.entireColumn.Autofit() 
# Another example:
$xlCellTypeLastCell = 11
$xl = new-object -com excel.application
$xl.visible=$true
$wb = $xl.workbooks.add()
$ws = $wb.worksheets | where {$_.name -eq "sheet1"} 
$used = $ws.usedRange
$lastCell = $used.SpecialCells($xlCellTypeLastCell)
$lastrow = $lastCell.row 
$ws.Cells.Item(2,1).FormulaR1C1 = "=CONCATENATE(C[+1],C[+2],C[+3])"
$range1= $ws.range("A2")
$r = $ws.Range("A2:A$lastrow")
$range1.AutoFill($r) | Out-Null
[void]$range1.entireColumn.Autofit() 
$wb.close()
$xl.quit()
 
# ________________________________________________________________________
#
# How to get a range and format it in excel?
  
# get-excelrange.ps1
# opens an existing workbook in Excel 2007, using PowerShell
# and turns a range bold # Thomas Lee - t...@psp.co.uk
# Create base object
$xl = new-object -comobject Excel.Application 
# make Excel visible
$xl.visible = $true
$xl.DisplayAlerts = $False
# open a workbook
$wb = $xl.workbooks.open("C:\Scripts\powershell\test.xls") 
  
# Get sheet1
$ws = $wb.worksheets | where {$_.name -eq "sheet1"} 
  
# Make A1-B1 bold
$range = $ws.range("A1:B1")
$range.font.bold = "true"
  
# Make A2-B2 italic
$range2 = $ws.range("A2:B2")
$range2.font.italic = "true"
  
# Set range to a value
$range4=$ws.range("3:3")
$range4.cells="Row 3"
 
# Set range to a Font size
$range3=$ws.range("A2:B2")
$Range3.font.size=24 
 
# now format an entire row
 
$range4=$ws.range("3:3")
 
$range4.font.italic="$true"
$range4.font.bold=$true
$range4.font.size=10
$range4.font.name="Comic Sans MS"
# now format a Range of cells
$ws.Range("D1:F5").NumberFormat = "#,##0.00"
 
# ______________________________________________________________________
#
# How do I add a comment to a cell in Excel? 
  
$xl = New-Object -com Excel.Application
$xl.visible = $true
$xl.DisplayAlerts = $False
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Cells.Item(1,1) = “A value in cell A1.”
[void]$ws.Range("A1").AddComment()
[void]$ws.Range("A1").comment.Visible = $False
[void]$ws.Range("A1").Comment.text("OldDog: `r this is a comment")
[void]$ws.Range("A2").Select 
 
# The 'r adds a line feed after the comment's author. This is required!
# ________________________________________________________________________
#
# How do I copy and Paste in Excel (special)?
 
$xl = New-Object -comobject Excel.Application
# Show Excel
$xl.visible = $true
$xl.DisplayAlerts = $False
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1) 
$xlPasteValues = -4163          # Values only, not formulas
$xlCellTypeLastCell = 11        # to find last used cell
$used = $ws.usedRange
$lastCell = $used.SpecialCells($xlCellTypeLastCell)
$row = $lastCell.row
$range = $ws.UsedRange
[void]$ws.Range("A8:F$row").Copy()
[void]$ws.Range("A8").PasteSpecial(-4163) 
 
# __________________________________________________________________________
#
# How do I Add Worksheets, name them and save as today's date? 
#****************************************** 
# get today's date and format it as a string. 
  
$m = (get-date).month
$d = (get-date).day
$y = [string] (get-date).year
$y = $y.substring($y.length - 2, 2)
$f = "C:\Scripts\" + $m + "-" + $d + "-" + $y + ".xlsx"
# Or
$Date = (Get-Date -format "MM-dd-yyyy")
$f = "C:\Scripts\$Date.xlsx"
 
# Create Excel.Application object
$xl = New-Object -comobject Excel.Application
# Show Excel
$xl.visible = $true
$xl.DisplayAlerts = $False
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1) 
# this will add 9 sheets to the default 3 for a total of 12 sheets 
for ($i = 0; $i -le 8; $i++) {
$ws = $wb.Sheets.Add() } 
# Now we name our new sheets 
$xl.Worksheets.item(1).name = "Jan"
$xl.Worksheets.item(2).name = "Feb"
$xl.Worksheets.item(3).name = "Mar"
$xl.Worksheets.item(4).name = "Apr"
$xl.Worksheets.item(5).name = "May"
$xl.Worksheets.item(6).name = "June"
$xl.Worksheets.item(7).name = "July"
$xl.Worksheets.item(8).name = "Aug"
$xl.Worksheets.item(9).name = "Sept"
$xl.Worksheets.item(10).name = "Oct"
$xl.Worksheets.item(11).name = "Nov"
$xl.Worksheets.item(12).name = "Dec"
# and here we save it 
$wb.SaveAs($F) 
 
#*******************************************
# How do I find duplicate entries in Excel?
# This Function creates a spreadsheet with some
# duplicate names and then highlights the Dups in Blue
  
Function XLFindDups {
$xlExpression = 2
$xlPasteFormats = -4122
$xlNone = -4142
$xlToRight = -4161
$xlToLeft = -4159
$xlDown = -4121
$xlShiftToRight = -4161
$xlFillDefault = 0
$xlSummaryAbove = 0
$xlSortValues = $xlPinYin = 1
$xlAscending = 1
$xlDescending = 2
$xlYes = 1
$xlTopToBottom = 1
$xlPasteValues = -4163          # Values only, not formulas
$xlCellTypeLastCell = 11        # to find last used cell 
$xl = new-object -comobject excel.application
$xl.Visible = $true
$xl.DisplayAlerts = $False
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.name = "Concatenate"
$ws.Tab.ColorIndex = 4 
$ws.Cells.Item(1,1) = "FirstName"
$ws.Cells.Item(1,2) = "MI"
$ws.Cells.Item(1,3) = "LastName"
$ws.Cells.Item(2,1) = "Jesse"
$ws.Cells.Item(2,2) = "L"
$ws.Cells.Item(2,3) = "Roberts"
$ws.Cells.Item(3,1) = "Mary"
$ws.Cells.Item(3,2) = "S"
$ws.Cells.Item(3,3) = "Talbert"
$ws.Cells.Item(4,1) = "Ben"
$ws.Cells.Item(4,2) = "N"
$ws.Cells.Item(4,3) = "Smith"
$ws.Cells.Item(5,1) = "Ed"
$ws.Cells.Item(5,2) = "S"
$ws.Cells.Item(5,3) = "Turner"
$ws.Cells.Item(6,1) = "Mary"
$ws.Cells.Item(6,2) = "S"
$ws.Cells.Item(6,3) = "Talbert"
$ws.Cells.Item(7,1) = "Jesse"
$ws.Cells.Item(7,2) = "L"
$ws.Cells.Item(7,3) = "Roberts"
$ws.Cells.Item(8,1) = "Joe"
$ws.Cells.Item(8,2) = "L"
$ws.Cells.Item(8,3) = "Smith"
$ws.Cells.Item(9,1) = "Ben"
$ws.Cells.Item(9,2) = "A"
$ws.Cells.Item(9,3) = "Smith"
$used = $ws.usedRange
$lastCell = $used.SpecialCells($xlCellTypeLastCell)
$lastrow = $lastCell.row 
$range4=$ws.range("2:2")
$range4.Select() | Out-Null
$xl.ActiveWindow.FreezePanes = $true
$ws.cells.EntireColumn.AutoFit() | Out-Null
$range1 = $ws.Range("A1").EntireColumn
$range1.Insert($xlShiftToRight) | Out-Null
$range1.Select() | Out-Null
$ws.Cells.Item(1, 1) = "Concat"
$r2 = $ws.Range("A2")
$r2.Select() | Out-Null
$ws.Cells.Item(2,1).FormulaR1C1 = "=CONCATENATE(C[+1],C[+2],C[+3])"
$range1= $ws.range("A2")
$r = $ws.Range("A2:A$lastrow")
$range1.AutoFill($r) | Out-Null
$range.EntireColumn.AutoFit() | Out-Null
$select = $range1.SpecialCells(11).Select()  | Out-Null
$ws.Range("A2:A$lastrow").Copy()| Out-Null
$ws1 = $wb.Sheets.Add()
$ws1.name = "FindDups"
$ws1 = $wb.worksheets | where {$_.name -eq "FindDups"}
$ws1.Tab.ColorIndex = 5
$ws1.Select() | Out-Null
[void]$ws1.Range("A2").PasteSpecial(-4163)
$ws1.Range("A1").Select() | Out-Null
$objRange = $xl.Range("B1").EntireColumn
[void] $objRange.Insert($xlShiftToRight) 
$ws1.Cells.Item(1, 2) = "Dups"
$range = $ws.range("B1:D$lastrow")
$range.copy() | Out-Null
[void]$ws1.Range("C1").PasteSpecial(-4163) 
$ws1.Cells.Item(2,2).FormulaR1C1 = "=COUNTIF(C[-1],RC[-1])>1"
$range1= $ws1.range("B2")
$range2 = $ws1.range("B2:B$lastrow")
[void]$range1.AutoFill($range2,$xlFillDefault) 
# Thnaks to Wolfgang Kais for the following:
$xl.Range("B2").Select() | Out-Null
$xl.Selection.FormatConditions.Delete()
$xl.Selection.FormatConditions.Add(2, 0, "=COUNTIF(A:A,A2)>1") | Out-Null
$xl.Selection.FormatConditions.Item(1).Interior.ColorIndex = 8
$xl.Selection.Copy() | Out-Null
$xl.Columns.Item("B:B").Select() | Out-Null
$xl.Range("B2").Activate() | Out-Null
$xl.Selection.PasteSpecial(-4122, -4142, $false, $false) | Out-Null
$r = $ws1.UsedRange
$r2 = $ws1.Range("B2")
$r3 = $ws1.Range("E2")
$r4 = $ws1.Range("C2") 
$a = $r.Sort($r2,$xlDescending,$r3,$null,$xlAscending, `
$r4,$xlAscending,$xlYes) 
$ws1.Application.ActiveWindow.FreezePanes=$true
[void]$ws1.cells.entireColumn.Autofit()
$s = $xl.Range("A1").EntireColumn
$s.Hidden = $true
$ws1.Range("B2").Select() | Out-Null
}
 
# ==============================================================================================
# How do I show all 57 Colors in Excel?
 
$xl = new-object -comobject excel.application
$xl.Visible = $true
$xl.DisplayAlerts = $False
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$row = 1
$i = 1
For($i = 1; $i -lt 57; $i++){
$ws.Cells.Item($row, 1) = "'$'ws.Cells.Item($row, 2).Font.ColorIndex = " + $row
$ws.Cells.Item($row, 2).Font.ColorIndex = $row
$ws.Cells.Item($row, 2) = "test " + $row
$row++
}
[void]$ws.cells.entireColumn.Autofit() 
  
# ==============================================================================================
# How do I insert a Hyperlink into an Excel Spreadsheet? 
  
function Release-Ref ($info) {
    foreach ( $p in $args ) {
        ([System.Runtime.InteropServices.Marshal]::ReleaseComObject(
        [System.__ComObject]$p) -gt 0)
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
} # End Function
 
Function XLHyperlinks {
$link = "http://www.microsoft.com/technet/scriptcenter"
$xl = new-object -comobject excel.application
$xl.Visible = $true
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1) 
$ws.Cells.Item(1, 1).Value() = "Script Center"
$r = $ws.Range("A1")
$objLink = $ws.Hyperlinks.Add($r, $link)
$a = Release-Ref $r $ws $wb $xl
} # End Function
#
# Link to another sheet
 
$xl = New-Object -comobject Excel.Application
 
$xl.Visible = $True
 
$wb = $xl.Workbooks.Add()
 
$wb.Worksheets.Item(1).Hyperlinks.Add( `
  $wb.Worksheets.Item(1).Cells.Item(1,1) , `
  "" , "Sheet2!C4", "", "Link to sheet2")
 
# ________________________________________________________________________ 
#
# How do I sum a column in Excel?
#
 
Function xlSum {
$range = $ws1.usedRange
$row = $range.rows.count # Takes you to the last used row
$Sumrow = $row + 1
$r = $ws1.Range("A2:A$row") # select the column to Add up
$functions = $xl.WorkSheetfunction
$ws.cells.item($Sumrow,1) = $functions.sum($r) # this uses the Excel sum function
$rangeString = $r.address().tostring() -replace "\$",'' # convert formula to Text
$ws.cells.item($Sumrow,2) = "Sum $rangeString" # Print formula in Cell B & last row + 1
$ws1.cells.item($Sumrow,1).Select()
$ws1.range("a${Sumrow}:b$Sumrow").font.bold = "true" # seperate the : from the $ with {}
$ws1.range("a${Sumrow}:b$Sumrow").font.size=12 # Changes the font size to 12 points
[void]$range.entireColumn.Autofit()
} # End Function
 
# ________________________________________________________________________
#
# How do I SubTotal a column in an Excel Worksheet?
#
#
#  Sample Spreadsheet
#          mon    tue    wed
# eggs        1     1      1
# ham         5     5      5
# spam       1     4      7
# spam        2     5      8
# spam        3     6      9
#
# Code to sub total
$xlSum = -4157
$range = $xl.range("A1:D6")
$range.Subtotal(1,-4157,(2,3,4),$true,$False,$true) 
 
#     Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(4), _
#     Replace:=True, PageBreaks:=False, SummaryBelowData:=True
# ________________________________________________________________________
#
# In this example Subtotals are sums grouped by each change in field 1 "Salesperson"
# with Subtotals of field 2 "Amount".
# Details of Subtotal function
# SubTotal GroupBy =2, Function =XLSum, TotalList =Array(1),
# Replace =False, PageBreaks =False, SummaryBelowData =$True 
 
Function XLSubtotals {
$xlExpression = 2
$xlPasteFormats = -4122
$xlNone = -4142
$xlToRight = -4161
$xlToLeft = -4159
$xlDown = -4121
$xlShiftToRight = -4161
$xlFillDefault = 0
$xlSummaryAbove = 0
$xlSortValues = $xlPinYin = 1
$xlAscending = 1
$xlDescending = 2
$xlYes = 1
$xlTopToBottom = 1
$xlPasteValues = -4163          # Values only, not formulas
$xlCellTypeLastCell = 11        # to find last used cell
$xlSum = -4157 
$xl = New-Object -comobject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $False
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws = $wb.Sheets.Add()
$xl.Worksheets.item(1).name = "Detail"
$xl.Worksheets.item(2).name = "ShowLevels1"
$xl.Worksheets.item(3).name = "ShowLevels2"
$xl.Worksheets.item(4).name = "ShowLevels3" 
$ws1 = $wb.worksheets | where {$_.name -eq "Detail"}
$ws2 = $wb.worksheets | where {$_.name -eq "ShowLevels1"} #------- Selects sheet 6
$ws3 = $wb.worksheets | where {$_.name -eq "ShowLevels2"} #------- Selects sheet 5
$ws4 = $wb.worksheets | where {$_.name -eq "ShowLevels3"} #------- Selects sheet 4 
$ws1.Tab.ColorIndex = 8
$ws1.Tab.ColorIndex = 5
$ws2.Tab.ColorIndex = 6
$ws3.Tab.ColorIndex = 7 
$ws1.Cells.Item(1,2) = "Amount"
$ws1.Cells.Item(1,1) = "SalesPerson"
$ws1.Cells.Item(2,2) = 7324
$ws1.Cells.Item(2,1) = "Jack"
$ws1.Cells.Item(3,2) = 294
$ws1.Cells.Item(3,1) = "Elizabeth"
$ws1.Cells.Item(4,2) = 41472
$ws1.Cells.Item(4,1) = "Renee"
$ws1.Cells.Item(5,2) = 25406
$ws1.Cells.Item(5,1) = "Elizabeth"
$ws1.Cells.Item(6,2) = 20480
$ws1.Cells.Item(6,1) = "Jack"
$ws1.Cells.Item(7,2) = 11294
$ws1.Cells.Item(7,1)= "Renee"
$ws1.Cells.Item(8,2) = 982040
$ws1.Cells.Item(8,1) = "Elizabeth"
$ws1.Cells.Item(9,2) = 2622368
$ws1.Cells.Item(9,1) = "Jack"
$ws1.Cells.Item(10,2) = 884144
$ws1.Cells.Item(10,1) = "Renee" 
$ws1.Range("B2").Select() | Out-Null
$ws1.Application.ActiveWindow.FreezePanes=$true
[void]$ws1.cells.entireColumn.Autofit() 
$ws1.Range("A1").Select() | Out-Null
$r = $ws.Range("A2:A10")
$r2 = $ws.Range("A2") # Sorts on Column B and leaves Header alone
$a = $r.sort($r2, $xlAscending) 
$range1 = $ws1.Range("A1:B1").EntireColumn
$Range1.Select() | Out-Null
#$ws.Range.SpecialCells(11)).Select()
$Range1.Copy()
$ws1.Range("A1").Select() | Out-Null 
$ws2.Select() | Out-Null
[void]$ws2.Range("A1").PasteSpecial(-4163)
$ws3.Select() | Out-Null
[void]$ws3.Range("A1").PasteSpecial(-4163)
$ws4.Select() | Out-Null
[void]$ws4.Range("A1").PasteSpecial(-4163) 
$ws2.Select() | Out-Null
$range = $xl.range("A1:B10")
$range.Subtotal(1,-4157,(2),$true,$False,$true)
$ws2.Outline.ShowLevels(1)
[void]$ws1.cells.entireColumn.Autofit()
$ws2.Range("A1").Select() | Out-Null 
$ws3.Select() | Out-Null
$range = $xl.range("A1:B10")
$range.Subtotal(1,-4157,(2),$true,$False,$true)
$ws3.Outline.ShowLevels(2)
[void]$ws1.cells.entireColumn.Autofit()
$ws3.Range("A1").Select() | Out-Null 
$ws4.Select() | Out-Null
$range = $xl.range("A1:B10")
$range.Subtotal(1,-4157,(2),$true,$False,$true)
$ws4.Outline.ShowLevels(3) 
$ws1.Select() | Out-Null
[void]$ws1.cells.entireColumn.Autofit()
$ws1.Range("A1").Select() | Out-Null
} # End Function
 
#---------------------------------------------------
#
#How do I set up Auto Filters in Excel?
#This function sets up a spreadsheet and then sets Auto filters 
  
Function XLAutoFilter {
$xlTop10Items = 3
$xlTop10Percent = 5
$xlBottom10Percent = 6
$xlBottom10Items = 4
$xlAnd = 1
$xlOr = 2
$xlNormal = -4143
$xlPasteValues = -4163          # Values only, not formulas
$xlCellTypeLastCell = 11        # to find last used cell 
$xl = New-Object -comobject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $False
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws = $wb.Sheets.Add()
$ws = $wb.Sheets.Add() 
$ws1 = $wb.worksheets | where {$_.name -eq "Sheet1"}        #------- Selects sheet 1
$ws2 = $wb.worksheets | where {$_.name -eq "Sheet2"}         #------- Selects sheet 2
$ws3 = $wb.worksheets | where {$_.name -eq "Sheet3"}         #------ Selects sheet 3
$ws4 = $wb.worksheets | where {$_.name -eq "Sheet4"}         #------- Selects sheet 4
$ws5 = $wb.worksheets | where {$_.name -eq "Sheet5"}        #------- Selects sheet 5 
$ws1.Tab.ColorIndex = 8
$ws2.Tab.ColorIndex = 7
$ws3.Tab.ColorIndex = 6
$ws4.Tab.ColorIndex = 5
$ws5.Tab.ColorIndex = 4 
$ws1.name = "Detail"
$ws2.name = "JackOnly"
$ws3.name = "Top2"
$ws4.name = "LowestHighest"
$ws5.name = "Top25Percent" 
$ws1.cells.Item(1,1) =  "Amount"
$ws1.cells.Item(1,2) =  "SalesPerson"
$ws1.cells.Item(2,1) = 1
$ws1.cells.Item(2,2) = "Jack"
$ws1.cells.Item(3,1) = 2
$ws1.cells.Item(3,2) = "Elizabeth"
$ws1.cells.Item(4,1) = 3
$ws1.cells.Item(4,2) = "Renee"
$ws1.cells.Item(5,1) = 4
$ws1.cells.Item(5,2) = "Elizabeth"
$ws1.cells.Item(6,1) = 5
$ws1.cells.Item(6,2) = "Jack"
$ws1.cells.Item(7,1) = 6
$ws1.cells.Item(7,2) = "Renee"
$ws1.cells.Item(8,1) = 7
$ws1.cells.Item(8,2) = "Elizabeth"
$ws1.cells.Item(9,1) = 8
$ws1.cells.Item(9,2) = "Jack"
$ws1.cells.Item(10,1) = 9
$ws1.cells.Item(10,2) = "Renee"
$ws1.cells.Item(11,1) = 10
$ws1.cells.Item(11,2) = "Jack"
$ws1.cells.Item(12,1) = 11
$ws1.cells.Item(12,2) = "Jack"
$ws1.cells.Item(13,1) = 12
$ws1.cells.Item(13,2) = "Elizabeth"
$ws1.cells.Item(14,1) = 13
$ws1.cells.Item(14,2) = "Renee"
$ws1.cells.Item(15,1) = 14
$ws1.cells.Item(15,2) = "Elizabeth"
$ws1.cells.Item(16,1) = 15
$ws1.cells.Item(16,2) = "Jack"
$ws1.cells.Item(17,1) = 16
$ws1.cells.Item(17,2) = "Renee"
$ws1.cells.Item(18,1) = 17
$ws1.cells.Item(18,2) = "Elizabeth"
$ws1.cells.Item(19,1) = 18
$ws1.cells.Item(19,2) = "Jack"
$ws1.cells.Item(20,1) = 19
$ws1.cells.Item(20,2) = "Renee"
$ws1.cells.Item(21,1) = 20
$ws1.cells.Item(21,2) = "Renee" 
$used = $ws1.usedRange
$lastCell = $used.SpecialCells($xlCellTypeLastCell)
$lastrow = $lastCell.row 
$r = $ws1.Range("A1:B$lastrow")
$ws1.Range("A1:B$lastrow").Copy() 
$ws2.Select() | Out-Null
[void]$ws2.Range("A1").PasteSpecial(-4163)
$ws3.Select() | Out-Null
[void]$ws3.Range("A1").PasteSpecial(-4163)
$ws4.Select() | Out-Null
[void]$ws4.Range("A1").PasteSpecial(-4163)
$ws5.Select() | Out-Null
[void]$ws5.Range("A1").PasteSpecial(-4163)
#
$ws5.Range("A1").Select()
# AutoFilter structure - Field, Criteria, Operator
#$xl.Selection.AutoFilter 1, "10", $xlTop10Items        #top 10
$xl.Range("A1").Select() | Out-Null
$xl.Selection.AutoFilter(1, "2", $xlTop10Items)            #top 2
#$xl.Selection.AutoFilter 1, "10", $xlTop10Percent        #top 10 percent
#$xl.Selection.AutoFilter 1, "25", $$xlTop10Percent        #top 25 percent
#$xl.Selection.AutoFilter 1, "5", $xlBottom10Items        #Lowest 5 Items
#$xl.Selection.AutoFilter 1, "10", $$xlBottom10Percent    #Bottom 10 percent
#$xl.Selection.AutoFilter 1, ">15"                        #size greater 15
#$xl.Selection.AutoFilter 1, ">19",XLOr , "<2"            #Lowest and Highest
#$xl.Selection.AutoFilter 2, "Jack"                        #Jack items only
$ws5.cells.Item.EntireColumn.AutoFit 
$ws2.Select()
$ws2.Range("A1").Select()
# AutoFilter structure - Field, Criteria, Operator
#$xl.Selection.AutoFilter 1, "10", $xlTop10Items        #top 10
#$xl.Selection.AutoFilter 1, "2", $xlTop10Items            #top 2
#$xl.Selection.AutoFilter 1, "10", $xlTop10Percent        #top 10 percent
#$xl.Selection.AutoFilter 1, "25", $xlTop10Percent        #top 25 percent
#$xl.Selection.AutoFilter 1, "5", $xlBottom10Items        #Lowest 5 Items
#$xl.Selection.AutoFilter 1, "10", $xlBottom10Percent    #Bottom 10 percent
#$xl.Selection.AutoFilter 1, ">15"                        #size greater 15
#$xl.Selection.AutoFilter 1, ">19",XLOr , "<2"            #Lowest and Highest
$xl.Selection.AutoFilter(2, "Jack")                        #Jack items only
$ws2.cells.Item.EntireColumn.AutoFit 
$ws4.Select()
$ws4.Range("A1").Select()
# AutoFilter structure - Field, Criteria, Operator
#$xl.Selection.AutoFilter 1, "10", $xlTop10Items        #top 10
#$xl.Selection.AutoFilter 1, "2", $xlTop10Items            #top 2
#$xl.Selection.AutoFilter 1, "10", $xlTop10Percent        #top 10 percent
#$xl.Selection.AutoFilter 1, "25", $xlTop10Percent        #top 25 percent
#$xl.Selection.AutoFilter 1, "5", $xlBottom10Items        #Lowest 5 Items
#$xl.Selection.AutoFilter 1, "10", $xlBottom10Percent    #Bottom 10 percent
#$xl.Selection.AutoFilter 1, ">15"                        #size greater 15
$xl.Selection.AutoFilter(1, ">19",$xlOr , "<2")            #Lowest and Highest
#$xl.Selection.AutoFilter 2, "Jack"                        #Jack items only
$ws4.cells.Item.EntireColumn.AutoFit
# "Top25Percent"
$ws5.Select()
$ws5.Range("A1").Select()
# AutoFilter structure - Field, Criteria, Operator
#$xl.Selection.AutoFilter 1, "10", $xlTop10Items        #top 10
#$xl.Selection.AutoFilter 1, "2", $xlTop10Items            #top 2
#$xl.Selection.AutoFilter 1, "10", $xlTop10Percent        #top 10 percent
$xl.Range("A1").Select() | Out-Null
$xl.Selection.AutoFilter(1,"25",$xlTop10Percent)        #top 25 percent
#$xl.Selection.AutoFilter 1, "5", $xlBottom10Items        #Lowest 5 Items
#$xl.Selection.AutoFilter 1, "10", $xlBottom10Percent    #Bottom 10 percent
#$xl.Selection.AutoFilter 1, ">15"                        #size greater 15
#$xl.Selection.AutoFilter 1, ">19",XLOr , "<2"            #Lowest and Highest
#$xl.Selection.AutoFilter 2, "Jack"                        #Jack items only
$ws5.cells.Item.EntireColumn.AutoFit 
} # End Function
 
#---------------------------------------------------
#
 
#
# How do I set up a complex Formula in Excel?
# Here is one that creates a Complex Formula and executes it 
  
Function XLFormula1 {
$xl = New-Object -comobject excel.application
$xl.Visible = $true
$xl.DisplayAlerts = $False 
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1) 
$ws.name = "ComplexFormula"
$ws.Tab.ColorIndex = 9
$row = 2
$lastrow = $row
$Col = 3
$Off = $Col - 1
$ws.Cells.Item(1, 1) = "FileName"
$ws.Cells.Item(1, 2) = "Folder"
$ws.Cells.Item(1, 3) = "FullPath"
$ws.Cells.Item(2, 3) = "c:\Folder1\FunctionFolder1\FunctionFolder2\File1.txt"
$ws.Cells.Item(3, 3) = "c:\Folder1\FunctionFolder1\FunctionFolder2\FunctionFolder3\File2.txt"
$ws.Cells.Item(4, 3) = "c:\Folder1\FunctionFolder1\FunctionFolder2\FunctionFolder3\FunctionFolder4\File3.txt"
$lastrow = 4 
#Filename
$Range1 = $ws.Range("A2")
$ws.Cells.Item(2,1).FormulaR1C1 = "=MID(C[2],FIND(CHAR(127),Substitute(C[2],""\"",CHAR(127),LEN(C[2])-LEN(Substitute(C[2],""\"",""""))))+1,254)"
# Note I used 254 as a hardcoded length.  No filename should ever reach this length.  However….
# You could get the length programatically with the following
# "=MID(C[2],FIND(CHAR(127),FunctionSubsTITUTE(C[2],""\"",CHAR(127),LEN(C[2])-LEN(Substitute(C[2],""\"",""""))))+1,LEN(C[2])-(FIND(CHAR(127),Substitute(C[2],""\"",CHAR(127),LEN(C[2])-LEN(Substitute(C[2],""\"",""""))))))" 
$range2 = $ws.Range("A2:A$lastrow")
[void]$range1.AutoFill($range2,$xlFillDefault) 
#Folder
$lastrow = 4
$ws.Range("B2").Select() | Out-Null
$ws.Cells.Item(2,2).FormulaR1C1 = "=LEFT(C[1],FIND(CHAR(127),Substitute(C[1],""\"",CHAR(127),LEN(C[1])-LEN(Substitute(C[1],""\"",""""))))-1)"
$R1 = $ws.Range("B2")
$r2 = $ws.Range("B2:B$lastrow")
[void]$R1.AutoFill($R2,$xlFillDefault)
$ws.Cells.Item.EntireColumn.AutoFit
$ws.Range("A2").Select() | Out-Null
} # End Function
 
#---------------------------------------------------
#
 
# How do I set up a Pivot Table in Excel?
# I was working with MBSA reports here.
#
Function Pivot {
$xlPivotTableVersion12     = 3
$xlPivotTableVersion10     = 1
$xlCount                 = -4112
$xlDescending             = 2
$xlDatabase                = 1
$xlHidden                  = 0
$xlRowField                = 1
$xlColumnField             = 2
$xlPageField               = 3
$xlDataField               = 4    
# R1C1 means Row 1 Column 1 or "A1"
# R65536C5 means Row 65536 Column E or "E65536"
$PivotTable = $wb.PivotCaches().Create($xlDatabase,"Report!R1C1:R65536C5",$xlPivotTableVersion10)
$PivotTable.CreatePivotTable("Pivot!R1C1") | Out-Null
[void]$ws3.Select()
$ws3.Cells.Item(1,1).Select()
$wb.ShowPivotTableFieldList = $true
$PivotFields = $ws3.PivotTables("PivotTable1").PivotFields("Server") # Worksheet Name is Server
$PivotFields.Orientation = $xlRowField
$PivotFields.Position = 1 
$PivotFields = $ws3.PivotTables("PivotTable1").PivotFields("KBID") # Column Header is KBID
$PivotFields.Orientation = $xlColumnField
$PivotFields.Position = 1 
$PivotFields = $ws3.PivotTables("PivotTable1").PivotFields("KBID") # The data is in this column
$PivotFields.Orientation=$xlDataField
$PivotFields.Caption = "Count of KB's"
$PivotFields.Function = $xlCount
 
$mainRng = $ws3.UsedRange.Cells 
$RowCount = $mainRng.Rows.Count  
$R = $RowCount
$R = $R - 1    
$mainRng.Select()
$objSearch = $mainRng.Find("Grand Total")
$objSearch.Select()
$C = $objSearch.Column
Write-Host $C $R # this is just so I can see what's happining 
$xlSummaryAbove = 0 
$xlSortValues = $xlPinYin = 1
$xlAscending = 1 
$xlDescending = 2 
$range1 = $ws3.UsedRange 
$range2 = $ws3.Cells.Item(3, $C) 
# one-column sort
[void]$range2.sort($range2, $xlDescending) # puts the highest numbers at the top
} # End Function