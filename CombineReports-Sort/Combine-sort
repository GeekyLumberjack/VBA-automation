Sheets("Blanks").Select    
Range("A1").Select    
Workbooks.Open Filename:= _        
  "143.71_PART1.csv"    
 Range("A1:N15000").Select    
 Selection.Copy    
 Windows("4DEC_143.71_CONSOLIDATED.xlsx").Activate
 ActiveSheet.Paste
 Application.CutCopyMode = False
 Range("A1").Select
 Windows("143.71_PART1.csv").Activate
 ActiveWindow.Close
 Workbooks.Open Filename:= _        
  "143.71_PART2.csv"
  Range("A1:N15000").Select
  Selection.Copy
  Windows("4DEC_143.71_CONSOLIDATED.xlsx").Activate
  ActiveSheet.Paste
  Application.CutCopyMode = False
  Range("A1").Select
  Selection.End(xlDown).Select
  ActiveCell.Offset(rowOffset:=1, columnOffset:=0).Activate
  Windows("143.71_PART2.csv").Activate
  ActiveWindow.Close
  Workbooks.Open Filename:= _        
    "143.71_PART3.csv"
  Range("A1:N15000").Select
  Selection.Copy
  Windows("4DEC_143.71_CONSOLIDATED.xlsx").Activate
  ActiveSheet.Paste
  Application.CutCopyMode = False
  Range("A1").Select
  Selection.End(xlDown).Select
  ActiveCell.Offset(rowOffset:=1, columnOffset:=0).Activate
  Windows("143.71_PART3.csv").Activate
  ActiveWindow.Close
  Workbooks.Open Filename:= _        
    "143.71_PART4.csv"
  Range("A1:N15000").Select
  Selection.Copy
  Windows("4DEC_143.71_CONSOLIDATED.xlsx").Activate
  ActiveSheet.Paste
  Application.CutCopyMode = False
  Range("A1").Select
  Selection.End(xlDown).Select
  ActiveCell.Offset(rowOffset:=1, columnOffset:=0).Activate
  Windows("143.71_PART4.csv").Activate
  ActiveWindow.Close
  Workbooks.Open Filename:= _        
    "143.71_PART5.csv"
   Range("A1:N15000").Select
   Selection.Copy
   Windows("4DEC_143.71_CONSOLIDATED.xlsx").Activate
   ActiveSheet.Paste
   Application.CutCopyMode = False
   Range("A1").Select
   Selection.End(xlDown).Select
   ActiveCell.Offset(rowOffset:=1, columnOffset:=0).Activate
   Windows("143.71_PART5.csv").Activate
   ActiveWindow.Close
   Workbooks.Open Filename:= _        
      "143.71_PART6.csv"    
   Range("A1:N15000").Select    
   Selection.Copy    
   Windows("4DEC_143.71_CONSOLIDATED.xlsx").Activate
   ActiveSheet.Paste
   Application.CutCopyMode = False
   Range("A1").Select    
   Selection.End(xlDown).Select
   ActiveCell.Offset(rowOffset:=1, columnOffset:=0).Activate
   Windows("143.71_PART6.csv").Activate
   ActiveWindow.Close
   Workbooks.Open Filename:= _        
     "143.71_PART7.csv"
   Range("A1:N15000").Select
   Selection.Copy
   Windows("4DEC_143.71_CONSOLIDATED.xlsx").Activate
   ActiveSheet.Paste
   Application.CutCopyMode = False
   Range("A1").Select
   Selection.End(xlDown).Select
   ActiveCell.Offset(rowOffset:=1, columnOffset:=0).Activate
   Windows("143.71_PART7.csv").Activate
   ActiveWindow.Close
   Workbooks.Open Filename:= _        
    "143.71_PART8.csv" 
   Range("A1:N15000").Select
   Selection.Copy
   Windows("4DEC_143.71_CONSOLIDATED.xlsx").Activate
   ActiveSheet.Paste
   Application.CutCopyMode = False
   Range("A1").Select
   Selection.End(xlDown).Select
   ActiveCell.Offset(rowOffset:=1, columnOffset:=0).Activate
   Windows("143.71_PART8.csv").Activate
   ActiveWindow.Close
   Range("A1").Select
   Selection.AutoFilter
   ActiveWindow.LargeScroll ToRight:=-1    ActiveSheet.Range("$A$1:$N$8881").AutoFilter Field:=1, Criteria1:= _        "IP Address"
   Rows("1282:1282").Select
   Range(Selection, Selection.End(xlDown)).Select
   Selection.Delete Shift:=xlUp
   Selection.AutoFilter
   Columns("D:K").Select
   Selection.Delete Shift:=xlToLeft
   Columns("E:F").Select
   Selection.Delete Shift:=xlToLeft
   Range("E1").Select
   ActiveCell.FormulaR1C1 = "NOTES"
   Range("E2").Select
   Workbooks.Open Filename:= _        "NETWORKING IPS_MASTER_27JUN2017_DO NOT DELETE.xlsx"
   Windows("4DEC_143.71_CONSOLIDATED.xlsx").Activate
   ActiveCell.FormulaR1C1 = _        "=VLOOKUP(RC[-4],'[NETWORKING IPS_MASTER_27JUN2017_DO NOT DELETE.xlsx]143.71'!C1:C5,5,FALSE)"
   Range("E2").Select
   Selection.AutoFill Destination:=Range("E2:E7595")
   Range("E2:E7595").Select
   Selection.Copy
   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _        :=False, Transpose:=False
   Range("E1").Select
   Application.CutCopyMode = False
   Columns("E:E").Select
   Selection.Replace What:="#N/A", Replacement:="", LookAt:=xlPart, _        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _        ReplaceFormat:=False
   'Workbooks.Open Filename:= _    
   '    "C:\DHCP\DHCP_Results.csv"    
   'Windows("4DEC_143.71_CONSOLIDATED.xlsx").Activate    
   'Range("B2").Select    
   'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],DHCP_Results.csv!C5:C6,2,FALSE)"    
   'Range("B2").Select    
   'Selection.AutoFill Destination:=Range("B2:B10000"), Type:=xlFillDefault    
   'Range("B2:B10000").Select    
   'Selection.Copy    
   'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _    
   '    :=False, Transpose:=False    
   Columns("B:B").Select    
   Selection.Replace What:="*\", Replacement:="", LookAt:=xlPart, _        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _        ReplaceFormat:=False    
   'Windows("DHCP_Results.csv").Activate    
   'ActiveWindow.Close    
   Range("A6").Select    
   Selection.AutoFilter   
   ActiveWindow.SmallScroll Down:=-6    
   Range("A2").Select    
   Windows("NETWORKING IPS_MASTER_27JUN2017_DO NOT DELETE.xlsx").Activate    
   ActiveWindow.Close    
   Cells(1, 1).Select   
   ActiveCell.EntireRow.Select    
   Selection.Copy    
   Sheets("NETWORKING").Activate    
   Cells(1, 1).Select    
   ActiveCell.EntireRow.Select    
   ActiveSheet.Paste   
   Sheets("WORKSTATIONS").Activate    
   Cells(1, 1).Select    
   ActiveCell.EntireRow.Select   
   ActiveSheet.Paste    
   Sheets("PRINTERS").Activate    
   Cells(1, 1).Select    
   ActiveCell.EntireRow.Select    
   ActiveSheet.Paste    
   Sheets("SERVERS").Activate   
   Cells(1, 1).Select    
   ActiveCell.EntireRow.Select   
   ActiveSheet.Paste    
   Sheets("MISC").Activate    
   Cells(1, 1).Select   
   ActiveCell.EntireRow.Select   
   ActiveSheet.Paste    
   Sheets("TDY").Activate    
   Cells(1, 1).Select   
   ActiveCell.EntireRow.Select    
   ActiveSheet.Paste    
   Sheets("CAMERA").Activate   
   Cells(1, 1).Select   
   ActiveCell.EntireRow.Select   
   ActiveSheet.Paste    
   Sheets("Blanks").Activate   
   Cells(1, 1).Select    
   Application.Run "'Combine-sort-DS.XLSB'!DS_Breakdown.DS_Breakdown"    
   Workbooks.Open Filename:= _        
    `C:\DHCP\DHCP_Results.csv"   
    Windows("4DEC_143.71_CONSOLIDATED.xlsx").Activate    
    Range("B2").Select    
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],DHCP_Results.csv!C5:C6,2,FALSE)"    
    Range("B2").Select    
    Selection.AutoFill Destination:=Range("B2:B10000"), Type:=xlFillDefault   
    Range("B2:B10000").Select    
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _        :=False, Transpose:=False    
    Columns("B:B").Select
    Selection.Replace What:="#N/A", Replacement:="", LookAt:=xlPart, _        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _        ReplaceFormat:=False    
    Windows("DHCP_Results.csv").Activate
    ActiveWindow.Close
    Application.Run "'Combine-sort-DS.XLSB'!DS_Breakdown.DS_Breakdown"
