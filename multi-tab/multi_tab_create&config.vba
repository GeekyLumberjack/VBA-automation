'' This script uses a function called CollectionContains that is defined in a different file. 
'' In order to use this script you will need to get the file called FunctionCollectionConatins located in VBA-automation repository.
Dim iavalist As Collection
Set iavalist = New Collection
Dim noncompliant As Integer
Cells(2, 1).Activate
While IsEmpty(ActiveCell.Value) = False
  If CollectionContains(iavalist, ActiveCell.Value) Then
    ActiveCell.Offset(1, 0).Activate
  Else
    iavalist.Add (ActiveCell.Value)
    ActiveCell.Offset(1, 0).Activate
  End If
Wend
For Each iava In iavalist
  Sheets.Add.Name = iava
  Sheets("RAW DATA").Activate
  Cells(1, 1).Activate
  Selection.AutoFilter
  ActiveSheet.Range("$A$1:$AK$15000").AutoFilter Field:=1, Criteria1:=iava
  Cells(1, 1).Select
  ActiveCell.EntireRow.Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Copy
  Sheets(iava).Activate
  Cells(1, 1).Activate
  ActiveSheet.Paste
  Application.CutCopyMode = False
Next
Sheets.Add.Name = "SUMMARY"
Cells(1, 1).Activate
ActiveCell.Value = "IAVM"
Cells(1, 2).Activate
ActiveCell.Value = "NIPR Affected"
Cells(1, 3).Activate
ActiveCell.Value = "NIPR Non Compliant"
Cells(1, 4).Activate
ActiveCell.Value = "NIPR Compliant"
Cells(1, 5).Activate
ActiveCell.Value = "% NIPR Compliant"
Cells(2, 1).Activate
For Each iavm In iavalist
  noncompliant = Worksheets(iavm).Range("A2:A15000").Cells.SpecialCells(xlCellTypeConstants).Count
  ActiveCell.Value = iavm
  ActiveCell.Offset(0, 1).Activate
  ActiveCell.Value = "19337"
  ActiveCell.Offset(0, 1).Activate
  ActiveCell.Value = noncompliant
  ActiveCell.Offset(1, -2).Activate
Next
Range("D2").Select
ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-1]"
Range("D2").Select
Selection.AutoFill Destination:=Range("D2:D200")
Range("D2:D200").Select
Selection.CopySelection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _     :=False, Transpose:=False 
Range("E2").Select
ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-3]"
Range("E2").Select
Selection.AutoFill Destination:=Range("E2:E200")
Range("E2:E200").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _    :=False, Transpose:=False
