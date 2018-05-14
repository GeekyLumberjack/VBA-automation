Dim findworkstations As Collection
Set findworkstations = New Collection
findworkstations.Add "ANYTHING YOU NEED TO FIND THAT'S A WORKSTATION"
Dim findprinters As Collection
Set findprinters = New Collection
findprinters.Add "ANYTHING YOU NEED TO FIND THAT'S A PRINTER"
Dim findservers As Collection
Set findservers = New Collection
findservers.Add "ANYTHING YOU NEED TO FIND THAT'S A SERVER"

Sheets("Blanks").Activate
Cells(1, 4).Activate
If IsEmpty(ActiveCell.Value) = True Then    
  ActiveCell.Value = "OS"
End If
Cells(1, 1).Select
Selection.AutoFilter
ActiveSheet.Range("$A$1:$AK$15000").AutoFilter Field:=5, Criteria1:="<>"
Range("A2:E10000").Select
Selection.Copy
Sheets("NETWORKING").Select
Cells(1, 1).Activate
While IsEmpty(ActiveCell.Value) = Flase
  ActiveCell.Offset(1, 0).Activate
Wend
ActiveSheet.Paste
Application.CutCopyMode = False
Sheets("Blanks").Activate
Selection.EntireRow.Delete
Selection.AutoFilter
For Each Workstation In findworkstations
  Cells(1, 1).Select
  Selection.AutoFilter
  ActiveSheet.Range("$A$1:$AK$15000").AutoFilter Field:=2, Criteria1:=Workstation    
  If ActiveSheet.AutoFilter.FilterMode = True Then
    Range("A2:E10000").Select
    Selection.Copy        
    Sheets("WORKSTATIONS").Select
    Cells(1, 1).Activate
    While IsEmpty(ActiveCell.Value) = Flase
      ActiveCell.Offset(1, 0).Activate
    Wend
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("Blanks").Activate
    Selection.EntireRow.Delete
  End If
Selection.AutoFilter
Next
For Each Printer In findprinters
  Cells(1, 1).Select
  Selection.AutoFilter
  ActiveSheet.Range("$A$1:$AK$15000").AutoFilter Field:=2, Criteria1:=Printer
  If ActiveSheet.AutoFilter.FilterMode = True Then
    Range("A2:E10000").Select
    Selection.Copy
    Sheets("PRINTERS").Select
    Cells(1, 1).Activate
    While IsEmpty(ActiveCell.Value) = Flase
      ActiveCell.Offset(1, 0).Activate
    Wend
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("Blanks").Activate
    Selection.EntireRow.Delete
  End If
Selection.AutoFilter
Next
For Each Server In findservers
  Cells(1, 1).Select
  Selection.AutoFilter
  ActiveSheet.Range("$A$1:$AK$15000").AutoFilter Field:=2, Criteria1:=Server
  If ActiveSheet.AutoFilter.FilterMode = True Then
    Range("A2:E10000").Select
    Selection.Copy
    Sheets("SERVERS").Select
    Cells(1, 1).Activate
    While IsEmpty(ActiveCell.Value) = Flase
      ActiveCell.Offset(1, 0).Activate
    Wend
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("Blanks").Activate
    Selection.EntireRow.Delete
  End If
Selection.AutoFilter
Next
