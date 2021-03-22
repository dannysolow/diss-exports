Sub DISS_Reformat()

ActiveSheet.Cells.Unmerge
ActiveSheet.Columns("A:Z").AutoFit

Range("I5").Value = "ACCESS"
Range("K5").Value = "ELIGIBILITY"
Range("L5").Value = "ELIGIBILITY DATE"
Range("M5").Value = "INVESTIGATION TYPE"
Range("O5").Value = "INVESTIGATION DATE"
Range("P5").Value = "CE STATUS"
Range("R5").Value = "DATE ENROLLED"

ActiveSheet.Range("1:4").EntireRow.Delete
ActiveSheet.Range("B:B,F:G,J:J,N:N,Q:Q,S:S,U:U").EntireColumn.Delete

lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
Rows(lastrow).Delete

lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
Rows(lastrow).Delete

lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
Rows(lastrow).Delete

lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
Rows(lastrow).Delete

End Sub
