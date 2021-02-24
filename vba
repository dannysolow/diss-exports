Sub DISS_Reformat()

ActiveSheet.Cells.Unmerge
ActiveSheet.Columns("A:Z").AutoFit

Range("I6").Value = "ACCESS"
Range("M6").Value = "ELIGIBILITY"
Range("O6").Value = "ELIGIBILITY DATE"
Range("Q6").Value = "INVESTIGATION TYPE"
Range("Q6").Value = "INVESTIGATION TYPE"
Range("R6").Value = "INVESTIGATION DATE"
Range("T6").Value = "CE STATUS"
Range("X6").Value = "DATE ENROLLED"

ActiveSheet.Range("1:5").EntireRow.Delete
ActiveSheet.Range("B:B,F:G,J:L,N:N,P:P,S:S,U:W,Y:Y").EntireColumn.Delete

End Sub
