Sub StatementTabling()

    With ActiveSheet
    
        Dim AccountName As String: AccountName = Cells(11, "A").Value
        Dim ws As Worksheet
        Dim cell As Range
        
        
        'Delete formatting data
        Rows("1:19").Delete
        FirstRow = .UsedRange.Cells(1).Row
        LastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        For CurrentRow = FirstRow To LastRow Step 1
            If Cells(CurrentRow, "A") Like "Account Number *" Then Rows(CurrentRow).Delete
            If Cells(CurrentRow, "A") Like "Statement Period *" Then Rows(CurrentRow).Delete
            If Cells(CurrentRow, "A") Like "Statement No. *" Then Rows(CurrentRow).Delete
            If Cells(CurrentRow, "A") Like "Transaction Details *" Then Rows(CurrentRow).Delete
            If Cells(CurrentRow, "A") Like "Date Transaction *" Then Rows(CurrentRow).Delete
            If Cells(CurrentRow, "A") Like "SUB TOTAL CARRIED *" Then Rows(CurrentRow).Delete
            If Cells(CurrentRow, "A") Like "##*CLOSING BALANCE*" Then
                Rows(CurrentRow + 1 & ":" & LastRow).Delete
                Exit For
            End If
        Next CurrentRow
                
        'Write the column names
        Cells(1, "A") = "Original"
        Cells(1, "B") = "Date"
        Cells(1, "C") = "Description"
        Cells(1, "D") = "Debit"
        Cells(1, "E") = "Credit"
        Cells(1, "F") = "Balance"
        Cells(1, "G") = "AccountNumber"
        
        
        'Fixes the spacing of the May dates
        FirstRow = .UsedRange.Cells(1).Row
        LastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        
        For CurrentRow = FirstRow + 1 To LastRow Step 1
            If Cells(CurrentRow, "A") Like "##MAY *" Then _
                'Only leaves the May part of the date
                Cells(CurrentRow, "A") = (Left(Cells(CurrentRow, "A"), 2) & " " & Right(Cells(CurrentRow, "A"), (Len(Cells(CurrentRow, "A")) - 2)))
            End If
        Next CurrentRow
        
        
        'Write the date
        FirstRow = .UsedRange.Cells(1).Row
        LastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        
        For CurrentRow = FirstRow + 1 To LastRow Step 1
            If Cells(CurrentRow, "A") Like "## ??? *" Then Cells(CurrentRow, "B") = Left(Cells(CurrentRow, "A"), 6)
        Next CurrentRow
        
'
'        'Write the account name as the rightmost column
'        FirstRow = .UsedRange.Cells(1).Row
'        LastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
'        For CurrentRow = FirstRow To LastRow Step 1
'            Cells(CurrentRow, "G") = AccountName
'        Next CurrentRow
    
    End With

End Sub
