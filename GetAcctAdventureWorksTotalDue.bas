Attribute VB_Name = "GetAcctAdventureWorksTotalDue"
Option Compare Database

'author Robleh Wais
Function GetAccountNumberByDueDate()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim startDate As String
    Dim endDate As String
    Dim percentage As String
    Dim storeName As String
    Dim strSQL As String

    ' Prompt the user for Start Date, End Date, Percentage, and Store Name
    startDate = InputBox("Enter Start Date (format: YYYY-MM-DD)")
    endDate = InputBox("Enter End Date (format: YYYY-MM-DD)")
    percentage = InputBox("Enter Percentage Threshold (e.g., 5 for 5%)")
    storeName = InputBox("Enter Store Name (leave blank for all stores)")

    ' Build the SQL string dynamically for the new stored procedure
    ' If a parameter is blank, pass NULL to the stored procedure
    strSQL = "EXEC [dbo].[GetAcctNumberAmountByDueDateNoTies] " & _
             "@StartDate = " & IIf(startDate = "", "NULL", "'" & startDate & "'") & ", " & _
             "@EndDate = " & IIf(endDate = "", "NULL", "'" & endDate & "'") & ", " & _
             "@Percentage = " & IIf(percentage = "", "NULL", percentage) & ", " & _
             "@StoreName = " & IIf(storeName = "", "NULL", "'" & storeName & "'") & ";"

    ' Reference to the current database
    Set db = CurrentDb()

    ' Use existing pass-through query, dynamically adjust the SQL
    Set qdf = db.QueryDefs("GetAcctNumberAmountByDueDateNoTies") ' Ensure this matches your actual pass-through query name
    qdf.SQL = strSQL

    ' Execute the pass-through query and return the results
    qdf.ReturnsRecords = True ' Ensure that the query will get records
    DoCmd.OpenQuery "GetAcctNumberAmountByDueDateNoTies" ' Name must match the qdf
    
    DoCmd.Close acQuery, "GetAcctNumberAmountByDueDateNoTies" 'close pass-thru query
    
    DoCmd.SetWarnings False 'turn off make table warnings
    
    DoCmd.OpenQuery "TestMKT_toBeModOrDel"
    
    DoCmd.Close acQuery, "TestMKT_toBeModOrDel"

       
        ' Clean up
    Set qdf = Nothing
    Set db = Nothing
End Function



Private Sub Command6_Click()
    Call GetAccountNumberByDueDate
End Sub

