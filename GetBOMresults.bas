Attribute VB_Name = "GetBOMresults"
Option Compare Database

Function GetBOMResults()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim productID As Integer
    Dim checkDate As String
    Dim strSQL As String

    ' Prompt the user for Product ID and Check Date
    productID = InputBox("Enter Start Product ID")
    checkDate = InputBox("Enter Check Date (format: YYYY-MM-DD)")

    ' Build the SQL string dynamically for the stored procedure
    strSQL = "EXEC [dbo].[uspGetBillOfMaterials1] @StartProductID = " & productID & ", " & _
             "@CheckDate = " & IIf(checkDate = "", "NULL", "'" & checkDate & "'") & ";"

    ' Reference to the current database
    Set db = CurrentDb()

    ' Use existing pass-through query, dynamically adjust the SQL
    Set qdf = db.QueryDefs("uspGetBillOfMaterials1SP") ' Replace with your actual pass-through query name
    qdf.SQL = strSQL

    ' Execute the pass-through query and return the results
    qdf.ReturnsRecords = True ' Ensure the query returns records
    DoCmd.OpenQuery "uspGetBillOfMaterials1SP" ' Ensure this matches the name of your pass-through query

    ' Clean up
    Set qdf = Nothing
    Set db = Nothing
End Function


Private Sub Command6_Click()
    Call GetBOMResults
End Sub

