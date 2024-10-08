USE AdventureWorks2022
GO

ALTER PROCEDURE dbo.GetAcctNumberAmountByDueDate
    @StartDate DATE, -- Start date parameter
    @EndDate DATE    -- End date parameter
AS
BEGIN
    SELECT 
        Sales.SalesOrderHeader.AccountNumber,  
        ROUND(Sales.SalesOrderHeader.SubTotal, 1) AS SubTotal, 
        ROUND(Sales.SalesOrderHeader.TaxAmt, 1) AS TaxAmt, 
        ROUND(Sales.SalesOrderHeader.Freight, 1) AS Freight,
        ROUND(Sales.SalesOrderHeader.TotalDue, 1) AS TotalDue,
        CONVERT(VARCHAR, Sales.SalesOrderHeader.DueDate, 101) AS DueDate,
        Sales.Store.Name AS StoreName,
        -- Calculate StoreNamePercentage
        ROUND((Sales.SalesOrderHeader.TotalDue * 100.0) / SUM(Sales.SalesOrderHeader.TotalDue) OVER (), 2) AS StoreNamePercentage
    FROM Sales.SalesOrderHeader
    INNER JOIN Sales.Customer 
        ON Sales.SalesOrderHeader.CustomerID = Sales.Customer.CustomerID
    INNER JOIN Sales.Store 
        ON Sales.Customer.StoreID = Sales.Store.BusinessEntityID
    WHERE Sales.SalesOrderHeader.DueDate BETWEEN @StartDate AND @EndDate
    ORDER BY Sales.SalesOrderHeader.DueDate ASC;
END
GO
