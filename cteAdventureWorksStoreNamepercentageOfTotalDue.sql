USE AdventureWorks2022
GO

ALTER PROCEDURE dbo.GetAcctNumberAmountByDueDate
    @StartDate DATE, -- Start date parameter
    @EndDate DATE    -- End date parameter
AS
BEGIN
    -- CTE to calculate individual store percentages
    WITH StoreTotals AS (
        SELECT 
            Sales.Store.Name AS StoreName,
            SUM(Sales.SalesOrderHeader.TotalDue) AS StoreTotalDue,
            SUM(SUM(Sales.SalesOrderHeader.TotalDue)) OVER () AS GrandTotalDue
        FROM Sales.SalesOrderHeader
        INNER JOIN Sales.Customer 
            ON Sales.SalesOrderHeader.CustomerID = Sales.Customer.CustomerID
        INNER JOIN Sales.Store 
            ON Sales.Customer.StoreID = Sales.Store.BusinessEntityID
        WHERE Sales.SalesOrderHeader.DueDate BETWEEN @StartDate AND @EndDate
        GROUP BY Sales.Store.Name
    )
    SELECT 
        StoreName,
        StoreTotalDue,
        -- Calculate percentage and scale so the total is exactly 100%
        CAST(ROUND((StoreTotalDue * 100.0) / GrandTotalDue, 2) AS NUMERIC(5, 2)) AS StoreNamePercentage
    FROM StoreTotals
    ORDER BY StoreNamePercentage DESC, StoreName ASC;
END
GO
