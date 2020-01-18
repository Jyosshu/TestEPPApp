using System;

namespace TestEppApp
{
    class Program
    {
        static void Main(string[] args)
        {
            //ExcelSpreadsheet.CreateSpreadsheet("JoshTest.xlsx", "Products Price History", productPriceHistoryQuery);

            //ExcelSpreadsheet.CreateSpreadsheet("SalesOrders.xlsx", "Sales Orders 2014", salesOrderQuery);

            ReadExistingWorkbook.ReadFileSaveAsNewFile("JoshTest.xlsx");

            //ReadExistingWorkbook.ReadWorkbook("JoshTest.xlsx", 0);
        }

        static readonly string productPriceHistoryQuery = @"Select prod.Name AS [Product Name]
, prod.ProductNumber AS [Product Number]
, prod.Color AS [Color]
, prod.StandardCost AS [Cost]
, prod.ListPrice AS [Current List Price]
, ISNULL(prod.Size, 'N/A') AS [Size]
, ISNULL(prod.SizeUnitMeasureCode, 'N/A') AS [Size UoM]
, listHis.StartDate AS [Start Date]
, listHis.EndDate AS [End Date]
, listHis.ListPrice [List Price]
From Production.Product prod
Inner Join Production.ProductListPriceHistory listHis On listHis.ProductID = prod.ProductID
Order By prod.ProductID, listHis.StartDate Asc";


        static readonly string salesOrderQuery = @"SELECT
h.SalesOrderID
, h.OrderDate
, h.SalesOrderNumber
, h.PurchaseOrderNumber
, h.SubTotal
, h.TaxAmt
, h.Freight
, TotalDue
, l.OrderQty
, l.UnitPrice
, p.ProductNumber
, p.Name [ProductName]
, a.AddressLine1
, a.AddressLine2
, a.City
, sp.StateProvinceCode
, a.PostalCode
, sp.CountryRegionCode
FROM Sales.SalesOrderHeader h
INNER JOIN Sales.SalesOrderDetail l ON l.SalesOrderID = h.SalesOrderID
INNER JOIN Person.Address a ON a.AddressID = h.ShipToAddressID
INNER JOIN Person.StateProvince sp ON sp.StateProvinceID = a.StateProvinceID
LEFT OUTER JOIN Production.Product p ON p.ProductID = l.ProductID
WHERE h.OrderDate > '2014-01-01'
AND sp.CountryRegionCode = 'US'";
    }
}
