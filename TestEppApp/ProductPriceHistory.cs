using System;
using OfficeOpenXml;
using Microsoft.Data.SqlClient;
using System.IO;
using System.Collections.Generic;

namespace TestEppApp
{
    public static class ProductPriceHistory
    {
        public static void CreateProductPriceHistory(string worksheetName)
        {
            Utils.OutputDir = new DirectoryInfo($"{AppDomain.CurrentDomain.BaseDirectory}SampleApp");

            var file = Utils.GetFileInfo("JoshTest.xlsx");

            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(worksheetName);

                if (worksheet != null)
                {
                    const int headerRow = 1;
                    const int startRow = 2;
                    int row = startRow;
                    int totCol = 0;
                    List<int> dateTimeColumns = new List<int>();


                    using (SqlConnection sqlConn = new SqlConnection(AppSettings.ConnectionString()))
                    {
                        sqlConn.Open();
                        using (SqlCommand sqlCmd = new SqlCommand(query, sqlConn))
                        {
                            using (SqlDataReader sqlReader = sqlCmd.ExecuteReader())
                            {
                                while (sqlReader.Read())
                                {
                                    int col = 1;
                                    totCol = sqlReader.FieldCount;

                                    for (int i = 0; i < totCol; i++)
                                    {
                                        if (sqlReader.GetValue(i) != null)
                                        {
                                            if (row - 2 == headerRow)
                                            {
                                                worksheet.Cells[headerRow, col].Value = sqlReader.GetName(i);

                                                if (sqlReader.GetDataTypeName(i) == "datetime")
                                                    dateTimeColumns.Add(i + 1);
                                            }
                                            worksheet.Cells[row, col].Value = sqlReader.GetValue(i);
                                        }

                                        col++;
                                    }
                                    row++;
                                }
                                sqlReader.Close();

                                worksheet.Cells[headerRow, 1, headerRow, totCol].Style.Font.Bold = true;
                                worksheet.Cells[headerRow, 1, headerRow, totCol].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                foreach (int i in dateTimeColumns)
                                {
                                    worksheet.Cells[startRow, i, row - 1, i].Style.Numberformat.Format = "YYYY-MM-DD";
                                }
                            }
                        }
                        sqlConn.Close();
                    }
                }

                excelPackage.Workbook.Properties.Title = worksheetName;
                excelPackage.Workbook.Properties.Author = "TestEppApp";

                excelPackage.Save();
            }
        }

        static readonly string query = @"Select prod.Name AS [Product Name]
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
    }
}

