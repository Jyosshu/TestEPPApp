using System;
using OfficeOpenXml;
using System.Data.SqlClient;
using System.IO;

namespace TestEppApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Utils.OutputDir = new DirectoryInfo($"{AppDomain.CurrentDomain.BaseDirectory}SampleApp");

            var file = Utils.GetFileInfo("JoshTest.xlsx");

            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Products Price History");

                if (worksheet != null)
                {
                    const int headerRow = 1;
                    const int startRow = 3;
                    int row = startRow;
                    int totCol = 0;

                    //worksheet.Cells["A1"].Value = "Product Name";
                    //worksheet.Cells["B1"].Value = "Product Number";
                    //worksheet.Cells["C1"].Value = "Color";
                    //worksheet.Cells["D1"].Value = "Standard Cost";
                    //worksheet.Cells["E1"].Value = "Current List Price";
                    //worksheet.Cells["F1"].Value = "Size";
                    //worksheet.Cells["G1"].Value = "Size UoM";
                    //worksheet.Cells["H1"].Value = "Start Date";
                    //worksheet.Cells["I1"].Value = "End Date";
                    //worksheet.Cells["J1"].Value = "List Price";
                    //worksheet.Cells["A1:J1"].Style.Font.Bold = true;
                    //worksheet.Cells["A1:J1"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;


                    using (SqlConnection sqlConn = new SqlConnection(AppSettings.DefaultConnection))
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

                                    for (int i = 0; i < sqlReader.FieldCount; i++)
                                    {
                                        if (sqlReader.GetValue(i) != null)
                                        {
                                            if (row - 2 == headerRow)
                                            {
                                                worksheet.Cells[headerRow, col].Value = sqlReader.GetName(i);
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
                                worksheet.Cells[startRow, 8, row - 1, 8].Style.Numberformat.Format = "YYYY-MM-DD";
                                worksheet.Cells[startRow, 9, row - 1, 9].Style.Numberformat.Format = "YYYY-MM-DD";
                            }
                        }
                        sqlConn.Close();
                    }
                }

                excelPackage.Workbook.Properties.Title = "Josh Test - Product Cost History";
                excelPackage.Workbook.Properties.Author = "Josh Wygle";

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
