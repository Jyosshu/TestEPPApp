using System;
using OfficeOpenXml;
using Microsoft.Data.SqlClient;
using System.IO;
using System.Collections.Generic;

namespace TestEppApp
{
    public static class ExcelSpreadsheet
    {
        public static void CreateSpreadsheet(string filename, string worksheetName, string sqlQuery)
        {
            Utils.OutputDir = new DirectoryInfo($"{AppDomain.CurrentDomain.BaseDirectory}SampleApp");

            var file = Utils.GetFileInfo(filename);

            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(worksheetName);

                if (worksheet != null)
                {
                    const int headerRow = 1;
                    const int startRow = 2;
                    int row = startRow;
                    int totalCol = 0;
                    List<int> dateTimeColumns = new List<int>();


                    using (SqlConnection sqlConn = new SqlConnection(AppSettings.ConnectionString()))
                    {
                        sqlConn.Open();
                        using (SqlCommand sqlCmd = new SqlCommand(sqlQuery, sqlConn))
                        {
                            using (SqlDataReader sqlReader = sqlCmd.ExecuteReader())
                            {
                                while (sqlReader.Read())
                                {
                                    int col = 1;
                                    totalCol = sqlReader.FieldCount;

                                    for (int i = 0; i < totalCol; i++)
                                    {
                                        if (sqlReader.GetValue(i) != null)
                                        {
                                            if (row - 1 == headerRow)
                                            {
                                                worksheet.Cells[headerRow, col].Value = sqlReader.GetName(i);

                                                if (sqlReader.GetDataTypeName(i) == "datetime")
                                                    dateTimeColumns.Add(i + 1); // difference of index of sqlReader and worksheet
                                            }
                                            worksheet.Cells[row, col].Value = sqlReader.GetValue(i);
                                        }

                                        col++;
                                    }
                                    row++;
                                }
                                sqlReader.Close();

                                // Setting Formating for Header row of the spreadsheet
                                worksheet.Cells[headerRow, 1, headerRow, totalCol].Style.Font.Bold = true;
                                worksheet.Cells[headerRow, 1, headerRow, totalCol].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

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
    }
}

