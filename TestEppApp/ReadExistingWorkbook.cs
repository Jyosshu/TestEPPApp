using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;
using System.IO;

namespace TestEppApp
{
    public static class ReadExistingWorkbook
    {
        public static void ReadFileSaveAsNewFile(string filename)
        {
            Utils.OutputDir = new DirectoryInfo($"{AppDomain.CurrentDomain.BaseDirectory}SampleApp");
            var file = Utils.GetFileInfo(filename, false);

            using (var excelPackage = new ExcelPackage(file))
            {
                var worksheet = excelPackage.Workbook.Worksheets[0];

                //int row = 1;
                int col = 11;

                var endCell = FindLastCell(worksheet);

                for (int i = 1; i < endCell.Row; i++)
                {
                    if (i == 1 && col == 11)
                    {
                        //wkst.Cells[row, col].Value = "Percentage Increase";
                        worksheet.Cells[i, col].Value = "Days At Price Point";
                        //worksheet.Cells[i, col].Style.Font.Bold = true;
                        //worksheet.Cells[i, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[i, col].StyleID = worksheet.Cells[i, col - 1].StyleID;
                    }
                    else if (i == 2)
                    {
                        continue;
                    }
                    else
                    {
                        if (worksheet.Cells[i, 1].Value != null && worksheet.Cells[i, col - 3].Value != null)
                        {
                            var tempDate = worksheet.Cells[i, col - 3].Value;
                            var tempDate2 = worksheet.Cells[i, col - 2].Value;

                            DateTime startDate = DateTime.FromOADate(Convert.ToDouble(tempDate));
                            DateTime endDate = (tempDate2 != null && tempDate2 != string.Empty) ? DateTime.FromOADate(Convert.ToDouble(tempDate2)) : DateTime.Now;
                            TimeSpan timeSpan = endDate - startDate;

                            worksheet.Cells[i, col].Value = timeSpan.Days;
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                

                excelPackage.SaveAs(Utils.GetFileInfo($"{filename}_{DateTime.Now.ToString("yyyy-MM-dd HHmmss")}.xlsx"));
            }                
        }

        private static ExcelCellAddress FindLastCell(ExcelWorksheet excelWorksheet)
        {
            ExcelCellAddress excelCellAddress = excelWorksheet.Cells.End;
            int maxRow = excelCellAddress.Row;
            int maxCol = excelCellAddress.Column;
            int row = 1;
            int col = 1;
            int counter = 10;
            ExcelCellAddress endCellAddress;

            for (int c = 1; c < maxCol; c++)
            {
                if (excelWorksheet.Cells[row, c].Value != null)
                {
                    col = c;
                }
            }

            for (int i = 1; i < maxRow; i++)
            {
                if (excelWorksheet.Cells[i, 1].Value != null || excelWorksheet.Cells[i, col].Value != null)
                {
                    row = i;
                }
                else
                {
                    counter--;

                    if (counter <= 0)
                    {
                        break;
                    }
                }
            }

            endCellAddress = new ExcelCellAddress(row, col);

            return endCellAddress;
        }
    }
}
