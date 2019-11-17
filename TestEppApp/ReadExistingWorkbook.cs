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

                var endCell = FindLastCellofWorksheet(worksheet);

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
                            DateTime endDate = (tempDate2 != null && tempDate2.ToString() != string.Empty) ? DateTime.FromOADate(Convert.ToDouble(tempDate2)) : DateTime.Now;
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

        public static void ReadWorkbook(string filename, int worksheetToRead)
        {
            Utils.OutputDir = new DirectoryInfo($"{AppDomain.CurrentDomain.BaseDirectory}SampleApp");
            var file = Utils.GetFileInfo(filename, false);

            using (var excelPackage = new ExcelPackage(file))
            {
                var worksheet = excelPackage.Workbook.Worksheets[worksheetToRead];

                var endCell = FindLastCellofWorksheet(worksheet);

                // Build a Dictionary of the header fields in row 1.
                Dictionary<string, int> worksheetDict = new Dictionary<string, int>();
                for (int i = 1; i <= endCell.Column; i++)
                {
                    worksheetDict.Add(worksheet.Cells[1, i].Value.ToString(), i);
                }

                //foreach(KeyValuePair<string, int> pair in worksheetDict)
                //{
                //    Console.WriteLine($"Key: {pair.Key}, Value: {pair.Value}");
                //}

                for (int i = 1; i <= endCell.Row; i++)
                {
                    Console.WriteLine("{0,-40}{1,16}{2,10}", 
                        worksheet.Cells[i, worksheetDict["Product Name"]].Value, 
                        worksheet.Cells[i, worksheetDict["Product Number"]].Value, 
                        worksheet.Cells[i, worksheetDict["Current List Price"]].Value);
                }
                
            }
        }

        private static ExcelCellAddress FindLastCellofWorksheet(ExcelWorksheet excelWorksheet)
        {
            // Get the MAX Cell address in the passed worksheet to use as upper limit for iteration.
            ExcelCellAddress endCellAddress = excelWorksheet.Cells.End;
            int row = 1;
            int col = 1;
            int counter = 10; // This value could be set higher or lower based on a worksheet.  I set it equal to 10 to not waste any time on rows that probably were empty.

            // Find the last header cell with data
            for (int c = 1; c < endCellAddress.Column; c++)
            {
                if (excelWorksheet.Cells[row, c].Value != null && excelWorksheet.Cells[row, c].Value.ToString() != string.Empty)
                {
                    col = c;
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

            counter = 10; // reset counter
            // Based on the 
            for (int i = 1; i < endCellAddress.Row; i++)
            {

                // Checking to see if the first or last columns in the row have data.
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

            return new ExcelCellAddress(row, col);
        }
    }
}
