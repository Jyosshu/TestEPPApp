using System;

namespace TestEppApp
{
    class Program
    {
        static void Main(string[] args)
        {
            //ProductPriceHistory.CreateProductPriceHistory();

            ReadExistingWorkbook.ReadFileSaveAsNewFile("JoshTest.xlsx");
        }
    }
}
