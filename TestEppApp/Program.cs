using System;

namespace TestEppApp
{
    class Program
    {
        static void Main(string[] args)
        {
            ProductPriceHistory.CreateProductPriceHistory("Products Price History");

            //ReadExistingWorkbook.ReadFileSaveAsNewFile("JoshTest.xlsx");

            ReadExistingWorkbook.ReadWorkbook("JoshTest.xlsx", 0);
        }
    }
}
