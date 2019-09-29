using System;
using System.Collections.Generic;
using System.Text;

namespace TestEppApp
{
    public static class AppSettings
    {
        public static string DefaultConnection
        {
            get => "Server=.\\SQLEXPRESS;Database=AdventureWorks2016CTP3;Trusted_Connection=True;MultipleActiveResultSets=true;";
        }

        
    }
}
