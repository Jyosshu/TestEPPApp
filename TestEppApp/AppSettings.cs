using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace TestEppApp
{
    public static class AppSettings
    {
        private static string DefaultConnection
        {
            get => "Server=.\\SQLEXPRESS;Database=AdventureWorks2016CTP3;Trusted_Connection=True;MultipleActiveResultSets=true;";
        }

        private static string RemoteConnection
        {
            get => "Server=JYOSSHU\\SQLEXPRESS,49172;Database=AdventureWorks2016CTP3;User Id=jwygle_macos;Password=dFSYrnIWjjteW8sQeaCe;";
        }

        private static bool isWindows = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);

        public static string ConnectionString()
        {
            if (isWindows == true)
            {
                return DefaultConnection;
            }
            else
            {
                return RemoteConnection;
            }
        }
    }
}
