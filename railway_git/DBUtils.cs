using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace test_railway
{
    class DBUtils
    {
        public static SqlConnection GetDBConnection()
        {
            string datasource = @".\SQLEXPRESS";

            string database = "railway";
            string username = "master";
            string password = "1234";

            return DBService.GetDBConnection(datasource, database, username, password);
        }
    }
}