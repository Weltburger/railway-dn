using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace test_railway
{
    static class GlobalData
    {
        public static SqlConnection conn;
        public static SqlCommand cmd;
        public static string sql;
        public static Excel.Application excelApp;
        public static Excel.Workbook workBook;
    }
}
