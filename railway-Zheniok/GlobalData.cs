using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
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
        public static Process excelProc;

        public static int stationSelected;
        public static string sql;

        public static Excel.Worksheet sheet;
        public static Excel.Range titleTable;
        public static Excel.Range cellsValuesTable;


        public static Excel.Application excelApp;
        public static Excel.Workbook workBook;

        // сохранение, закрытие и удаление процесса файла Excel
        public static void processClose()
        {
            workBook.Close(true); //сохраняем и закрываем файл
            excelApp.Quit();
            releaseObject(sheet);
            releaseObject(workBook);
            releaseObject(excelApp);
            excelProc.Kill();
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception)
            {
                obj = null;
                //MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
