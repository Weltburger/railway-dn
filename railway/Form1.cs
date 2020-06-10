using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Data.Common;
using Excel = Microsoft.Office.Interop.Excel;

namespace test_railway
{
    public partial class Form1 : Form
    {
        SqlConnection conn;
        Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            conn = DBUtils.GetDBConnection();
            conn.Open();

            if (conn.State == ConnectionState.Open)
            {
                MessageBox.Show("всьо чьотка!");
            }
            else
            {
                MessageBox.Show("ашыбачька...");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string sql = "select st.NAME as 'Станция (для заголовка)', " +
                "LIST_NO as 'Номер ПЛ (для заголовка)'," +
                "CAR_NO as 'Номер вагона', " +
                "BUILT_YEAR as 'Год постройки', " +
                "CAR_TYPE as 'Род вагона', " +
                "CAR_LOCATION as 'Дислокация', " +
                "ADM_CODE as 'Код страны-собств.', " +
                "[OWNER] as 'Собственник', " +
                "CASE " +
                    "WHEN IS_LOADED = 0 THEN 'негруженный' " +
                    "WHEN IS_LOADED = 1 THEN 'груженный' " +
                    "END as 'Состояние', " +
                "CASE " +
                    "WHEN IS_WORKING = 0 THEN 'нерабочий' " +
                    "WHEN IS_WORKING = 1 THEN 'рабочий' " +
                    "else 'не определено'" +
                    "END as 'Парк', " +
                "CASE " +
                    "WHEN NON_WORKING_STATE = 1 THEN 'неисправный' " +
                    "WHEN NON_WORKING_STATE = 2 THEN 'резерв' " +
                    "WHEN NON_WORKING_STATE = 3 THEN 'ДЛЗО' " +
                    "WHEN NON_WORKING_STATE = 4 THEN 'СТН' " +
                    "WHEN NON_WORKING_STATE = 5 THEN 'Поврежден по акту ВУ-25' " +
                    "else 'не определено'" +
                "END as 'Категория НРП'" +
                "from CAR_CENSUS_LISTS ccl " +
                "INNER JOIN STATIONS st ON st.ESR = ccl.LOCATION_ESR " +
                "WHERE st.ESR = 480403";

            // Создать объект Command.
            SqlCommand cmd = new SqlCommand();

            // Сочетать Command с Connection.
            cmd.Connection = conn;
            cmd.CommandText = sql;

            using (DbDataReader reader = cmd.ExecuteReader())
            {
                if (reader.HasRows)
                {
                    int idRec = 0;
                    while (reader.Read())
                    {
                        idRec++;
                        string station = reader.GetString(0);
                        int listNO = Convert.ToInt32(reader.GetValue(1));
                        long carNO = Convert.ToInt64(reader.GetValue(2));
                        int builtYear = Convert.ToInt32(reader.GetValue(3));
                        string carType = reader.GetString(4);
                        string carLoc = reader.GetString(5);
                        int admCode = Convert.ToInt32(reader.GetValue(6));
                        string owner = reader.GetString(7);
                        string isLoaded = reader.GetString(8);
                        string isWorking = reader.GetString(9);
                        string workState = reader.GetString(10);






                        //Индекс столбца Mng_Id в команде SQL.
                        //int mngIdIndex = reader.GetOrdinal("Mng_Id");
                        //long? mngId = null;
                        //if (!reader.IsDBNull(mngIdIndex))
                        //{
                        //    mngId = Convert.ToInt64(reader.GetValue(mngIdIndex));
                        //}
                        //Console.WriteLine("--------------------");
                        //Console.WriteLine("empIdIndex:" + empIdIndex);
                        //Console.WriteLine("EmpId:" + empId);
                        //Console.WriteLine("EmpNo:" + empNo);
                        //Console.WriteLine("EmpName:" + empName);
                        //Console.WriteLine("MngId:" + mngId);

                        //MessageBox.Show(
                        //    idRec + " " +
                        //    station + " " +
                        //    listNO + " " +
                        //    carNO + " " +
                        //    builtYear + " " +
                        //    carType + " " +
                        //    carLoc + " " +
                        //    admCode + " " +
                        //    owner + " " +
                        //    isLoaded + " " +
                        //    isWorking + " " +
                        //    workState);
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ex.SheetsInNewWorkbook = 1;
            Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);
            sheet.Name = "Отчет";

            for (int i = 1; i <= 3; i++) // строки
            {
                for (int j = 1; j <= 4; j++) // столбцы
                    sheet.Cells[i, j] = String.Format("хуяк {0} {1}", i, j);
            }

            // Выделяем диапазон ячеек от H1 до K1
            Excel.Range _excelCells1 = (Excel.Range)sheet.get_Range("A1", "J1").Cells;
            // Производим объединение
            _excelCells1.Merge(Type.Missing);
            sheet.Cells[1, 1] = "Общие";

            ex.Visible = true;
        }
    }
}
