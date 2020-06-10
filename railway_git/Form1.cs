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
        string station;
        int listNO;
        long carNO;
        int builtYear;
        string carType;
        string carLoc;
        int admCode;
        string owner;
        string isLoaded;
        string isWorking;
        string workState;
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
                "WHERE st.ESR = 480403 AND ccl.LIST_NO = 1";

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
                        station = reader.GetString(0);
                        listNO = Convert.ToInt32(reader.GetValue(1));
                        carNO = Convert.ToInt64(reader.GetValue(2));
                        builtYear = Convert.ToInt32(reader.GetValue(3));
                        carType = reader.GetString(4);
                        carLoc = reader.GetString(5);
                        admCode = Convert.ToInt32(reader.GetValue(6));
                        owner = reader.GetString(7);
                        isLoaded = reader.GetString(8);
                        isWorking = reader.GetString(9);
                        workState = reader.GetString(10);






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
            ex.DisplayAlerts = false;
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
            _excelCells1.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            _excelCells1.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

            Excel.Range _excelCells2 = (Excel.Range)sheet.get_Range("A2", "J2").Cells;
            _excelCells2.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            _excelCells2.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            _excelCells2.WrapText = true; // перенос текста в ячейках

            sheet.Columns[1].ColumnWidth = 4;
            sheet.Columns[2].ColumnWidth = 8;
            sheet.Columns[3].ColumnWidth = 11;
            sheet.Columns[4].ColumnWidth = 8;
            sheet.Columns[5].ColumnWidth = 12;
            sheet.Columns[6].ColumnWidth = 8;
            sheet.Columns[7].ColumnWidth = 12;
            sheet.Columns[8].ColumnWidth = 11;
            sheet.Columns[9].ColumnWidth = 6;
            sheet.Columns[10].ColumnWidth = 11;

            _excelCells2.Borders.ColorIndex = 0;
            _excelCells2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            _excelCells2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

            // Производим объединение
            _excelCells1.Merge(Type.Missing);
            sheet.Cells[1, 1] = "Переписной лист №";
            sheet.Cells[2, 1] = "№ п/п";
            sheet.Cells[2, 2] = "Номер вагона";
            sheet.Cells[2, 3] = "Год постройки";
            sheet.Cells[2, 4] = "Род вагона";
            sheet.Cells[2, 5] = "Дислокация";
            sheet.Cells[2, 6] = "Код страны-собств.";
            sheet.Cells[2, 7] = "Собственник вагона";
            sheet.Cells[2, 8] = "Состояние";
            sheet.Cells[2, 9] = "Парк";
            sheet.Cells[2, 10] = "Категория НРП";

            ex.Visible = true;
        }
    }
}
