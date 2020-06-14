using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Data.Common;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace test_railway
{
    public partial class Form1 : Form
    {
        SqlConnection conn;
        SqlCommand cmd;
        string sql;

        string stationESR;
        string stationName;
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

        string esrStation;
        Excel.Application excelApp;
        Excel.Range valuesTable;
        ReportResults rptres;
        public Form1()
        {
            InitializeComponent();
            excelApp = new Excel.Application();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            conn = DBUtils.GetDBConnection();
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                MessageBox.Show("Успешное подключение к базе данных!");
                label5.ForeColor = Color.Green;
                label5.Text = "Подключена";
                cmd = new SqlCommand();

                sql = "select ESR, NAME FROM STATIONS";

                cmd.Connection = conn;
                cmd.CommandText = sql;

                listView1.View = View.Details;
                listView1.ListViewItemSorter = new ListViewColumnComparer(0);

                listView2.View = View.Details;
                listView2.ListViewItemSorter = new ListViewColumnComparer(0);
                using (DbDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            listView1.Columns[0].Text = "ESR";
                            listView1.Columns[1].Text = "Название станции";
                            listView1.Columns[1].Width = 120;
                            string[] row = { reader.GetInt32(0).ToString(), reader.GetValue(1).ToString() };
                            var listViewItem = new ListViewItem(row);
                            listView1.Items.Add(listViewItem);

                            listView2.Columns[0].Text = "ESR";
                            listView2.Columns[1].Text = "Название станции";
                            listView2.Columns[1].Width = 120;
                            listViewItem = new ListViewItem(row);
                            listView2.Items.Add(listViewItem);

                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Установить соединение с базой данных не удалось!");
                label5.ForeColor = Color.Red;
                label5.Text = "Отключена";
            }
        }
        private void Types() // формирование таблицы по родам вагона
        {
            rptres.CreateNewWorkbook(1);

            rptres.SetSheet(1, "По родам вагона");
            rptres.CreateCarcass("По родам вагона");
            rptres.ExSqlType(rptres.ESR);
            rptres.FillCarcass();

            rptres.SaveDocument();
        }
        private void WorkingPark() // формирование таблицы по рабочему парку
        {
            rptres.CreateNewWorkbook(3);

            excelApp.Worksheets.get_Item(1);
            rptres.SetSheet(1, "Рабочий парк");
            rptres.CreateCarcass("Всего рабочий парк");
            rptres.ExSQLWorkingParkResult(rptres.ESR);
            rptres.FillCarcass();


            excelApp.Worksheets.get_Item(2);
            rptres.SetSheet(2, "Груженых");
            rptres.CreateCarcass("Груженых");
            rptres.ExSqlWorkingParkLoaded(rptres.ESR, 1);
            rptres.FillCarcass();


            excelApp.Worksheets.get_Item(3);
            rptres.SetSheet(3, "Порожних");
            rptres.CreateCarcass("Порожних");
            rptres.ExSqlWorkingParkLoaded(rptres.ESR, 0);
            rptres.FillCarcass();

            rptres.SaveDocument();
        }

        private void NonWorkingPark(){
            rptres.CreateNewWorkbook(6);

            excelApp.Worksheets.get_Item(1);
            rptres.SetSheet(1, "Не рабочий парк");
            rptres.CreateCarcass("Всего НРП");
            rptres.ExSqlNonWorkingParkResult(rptres.ESR);
            rptres.FillCarcass();

            excelApp.Worksheets.get_Item(2);
            rptres.SetSheet(2, "Неисправных");
            rptres.CreateCarcass("Неисправных");
            rptres.ExSqlNonWorkingPark(rptres.ESR, 1);
            rptres.FillCarcass();

            excelApp.Worksheets.get_Item(3);
            rptres.SetSheet(3, "Резерв");
            rptres.CreateCarcass("Резерв");
            rptres.ExSqlNonWorkingPark(rptres.ESR, 2);
            rptres.FillCarcass();

            excelApp.Worksheets.get_Item(4);
            rptres.SetSheet(4, "ДЛЗО");
            rptres.CreateCarcass("ДЛЗО");
            rptres.ExSqlNonWorkingPark(rptres.ESR, 3);
            rptres.FillCarcass();

            excelApp.Worksheets.get_Item(5);
            rptres.SetSheet(5, "СТН");
            rptres.CreateCarcass("СТН");
            rptres.ExSqlNonWorkingPark(rptres.ESR, 4);
            rptres.FillCarcass();

            excelApp.Worksheets.get_Item(6);
            rptres.SetSheet(6, "Поврежден по акту ВУ-25");
            rptres.CreateCarcass("Поврежден по акту ВУ-25");
            rptres.ExSqlNonWorkingPark(rptres.ESR, 5);
            rptres.FillCarcass();

            rptres.SaveDocument();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            rptres = new ReportResults(excelApp);
            rptres.conn = conn;
            rptres.cmd.Connection = conn;
            if (lists.Items.Count == 0) {
                MessageBox.Show("Не найдены листы по данной станции");
                return;
            }
            if (listView1.SelectedItems.Count == 0)
            {
                MessageBox.Show("Выберите станцию");
            }
            else if (lists.Text == "") { 
                MessageBox.Show("Выберите лист");
            }
            else
            {
                rptres.CreateNewWorkbook(1);
                int firstIndex = listView1.SelectedIndices[0];
                rptres.ESR = listView1.Items[firstIndex].SubItems[0].Text;
                rptres.SetFileName(listView1.Items[firstIndex].SubItems[1].Text);

                sql = "select st.ESR, st.NAME as 'Станция (для заголовка)', " +
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
                "WHERE st.ESR = " + rptres.ESR + " and ccl.LIST_NO = " + lists.Text + "";

                cmd.CommandText = sql;

                using (DbDataReader reader = cmd.ExecuteReader())
                {
                    int idRec = 0;
                    if (reader.HasRows)
                    {

                        rptres.SetCellValue(3, 1, "№ п/п");
                        rptres.SetCellValue(3, 2, "Номер вагона");
                        rptres.SetCellValue(3, 3, "Год постройки");
                        rptres.SetCellValue(3, 4, "Род вагона");
                        rptres.SetCellValue(3, 5, "Дислокация");
                        rptres.SetCellValue(3, 6, "Код страны-собств.");
                        rptres.SetCellValue(3, 7, "Собственник вагона");
                        rptres.SetCellValue(3, 8, "Состояние");
                        rptres.SetCellValue(3, 9, "Парк");
                        rptres.SetCellValue(3, 10, "Категория НРП");
                        while (reader.Read())
                        {
                            idRec++;

                            listNO = Convert.ToInt32(reader.GetValue(2));
                            stationESR = reader.GetValue(0).ToString();
                            stationName = reader.GetString(1);

                            carNO = Convert.ToInt64(reader.GetValue(3));
                            builtYear = Convert.ToInt32(reader.GetValue(4));
                            carType = reader.GetString(5);
                            carLoc = reader.GetString(6);
                            admCode = Convert.ToInt32(reader.GetValue(7));
                            owner = reader.GetString(8);
                            isLoaded = reader.GetString(9);
                            isWorking = reader.GetString(10);
                            workState = reader.GetString(11);

                            rptres.SetCellValue(1, 1, "Переписной лист №" + listNO + " станция " + stationName + " (" + stationESR + ")");
                            rptres.SetCellValue(3 + idRec, 1, idRec);
                            rptres.SetCellValue(3 + idRec, 2, carNO);
                            rptres.SetCellValue(3 + idRec, 3, builtYear);
                            rptres.SetCellValue(3 + idRec, 4, carType);
                            rptres.SetCellValue(3 + idRec, 5, carLoc);
                            rptres.SetCellValue(3 + idRec, 6, admCode);
                            rptres.SetCellValue(3 + idRec, 7, owner);
                            rptres.SetCellValue(3 + idRec, 8, isLoaded);
                            rptres.SetCellValue(3 + idRec, 9, isWorking);
                            rptres.SetCellValue(3 + idRec, 10, workState);

                        }

                        valuesTable = rptres.GetRange("A3", "J" + (idRec + 3).ToString());
                        valuesTable.HorizontalAlignment = Excel.Constants.xlCenter;
                        valuesTable.VerticalAlignment = Excel.Constants.xlCenter;
                        valuesTable.WrapText = true; // перенос текста в ячейках
                        valuesTable.Borders.ColorIndex = 0;
                        valuesTable.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        valuesTable.Borders.Weight = Excel.XlBorderWeight.xlThin;

                        rptres.SetTitle("A1", "K1");

                        rptres.SetFileName("Переписной лист по станции " + stationName + " (" + stationESR + ").xlsx");
                        rptres.SaveDocument();
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            rptres = new ReportResults(excelApp);
            rptres.conn = conn;
            rptres.cmd.Connection = conn;
            if (listView2.SelectedItems.Count == 0)
            {
                MessageBox.Show("Выберите станцию");
            }
            else
            {
                int firstIndex = listView2.SelectedIndices[0];
                rptres.ESR = listView2.Items[firstIndex].SubItems[0].Text;
                string StationName = listView2.Items[firstIndex].SubItems[1].Text;
                rptres.SetFileName("Итоги переписи по станции " + StationName + " (" + rptres.ESR + ").xlsx");
                if (radioButton1.Checked)
                {
                    Types();
                }
                else if (radioButton2.Checked)
                {
                    WorkingPark();
                }
                else if (radioButton3.Checked)
                {
                    NonWorkingPark();
                }
                else
                {
                    MessageBox.Show("Выберите пункт");
                }
            }
        }

        private void listView2_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            
            if (listView2.SelectedIndices.Count > 0)
               esrStation = listView2.SelectedItems[0].Text;
            else
                esrStation = "0";
        }

        private void listView1_ItemSelectionChanged_1(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            lists.Items.Clear();
            if (listView1.SelectedIndices.Count > 0)
            {
                esrStation = listView1.SelectedItems[0].Text;
                sql = "select LIST_NO from CAR_CENSUS_LISTS ccl " +
                    "INNER JOIN STATIONS st on st.ESR = ccl.LOCATION_ESR " +
                    "where st.ESR = " + esrStation + " " +
                    "group by LIST_NO";

                cmd.Connection = conn;
                cmd.CommandText = sql;

                using (DbDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            lists.Items.Add(reader.GetValue(0).ToString());
                        }
                    }
                }
            }
        }
    }
    public class ReportResults
    {
        public string ESR;
        public Excel.Application excelApp;
        public Excel.Workbook workBook;
        public Excel.Worksheet sheet;
        public Excel.Range range;

        public SqlCommand cmd;
        public SqlConnection conn;
        public string sql;

        public string fname = "";
        public ReportResults(Excel.Application exApp)
        {
            excelApp = exApp;
            cmd = new SqlCommand();
        }
        public void CreateNewWorkbook(int countSheets) {
            excelApp.SheetsInNewWorkbook = countSheets;
            excelApp.Workbooks.Add(Type.Missing);
            sheet = (Excel.Worksheet)excelApp.Worksheets.get_Item(1);
            excelApp.DisplayAlerts = false;
        }
        public void SetSheet(int NmbWorksheet, string SheetName)
        {
            sheet = (Excel.Worksheet)excelApp.Worksheets.get_Item(NmbWorksheet);
            sheet.Name = SheetName;
        }
        public Excel.Range GetRange(string CellFrom, string CellTo)
        {
            return sheet.get_Range(CellFrom, CellTo).Cells;
        }
        public Excel.Range GetRange(string CellFromTo)
        {
            return sheet.get_Range(CellFromTo).Cells;
        }
        public void SetTitle(string CellFrom, string CellTo)
        {
            range = GetRange(CellFrom, CellTo);
            range.HorizontalAlignment = Excel.Constants.xlCenter;
            range.VerticalAlignment = Excel.Constants.xlCenter;
            range.Font.Bold = true;
            range.Merge(Type.Missing);
        }
        public void MergeCells(string CellFrom, string CellTo)
        {
            range = GetRange(CellFrom, CellTo);
            range.HorizontalAlignment = Excel.Constants.xlCenter;
            range.VerticalAlignment = Excel.Constants.xlCenter;
            range.WrapText = true; // перенос текста в ячейках
            range.Borders.ColorIndex = 0;
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.Weight = Excel.XlBorderWeight.xlThin;
            range.Merge(Type.Missing);
        }

        public void SetCellValue(int CellFrom, int CellTo, string Value)
        {
            sheet.Cells[CellFrom, CellTo] = Value;
        }
        public void SetCellValue(int CellFrom, int CellTo, int Value)
        {
            sheet.Cells[CellFrom, CellTo] = Value;
        }
        public void SetCellValue(int CellFrom, int CellTo, float Value)
        {
            sheet.Cells[CellFrom, CellTo] = Value;
        }
        public void SetCellValue(int CellFrom, int CellTo, double Value)
        {
            sheet.Cells[CellFrom, CellTo] = Value;
        }

        public void SetRowHeight(int RowNmb, int Value)
        {
            sheet.Rows[RowNmb].RowHeight = Value;
        }
        public void SetColumnWidth(int ColNmb, int Value)
        {
            sheet.Columns[ColNmb].ColumnWidth = Value;
        }
        public void ExecuteSql(string _sql)
        {
            sql = _sql;
            cmd.Connection = conn;
            cmd.CommandText = sql;
        }
        public void ExSqlType(string ESR_SQL) {
            sql = "SELECT DISTINCT " +
                    "st.ESR, " +
                    "st.[NAME], " +
                    "ccl.LIST_NO, " +
                    "(SELECT COUNT(CASE CAR_TYPE WHEN 'КР-20' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO)  AS \"КР -20\", " +
                    "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПЛ-40' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO)  AS \"ПЛ -40\", " +
                        "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПВ-60' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO)  AS \"ПВ-60\", " +
                        "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЦС-70' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO)  AS \"ЦС -70\", " +
                        "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПР-90' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO)  AS \"ПР -90\", " +
                        "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЦМВ-93' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO)  AS \"ЦМВ -93\", " +
                        "(SELECT COUNT(CASE CAR_TYPE WHEN 'ФТ-94' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO)  AS \"ФТ -94\", " +
                        "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЗРВ-95' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO)  AS \"ЗРВ -95\", " +
                        "(SELECT COUNT(CASE CAR_TYPE WHEN 'МВЗ-92' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO)  AS \"МВЗ -92\" " +
                    "FROM " +
                    "CAR_CENSUS_LISTS ccl " +
                    "JOIN STATIONS st ON st.ESR = ccl.LOCATION_ESR " +
                    "WHERE " +
                    "st.ESR = " + ESR_SQL + "";
            ExecuteSql(sql);
        }
        public void ExSqlNonWorkingPark(string ESR_SQL, int nonWorkingState) {
            sql = "SELECT DISTINCT " +
                            "st.ESR, " +
                            "st.[NAME], " +
                            "ccl.LIST_NO, " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'КР-20' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0 AND ccll.NON_WORKING_STATE = " + nonWorkingState.ToString() + ")  AS \"КР -20\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПЛ-40' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0 AND ccll.NON_WORKING_STATE = " + nonWorkingState.ToString() + ")  AS \"ПЛ -40\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПВ-60' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0 AND ccll.NON_WORKING_STATE = " + nonWorkingState.ToString() + ")  AS \"ПВ-60\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЦС-70' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0 AND ccll.NON_WORKING_STATE = " + nonWorkingState.ToString() + ")  AS \"ЦС -70\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПР-90' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0 AND ccll.NON_WORKING_STATE = " + nonWorkingState.ToString() + ")  AS \"ПР -90\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЦМВ-93' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0 AND ccll.NON_WORKING_STATE = " + nonWorkingState.ToString() + ")  AS \"ЦМВ -93\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'ФТ-94' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0 AND ccll.NON_WORKING_STATE = " + nonWorkingState.ToString() + ")  AS \"ФТ -94\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЗРВ-95' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0 AND ccll.NON_WORKING_STATE = " + nonWorkingState.ToString() + ")  AS \"ЗРВ -95\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'МВЗ-92' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0 AND ccll.NON_WORKING_STATE = " + nonWorkingState.ToString() + ")  AS \"МВЗ -92\" " +
                            "FROM " +
                            "CAR_CENSUS_LISTS ccl " +
                            "JOIN STATIONS st ON st.ESR = ccl.LOCATION_ESR " +
                            "WHERE " +
                            "st.ESR = " + ESR_SQL + "";
            ExecuteSql(sql);
        }
        public void ExSqlNonWorkingParkResult(string ESR_SQL)
        {
            sql = "SELECT DISTINCT " +
                            "st.ESR, " +
                            "st.[NAME], " +
                            "ccl.LIST_NO, " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'КР-20' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0)  AS \"КР -20\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПЛ-40' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0)  AS \"ПЛ -40\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПВ-60' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0)  AS \"ПВ-60\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЦС-70' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0)  AS \"ЦС -70\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПР-90' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0)  AS \"ПР -90\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЦМВ-93' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0)  AS \"ЦМВ -93\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'ФТ-94' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0)  AS \"ФТ -94\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЗРВ-95' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0)  AS \"ЗРВ -95\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'МВЗ-92' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0)  AS \"МВЗ -92\" " +
                            "FROM " +
                            "CAR_CENSUS_LISTS ccl " +
                            "JOIN STATIONS st ON st.ESR = ccl.LOCATION_ESR " +
                            "WHERE " +
                            "st.ESR = " + ESR_SQL + "";
            ExecuteSql(sql);
        }
        public void ExSqlWorkingParkLoaded(string ESR_SQL, int isLoaded)
        {
            sql = "SELECT DISTINCT " +
                            "st.ESR, " +
                            "st.[NAME], " +
                            "ccl.LIST_NO, " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'КР-20' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1 AND ccll.IS_LOADED = " + isLoaded.ToString() + ")  AS \"КР -20\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПЛ-40' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1 AND ccll.IS_LOADED = " + isLoaded.ToString() + ")  AS \"ПЛ -40\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПВ-60' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1 AND ccll.IS_LOADED = " + isLoaded.ToString() + ")  AS \"ПВ-60\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЦС-70' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1 AND ccll.IS_LOADED = " + isLoaded.ToString() + ")  AS \"ЦС -70\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПР-90' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1 AND ccll.IS_LOADED = " + isLoaded.ToString() + ")  AS \"ПР -90\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЦМВ-93' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1 AND ccll.IS_LOADED = " + isLoaded.ToString() + ")  AS \"ЦМВ -93\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'ФТ-94' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1 AND ccll.IS_LOADED = " + isLoaded.ToString() + ")  AS \"ФТ -94\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЗРВ-95' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1 AND ccll.IS_LOADED = " + isLoaded.ToString() + ")  AS \"ЗРВ -95\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'МВЗ-92' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1 AND ccll.IS_LOADED = " + isLoaded.ToString() + ")  AS \"МВЗ -92\" " +
                            "FROM " +
                            "CAR_CENSUS_LISTS ccl " +
                            "JOIN STATIONS st ON st.ESR = ccl.LOCATION_ESR " +
                            "WHERE " +
                            "st.ESR = " + ESR_SQL + "";
            ExecuteSql(sql);
        }
        public void ExSQLWorkingParkResult(string ESR_SQL) {
            sql = "SELECT DISTINCT " +
                                "st.ESR, " +
                                "st.[NAME], " +
                                "ccl.LIST_NO, " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'КР-20' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1)  AS \"КР -20\", " +
                                "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПЛ-40' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1)  AS \"ПЛ -40\", " +
                                    "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПВ-60' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1)  AS \"ПВ-60\", " +
                                    "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЦС-70' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1)  AS \"ЦС -70\", " +
                                    "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПР-90' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1)  AS \"ПР -90\", " +
                                    "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЦМВ-93' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1)  AS \"ЦМВ -93\", " +
                                    "(SELECT COUNT(CASE CAR_TYPE WHEN 'ФТ-94' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1)  AS \"ФТ -94\", " +
                                    "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЗРВ-95' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1)  AS \"ЗРВ -95\", " +
                                    "(SELECT COUNT(CASE CAR_TYPE WHEN 'МВЗ-92' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1)  AS \"МВЗ -92\" " +
                                "FROM " +
                                "CAR_CENSUS_LISTS ccl " +
                                "JOIN STATIONS st ON st.ESR = ccl.LOCATION_ESR " +
                                "WHERE " +
                                "st.ESR = " + ESR_SQL + "";
            ExecuteSql(sql);
        }

        public void SaveDocument()
        {
            SaveFileDialog fileDialog = new SaveFileDialog();
            fileDialog.Filter = "Microsoft Excel (*.xlxs)|*.*";
            fileDialog.FileName = fname;
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                excelApp.Application.ActiveWorkbook.SaveAs(
                    fileDialog.FileName,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Excel.XlSaveAsAccessMode.xlShared,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing);

                if (MessageBox.Show("Файл успешно сохранен!\n" +
                    "\nОткрыть этот файл?", "Сообщение", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    excelApp.Visible = true;
                }
                else
                {
                    excelApp.Application.ActiveWorkbook.Close(true, Type.Missing, Type.Missing);
                }
            }
            else
            {
                MessageBox.Show("Файл не был сохранен...");
            }
        }

        public void CreateCarcass(string name)
        {
            SetTitle("A1", "K1");

            MergeCells("A3", "A7");
            MergeCells("B3", "B7");
            MergeCells("C3", "K4");
            MergeCells("C5", "C7");
            MergeCells("C5", "C7");
            MergeCells("D5", "D7");
            MergeCells("E5", "E7");
            MergeCells("F5", "F7");
            MergeCells("G5", "G7");
            MergeCells("G5", "G7");
            MergeCells("H5", "K5");
            MergeCells("H6", "H7");
            MergeCells("I6", "I7");
            MergeCells("J6", "J7");
            MergeCells("K6", "K7");

            SetCellValue(1, 1, "Переписной лист №");
            SetCellValue(3, 1, "№ листа");
            SetCellValue(3, 2, "Всего преписано вагонов");
            SetCellValue(3, 3, name);
            SetCellValue(5, 3, "КР-20");
            SetCellValue(5, 4, "ПЛ-40");
            SetCellValue(5, 5, "ПВ-60");
            SetCellValue(5, 6, "ЦС-70");
            SetCellValue(5, 7, "ПР-90");
            SetCellValue(5, 8, "в т.ч.");
            SetCellValue(6, 8, "ЦМВ-93");
            SetCellValue(6, 9, "ФТ-94");
            SetCellValue(6, 10, "ЗРВ-95");
            SetCellValue(6, 11, "МВЗ-95");

            range = GetRange("A25", "K25");
            range.Font.Bold = true;

            sheet.Rows.RowHeight = 25;
            SetRowHeight(1, 40);
            SetRowHeight(3, 10);
            SetRowHeight(4, 10);
            SetRowHeight(5, 15);
            SetRowHeight(6, 15);
            SetRowHeight(7, 15);

            SetColumnWidth(1, 6);
            SetColumnWidth(2, 10);
            SetColumnWidth(3, 7);
            SetColumnWidth(4, 7);
            SetColumnWidth(5, 7);
            SetColumnWidth(6, 7);
            SetColumnWidth(7, 7);
            SetColumnWidth(8, 8);
            SetColumnWidth(9, 8);
            SetColumnWidth(10, 8);
            SetColumnWidth(11, 8);
        }
        public void SetFileName(string newFname) {
            fname = "";
            fname = newFname;
        }
        public void FillCarcass()
        {
            string listNO, KR20, PL40, PV60, CS70, PR90, CMV93, FT94, ZRV95, MVZ93, ESR, stationName;
            using (DbDataReader reader = cmd.ExecuteReader())
            {
                int idRec = 0;
                Excel.Range rng;
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        idRec++;
                        ESR = reader.GetValue(0).ToString();
                        stationName = reader.GetValue(1).ToString();

                        listNO = reader.GetValue(2).ToString();
                        KR20 = reader.GetValue(3).ToString();
                        PL40 = reader.GetValue(4).ToString();
                        PV60 = reader.GetValue(5).ToString();
                        CS70 = reader.GetValue(6).ToString();
                        PR90 = reader.GetValue(7).ToString();
                        CMV93 = reader.GetValue(8).ToString();
                        FT94 = reader.GetValue(9).ToString();
                        ZRV95 = reader.GetValue(10).ToString();
                        MVZ93 = reader.GetValue(11).ToString();

                        SetCellValue(1, 1, "Итоги переписи по станции " + stationName + " (" + ESR + ")");
                        SetCellValue(7 + idRec, 1, listNO);
                        SetCellValue(7 + idRec, 3, KR20);
                        SetCellValue(7 + idRec, 4, PL40);
                        SetCellValue(7 + idRec, 5, PV60);
                        SetCellValue(7 + idRec, 6, CS70);
                        SetCellValue(7 + idRec, 7, PR90);
                        SetCellValue(7 + idRec, 8, CMV93);
                        SetCellValue(7 + idRec, 9, FT94);
                        SetCellValue(7 + idRec, 10, ZRV95);
                        SetCellValue(7 + idRec, 11, MVZ93);

                        rng = GetRange("C" + (7 + idRec).ToString() + ":" + "K" + (7 + idRec).ToString());
                        SetCellValue(7 + idRec, 2, excelApp.WorksheetFunction.Sum(rng)); //вычисляем сумму ячеек
                    }

                    range = GetRange("A8", "K" + (idRec + 8).ToString());
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    range.VerticalAlignment = Excel.Constants.xlCenter;
                    range.WrapText = true; // перенос текста в ячейках
                    range.Borders.ColorIndex = 0;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders.Weight = Excel.XlBorderWeight.xlThin;
                    SetCellValue(idRec + 8, 1, "Всего");

                    rng = GetRange("B8:B" + (7 + idRec).ToString());
                    SetCellValue(8 + idRec, 2, excelApp.WorksheetFunction.Sum(rng)); //вычисляем сумму ячеек

                    rng = GetRange("C8:C" + (7 + idRec).ToString());
                    SetCellValue(8 + idRec, 3, excelApp.WorksheetFunction.Sum(rng)); //вычисляем сумму ячеек      

                    rng = GetRange("D8:D" + (7 + idRec).ToString());
                    SetCellValue(8 + idRec, 4, excelApp.WorksheetFunction.Sum(rng)); //вычисляем сумму ячеек

                    rng = GetRange("E8:E" + (7 + idRec).ToString());
                    SetCellValue(8 + idRec, 5, excelApp.WorksheetFunction.Sum(rng)); //вычисляем сумму ячеек

                    rng = GetRange("F8:F" + (7 + idRec).ToString());
                    SetCellValue(8 + idRec, 6, excelApp.WorksheetFunction.Sum(rng)); //вычисляем сумму ячеек

                    rng = GetRange("G8:G" + (7 + idRec).ToString());
                    SetCellValue(8 + idRec, 7, excelApp.WorksheetFunction.Sum(rng)); //вычисляем сумму ячеек

                    rng = GetRange("H8:H" + (7 + idRec).ToString());
                    SetCellValue(8 + idRec, 8, excelApp.WorksheetFunction.Sum(rng)); //вычисляем сумму ячеек

                    rng = GetRange("I8:I" + (7 + idRec).ToString());
                    SetCellValue(8 + idRec, 9, excelApp.WorksheetFunction.Sum(rng)); //вычисляем сумму ячеек

                    rng = GetRange("J8:J" + (7 + idRec).ToString());
                    SetCellValue(8 + idRec, 10, excelApp.WorksheetFunction.Sum(rng)); //вычисляем сумму ячеек
                    rng = GetRange("K8:K" + (7 + idRec).ToString());
                    SetCellValue(8 + idRec, 11, excelApp.WorksheetFunction.Sum(rng)); //вычисляем сумму ячеек
                }
            }
        }
    }
}
