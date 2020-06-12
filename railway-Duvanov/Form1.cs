using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Data.Common;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Aspose.Cells.Drawing;
using System.Collections;

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
        string listNum;

        Excel.Application excelApp;//
        Excel.Workbook workBook;//
        Excel.Worksheet sheet;//
        Excel.Range title;//
        Excel.Range columnsTable;//
        Excel.Range valuesTable;//
        ReportResults rptres;
        public Form1()
        {
            InitializeComponent();
            excelApp = new Excel.Application();
            /*excelApp = new Microsoft.Office.Interop.Excel.Application();
            workBook = excelApp.Workbooks.Add(Type.Missing);
            sheet = (Excel.Worksheet)excelApp.Worksheets.get_Item(1);
            excelApp.SheetsInNewWorkbook = 1;
            excelApp.DisplayAlerts = false;
            sheet.Name = "Отчет";//

            title = (Excel.Range)sheet.get_Range("A1", "J1").Cells;
            title.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            title.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            title.Font.Bold = true;
            title.Merge(Type.Missing);//

            columnsTable = (Excel.Range)sheet.get_Range("A2", "J2").Cells;
            columnsTable.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            columnsTable.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            columnsTable.WrapText = true; // перенос текста в ячейках
            columnsTable.Borders.ColorIndex = 0;
            columnsTable.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            columnsTable.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

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

            sheet.Rows.RowHeight = 25;
            sheet.Rows[1].RowHeight = 40;
            sheet.Rows[2].RowHeight = 50;
            sheet.Columns[1].ColumnWidth = 4;
            sheet.Columns[2].ColumnWidth = 9;
            sheet.Columns[3].ColumnWidth = 11;
            sheet.Columns[4].ColumnWidth = 7;
            sheet.Columns[5].ColumnWidth = 13;
            sheet.Columns[6].ColumnWidth = 8;
            sheet.Columns[7].ColumnWidth = 14;
            sheet.Columns[8].ColumnWidth = 14;
            sheet.Columns[9].ColumnWidth = 12;
            sheet.Columns[10].ColumnWidth = 14;
            */
        }

        private void Types() // формирование таблицы по родам вагона
        {
            string ESR_SQL = "480403";
            //int listNO, KR20, PL40, PV60, CS70, PR90, CMV93, FT94, ZRV95, MVZ93, ESR;

            excelApp.Worksheets.get_Item(1);
            rptres.SetSheet(1, "По родам вагона");
            rptres.CreateCarcass("По родам вагона");

            rptres.ExecuteSql("SELECT st.ESR, st.[NAME], ccl.LIST_NO," +
                "COUNT(CASE CAR_TYPE WHEN 'КР-20' then CAR_TYPE END) AS \"ЗРВ-95\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ПЛ-40' then CAR_TYPE END) AS \"ПЛ-40\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ПВ-60' then CAR_TYPE END) AS \"ПВ-60\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ЦС-70' then CAR_TYPE END) AS \"ЦС-70\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ПР-90' then CAR_TYPE END) AS \"ПР-90\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ЦМВ-93' then CAR_TYPE END) AS \"ЦМВ-93\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ФТ-94' then CAR_TYPE END) AS \"ФТ-94\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ЗРВ-95' then CAR_TYPE END) AS \"ЗРВ-95\", " +
                "COUNT(CASE CAR_TYPE WHEN 'МВЗ-92' then CAR_TYPE END) AS \"МВЗ-92\" " +
                "FROM CAR_CENSUS_LISTS ccl " +
                "INNER JOIN STATIONS st ON st.ESR = ccl.LOCATION_ESR " +
                "WHERE st.ESR = " + ESR_SQL +
                "GROUP BY st.ESR, st.NAME, ccl.LIST_NO");

            rptres.FillCarcass();

            rptres.SaveDocument(rptres.fname);
        }
        private void WorkingPark() // формирование таблицы по рабочему парку
        {
            string ESR_SQL = "480403";

            excelApp.Worksheets.get_Item(1);
            rptres.SetSheet(1, "Рабочий парк");
            rptres.CreateCarcass("Рабочий парк");

            rptres.ExecuteSql("SELECT st.ESR, st.[NAME], ccl.LIST_NO," +
                "COUNT(CASE CAR_TYPE WHEN 'КР-20' then CAR_TYPE END) AS \"ЗРВ-95\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ПЛ-40' then CAR_TYPE END) AS \"ПЛ-40\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ПВ-60' then CAR_TYPE END) AS \"ПВ-60\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ЦС-70' then CAR_TYPE END) AS \"ЦС-70\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ПР-90' then CAR_TYPE END) AS \"ПР-90\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ЦМВ-93' then CAR_TYPE END) AS \"ЦМВ-93\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ФТ-94' then CAR_TYPE END) AS \"ФТ-94\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ЗРВ-95' then CAR_TYPE END) AS \"ЗРВ-95\", " +
                "COUNT(CASE CAR_TYPE WHEN 'МВЗ-92' then CAR_TYPE END) AS \"МВЗ-92\" " +
                "FROM CAR_CENSUS_LISTS ccl " +
                "INNER JOIN STATIONS st ON st.ESR = ccl.LOCATION_ESR " +
                "WHERE ccl.IS_WORKING=1 AND st.ESR = " + ESR_SQL +
                "GROUP BY st.ESR, st.NAME, ccl.LIST_NO");

            rptres.FillCarcass();

            excelApp.Worksheets.get_Item(2);

            rptres.SetSheet(2, "Груженых");
            rptres.CreateCarcass("Груженых");

            rptres.ExecuteSql("SELECT st.ESR, st.[NAME], ccl.LIST_NO," +
                "COUNT(CASE CAR_TYPE WHEN 'КР-20' then CAR_TYPE END) AS \"ЗРВ-95\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ПЛ-40' then CAR_TYPE END) AS \"ПЛ-40\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ПВ-60' then CAR_TYPE END) AS \"ПВ-60\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ЦС-70' then CAR_TYPE END) AS \"ЦС-70\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ПР-90' then CAR_TYPE END) AS \"ПР-90\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ЦМВ-93' then CAR_TYPE END) AS \"ЦМВ-93\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ФТ-94' then CAR_TYPE END) AS \"ФТ-94\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ЗРВ-95' then CAR_TYPE END) AS \"ЗРВ-95\", " +
                "COUNT(CASE CAR_TYPE WHEN 'МВЗ-92' then CAR_TYPE END) AS \"МВЗ-92\" " +
                "FROM CAR_CENSUS_LISTS ccl " +
                "INNER JOIN STATIONS st ON st.ESR = ccl.LOCATION_ESR " +
                "WHERE ccl.IS_WORKING=1 AND ccl.IS_LOADED=1 AND st.ESR = " + ESR_SQL +
                "GROUP BY st.ESR, st.NAME, ccl.LIST_NO");

            rptres.FillCarcass();

            excelApp.Worksheets.get_Item(3);

            rptres.SetSheet(3, "Порожних");
            rptres.CreateCarcass("Порожних");

            rptres.ExecuteSql("SELECT st.ESR, st.[NAME], ccl.LIST_NO," +
                "COUNT(CASE CAR_TYPE WHEN 'КР-20' then CAR_TYPE END) AS \"ЗРВ-95\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ПЛ-40' then CAR_TYPE END) AS \"ПЛ-40\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ПВ-60' then CAR_TYPE END) AS \"ПВ-60\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ЦС-70' then CAR_TYPE END) AS \"ЦС-70\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ПР-90' then CAR_TYPE END) AS \"ПР-90\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ЦМВ-93' then CAR_TYPE END) AS \"ЦМВ-93\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ФТ-94' then CAR_TYPE END) AS \"ФТ-94\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ЗРВ-95' then CAR_TYPE END) AS \"ЗРВ-95\", " +
                "COUNT(CASE CAR_TYPE WHEN 'МВЗ-92' then CAR_TYPE END) AS \"МВЗ-92\" " +
                "FROM CAR_CENSUS_LISTS ccl " +
                "INNER JOIN STATIONS st ON st.ESR = ccl.LOCATION_ESR " +
                "WHERE ccl.IS_WORKING=1  AND ccl.IS_LOADED=0 AND st.ESR = " + ESR_SQL +
                "GROUP BY st.ESR, st.NAME, ccl.LIST_NO");

            rptres.FillCarcass();

            rptres.SaveDocument(rptres.fname);
        }

        private void NonWorkingPark(){
            string ESR_SQL = "480403";

            excelApp.Worksheets.get_Item(1);
            rptres.SetSheet(1, "Не рабочий парк");
            rptres.CreateCarcass("Не рабочий парк");

            rptres.ExecuteSql("SELECT st.ESR, st.[NAME], ccl.LIST_NO," +
                "COUNT(CASE CAR_TYPE WHEN 'КР-20' then CAR_TYPE END) AS \"ЗРВ-95\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ПЛ-40' then CAR_TYPE END) AS \"ПЛ-40\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ПВ-60' then CAR_TYPE END) AS \"ПВ-60\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ЦС-70' then CAR_TYPE END) AS \"ЦС-70\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ПР-90' then CAR_TYPE END) AS \"ПР-90\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ЦМВ-93' then CAR_TYPE END) AS \"ЦМВ-93\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ФТ-94' then CAR_TYPE END) AS \"ФТ-94\", " +
                "COUNT(CASE CAR_TYPE WHEN 'ЗРВ-95' then CAR_TYPE END) AS \"ЗРВ-95\", " +
                "COUNT(CASE CAR_TYPE WHEN 'МВЗ-92' then CAR_TYPE END) AS \"МВЗ-92\" " +
                "FROM CAR_CENSUS_LISTS ccl " +
                "INNER JOIN STATIONS st ON st.ESR = ccl.LOCATION_ESR " +
                "WHERE ccl.IS_WORKING = 0 AND ccl.NON_WORKING_STATE=1 AND st.ESR = "+ ESR_SQL+
                "GROUP BY st.ESR, st.NAME, ccl.LIST_NO");

            rptres.FillCarcass();

            rptres.SaveDocument(rptres.fname);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            /*conn = DBUtils.GetDBConnection();
            conn.Open();

            if (conn.State == ConnectionState.Open)
            {
                MessageBox.Show("всьо чьотка!");
            }
            else
            {
                MessageBox.Show("ашыбачька...");
            }*/
        }

        private void button2_Click(object sender, EventArgs e)
        {
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
                "WHERE st.ESR = " + esrStation + " and ccl.LIST_NO = " + listNum + "";

            cmd.CommandText = sql;

            using (DbDataReader reader = cmd.ExecuteReader())
            {
                int idRec = 0;
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        idRec++;
                        stationESR = reader.GetValue(0).ToString();
                        stationName = reader.GetString(1);
                        listNO = Convert.ToInt32(reader.GetValue(2));
                        carNO = Convert.ToInt64(reader.GetValue(3));
                        builtYear = Convert.ToInt32(reader.GetValue(4));
                        carType = reader.GetString(5);
                        carLoc = reader.GetString(6);
                        admCode = Convert.ToInt32(reader.GetValue(7));
                        owner = reader.GetString(8);
                        isLoaded = reader.GetString(9);
                        isWorking = reader.GetString(10);
                        workState = reader.GetString(11);

                        sheet.Cells[1, 1] = "Переписной лист №" + listNO + " станция " + stationName + " (" + stationESR + ")";
                        sheet.Cells[2 + idRec, 1] = idRec;
                        sheet.Cells[2 + idRec, 2] = carNO;
                        sheet.Cells[2 + idRec, 3] = builtYear;
                        sheet.Cells[2 + idRec, 4] = carType;
                        sheet.Cells[2 + idRec, 5] = carLoc;
                        sheet.Cells[2 + idRec, 6] = admCode;
                        sheet.Cells[2 + idRec, 7] = owner;
                        sheet.Cells[2 + idRec, 8] = isLoaded;
                        sheet.Cells[2 + idRec, 9] = isWorking;
                        sheet.Cells[2 + idRec, 10] = workState;

                    }

                    valuesTable = (Excel.Range)sheet.get_Range("A3", "J" + (idRec + 2).ToString()).Cells;
                    valuesTable.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    valuesTable.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    valuesTable.WrapText = true; // перенос текста в ячейках
                    valuesTable.Borders.ColorIndex = 0;
                    valuesTable.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    valuesTable.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    // Реализовать сохранение документов по каждому номеру переписного листа - listNO

                    SaveFileDialog fileDialog = new SaveFileDialog();
                    fileDialog.FileName = "Переписной лист №" + listNO + " - станция " + stationName + " (" + stationESR + ").xlsx";
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
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            conn = DBUtils.GetDBConnection();
            conn.Open();
            //rptres.conn = conn;
            if (conn.State == ConnectionState.Open)
            {
                MessageBox.Show("всьо чьотка!");

                cmd = new SqlCommand();

                sql = "select NAME, ESR FROM STATIONS";

                cmd.Connection = conn;
                //rptres.cmd.Connection = conn;
                cmd.CommandText = sql;

                listView1.View = View.Details;
                listView1.ListViewItemSorter = new ListViewColumnComparer(0);

                using (DbDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            listView1.Items.Add(new ListViewItem(
                                new string[] 
                                { 
                                    reader.GetString(0),
                                    reader.GetValue(1).ToString()
                                }));

                            //stations.Items.Add(reader.GetString(0) + " ("+ reader.GetValue(1).ToString() + ")");
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("ашыбачька...");
            }
        }

        private void listView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            lists.Items.Clear();
            esrStation = listView1.Items[e.ItemIndex].SubItems[1].Text;

            sql = "select LIST_NO from CAR_CENSUS_LISTS ccl " +
                "INNER JOIN STATIONS st on st.ESR = ccl.LOCATION_ESR " +
                "where st.ESR = "+ esrStation +" " +
                "group by LIST_NO";

            cmd.Connection = conn;
            rptres.cmd.Connection = conn;
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

        private void lists_SelectedIndexChanged(object sender, EventArgs e)
        {
            listNum = lists.Text;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            rptres = new ReportResults(excelApp);
            rptres.conn = conn;
            rptres.cmd.Connection = conn;
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
        }

    }
    public class ReportResults
    {
        public Excel.Application excelApp;
        public Excel.Workbook workBook;
        public Excel.Worksheet sheet;
        public Excel.Range range;
        public Excel.Sheets excelsheets;
        public int countExcelSheets;

        public SqlCommand cmd;
        public SqlConnection conn;
        public string sql;

        public string fname = "";
        public ReportResults(Excel.Application exApp)
        {
            excelApp = exApp;
            excelApp.SheetsInNewWorkbook = 3;
            workBook = excelApp.Workbooks.Add(Type.Missing);
            sheet = (Excel.Worksheet)excelApp.Worksheets.get_Item(1);
            excelApp.DisplayAlerts = false;
            excelsheets = workBook.Worksheets;
            countExcelSheets = excelsheets.Count;
            cmd = new SqlCommand();
        }
        public void SetSheet(int NmbWorksheet, string SheetName)
        {
            if (countExcelSheets < NmbWorksheet)
            {
                excelApp.SheetsInNewWorkbook = NmbWorksheet;
            }
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

        public void SaveDocument(string DocumentName)
        {
            SaveFileDialog fileDialog = new SaveFileDialog();
            fileDialog.FileName = DocumentName;
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

        public void FillCarcass()
        {
            string listNO, KR20, PL40, PV60, CS70, PR90, CMV93, FT94, ZRV95, MVZ93, ESR;
            string stationName;
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

                        fname = "Итоги переписи по станции " + stationName + " (" + ESR + ").xlsx";

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

        public void ChangeWorkSheet(int Nmb)
        {
            sheet = excelsheets.get_Item(Nmb);
        }
    }
}
