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

        Excel.Application excelApp;
        Excel.Workbook workBook;
        Excel.Worksheet sheet;
        Excel.Range title;
        Excel.Range columnsTable;
        Excel.Range valuesTable;

        public Form1()
        {
            InitializeComponent();

            excelApp = new Microsoft.Office.Interop.Excel.Application();
            workBook = excelApp.Workbooks.Add(Type.Missing);
            sheet = (Excel.Worksheet)excelApp.Worksheets.get_Item(1);
            excelApp.SheetsInNewWorkbook = 1;
            excelApp.DisplayAlerts = false;
            sheet.Name = "Отчет";

            title = (Excel.Range)sheet.get_Range("A1", "J1").Cells;
            title.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            title.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            title.Font.Bold = true;
            title.Merge(Type.Missing);

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

                        //MessageBox.Show((2 + idRec).ToString());

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

            if (conn.State == ConnectionState.Open)
            {
                MessageBox.Show("всьо чьотка!");

                cmd = new SqlCommand();

                sql = "select NAME, ESR FROM STATIONS";

                cmd.Connection = conn;
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
            if (radioButton1.Checked)
            {
                types();
            }
            else if (radioButton2.Checked)
            {

            }
            else if (radioButton3.Checked)
            {

            }
        }



























        private void types() // формирование таблицы по родам вагона
        {
            cmd = new SqlCommand();

            int listNO, KR20, PL40, PV60, CS70, PR90, CMV93, FT94, ZRV95, MVZ93, ESR;
            string stationName;

            Excel.Application excelApp;
            Excel.Workbook workBook;
            Excel.Worksheet sheet;
            Excel.Range title;
            Excel.Range range;
            Excel.Range valuesTable;

            //--------------------------------------------------------------------------

            excelApp = new Microsoft.Office.Interop.Excel.Application();
            workBook = excelApp.Workbooks.Add(Type.Missing);
            sheet = (Excel.Worksheet)excelApp.Worksheets.get_Item(1);
            excelApp.SheetsInNewWorkbook = 1;
            excelApp.DisplayAlerts = false;
            sheet.Name = "По родам вагона";

            title = (Excel.Range)sheet.get_Range("A1", "K1").Cells;
            title.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            title.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            title.Font.Bold = true;
            title.Merge(Type.Missing);

            range = (Excel.Range)sheet.get_Range("A3", "A7").Cells;
            cellsMerge(range);

            //cellsMerge("B3", "B7");
            range = (Excel.Range)sheet.get_Range("B3", "B7").Cells;
            cellsMerge(range);
            //cellsMerge("C3", "K4");
            range = (Excel.Range)sheet.get_Range("C3", "K4").Cells;
            cellsMerge(range);

            //cellsMerge("C5", "C7");
            range = (Excel.Range)sheet.get_Range("C5", "C7").Cells;
            cellsMerge(range);
            //cellsMerge("D5", "D7");
            range = (Excel.Range)sheet.get_Range("D5", "D7").Cells;
            cellsMerge(range);
            //cellsMerge("E5", "E7");
            range = (Excel.Range)sheet.get_Range("E5", "E7").Cells;
            cellsMerge(range);
            //cellsMerge("F5", "F7");
            range = (Excel.Range)sheet.get_Range("F5", "F7").Cells;
            cellsMerge(range);
            //cellsMerge("G5", "G7");
            range = (Excel.Range)sheet.get_Range("G5", "G7").Cells;
            cellsMerge(range);

            //cellsMerge("H5", "K5");
            range = (Excel.Range)sheet.get_Range("H5", "K5").Cells;
            cellsMerge(range);
            //cellsMerge("H6", "H7");
            range = (Excel.Range)sheet.get_Range("H6", "H7").Cells;
            cellsMerge(range);
            //cellsMerge("I6", "I7");
            range = (Excel.Range)sheet.get_Range("I6", "I7").Cells;
            cellsMerge(range);
            //cellsMerge("J6", "J7");
            range = (Excel.Range)sheet.get_Range("J6", "J7").Cells;
            cellsMerge(range);
            //cellsMerge("K6", "K7");
            range = (Excel.Range)sheet.get_Range("K6", "K7").Cells;
            cellsMerge(range);

            /*valuesTable = (Excel.Range)sheet.get_Range("A8", "K25").Cells;
            valuesTable.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            valuesTable.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            valuesTable.WrapText = true; // перенос текста в ячейках
            valuesTable.Borders.ColorIndex = 0;
            valuesTable.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            valuesTable.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;*/

            sheet.Cells[1, 1] = "Переписной лист №"; ///////////////////////////////////////////////////////////
            sheet.Cells[3, 1] = "№ листа";
            sheet.Cells[3, 2] = "Всего преписано вагонов";
            sheet.Cells[3, 3] = "По родам вагона";
            sheet.Cells[5, 3] = "КР-20";
            sheet.Cells[5, 4] = "ПЛ-40";
            sheet.Cells[5, 5] = "ПВ-60";
            sheet.Cells[5, 6] = "ЦС-70";
            sheet.Cells[5, 7] = "ПР-90";
            sheet.Cells[5, 8] = "в т.ч.";
            sheet.Cells[6, 8] = "ЦМВ-93";
            sheet.Cells[6, 9] = "ФИ-94";
            sheet.Cells[6, 10] = "ЗРВ-95";
            sheet.Cells[6, 11] = "МВЗ-95";


            range = (Excel.Range)sheet.get_Range("A25", "K25").Cells;
            range.Font.Bold = true;

            sheet.Rows.RowHeight = 25;
            sheet.Rows[1].RowHeight = 40;
            sheet.Rows[3].RowHeight = 10;
            sheet.Rows[4].RowHeight = 10;
            sheet.Rows[5].RowHeight = 15;
            sheet.Rows[6].RowHeight = 15;
            sheet.Rows[7].RowHeight = 15;
            sheet.Columns[1].ColumnWidth = 6;
            sheet.Columns[2].ColumnWidth = 10;
            sheet.Columns[3].ColumnWidth = 7;
            sheet.Columns[4].ColumnWidth = 7;
            sheet.Columns[5].ColumnWidth = 7;
            sheet.Columns[6].ColumnWidth = 7;
            sheet.Columns[7].ColumnWidth = 7;
            sheet.Columns[8].ColumnWidth = 8;
            sheet.Columns[9].ColumnWidth = 8;
            sheet.Columns[10].ColumnWidth = 8;
            sheet.Columns[11].ColumnWidth = 8;

            sql = "SELECT st.ESR, st.[NAME], ccl.LIST_NO," +
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
                "WHERE st.ESR = 480403 " +
                "GROUP BY st.ESR, st.NAME, ccl.LIST_NO";

            cmd.Connection = conn;
            cmd.CommandText = sql;

            using (DbDataReader reader = cmd.ExecuteReader())
            {
                int idRec = 0;
                Excel.Range rng;
                string fname = "";
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        idRec++;
                        ESR = Convert.ToInt32(reader.GetValue(0));
                        stationName = reader.GetValue(1).ToString();
                        listNO = Convert.ToInt32(reader.GetValue(2));
                        KR20 = Convert.ToInt32(reader.GetValue(3));
                        PL40 = Convert.ToInt32(reader.GetValue(4));
                        PV60 = Convert.ToInt32(reader.GetValue(5));
                        CS70 = Convert.ToInt32(reader.GetValue(6));
                        PR90 = Convert.ToInt32(reader.GetValue(7));
                        CMV93 = Convert.ToInt32(reader.GetValue(8));
                        FT94 = Convert.ToInt32(reader.GetValue(9));
                        ZRV95 = Convert.ToInt32(reader.GetValue(10));
                        MVZ93 = Convert.ToInt32(reader.GetValue(11));

                        fname = "Итоги переписи по станции " + stationName + " (" + ESR + ").xlsx";

                        sheet.Cells[1, 1] = "Итоги переписи по станции " + stationName + " (" + ESR + ")"; 
                        sheet.Cells[7 + idRec, 1] = listNO;
                        sheet.Cells[7 + idRec, 3] = KR20;
                        sheet.Cells[7 + idRec, 4] = PL40;
                        sheet.Cells[7 + idRec, 5] = PV60;
                        sheet.Cells[7 + idRec, 6] = CS70;
                        sheet.Cells[7 + idRec, 7] = PR90;
                        sheet.Cells[7 + idRec, 8] = CMV93;
                        sheet.Cells[7 + idRec, 9] = FT94;
                        sheet.Cells[7 + idRec, 10] = ZRV95;
                        sheet.Cells[7 + idRec, 11] = MVZ93;

                        rng = (Excel.Range)sheet.get_Range("C" + (7 + idRec).ToString() + ":" + "K" + (7 + idRec).ToString()).Cells;
                        sheet.Cells[7 + idRec, 2] = excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек
                    }

                    valuesTable = (Excel.Range)sheet.get_Range("A8", "K" + (idRec + 8).ToString()).Cells;
                    valuesTable.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    valuesTable.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    valuesTable.WrapText = true; // перенос текста в ячейках
                    valuesTable.Borders.ColorIndex = 0;
                    valuesTable.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    valuesTable.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    sheet.Cells[idRec + 8, 1] = "Всего";


                    rng = (Excel.Range)sheet.get_Range("B8:B" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 2] = excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("C8:C" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 3] = excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек      

                    rng = (Excel.Range)sheet.get_Range("D8:D" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 4] = excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("E8:E" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 5] = excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("F8:F" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 6] = excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("G8:G" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 7] = excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("H8:H" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 8] = excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("I8:I" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 9] = excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("J8:J" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 10] = excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("K8:K" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 11] = excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек


                    /*//Excel.Range formulaRange = (Excel.Range)sheet.get_Range(sheet.Cells[8, 3], sheet.Cells[8, 11]);
                    Excel.Range formulaRange = (Excel.Range)sheet.get_Range("C8", "K8").Cells;
                    string adder = formulaRange.get_Address(1, 1, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    //Одна ячейка как диапазон
                    Excel.Range r = sheet.Cells[8, 2] as Excel.Range;
                    //Задаем формулу суммы
                    r.Formula = String.Format("=СУММ({0})", adder);*/


                    SaveFileDialog fileDialog = new SaveFileDialog();
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
            }

        }

        private void cellsMerge(Excel.Range range/*, string start, string end*/)
        {
            /*Excel.Range range;
            range = (Excel.Range)sheet.get_Range(start, end).Cells;*/
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            range.WrapText = true; // перенос текста в ячейках
            range.Borders.ColorIndex = 0;
            range.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            range.Merge(Type.Missing);
        }
    }
}
