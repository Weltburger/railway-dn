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
    }
}
