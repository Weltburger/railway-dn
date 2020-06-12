using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace test_railway
{
    class CensusList
    {
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

        string stationName;
        string esrStation;
        string listNum;

        Excel.Range columnsTable;

        public CensusList() 
        {
        }

        public void CreateDocument() 
        {
            GlobalData.excelApp = new Microsoft.Office.Interop.Excel.Application();
            GlobalData.workBook = GlobalData.excelApp.Workbooks.Add(Type.Missing);
            GlobalData.excelApp.SheetsInNewWorkbook = 1;
            GlobalData.excelApp.DisplayAlerts = false;
            GlobalData.excelProc = Process.GetProcessesByName("EXCEL").Last();

            GlobalData.sheet = (Excel.Worksheet)GlobalData.excelApp.Worksheets.get_Item(1);
            GlobalData.sheet.Name = "Переписной лист";

            GlobalData.titleTable = (Excel.Range)GlobalData.sheet.get_Range("A1", "J1").Cells;
            GlobalData.titleTable.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            GlobalData.titleTable.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            GlobalData.titleTable.Font.Bold = true;
            GlobalData.titleTable.Merge(Type.Missing);

            columnsTable = (Excel.Range)GlobalData.sheet.get_Range("A2", "J2").Cells;
            columnsTable.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            columnsTable.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            columnsTable.WrapText = true; // перенос текста в ячейках
            columnsTable.Borders.ColorIndex = 0;
            columnsTable.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            columnsTable.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

            GlobalData.sheet.Cells[1, 1] = "Переписной лист №";
            GlobalData.sheet.Cells[2, 1] = "№ п/п";
            GlobalData.sheet.Cells[2, 2] = "Номер вагона";
            GlobalData.sheet.Cells[2, 3] = "Год постройки";
            GlobalData.sheet.Cells[2, 4] = "Род вагона";
            GlobalData.sheet.Cells[2, 5] = "Дислокация";
            GlobalData.sheet.Cells[2, 6] = "Код страны-собств.";
            GlobalData.sheet.Cells[2, 7] = "Собственник вагона";
            GlobalData.sheet.Cells[2, 8] = "Состояние";
            GlobalData.sheet.Cells[2, 9] = "Парк";
            GlobalData.sheet.Cells[2, 10] = "Категория НРП";

            GlobalData.sheet.Rows.RowHeight = 25;
            GlobalData.sheet.Rows[1].RowHeight = 40;
            GlobalData.sheet.Rows[2].RowHeight = 50;
            GlobalData.sheet.Columns[1].ColumnWidth = 4;
            GlobalData.sheet.Columns[2].ColumnWidth = 9;
            GlobalData.sheet.Columns[3].ColumnWidth = 11;
            GlobalData.sheet.Columns[4].ColumnWidth = 7;
            GlobalData.sheet.Columns[5].ColumnWidth = 13;
            GlobalData.sheet.Columns[6].ColumnWidth = 8;
            GlobalData.sheet.Columns[7].ColumnWidth = 14;
            GlobalData.sheet.Columns[8].ColumnWidth = 14;
            GlobalData.sheet.Columns[9].ColumnWidth = 12;
            GlobalData.sheet.Columns[10].ColumnWidth = 14;

            GlobalData.sql = "select st.ESR, st.NAME as 'Станция (для заголовка)', " +
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

            GlobalData.cmd.CommandText = GlobalData.sql;

            using (DbDataReader reader = GlobalData.cmd.ExecuteReader())
            {
                int idRec = 0;
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        idRec++;
                        esrStation = reader.GetValue(0).ToString();
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

                        GlobalData.sheet.Cells[1, 1] = "Переписной лист №" + listNO + " станция " + stationName + " (" + esrStation + ")";
                        GlobalData.sheet.Cells[2 + idRec, 1] = idRec;
                        GlobalData.sheet.Cells[2 + idRec, 2] = carNO;
                        GlobalData.sheet.Cells[2 + idRec, 3] = builtYear;
                        GlobalData.sheet.Cells[2 + idRec, 4] = carType;
                        GlobalData.sheet.Cells[2 + idRec, 5] = carLoc;
                        GlobalData.sheet.Cells[2 + idRec, 6] = admCode;
                        GlobalData.sheet.Cells[2 + idRec, 7] = owner;
                        GlobalData.sheet.Cells[2 + idRec, 8] = isLoaded;
                        GlobalData.sheet.Cells[2 + idRec, 9] = isWorking;
                        GlobalData.sheet.Cells[2 + idRec, 10] = workState;

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

                    GlobalData.cellsValuesTable = (Excel.Range)GlobalData.sheet.get_Range("A3", "J" + (idRec + 2).ToString()).Cells;
                    GlobalData.cellsValuesTable.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    GlobalData.cellsValuesTable.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    GlobalData.cellsValuesTable.WrapText = true; // перенос текста в ячейках
                    GlobalData.cellsValuesTable.Borders.ColorIndex = 0;
                    GlobalData.cellsValuesTable.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    GlobalData.cellsValuesTable.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    // Реализовать сохранение документов по каждому номеру переписного листа - listNO

                    SaveFileDialog fileDialog = new SaveFileDialog();
                    fileDialog.FileName = "Переписной лист №" + listNO + " - станция " + stationName + " (" + esrStation + ").xlsx";
                    if (fileDialog.ShowDialog() == DialogResult.OK)
                    {
                        GlobalData.excelApp.Application.ActiveWorkbook.SaveAs(
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
                            Process.Start(fileDialog.FileName); // открытие сохраненного документа как отдельное приложение
                            GlobalData.processClose();
                        }
                        else
                        {
                            GlobalData.processClose();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Файл не был сохранен...");
                    }
                }
            }
        }


        public void SetEsrStation(string esrStation)
        {
            this.esrStation = esrStation;
        }

        public void SetNumList(string listNum)
        {
            this.listNum = listNum;
        }
    }
}
