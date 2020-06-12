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
    class CensusResults
    {
        int idRec = 0, listNO, KR20, PL40, PV60, CS70, PR90, CMV93, FT94, ZRV95, MVZ93, ESR;
        string stationName;

        Excel.Worksheet sheet;
        Excel.Range title;
        Excel.Range range;
        Excel.Range valuesTable;

        public CensusResults()
        {
        }

        public void CreateTypesWagones()
        {
            GlobalData.excelApp = new Microsoft.Office.Interop.Excel.Application();
            GlobalData.workBook = GlobalData.excelApp.Workbooks.Add(Type.Missing);
            GlobalData.excelApp.SheetsInNewWorkbook = 1;
            GlobalData.excelApp.DisplayAlerts = false;
            sheet = (Excel.Worksheet)GlobalData.excelApp.Worksheets.get_Item(1);

            sheet.Name = "По родам вагона";

            { 
                title = (Excel.Range)sheet.get_Range("A1", "K1").Cells;
                title.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                title.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                title.Font.Bold = true;
                title.Merge(Type.Missing);
           
                range = (Excel.Range)sheet.get_Range("A3", "A7").Cells;
                cellsMerge(range);
                range = (Excel.Range)sheet.get_Range("B3", "B7").Cells;
                cellsMerge(range);
                range = (Excel.Range)sheet.get_Range("C3", "K4").Cells;
                cellsMerge(range);

                range = (Excel.Range)sheet.get_Range("C5", "C7").Cells;
                cellsMerge(range);

                range = (Excel.Range)sheet.get_Range("D5", "D7").Cells;
                cellsMerge(range);

                range = (Excel.Range)sheet.get_Range("E5", "E7").Cells;
                cellsMerge(range);

                range = (Excel.Range)sheet.get_Range("F5", "F7").Cells;
                cellsMerge(range);

                range = (Excel.Range)sheet.get_Range("G5", "G7").Cells;
                cellsMerge(range);

                range = (Excel.Range)sheet.get_Range("H5", "K5").Cells;
                cellsMerge(range);

                range = (Excel.Range)sheet.get_Range("H6", "H7").Cells;
                cellsMerge(range);

                range = (Excel.Range)sheet.get_Range("I6", "I7").Cells;
                cellsMerge(range);

                range = (Excel.Range)sheet.get_Range("J6", "J7").Cells;
                cellsMerge(range);

                range = (Excel.Range)sheet.get_Range("K6", "K7").Cells;
                cellsMerge(range);

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
                sheet.Cells[6, 9] = "ФТ-94";
                sheet.Cells[6, 10] = "ЗРВ-95";
                sheet.Cells[6, 11] = "МВЗ-95";

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
            }

            

            GlobalData.sql = "SELECT st.ESR, st.[NAME], ccl.LIST_NO," +
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


            GlobalData.cmd.Connection = GlobalData.conn;
            GlobalData.cmd.CommandText = GlobalData.sql;

            using (DbDataReader reader = GlobalData.cmd.ExecuteReader())
            {
                Excel.Range rng;
                idRec = 0;

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

                        sheet.Cells[1, 1] = "Итоги переписи по станции " + stationName + " (" + ESR.ToString() + ")";
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
                        sheet.Cells[7 + idRec, 2] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек
                    }

                    valuesTable = (Excel.Range)sheet.get_Range("A8", "K" + (idRec + 8).ToString()).Cells;
                    valuesTable.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    valuesTable.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    valuesTable.WrapText = true; // перенос текста в ячейках
                    valuesTable.Borders.ColorIndex = 0;
                    valuesTable.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    valuesTable.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    sheet.Cells[8 + idRec, 1] = "Всего";

                    // выделение жирным строку ВСЕГО
                    range = (Excel.Range)sheet.get_Range("A"+ (8 + idRec).ToString() + "", "K"+ (8 + idRec).ToString() + "").Cells;
                    range.Font.Bold = true;

                    rng = (Excel.Range)sheet.get_Range("B8:B" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 2] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("C8:C" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 3] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек      

                    rng = (Excel.Range)sheet.get_Range("D8:D" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 4] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("E8:E" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 5] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("F8:F" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 6] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("G8:G" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 7] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("H8:H" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 8] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("I8:I" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 9] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("J8:J" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 10] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("K8:K" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 11] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек



                    // сохранение файла
                    SaveFileDialog fileDialog = new SaveFileDialog();
                    fileDialog.FileName = "Итоги переписи по станции " + stationName + " (" + ESR.ToString() + ").xlsx";
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
                            GlobalData.excelApp.Visible = true;
                        }
                        else
                        {
                            GlobalData.excelApp.Application.ActiveWorkbook.Close(true, Type.Missing, Type.Missing);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Файл не был сохранен...");
                    }
                }
            }
        }

        public void NonWorking()
        {
            //GlobalData.excelApp = new Excel.Application();
            GlobalData.excelApp.SheetsInNewWorkbook = 3;
            GlobalData.workBook = GlobalData.excelApp.Workbooks.Add(Type.Missing);
            GlobalData.excelApp.DisplayAlerts = false;

            pattern("ДЛЗО", "3", 1);
            pattern("СТН", "4", 2);
            pattern("Поврежден по акту ВУ-25", "5", 3);

            save();
            
            GlobalData.excelApp.Workbooks.Close();
            GlobalData.excelApp.Quit();
        }

        private void pattern(string name, string state, int sheetNum)
        {
            sheet = (Excel.Worksheet)GlobalData.excelApp.Worksheets.get_Item(sheetNum);
            sheet.Name = name;

            {
                title = (Excel.Range)sheet.get_Range("A1", "K1").Cells;
                title.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                title.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                title.Font.Bold = true;
                title.Merge(Type.Missing);

                range = (Excel.Range)sheet.get_Range("A3", "A7").Cells;
                cellsMerge(range);
                range = (Excel.Range)sheet.get_Range("B3", "B7").Cells;
                cellsMerge(range);
                range = (Excel.Range)sheet.get_Range("C3", "K4").Cells;
                cellsMerge(range);

                range = (Excel.Range)sheet.get_Range("C5", "C7").Cells;
                cellsMerge(range);

                range = (Excel.Range)sheet.get_Range("D5", "D7").Cells;
                cellsMerge(range);

                range = (Excel.Range)sheet.get_Range("E5", "E7").Cells;
                cellsMerge(range);

                range = (Excel.Range)sheet.get_Range("F5", "F7").Cells;
                cellsMerge(range);

                range = (Excel.Range)sheet.get_Range("G5", "G7").Cells;
                cellsMerge(range);

                range = (Excel.Range)sheet.get_Range("H5", "K5").Cells;
                cellsMerge(range);

                range = (Excel.Range)sheet.get_Range("H6", "H7").Cells;
                cellsMerge(range);

                range = (Excel.Range)sheet.get_Range("I6", "I7").Cells;
                cellsMerge(range);

                range = (Excel.Range)sheet.get_Range("J6", "J7").Cells;
                cellsMerge(range);

                range = (Excel.Range)sheet.get_Range("K6", "K7").Cells;
                cellsMerge(range);

                sheet.Cells[3, 1] = "№ листа";
                sheet.Cells[3, 2] = "Всего";
                sheet.Cells[3, 3] = name;
                sheet.Cells[5, 3] = "КР-20";
                sheet.Cells[5, 4] = "ПЛ-40";
                sheet.Cells[5, 5] = "ПВ-60";
                sheet.Cells[5, 6] = "ЦС-70";
                sheet.Cells[5, 7] = "ПР-90";
                sheet.Cells[5, 8] = "в т.ч.";
                sheet.Cells[6, 8] = "ЦМВ-93";
                sheet.Cells[6, 9] = "ФТ-94";
                sheet.Cells[6, 10] = "ЗРВ-95";
                sheet.Cells[6, 11] = "МВЗ-95";

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
            }



            GlobalData.sql = "SELECT st.ESR, st.[NAME], ccl.LIST_NO," +
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
                "WHERE ccl.IS_WORKING = 0 AND ccl.NON_WORKING_STATE = " + state + " AND st.ESR = 480403 " +
                "GROUP BY st.ESR, st.NAME, ccl.LIST_NO";


            GlobalData.cmd.Connection = GlobalData.conn;
            GlobalData.cmd.CommandText = GlobalData.sql;

            using (DbDataReader reader = GlobalData.cmd.ExecuteReader())
            {
                Excel.Range rng;
                idRec = 0;

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

                        sheet.Cells[1, 1] = "Итоги переписи по станции " + stationName + " (" + ESR.ToString() + ")";
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
                        sheet.Cells[7 + idRec, 2] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек
                    }

                    valuesTable = (Excel.Range)sheet.get_Range("A8", "K" + (idRec + 8).ToString()).Cells;
                    valuesTable.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    valuesTable.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    valuesTable.WrapText = true; // перенос текста в ячейках
                    valuesTable.Borders.ColorIndex = 0;
                    valuesTable.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    valuesTable.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    sheet.Cells[8 + idRec, 1] = "Всего";

                    // выделение жирным строку ВСЕГО
                    range = (Excel.Range)sheet.get_Range("A" + (8 + idRec).ToString() + "", "K" + (8 + idRec).ToString() + "").Cells;
                    range.Font.Bold = true;

                    rng = (Excel.Range)sheet.get_Range("B8:B" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 2] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("C8:C" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 3] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек      

                    rng = (Excel.Range)sheet.get_Range("D8:D" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 4] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("E8:E" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 5] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("F8:F" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 6] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("G8:G" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 7] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("H8:H" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 8] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("I8:I" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 9] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("J8:J" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 10] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)sheet.get_Range("K8:K" + (7 + idRec).ToString()).Cells;
                    sheet.Cells[8 + idRec, 11] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                }
            }
        }

        private void save()
        {
            // сохранение файла
            SaveFileDialog fileDialog = new SaveFileDialog();
            fileDialog.FileName = "Итоги переписи по станции " + stationName + " (" + ESR.ToString() + ").xlsx";
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
                    Process.Start(fileDialog.FileName);
                }
                else
                {
                    GlobalData.excelApp.Application.ActiveWorkbook.Close(true, Type.Missing, Type.Missing);
                }
            }
            else
            {
                MessageBox.Show("Файл не был сохранен...");
            }
        }

        private void cellsMerge(Excel.Range range)
        {
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
