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

        private enum ReportType 
        {
            По_родам_вагона ,Рабочий_парк, Нерабочий_парк
        }

        private enum TypeScheme 
        {
            Общая, По_работоспособности
        }

        int idRec = 0, listNO, KR20, PL40, PV60, CS70, PR90, CMV93, FT94, ZRV95, MVZ93, ESR;
        string stationName;

        Excel.Range range;

        public CensusResults()
        {
        }

        public void CreateReport() 
        {
            GlobalData.excelApp = new Excel.Application();
            GlobalData.excelApp.SheetsInNewWorkbook = 10;
            GlobalData.workBook = GlobalData.excelApp.Workbooks.Add(Type.Missing);
            GlobalData.excelApp.DisplayAlerts = false;

            CreateResultTable(ReportType.По_родам_вагона, "По родам вагона", true, 0, 1, GlobalData.stationSelected);
            CreateResultTable(ReportType.Рабочий_парк, "Рабочий парк всего", true, 0, 2, GlobalData.stationSelected);
            CreateResultTable(ReportType.Рабочий_парк, "Груженых", false, 1, 3, GlobalData.stationSelected);
            CreateResultTable(ReportType.Рабочий_парк, "Порожних", false, 0, 4, GlobalData.stationSelected);
            CreateResultTable(ReportType.Нерабочий_парк, "Всего НРП", true, 0, 5, GlobalData.stationSelected);
            CreateResultTable(ReportType.Нерабочий_парк, "Неисправно", false, 1, 6, GlobalData.stationSelected);
            CreateResultTable(ReportType.Нерабочий_парк, "Резерв", false, 2, 7, GlobalData.stationSelected);
            CreateResultTable(ReportType.Нерабочий_парк, "ДЛЗО", false, 3, 8, GlobalData.stationSelected);
            CreateResultTable(ReportType.Нерабочий_парк, "СТН", false, 4, 9, GlobalData.stationSelected);
            CreateResultTable(ReportType.Нерабочий_парк, "Поврежден по акту ВУ-25", false, 5, 10, GlobalData.stationSelected);

            save();
        }

        private void CreateResultTable(ReportType reportType, string name, bool generalTable, int state, int sheetNum, int stationSelected)
        {
            GlobalData.excelProc = Process.GetProcessesByName("EXCEL").Last();
            GlobalData.sheet = (Excel.Worksheet)GlobalData.excelApp.Worksheets.get_Item(sheetNum);
            GlobalData.sheet.Name = name;

            if (reportType == ReportType.По_родам_вагона) 
            {
                // создание разметки таблицы
                CreateTableScheme(TypeScheme.Общая, name);

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
                "WHERE st.ESR = " + stationSelected.ToString() + " " +
                "GROUP BY st.ESR, st.NAME, ccl.LIST_NO";
            }
            else if (reportType == ReportType.Нерабочий_парк)
            {
                // создание разметки таблицы
                CreateTableScheme(TypeScheme.По_работоспособности, name);

                if (generalTable == false)
                {
                    GlobalData.sql = "SELECT DISTINCT " +
                        "st.ESR, " +
                        "st.[NAME], " +
                        "ccl.LIST_NO, " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'КР-20' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0 AND ccll.NON_WORKING_STATE = " + state.ToString() + ")  AS \"КР -20\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПЛ-40' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0 AND ccll.NON_WORKING_STATE = " + state.ToString() + ")  AS \"ПЛ -40\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПВ-60' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0 AND ccll.NON_WORKING_STATE = " + state.ToString() + ")  AS \"ПВ-60\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЦС-70' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0 AND ccll.NON_WORKING_STATE = " + state.ToString() + ")  AS \"ЦС -70\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПР-90' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0 AND ccll.NON_WORKING_STATE = " + state.ToString() + ")  AS \"ПР -90\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЦМВ-93' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0 AND ccll.NON_WORKING_STATE = " + state.ToString() + ")  AS \"ЦМВ -93\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ФТ-94' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0 AND ccll.NON_WORKING_STATE = " + state.ToString() + ")  AS \"ФТ -94\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЗРВ-95' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0 AND ccll.NON_WORKING_STATE = " + state.ToString() + ")  AS \"ЗРВ -95\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'МВЗ-92' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 0 AND ccll.NON_WORKING_STATE = " + state.ToString() + ")  AS \"МВЗ -92\" " +
                        "FROM " +
                        "CAR_CENSUS_LISTS ccl " +
                        "JOIN STATIONS st ON st.ESR = ccl.LOCATION_ESR " +
                        "WHERE " +
                        "st.ESR = " + stationSelected.ToString() + "";

                }
                else
                {
                    GlobalData.sql = "SELECT DISTINCT " +
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
                        "st.ESR = " + stationSelected.ToString() + "";
                }
            }
            else if (reportType == ReportType.Рабочий_парк)
            {
                // создание разметки таблицы
                CreateTableScheme(TypeScheme.По_работоспособности, name);

                if (generalTable == false)
                {
                    GlobalData.sql = "SELECT DISTINCT " +
                        "st.ESR, " +
                        "st.[NAME], " +
                        "ccl.LIST_NO, " +
                        "(SELECT COUNT(CASE CAR_TYPE WHEN 'КР-20' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1 AND ccll.IS_LOADED = " + state.ToString() + ")  AS \"КР-20\", " +
                        "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПЛ-40' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1 AND ccll.IS_LOADED = " + state.ToString() + ")  AS \"ПЛ-40\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПВ-60' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1 AND ccll.IS_LOADED = " + state.ToString() + ")  AS \"ПВ-60\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЦС-70' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1 AND ccll.IS_LOADED = " + state.ToString() + ")  AS \"ЦС-70\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПР-90' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1 AND ccll.IS_LOADED = " + state.ToString() + ")  AS \"ПР-90\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЦМВ-93' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1 AND ccll.IS_LOADED = " + state.ToString() + ")  AS \"ЦМВ-93\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ФТ-94' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1 AND ccll.IS_LOADED = " + state.ToString() + ")  AS \"ФТ-94\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЗРВ-95' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1 AND ccll.IS_LOADED = " + state.ToString() + ")  AS \"ЗРВ-95\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'МВЗ-92' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1 AND ccll.IS_LOADED = " + state.ToString() + ")  AS \"МВЗ-92\" " +
                        "FROM " +
                        "CAR_CENSUS_LISTS ccl " +
                        "JOIN STATIONS st ON st.ESR = ccl.LOCATION_ESR " +
                        "WHERE " +
                        "st.ESR = " + stationSelected.ToString() + "";
                }
                else
                {
                    GlobalData.sql = "SELECT DISTINCT " +
                        "st.ESR, " +
                        "st.[NAME], " +
                        "ccl.LIST_NO, " +
                        "(SELECT COUNT(CASE CAR_TYPE WHEN 'КР-20' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1)  AS \"КР-20\", " +
                        "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПЛ-40' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1)  AS \"ПЛ-40\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПВ-60' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1)  AS \"ПВ-60\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЦС-70' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1)  AS \"ЦС-70\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ПР-90' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1)  AS \"ПР-90\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЦМВ-93' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1)  AS \"ЦМВ-93\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ФТ-94' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1)  AS \"ФТ-94\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'ЗРВ-95' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1)  AS \"ЗРВ-95\", " +
                            "(SELECT COUNT(CASE CAR_TYPE WHEN 'МВЗ-92' THEN CAR_TYPE END) FROM CAR_CENSUS_LISTS ccll WHERE ccl.LOCATION_ESR = ccll.LOCATION_ESR AND ccl.LIST_NO = ccll.LIST_NO AND ccll.IS_WORKING = 1)  AS \"МВЗ-92\" " +
                        "FROM " +
                        "CAR_CENSUS_LISTS ccl " +
                        "JOIN STATIONS st ON st.ESR = ccl.LOCATION_ESR " +
                        "WHERE " +
                        "st.ESR = " + stationSelected.ToString() + "";
                }
            }

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

                        GlobalData.sheet.Cells[1, 1] = "Итоги переписи по станции " + stationName + " (" + ESR.ToString() + ")";
                        GlobalData.sheet.Cells[7 + idRec, 1] = listNO;
                        GlobalData.sheet.Cells[7 + idRec, 3] = KR20;
                        GlobalData.sheet.Cells[7 + idRec, 4] = PL40;
                        GlobalData.sheet.Cells[7 + idRec, 5] = PV60;
                        GlobalData.sheet.Cells[7 + idRec, 6] = CS70;
                        GlobalData.sheet.Cells[7 + idRec, 7] = PR90;
                        GlobalData.sheet.Cells[7 + idRec, 8] = CMV93;
                        GlobalData.sheet.Cells[7 + idRec, 9] = FT94;
                        GlobalData.sheet.Cells[7 + idRec, 10] = ZRV95;
                        GlobalData.sheet.Cells[7 + idRec, 11] = MVZ93;

                        rng = (Excel.Range)GlobalData.sheet.get_Range("C" + (7 + idRec).ToString() + ":" + "K" + (7 + idRec).ToString()).Cells;
                        GlobalData.sheet.Cells[7 + idRec, 2] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек
                    }

                    GlobalData.cellsValuesTable = (Excel.Range)GlobalData.sheet.get_Range("A8", "K" + (idRec + 8).ToString()).Cells;
                    GlobalData.cellsValuesTable.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    GlobalData.cellsValuesTable.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    GlobalData.cellsValuesTable.WrapText = true; // перенос текста в ячейках
                    GlobalData.cellsValuesTable.Borders.ColorIndex = 0;
                    GlobalData.cellsValuesTable.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    GlobalData.cellsValuesTable.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    GlobalData.sheet.Cells[8 + idRec, 1] = "Всего";

                    // выделение жирным строку ВСЕГО
                    range = (Excel.Range)GlobalData.sheet.get_Range("A" + (8 + idRec).ToString() + "", "K" + (8 + idRec).ToString() + "").Cells;
                    range.Font.Bold = true;

                    rng = (Excel.Range)GlobalData.sheet.get_Range("B8:B" + (7 + idRec).ToString()).Cells;
                    GlobalData.sheet.Cells[8 + idRec, 2] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)GlobalData.sheet.get_Range("C8:C" + (7 + idRec).ToString()).Cells;
                    GlobalData.sheet.Cells[8 + idRec, 3] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек      

                    rng = (Excel.Range)GlobalData.sheet.get_Range("D8:D" + (7 + idRec).ToString()).Cells;
                    GlobalData.sheet.Cells[8 + idRec, 4] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)GlobalData.sheet.get_Range("E8:E" + (7 + idRec).ToString()).Cells;
                    GlobalData.sheet.Cells[8 + idRec, 5] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)GlobalData.sheet.get_Range("F8:F" + (7 + idRec).ToString()).Cells;
                    GlobalData.sheet.Cells[8 + idRec, 6] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)GlobalData.sheet.get_Range("G8:G" + (7 + idRec).ToString()).Cells;
                    GlobalData.sheet.Cells[8 + idRec, 7] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)GlobalData.sheet.get_Range("H8:H" + (7 + idRec).ToString()).Cells;
                    GlobalData.sheet.Cells[8 + idRec, 8] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)GlobalData.sheet.get_Range("I8:I" + (7 + idRec).ToString()).Cells;
                    GlobalData.sheet.Cells[8 + idRec, 9] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)GlobalData.sheet.get_Range("J8:J" + (7 + idRec).ToString()).Cells;
                    GlobalData.sheet.Cells[8 + idRec, 10] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                    rng = (Excel.Range)GlobalData.sheet.get_Range("K8:K" + (7 + idRec).ToString()).Cells;
                    GlobalData.sheet.Cells[8 + idRec, 11] = GlobalData.excelApp.WorksheetFunction.Sum(rng); //вычисляем сумму ячеек

                }
            }
        }

        private void CreateTableScheme(TypeScheme typeScheme, string name) 
        {
            
            GlobalData.titleTable = (Excel.Range)GlobalData.sheet.get_Range("A1", "K1").Cells;
            GlobalData.titleTable.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            GlobalData.titleTable.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            GlobalData.titleTable.Font.Bold = true;
            GlobalData.titleTable.Merge(Type.Missing);

            range = (Excel.Range)GlobalData.sheet.get_Range("A3", "A7").Cells;
            cellsMerge(range);
            range = (Excel.Range)GlobalData.sheet.get_Range("B3", "B7").Cells;
            cellsMerge(range);
            range = (Excel.Range)GlobalData.sheet.get_Range("C3", "K4").Cells;
            cellsMerge(range);

            range = (Excel.Range)GlobalData.sheet.get_Range("C5", "C7").Cells;
            cellsMerge(range);

            range = (Excel.Range)GlobalData.sheet.get_Range("D5", "D7").Cells;
            cellsMerge(range);

            range = (Excel.Range)GlobalData.sheet.get_Range("E5", "E7").Cells;
            cellsMerge(range);

            range = (Excel.Range)GlobalData.sheet.get_Range("F5", "F7").Cells;
            cellsMerge(range);

            range = (Excel.Range)GlobalData.sheet.get_Range("G5", "G7").Cells;
            cellsMerge(range);

            range = (Excel.Range)GlobalData.sheet.get_Range("H5", "K5").Cells;
            cellsMerge(range);

            range = (Excel.Range)GlobalData.sheet.get_Range("H6", "H7").Cells;
            cellsMerge(range);

            range = (Excel.Range)GlobalData.sheet.get_Range("I6", "I7").Cells;
            cellsMerge(range);

            range = (Excel.Range)GlobalData.sheet.get_Range("J6", "J7").Cells;
            cellsMerge(range);

            range = (Excel.Range)GlobalData.sheet.get_Range("K6", "K7").Cells;
            cellsMerge(range);

            GlobalData.sheet.Cells[3, 1] = "№ листа";
            if (typeScheme == TypeScheme.Общая)
            {
                GlobalData.sheet.Cells[3, 2] = "Всего преписано вагонов";
            }
            else 
            {
                GlobalData.sheet.Cells[3, 2] = "Всего";
            }
            GlobalData.sheet.Cells[3, 3] = name;
            GlobalData.sheet.Cells[5, 3] = "КР-20";
            GlobalData.sheet.Cells[5, 4] = "ПЛ-40";
            GlobalData.sheet.Cells[5, 5] = "ПВ-60";
            GlobalData.sheet.Cells[5, 6] = "ЦС-70";
            GlobalData.sheet.Cells[5, 7] = "ПР-90";
            GlobalData.sheet.Cells[5, 8] = "в т.ч.";
            GlobalData.sheet.Cells[6, 8] = "ЦМВ-93";
            GlobalData.sheet.Cells[6, 9] = "ФТ-94";
            GlobalData.sheet.Cells[6, 10] = "ЗРВ-95";
            GlobalData.sheet.Cells[6, 11] = "МВЗ-95";

            GlobalData.sheet.Rows.RowHeight = 25;
            GlobalData.sheet.Rows[1].RowHeight = 40;
            GlobalData.sheet.Rows[3].RowHeight = 10;
            GlobalData.sheet.Rows[4].RowHeight = 10;
            GlobalData.sheet.Rows[5].RowHeight = 15;
            GlobalData.sheet.Rows[6].RowHeight = 15;
            GlobalData.sheet.Rows[7].RowHeight = 15;
            GlobalData.sheet.Columns[1].ColumnWidth = 6;
            GlobalData.sheet.Columns[2].ColumnWidth = 10;
            GlobalData.sheet.Columns[3].ColumnWidth = 7;
            GlobalData.sheet.Columns[4].ColumnWidth = 7;
            GlobalData.sheet.Columns[5].ColumnWidth = 7;
            GlobalData.sheet.Columns[6].ColumnWidth = 7;
            GlobalData.sheet.Columns[7].ColumnWidth = 7;
            GlobalData.sheet.Columns[8].ColumnWidth = 8;
            GlobalData.sheet.Columns[9].ColumnWidth = 8;
            GlobalData.sheet.Columns[10].ColumnWidth = 8;
            GlobalData.sheet.Columns[11].ColumnWidth = 8;
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
